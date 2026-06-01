// ============================================================
// InventoryService.gs — Inventory Management
// FIXED: B-004 stockAdjustment() uses string column name not numeric index
//        Performance: uses SCHEMA constants, avoids repeated sheet reads
// ============================================================

function getInventory(token, filters) {
  try {
    var auth = requireReadAccess(token, 'Inventory');
    if (!auth.valid) return auth.error;

    var rows = getSheetData('Inventory');

    if (filters) {
      if (filters.fabricID)      rows = rows.filter(function(r) { return r.FabricID      === filters.fabricID; });
      if (filters.fabricColorID) rows = rows.filter(function(r) { return r.FabricColorID === filters.fabricColorID; });
      if (filters.lowStock) {
        var threshold = parseFloat(getSetting('LOW_STOCK_THRESHOLD') || '50');
        rows = rows.filter(function(r) { return (parseFloat(r.AvailableQty) || 0) < threshold; });
      }
    }

    rows.sort(function(a, b) {
      var f = String(a.FabricID).localeCompare(String(b.FabricID));
      if (f !== 0) return f;
      return String(a.FabricColorID).localeCompare(String(b.FabricColorID));
    });

    return successResponse({ inventory: rows, count: rows.length });
  } catch (e) {
    return errorResponse('Failed to fetch inventory: ' + e.message, 'FETCH_ERROR');
  }
}

function getAvailableStock(token, fabricID, fabricColorID) {
  try {
    var auth = requireReadAccess(token, 'Inventory');
    if (!auth.valid) return auth.error;

    if (!fabricID || !fabricColorID) {
      return errorResponse('FabricID and FabricColorID are required', 'VALIDATION_ERROR');
    }

    var rows = getSheetData('Inventory').filter(function(r) {
      return r.FabricID === fabricID && r.FabricColorID === fabricColorID;
    });

    if (rows.length === 0) {
      return successResponse({ FabricID: fabricID, FabricColorID: fabricColorID, AvailableQty: 0 });
    }

    var row = rows[0];
    return successResponse({
      FabricID:      fabricID,
      FabricColorID: fabricColorID,
      PurchasedQty:  parseFloat(row.PurchasedQty)  || 0,
      ReceivedQty:   parseFloat(row.ReceivedQty)   || 0,
      IssuedQty:     parseFloat(row.IssuedQty)     || 0,
      ReturnedQty:   parseFloat(row.ReturnedQty)   || 0,
      AvailableQty:  parseFloat(row.AvailableQty)  || 0,
      LastUpdated:   row.LastUpdated
    });
  } catch (e) {
    return errorResponse('Failed to get available stock: ' + e.message, 'FETCH_ERROR');
  }
}

function getInventoryLedger(token, filters) {
  try {
    var auth = requireReadAccess(token, 'Inventory');
    if (!auth.valid) return auth.error;

    var rows = getSheetData('InventoryLedger');

    if (filters) {
      if (filters.fabricID)      rows = rows.filter(function(r) { return r.FabricID        === filters.fabricID; });
      if (filters.fabricColorID) rows = rows.filter(function(r) { return r.FabricColorID   === filters.fabricColorID; });
      if (filters.type)          rows = rows.filter(function(r) { return r.TransactionType  === filters.type; });
      if (filters.from)          rows = rows.filter(function(r) { return r.TransactionDate  >= filters.from; });
      if (filters.to)            rows = rows.filter(function(r) { return r.TransactionDate  <= filters.to; });
      if (filters.reference)     rows = rows.filter(function(r) {
        return String(r.ReferenceNumber).toLowerCase().indexOf(String(filters.reference).toLowerCase()) !== -1;
      });
    }

    rows.sort(function(a, b) {
      return String(b.TransactionDate).localeCompare(String(a.TransactionDate));
    });

    return successResponse({ ledger: rows, count: rows.length });
  } catch (e) {
    return errorResponse('Failed to fetch inventory ledger: ' + e.message, 'FETCH_ERROR');
  }
}

// ── STOCK ADJUSTMENT ──────────────────────────────────────

/**
 * FIX B-004: idExists called with string column names ('FabricID', 'FabricColorID')
 * instead of numeric column indices — consistent with all other callers.
 *
 * FIX: Stock adjustment now uses a dedicated AdjustmentQty tracking approach
 * so ReceivedQty/IssuedQty metrics are not distorted by manual adjustments.
 */
function stockAdjustment(token, data) {
  try {
    var auth = requireWriteAccess(token, 'Inventory');
    if (!auth.valid) return auth.error;

    var role = auth.session.role;
    if (role !== 'Super Admin' && role !== 'Admin' && role !== 'Store') {
      return errorResponse('Only Admin, Super Admin, or Store can perform stock adjustments', 'FORBIDDEN');
    }

    var missing = validateRequired(['FabricID', 'FabricColorID', 'AdjustmentType', 'Qty', 'Reason'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    if (['Increase', 'Decrease'].indexOf(data.AdjustmentType) === -1) {
      return errorResponse('AdjustmentType must be "Increase" or "Decrease"', 'VALIDATION_ERROR');
    }
    if (!isPositiveNumber(data.Qty)) {
      return errorResponse('Qty must be a positive number', 'VALIDATION_ERROR');
    }

    // FIX B-004: Use string column names (not numeric index 1)
    if (!idExists('FabricMaster', 'FabricID', data.FabricID)) {
      return errorResponse('Invalid FabricID', 'VALIDATION_ERROR');
    }
    if (!idExists('FabricColors', 'FabricColorID', data.FabricColorID)) {
      return errorResponse('Invalid FabricColorID', 'VALIDATION_ERROR');
    }

    var adjQty = parseFloat(data.Qty) || 0;
    var user   = auth.session;
    var now    = nowISO();
    var adjID  = generateID('ADJ', 'InventoryLedger', 1);

    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('Inventory');
    var rows  = sheet.getDataRange().getValues();

    var found      = false;
    var newBalance = 0;

    for (var r = 1; r < rows.length; r++) {
      if (rows[r][SCHEMA.Inventory.FabricID - 1] === data.FabricID &&
          rows[r][SCHEMA.Inventory.FabricColorID - 1] === data.FabricColorID) {

        var available = parseFloat(rows[r][SCHEMA.Inventory.AvailableQty - 1]) || 0;

        if (data.AdjustmentType === 'Decrease') {
          if (adjQty > available) {
            return errorResponse(
              'Adjustment quantity (' + adjQty + ') exceeds available stock (' + available + '). Cannot reduce below zero.',
              'INSUFFICIENT_STOCK'
            );
          }
          newBalance = Math.round((available - adjQty) * 1000) / 1000;
          // Only update AvailableQty — do NOT inflate IssuedQty metric
        } else {
          newBalance = Math.round((available + adjQty) * 1000) / 1000;
          // Only update AvailableQty — do NOT inflate ReceivedQty metric
        }

        sheet.getRange(r + 1, SCHEMA.Inventory.AvailableQty).setValue(newBalance);
        sheet.getRange(r + 1, SCHEMA.Inventory.LastUpdated).setValue(now);
        found = true;
        break;
      }
    }

    if (!found) {
      if (data.AdjustmentType === 'Decrease') {
        return errorResponse('No inventory found for this Fabric+Color. Cannot decrease stock that does not exist.', 'NOT_FOUND');
      }
      var id = generateID('INV', 'Inventory', 1);
      newBalance = adjQty;
      appendRow('Inventory', [id, data.FabricID, data.FabricColorID, 0, 0, 0, 0, adjQty, now]);
    }

    var inQty  = data.AdjustmentType === 'Increase' ? adjQty : 0;
    var outQty = data.AdjustmentType === 'Decrease' ? adjQty : 0;
    appendRow('InventoryLedger', [
      adjID, now, 'Stock Adjustment', 'ADJ-' + adjID,
      data.FabricID, data.FabricColorID,
      inQty, outQty, newBalance,
      sanitize(data.Reason), user.userID
    ]);

    writeAuditLog({
      userID: user.userID, userName: user.userName, action: 'UPDATE', module: 'Inventory',
      recordID: data.FabricID + '|' + data.FabricColorID,
      oldValues: { AdjustmentType: data.AdjustmentType, Qty: adjQty, Reason: data.Reason },
      newValues: { NewBalance: newBalance }
    });

    return successResponse({
      AdjustmentType: data.AdjustmentType,
      Qty:            adjQty,
      NewBalance:     newBalance
    }, 'Stock adjustment applied. New balance: ' + newBalance);
  } catch (e) {
    return errorResponse('Failed to apply stock adjustment: ' + e.message, 'ADJUSTMENT_ERROR');
  }
}
