// ============================================================
// PurchaseService.gs — Purchase Orders
// FIXED: B-005 cancelPO() reverses Inventory.PurchasedQty
//        B-018 removed silent fuzzy Fabric name resolution
//        B-010 uses SCHEMA constants for column indices
// Added: PO_TRANSITIONS includes new statuses
// ============================================================

var PO_TRANSITIONS = {
  'Draft':              ['Approved', 'Cancelled'],
  'Approved':           ['Partially Received', 'Fully Received', 'Over Received', 'Cancelled'],
  'Partially Received': ['Fully Received', 'Completed', 'Over Received'],
  'Over Received':      ['Completed'],
  'Fully Received':     ['Completed'],
  'Completed':          [],
  'Cancelled':          []
};

// ── GET PURCHASE ORDERS ────────────────────────────────────

function getPurchaseOrders(token, filters) {
  try {
    var auth = requireReadAccess(token, 'PurchaseOrders');
    if (!auth.valid) return auth.error;

    var orders = getSheetData('PurchaseOrders');

    if (filters) {
      if (filters.vendorID) orders = orders.filter(function(r) { return r.VendorID === filters.vendorID; });
      if (filters.status)   orders = orders.filter(function(r) { return r.Status   === filters.status; });
      if (filters.from)     orders = orders.filter(function(r) { return r.PODate   >= filters.from; });
      if (filters.to)       orders = orders.filter(function(r) { return r.PODate   <= filters.to; });
    }

    orders.sort(function(a, b) { return String(b.POID).localeCompare(String(a.POID)); });
    return successResponse({ orders: orders, count: orders.length });
  } catch (e) {
    return errorResponse('Failed to fetch purchase orders: ' + e.message, 'FETCH_ERROR');
  }
}

function getPOByID(token, poID) {
  try {
    var auth = requireReadAccess(token, 'PurchaseOrders');
    if (!auth.valid) return auth.error;

    if (!poID) return errorResponse('POID required', 'VALIDATION_ERROR');

    var orders = getSheetData('PurchaseOrders').filter(function(r) { return r.POID === poID; });
    if (orders.length === 0) return errorResponse('Purchase Order not found: ' + poID, 'NOT_FOUND');

    var po    = orders[0];
    var items = getSheetData('PurchaseOrderItems').filter(function(r) { return r.POID === poID; });

    return successResponse({ po: po, items: items });
  } catch (e) {
    return errorResponse('Failed to fetch PO: ' + e.message, 'FETCH_ERROR');
  }
}

// ── CREATE PURCHASE ORDER ──────────────────────────────────

function createPO(token, data) {
  try {
    var auth = requireWriteAccess(token, 'PurchaseOrders');
    if (!auth.valid) return auth.error;

    var missing = validateRequired(['VendorID', 'PODate'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    if (!idExists('VendorMaster', 'VendorID', data.VendorID)) {
      return errorResponse('Invalid VendorID: ' + data.VendorID, 'VALIDATION_ERROR');
    }

    var vendors = getSheetData('VendorMaster').filter(function(r) { return r.VendorID === data.VendorID; });
    if (vendors.length === 0 || (vendors[0].VendorType !== 'Fabric Vendor' && vendors[0].VendorType !== 'Both')) {
      return errorResponse('Vendor is not a Fabric Vendor', 'VALIDATION_ERROR');
    }

    if (!Array.isArray(data.items) || data.items.length === 0) {
      return errorResponse('At least one item line is required', 'VALIDATION_ERROR');
    }

    // Pre-load master data for validation (avoids repeated getSheetData inside loop)
    var fabricMasterMap = {};
    getSheetData('FabricMaster').forEach(function(f) { fabricMasterMap[f.FabricID] = f; });

    var fabricColorMap = {};
    getSheetData('FabricColors').forEach(function(c) { fabricColorMap[c.FabricColorID] = c; });

    var seenKeys = {};
    for (var i = 0; i < data.items.length; i++) {
      var item    = data.items[i];
      var lineNum = i + 1;

      var missingItem = validateRequired(['FabricID', 'FabricColorID', 'OrderedQty', 'UOM', 'Rate'], item);
      if (missingItem.length) {
        return errorResponse('Line ' + lineNum + ': Missing ' + missingItem.join(', '), 'VALIDATION_ERROR');
      }

      // FIX B-018: Removed silent fuzzy name resolution. FabricID must be exact.
      if (!fabricMasterMap[item.FabricID]) {
        return errorResponse('Line ' + lineNum + ': Fabric not found: ' + item.FabricID + '. Use the exact FabricID from Fabric Master.', 'VALIDATION_ERROR');
      }

      var fcEntry = fabricColorMap[item.FabricColorID];
      if (!fcEntry) {
        return errorResponse('Line ' + lineNum + ': Invalid FabricColorID: ' + item.FabricColorID, 'VALIDATION_ERROR');
      }
      if (fcEntry.FabricID !== item.FabricID) {
        return errorResponse('Line ' + lineNum + ': FabricColor does not belong to selected Fabric', 'VALIDATION_ERROR');
      }

      if (!isPositiveNumber(item.OrderedQty)) {
        return errorResponse('Line ' + lineNum + ': OrderedQty must be positive', 'VALIDATION_ERROR');
      }
      if (['Meters', 'Yards'].indexOf(item.UOM) === -1) {
        return errorResponse('Line ' + lineNum + ': Invalid UOM. Must be Meters or Yards', 'VALIDATION_ERROR');
      }
      if (parseFloat(item.Rate) < 0) {
        return errorResponse('Line ' + lineNum + ': Rate cannot be negative', 'VALIDATION_ERROR');
      }

      var key = item.FabricID + '|' + item.FabricColorID;
      if (seenKeys[key]) {
        return errorResponse('Duplicate line at ' + lineNum + ': same Fabric + Color combination', 'VALIDATION_ERROR');
      }
      seenKeys[key] = true;
    }

    var now         = nowISO();
    var user        = auth.session;
    var poID        = generateYearID('PO', 'PurchaseOrders', 1);
    var totalAmount = data.items.reduce(function(sum, item) {
      return sum + Math.round((parseFloat(item.OrderedQty) || 0) * (parseFloat(item.Rate) || 0) * 100) / 100;
    }, 0);

    appendRow('PurchaseOrders', [
      poID, poID, data.PODate, data.VendorID, sanitize(data.Remarks || ''),
      'Draft', Math.round(totalAmount * 100) / 100, '', '', user.userID, now, '', ''
    ]);

    data.items.forEach(function(item) {
      var qty    = parseFloat(item.OrderedQty) || 0;
      var rate   = parseFloat(item.Rate) || 0;
      var amount = Math.round(qty * rate * 100) / 100;
      var itemID = generateID('POI', 'PurchaseOrderItems', 1);

      // PurchaseOrderItems: POItemID,POID,FabricID,FabricColorID,OrderedQty,ReceivedQty,PendingQty,UOM,Width,Rate,Amount[,ExcessQty,ItemStatus]
      appendRow('PurchaseOrderItems', [
        itemID, poID, item.FabricID, item.FabricColorID,
        qty, 0, qty, item.UOM, parseFloat(item.Width) || 0, rate, amount,
        0,        // ExcessQty
        'Pending' // ItemStatus
      ]);

      _updateInventoryPurchasedQty(item.FabricID, item.FabricColorID, qty);
    });

    writeAuditLog({ userID: user.userID, userName: user.userName, action: 'CREATE', module: 'PurchaseOrders', recordID: poID, oldValues: '', newValues: data });
    return successResponse({ POID: poID, totalAmount: totalAmount }, 'Purchase Order created: ' + poID);
  } catch (e) {
    return errorResponse('Failed to create PO: ' + e.message, 'CREATE_ERROR');
  }
}

// ── UPDATE PO (Draft only) ─────────────────────────────────

function updatePO(token, poID, data) {
  try {
    var auth = requireWriteAccess(token, 'PurchaseOrders');
    if (!auth.valid) return auth.error;

    if (!poID) return errorResponse('POID required', 'VALIDATION_ERROR');

    var rowIndex = findRowByID('PurchaseOrders', 1, poID);
    if (!rowIndex) return errorResponse('Purchase Order not found: ' + poID, 'NOT_FOUND');

    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('PurchaseOrders');
    var old   = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
    var status = old[SCHEMA.PurchaseOrders.Status - 1];

    if (status !== 'Draft') {
      return errorResponse('Only Draft POs can be edited. Current status: ' + status, 'INVALID_STATUS');
    }

    var now = nowISO();
    if (data.PODate)   sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.PODate).setValue(data.PODate);
    if (data.VendorID) {
      if (!idExists('VendorMaster', 'VendorID', data.VendorID)) return errorResponse('Invalid VendorID', 'VALIDATION_ERROR');
      sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.VendorID).setValue(data.VendorID);
    }
    if (data.Remarks !== undefined) sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.Remarks).setValue(sanitize(data.Remarks));
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.UpdatedBy).setValue(auth.session.userID);
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.UpdatedAt).setValue(now);

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'UPDATE', module: 'PurchaseOrders', recordID: poID, oldValues: old, newValues: data });
    return successResponse({ POID: poID }, 'Purchase Order updated');
  } catch (e) {
    return errorResponse('Failed to update PO: ' + e.message, 'UPDATE_ERROR');
  }
}

// ── APPROVE PO ─────────────────────────────────────────────

function approvePO(token, poID) {
  try {
    var auth = requireWriteAccess(token, 'ApprovePO');
    if (!auth.valid) return auth.error;

    var role = auth.session.role;
    if (role !== 'Super Admin' && role !== 'Admin') {
      return errorResponse('Only Admin or Super Admin can approve Purchase Orders', 'FORBIDDEN');
    }

    if (!poID) return errorResponse('POID required', 'VALIDATION_ERROR');

    var rowIndex = findRowByID('PurchaseOrders', 1, poID);
    if (!rowIndex) return errorResponse('Purchase Order not found: ' + poID, 'NOT_FOUND');

    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('PurchaseOrders');
    var old   = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
    var status = old[SCHEMA.PurchaseOrders.Status - 1];

    if (status !== 'Draft') {
      return errorResponse('Only Draft POs can be approved. Current status: ' + status, 'INVALID_STATUS');
    }

    var now = nowISO();
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.Status).setValue('Approved');
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.ApprovedBy).setValue(auth.session.userID);
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.ApprovedAt).setValue(now);
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.UpdatedBy).setValue(auth.session.userID);
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.UpdatedAt).setValue(now);

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'APPROVE', module: 'PurchaseOrders', recordID: poID, oldValues: { Status: 'Draft' }, newValues: { Status: 'Approved' } });
    return successResponse({ POID: poID, status: 'Approved' }, 'Purchase Order approved: ' + poID);
  } catch (e) {
    return errorResponse('Failed to approve PO: ' + e.message, 'APPROVE_ERROR');
  }
}

// ── CANCEL PO ──────────────────────────────────────────────

// FIX B-005: Reverses Inventory.PurchasedQty for all items when PO is cancelled
function cancelPO(token, poID, reason) {
  try {
    var auth = requireWriteAccess(token, 'PurchaseOrders');
    if (!auth.valid) return auth.error;

    if (!poID) return errorResponse('POID required', 'VALIDATION_ERROR');

    var rowIndex = findRowByID('PurchaseOrders', 1, poID);
    if (!rowIndex) return errorResponse('Purchase Order not found: ' + poID, 'NOT_FOUND');

    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('PurchaseOrders');
    var old   = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
    var status = old[SCHEMA.PurchaseOrders.Status - 1];

    var allowed = PO_TRANSITIONS[status] || [];
    if (allowed.indexOf('Cancelled') === -1) {
      return errorResponse('PO in status "' + status + '" cannot be cancelled', 'INVALID_STATUS');
    }

    var now = nowISO();
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.Status).setValue('Cancelled');
    if (reason) sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.Remarks).setValue(old[SCHEMA.PurchaseOrders.Remarks - 1] + ' [Cancelled: ' + sanitize(reason) + ']');
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.UpdatedBy).setValue(auth.session.userID);
    sheet.getRange(rowIndex, SCHEMA.PurchaseOrders.UpdatedAt).setValue(now);

    // FIX B-005: Reverse PurchasedQty in Inventory for all PO items
    if (status === 'Draft' || status === 'Approved') {
      var poItems = getSheetData('PurchaseOrderItems').filter(function(r) { return r.POID === poID; });
      poItems.forEach(function(item) {
        var orderedQty = parseFloat(item.OrderedQty) || 0;
        if (orderedQty > 0) {
          _reverseInventoryPurchasedQty(item.FabricID, item.FabricColorID, orderedQty);
        }
      });
    }

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'UPDATE', module: 'PurchaseOrders', recordID: poID, oldValues: { Status: status }, newValues: { Status: 'Cancelled', reason: reason } });
    return successResponse({ POID: poID }, 'Purchase Order cancelled');
  } catch (e) {
    return errorResponse('Failed to cancel PO: ' + e.message, 'CANCEL_ERROR');
  }
}

// ── Internal helpers ───────────────────────────────────────

function _updateInventoryPurchasedQty(fabricID, fabricColorID, qty) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Inventory');
  var rows  = getSheetData('Inventory').filter(function(r) {
    return r.FabricID === fabricID && r.FabricColorID === fabricColorID;
  });

  if (rows.length === 0) {
    var id = generateID('INV', 'Inventory', 1);
    appendRow('Inventory', [id, fabricID, fabricColorID, qty, 0, 0, 0, 0, nowISO()]);
  } else {
    var allData = sheet.getDataRange().getValues();
    for (var r = 1; r < allData.length; r++) {
      if (allData[r][SCHEMA.Inventory.FabricID - 1] === fabricID &&
          allData[r][SCHEMA.Inventory.FabricColorID - 1] === fabricColorID) {
        var current = parseFloat(allData[r][SCHEMA.Inventory.PurchasedQty - 1]) || 0;
        sheet.getRange(r + 1, SCHEMA.Inventory.PurchasedQty).setValue(Math.round((current + qty) * 1000) / 1000);
        sheet.getRange(r + 1, SCHEMA.Inventory.LastUpdated).setValue(nowISO());
        break;
      }
    }
  }
}

// FIX B-005: New function to reverse PurchasedQty on PO cancellation
function _reverseInventoryPurchasedQty(fabricID, fabricColorID, qty) {
  try {
    var ss      = getSpreadsheet();
    var sheet   = ss.getSheetByName('Inventory');
    var allData = sheet.getDataRange().getValues();
    for (var r = 1; r < allData.length; r++) {
      if (allData[r][SCHEMA.Inventory.FabricID - 1] === fabricID &&
          allData[r][SCHEMA.Inventory.FabricColorID - 1] === fabricColorID) {
        var current = parseFloat(allData[r][SCHEMA.Inventory.PurchasedQty - 1]) || 0;
        var newVal  = Math.max(0, Math.round((current - qty) * 1000) / 1000);
        sheet.getRange(r + 1, SCHEMA.Inventory.PurchasedQty).setValue(newVal);
        sheet.getRange(r + 1, SCHEMA.Inventory.LastUpdated).setValue(nowISO());
        break;
      }
    }
  } catch (e) {
    Logger.log('_reverseInventoryPurchasedQty error: ' + e.message);
  }
}
