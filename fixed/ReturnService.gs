// ============================================================
// ReturnService.gs — Returns Management
// FIXED: B-008 InventoryLedger BalanceQty now writes running balance
//        B-016 generateID called with sheet name for sequential IDs
// ============================================================

const RETURN_TYPES = ['Customer Return', 'Fabric Return', 'Fabricator Return'];

function getReturns(token, filters) {
  var auth = requireReadAccess(token, 'Returns');
  if (!auth.valid) return auth.error;
  try {
    filters = filters || {};
    var rows = getSheetData('Returns');

    if (filters.returnType) rows = rows.filter(function(r) { return r.ReturnType === filters.returnType; });
    if (filters.status)     rows = rows.filter(function(r) { return r.Status     === filters.status; });

    var buyerMap  = {}; getSheetData('BuyerMaster').forEach(function(b) { buyerMap[b.BuyerID]   = b.BuyerName; });
    var vendorMap = {}; getSheetData('VendorMaster').forEach(function(v) { vendorMap[v.VendorID] = v.VendorName; });

    var enriched = rows.map(function(r) {
      var refName = buyerMap[r.BuyerID] || vendorMap[r.BuyerID] || r.BuyerID || '';
      return Object.assign({}, r, { ReferenceName: refName });
    });

    enriched.sort(function(a, b) { return String(b.ReturnID).localeCompare(String(a.ReturnID)); });
    return successResponse(enriched);
  } catch (e) {
    return errorResponse('Failed to get returns: ' + e.message);
  }
}

function getReturnByID(token, returnID) {
  var auth = requireReadAccess(token, 'Returns');
  if (!auth.valid) return auth.error;
  try {
    var ret = findRowByID('Returns', 'ReturnID', returnID);
    if (!ret) return errorResponse('Return not found');

    var fabricMap = {}; getSheetData('FabricMaster').forEach(function(f) { fabricMap[f.FabricID]     = f.DesignName; });
    var colorMap  = {}; getSheetData('FabricColors').forEach(function(c) { colorMap[c.FabricColorID] = c.ColorName; });
    var buyerMap  = {}; getSheetData('BuyerMaster').forEach(function(b) { buyerMap[b.BuyerID]        = b.BuyerName; });
    var vendorMap = {}; getSheetData('VendorMaster').forEach(function(v) { vendorMap[v.VendorID]     = v.VendorName; });

    ret.ReferenceName = buyerMap[ret.BuyerID] || vendorMap[ret.BuyerID] || ret.BuyerID || '';

    var items = getSheetData('ReturnItems')
      .filter(function(i) { return i.ReturnID === returnID; })
      .map(function(i) {
        return Object.assign({}, i, {
          FabricName: fabricMap[i.FabricID]     || i.FabricID     || '',
          ColorName:  colorMap[i.FabricColorID] || i.FabricColorID || ''
        });
      });

    ret.Items = items;
    return successResponse(ret);
  } catch (e) {
    return errorResponse('Failed to get return: ' + e.message);
  }
}

function createReturn(token, data) {
  var auth = requireWriteAccess(token, 'Returns');
  if (!auth.valid) return auth.error;
  try {
    if (!data.ReturnType) return errorResponse('Return type is required');
    if (!data.ReturnDate) return errorResponse('Return date is required');
    if (RETURN_TYPES.indexOf(data.ReturnType) === -1) return errorResponse('Invalid return type');
    if (!Array.isArray(data.items) || data.items.length === 0) return errorResponse('At least one item is required');

    var now      = nowISO();
    var returnID = generateYearID('RTN', 'Returns', 1);

    appendRow('Returns', {
      ReturnID:    returnID,
      ReturnDate:  data.ReturnDate,
      ReturnType:  data.ReturnType,
      ReferenceID: sanitizeString(data.ReferenceID || ''),
      BuyerID:     sanitizeString(data.BuyerID     || ''),
      Reason:      sanitizeString(data.Reason      || ''),
      Status:      'Open',
      CreatedBy:   auth.session.userID,
      CreatedAt:   now,
      UpdatedBy:   '',
      UpdatedAt:   ''
    });

    data.items.forEach(function(item) {
      // FIX B-016: Pass sheet name so IDs are sequential, not timestamp-random
      var itemID = generateID('RTNI', 'ReturnItems', 1);
      appendRow('ReturnItems', {
        ReturnItemID:  itemID,
        ReturnID:      returnID,
        FabricID:      item.FabricID      || '',
        FabricColorID: item.FabricColorID || '',
        BundleID:      item.BundleID      || '',
        Qty:           parseFloat(item.Qty) || 0,
        UOM:           item.UOM            || 'Meters',
        Remarks:       sanitizeString(item.Remarks || '')
      });

      if (data.ReturnType === 'Fabric Return' && item.FabricID && item.FabricColorID) {
        _addReturnedStockToInventory(item.FabricID, item.FabricColorID, parseFloat(item.Qty) || 0, returnID, auth.session.userID, now);
      }
    });

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'CREATE', module: 'Returns', recordID: returnID, oldValues: '', newValues: data });
    return successResponse({ returnID: returnID }, 'Return created: ' + returnID);
  } catch (e) {
    return errorResponse('Failed to create return: ' + e.message);
  }
}

function updateReturnStatus(token, returnID, newStatus) {
  var auth = requireWriteAccess(token, 'Returns');
  if (!auth.valid) return auth.error;
  try {
    var ret = findRowByID('Returns', 'ReturnID', returnID);
    if (!ret) return errorResponse('Return not found');

    var allowed = { Open: ['Processed', 'Cancelled'], Processed: [], Cancelled: [] };
    if (!allowed[ret.Status] || allowed[ret.Status].indexOf(newStatus) === -1) {
      return errorResponse('Cannot change status from ' + ret.Status + ' to ' + newStatus);
    }

    var before    = Object.assign({}, ret);
    ret.Status    = newStatus;
    ret.UpdatedBy = auth.session.userID;
    ret.UpdatedAt = nowISO();
    updateRowByID('Returns', 'ReturnID', returnID, ret);

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'STATUS_CHANGE', module: 'Returns', recordID: returnID, oldValues: before, newValues: ret });
    return successResponse({ returnID: returnID, status: newStatus }, 'Status updated to ' + newStatus);
  } catch (e) {
    return errorResponse('Failed to update return status: ' + e.message);
  }
}

// ── Internal: Add returned fabric back to stock ────────────

// FIX B-008: BalanceQty in InventoryLedger now writes the new running
// AvailableQty (after return), not just the returned quantity.
function _addReturnedStockToInventory(fabricID, fabricColorID, qty, returnID, userID, now) {
  try {
    var invRows  = getSheetData('Inventory');
    var existing = invRows.filter(function(r) {
      return r.FabricID === fabricID && r.FabricColorID === fabricColorID;
    });

    var newBalance;

    if (existing.length > 0) {
      var inv         = existing[0];
      var newReturned = (Number(inv.ReturnedQty)  || 0) + qty;
      var newAvail    = (Number(inv.AvailableQty) || 0) + qty;
      newBalance      = newAvail;
      inv.ReturnedQty  = newReturned;
      inv.AvailableQty = newAvail;
      inv.LastUpdated  = now;
      updateRowByID('Inventory', 'InventoryID', inv.InventoryID, inv);
    } else {
      var invID = generateID('INV', 'Inventory', 1);
      newBalance = qty;
      appendRow('Inventory', {
        InventoryID:   invID,
        FabricID:      fabricID,
        FabricColorID: fabricColorID,
        PurchasedQty:  0,
        ReceivedQty:   0,
        IssuedQty:     0,
        ReturnedQty:   qty,
        AvailableQty:  qty,
        LastUpdated:   now
      });
    }

    // FIX B-008: BalanceQty = new running AvailableQty, not just returned qty
    var ledgerID = generateID('LGR', 'InventoryLedger', 1);
    appendRow('InventoryLedger', {
      LedgerID:        ledgerID,
      TransactionDate: now,
      TransactionType: 'Fabric Return',
      ReferenceNumber: returnID,
      FabricID:        fabricID,
      FabricColorID:   fabricColorID,
      InQty:           qty,
      OutQty:          0,
      BalanceQty:      newBalance,   // FIX: running balance after return
      Remarks:         'Return receipt: ' + returnID,
      CreatedBy:       userID
    });
  } catch (e) {
    Logger.log('_addReturnedStockToInventory error: ' + e.message);
  }
}
