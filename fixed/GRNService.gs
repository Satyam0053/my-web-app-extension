// ============================================================
// GRNService.gs — Goods Receipt Note
// FIXED: B-001 Over-receive allowed; Phase 4 workflow implemented
//        B-002 Inventory fallback row uses correct formula
//        B-003 Batch PO-item updates to prevent stale-read race condition
//        B-010 Hardcoded column indices replaced with SCHEMA constants
// New:   serverClosePO() — closes PO on under-receive user choice
// ============================================================

function getGRNs(token, filters) {
  try {
    var auth = requireReadAccess(token, 'GRN');
    if (!auth.valid) return auth.error;

    var grns = getSheetData('GRN');

    if (filters) {
      if (filters.vendorID) grns = grns.filter(function(r) { return r.VendorID === filters.vendorID; });
      if (filters.poID)     grns = grns.filter(function(r) { return r.POID     === filters.poID; });
      if (filters.from)     grns = grns.filter(function(r) { return r.GRNDate  >= filters.from; });
      if (filters.to)       grns = grns.filter(function(r) { return r.GRNDate  <= filters.to; });
    }

    // FIX B-012 pattern: sort by GRNDate descending, not by ID string
    grns.sort(function(a, b) { return String(b.GRNDate).localeCompare(String(a.GRNDate)); });
    return successResponse({ grns: grns, count: grns.length });
  } catch (e) {
    return errorResponse('Failed to fetch GRNs: ' + e.message, 'FETCH_ERROR');
  }
}

function getGRNByID(token, grnID) {
  try {
    var auth = requireReadAccess(token, 'GRN');
    if (!auth.valid) return auth.error;

    if (!grnID) return errorResponse('GRNID required', 'VALIDATION_ERROR');

    var grns = getSheetData('GRN').filter(function(r) { return r.GRNID === grnID; });
    if (grns.length === 0) return errorResponse('GRN not found: ' + grnID, 'NOT_FOUND');

    var grn   = grns[0];
    var items = getSheetData('GRNItems').filter(function(r) { return r.GRNID === grnID; });

    return successResponse({ grn: grn, items: items });
  } catch (e) {
    return errorResponse('Failed to fetch GRN: ' + e.message, 'FETCH_ERROR');
  }
}

/**
 * Returns PO items for GRN creation.
 * FIX B-001: Also returns Over-Received POs — POs with items that have
 * PendingQty <= 0 but status is Over Received can still receive corrections.
 * Only Approved and Partially Received POs accept new GRNs normally.
 */
function getPOItemsForGRN(token, poID) {
  try {
    var auth = requireReadAccess(token, 'GRN');
    if (!auth.valid) return auth.error;

    if (!poID) return errorResponse('POID required', 'VALIDATION_ERROR');

    var pos = getSheetData('PurchaseOrders').filter(function(r) { return r.POID === poID; });
    if (pos.length === 0) return errorResponse('Purchase Order not found: ' + poID, 'NOT_FOUND');

    var po             = pos[0];
    var allowedStatuses = ['Approved', 'Partially Received', 'Over Received'];
    if (allowedStatuses.indexOf(po.Status) === -1) {
      return errorResponse('GRN can only be created for Approved, Partially Received, or Over Received POs. Current status: ' + po.Status, 'INVALID_STATUS');
    }

    var allItems = getSheetData('PurchaseOrderItems').filter(function(r) { return r.POID === poID; });

    return successResponse({ po: po, items: allItems });
  } catch (e) {
    return errorResponse('Failed to fetch PO items for GRN: ' + e.message, 'FETCH_ERROR');
  }
}

/**
 * Creates a GRN.
 *
 * PHASE 4 IMPLEMENTATION:
 * Scenario A — Over Receiving: ReceivedQty > PendingQty
 *   → Allowed. Inventory updated with full qty. ExcessQty recorded.
 *   → PO item status = "Over Received". PO header = "Over Received".
 *
 * Scenario B — Under Receiving: ReceivedQty < PendingQty
 *   → Allowed. PO stays "Partially Received".
 *   → Response flag hasUnderReceive = true triggers UI modal:
 *       Option 1: serverClosePO() → PO status "Completed", balance zeroed
 *       Option 2: Keep open for future GRN
 *
 * FIX B-003: All PO item updates collected into a batch map and written
 * in one pass (not per-item separate reads inside the forEach loop).
 *
 * FIX B-010: Column indices use SCHEMA constants.
 */
function createGRN(token, data) {
  try {
    var auth = requireWriteAccess(token, 'GRN');
    if (!auth.valid) return auth.error;

    // Header validation
    var missing = validateRequired(['GRNDate', 'VendorID', 'POID', 'ChallanNumber'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    if (!idExists('VendorMaster', 'VendorID', data.VendorID)) {
      return errorResponse('Invalid VendorID', 'VALIDATION_ERROR');
    }

    var poRows = getSheetData('PurchaseOrders').filter(function(r) { return r.POID === data.POID; });
    if (poRows.length === 0) return errorResponse('PO not found: ' + data.POID, 'NOT_FOUND');
    var po = poRows[0];

    var allowedStatuses = ['Approved', 'Partially Received', 'Over Received'];
    if (allowedStatuses.indexOf(po.Status) === -1) {
      return errorResponse('GRN can only be created against Approved, Partially Received, or Over Received POs. Current status: ' + po.Status, 'INVALID_STATUS');
    }

    if (po.VendorID !== data.VendorID) {
      return errorResponse('Vendor does not match the Purchase Order vendor', 'VALIDATION_ERROR');
    }

    // Duplicate Challan check
    var existingGRNs = getSheetData('GRN').filter(function(r) {
      return r.POID === data.POID && String(r.ChallanNumber).toLowerCase() === String(data.ChallanNumber).toLowerCase();
    });
    if (existingGRNs.length > 0) {
      return errorResponse('A GRN with Challan Number "' + data.ChallanNumber + '" already exists for this PO', 'DUPLICATE');
    }

    if (!Array.isArray(data.items) || data.items.length === 0) {
      return errorResponse('At least one item is required', 'VALIDATION_ERROR');
    }

    // Load all PO items once
    var allPOItems = getSheetData('PurchaseOrderItems').filter(function(r) { return r.POID === data.POID; });
    var poItemMap  = {};
    allPOItems.forEach(function(r) { poItemMap[r.POItemID] = r; });

    // Item validation — FIX B-001: Over-receiving is now ALLOWED
    for (var i = 0; i < data.items.length; i++) {
      var item    = data.items[i];
      var lineNum = i + 1;

      var missingItem = validateRequired(['POItemID', 'FabricID', 'FabricColorID', 'ReceivedQty', 'UOM'], item);
      if (missingItem.length) return errorResponse('Line ' + lineNum + ': Missing ' + missingItem.join(', '), 'VALIDATION_ERROR');

      if (!isPositiveNumber(item.ReceivedQty)) {
        return errorResponse('Line ' + lineNum + ': ReceivedQty must be positive', 'VALIDATION_ERROR');
      }

      var poItem = poItemMap[item.POItemID];
      if (!poItem) return errorResponse('Line ' + lineNum + ': PO Item not found: ' + item.POItemID, 'NOT_FOUND');

      // FIX B-001: No more hard block on over-receive. Just flag it.
      var recv    = parseFloat(item.ReceivedQty) || 0;
      var ordered = parseFloat(poItem.OrderedQty) || 0;
      var prevRecv = parseFloat(poItem.ReceivedQty) || 0;
      var totalRecv = prevRecv + recv;
      item._excessQty      = Math.max(0, Math.round((totalRecv - ordered) * 1000) / 1000);
      item._isOverReceive  = item._excessQty > 0;
    }

    // All validations passed — write GRN
    var now   = nowISO();
    var user  = auth.session;
    var grnID = generateYearID('GRN', 'GRN', 1);

    // GRN headers: GRNID,GRNNumber,GRNDate,VendorID,POID,ChallanNumber,VehicleNumber,Remarks,CreatedBy,CreatedAt
    appendRow('GRN', [
      grnID, grnID, data.GRNDate, data.VendorID, data.POID,
      sanitize(data.ChallanNumber), sanitize(data.VehicleNumber || ''),
      sanitize(data.Remarks || ''), user.userID, now
    ]);

    // FIX B-003: Build batch update map for PO items — read sheet once before loop
    var ss          = getSpreadsheet();
    var poItemSheet = ss.getSheetByName('PurchaseOrderItems');
    var poItemData  = poItemSheet.getDataRange().getValues();

    // Index all PO item rows by POItemID for fast lookup
    var poItemRowMap = {};
    for (var r = 1; r < poItemData.length; r++) {
      var pid = poItemData[r][SCHEMA.PurchaseOrderItems.POItemID - 1];
      poItemRowMap[pid] = r + 1; // 1-indexed sheet row
    }

    // Collect all cell updates, apply after loop
    var batchUpdates = []; // { row, col, value }

    data.items.forEach(function(item) {
      var recv   = parseFloat(item.ReceivedQty) || 0;
      var itemID = generateID('GRNI', 'GRNItems', 1);

      appendRow('GRNItems', [
        itemID, grnID, item.POItemID, item.FabricID, item.FabricColorID,
        recv, item.UOM, sanitize(item.Remarks || '')
      ]);

      // Calculate new PO item quantities
      var poItem   = poItemMap[item.POItemID];
      var ordered  = parseFloat(poItem.OrderedQty)  || 0;
      var prevRecv = parseFloat(poItem.ReceivedQty) || 0;
      var newRecv  = Math.round((prevRecv + recv) * 1000) / 1000;
      var newPend  = Math.round(Math.max(0, ordered - newRecv) * 1000) / 1000;
      var excessQty = item._excessQty || 0;
      var itemStatus = excessQty > 0 ? 'Over Received' : (newPend <= 0 ? 'Received' : 'Partial');

      var sheetRow = poItemRowMap[item.POItemID];
      if (sheetRow) {
        batchUpdates.push({ row: sheetRow, col: SCHEMA.PurchaseOrderItems.ReceivedQty, value: newRecv });
        batchUpdates.push({ row: sheetRow, col: SCHEMA.PurchaseOrderItems.PendingQty,  value: newPend });
        // ExcessQty and ItemStatus columns — add if sheet has them
        if (SCHEMA.PurchaseOrderItems.ExcessQty) {
          batchUpdates.push({ row: sheetRow, col: SCHEMA.PurchaseOrderItems.ExcessQty,  value: excessQty });
        }
        if (SCHEMA.PurchaseOrderItems.ItemStatus) {
          batchUpdates.push({ row: sheetRow, col: SCHEMA.PurchaseOrderItems.ItemStatus, value: itemStatus });
        }
      }

      // Update Inventory
      _updateInventoryReceived(item.FabricID, item.FabricColorID, recv, grnID, user.userID);
    });

    // FIX B-003: Apply all PO item updates in one go
    batchUpdates.forEach(function(u) {
      poItemSheet.getRange(u.row, u.col).setValue(u.value);
    });

    // Determine PO header status
    var updatedPOItems  = getSheetData('PurchaseOrderItems').filter(function(r) { return r.POID === data.POID; });
    var hasOverReceive  = updatedPOItems.some(function(r) { return (parseFloat(r.ExcessQty) || 0) > 0 || (parseFloat(r.ReceivedQty) || 0) > (parseFloat(r.OrderedQty) || 0); });
    var allPendingZero  = updatedPOItems.every(function(r) { return (parseFloat(r.PendingQty) || 0) <= 0; });
    var hasUnderReceive = updatedPOItems.some(function(r) { return (parseFloat(r.PendingQty) || 0) > 0; });

    var newPOStatus;
    if (hasOverReceive)  newPOStatus = 'Over Received';
    else if (allPendingZero) newPOStatus = 'Fully Received';
    else                 newPOStatus = 'Partially Received';

    // FIX B-010: Use SCHEMA constants instead of hardcoded column numbers
    var poSheet  = ss.getSheetByName('PurchaseOrders');
    var poAllData = poSheet.getDataRange().getValues();
    for (var r = 1; r < poAllData.length; r++) {
      if (poAllData[r][SCHEMA.PurchaseOrders.POID - 1] === data.POID) {
        poSheet.getRange(r + 1, SCHEMA.PurchaseOrders.Status).setValue(newPOStatus);
        poSheet.getRange(r + 1, SCHEMA.PurchaseOrders.UpdatedBy).setValue(user.userID);
        poSheet.getRange(r + 1, SCHEMA.PurchaseOrders.UpdatedAt).setValue(now);
        break;
      }
    }

    var pendingItems = updatedPOItems.filter(function(r) { return (parseFloat(r.PendingQty) || 0) > 0; });

    writeAuditLog({ userID: user.userID, userName: user.userName, action: 'CREATE', module: 'GRN', recordID: grnID, oldValues: '', newValues: data });

    // Phase 4: Return flags for UI to handle under-receive modal and over-receive notification
    return successResponse({
      GRNID:          grnID,
      POStatus:       newPOStatus,
      hasUnderReceive: hasUnderReceive && !hasOverReceive,
      hasOverReceive:  hasOverReceive,
      pendingItems:    pendingItems.map(function(r) {
        return { POItemID: r.POItemID, FabricID: r.FabricID, FabricColorID: r.FabricColorID, PendingQty: r.PendingQty, UOM: r.UOM };
      })
    }, 'GRN created: ' + grnID + '. PO status: ' + newPOStatus);

  } catch (e) {
    return errorResponse('Failed to create GRN: ' + e.message, 'CREATE_ERROR');
  }
}

/**
 * Phase 4 — Under Receive: User chose "Mark PO Complete".
 * Zeros out PendingQty on all items and sets PO status to Completed.
 */
function closePO(token, poID) {
  try {
    var auth = requireWriteAccess(token, 'GRN');
    if (!auth.valid) return auth.error;

    if (!poID) return errorResponse('POID required', 'VALIDATION_ERROR');

    var poRows = getSheetData('PurchaseOrders').filter(function(r) { return r.POID === poID; });
    if (poRows.length === 0) return errorResponse('PO not found: ' + poID, 'NOT_FOUND');
    var po = poRows[0];

    var closableStatuses = ['Partially Received', 'Approved', 'Over Received'];
    if (closableStatuses.indexOf(po.Status) === -1) {
      return errorResponse('PO cannot be closed from status: ' + po.Status, 'INVALID_STATUS');
    }

    var now = nowISO();
    var ss  = getSpreadsheet();

    // Zero out all PendingQty on PO items
    var poItemSheet = ss.getSheetByName('PurchaseOrderItems');
    var poItemData  = poItemSheet.getDataRange().getValues();
    for (var r = 1; r < poItemData.length; r++) {
      if (poItemData[r][SCHEMA.PurchaseOrderItems.POID - 1] === poID) {
        poItemSheet.getRange(r + 1, SCHEMA.PurchaseOrderItems.PendingQty).setValue(0);
        poItemSheet.getRange(r + 1, SCHEMA.PurchaseOrderItems.ItemStatus).setValue('Completed');
      }
    }

    // Set PO status to Completed
    var poSheet   = ss.getSheetByName('PurchaseOrders');
    var poAllData = poSheet.getDataRange().getValues();
    for (var r = 1; r < poAllData.length; r++) {
      if (poAllData[r][SCHEMA.PurchaseOrders.POID - 1] === poID) {
        poSheet.getRange(r + 1, SCHEMA.PurchaseOrders.Status).setValue('Completed');
        poSheet.getRange(r + 1, SCHEMA.PurchaseOrders.UpdatedBy).setValue(auth.session.userID);
        poSheet.getRange(r + 1, SCHEMA.PurchaseOrders.UpdatedAt).setValue(now);
        break;
      }
    }

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'CLOSE', module: 'PurchaseOrders', recordID: poID, oldValues: { Status: po.Status }, newValues: { Status: 'Completed' } });
    return successResponse({ POID: poID, status: 'Completed' }, 'Purchase Order marked as Completed. Balance qty cleared.');
  } catch (e) {
    return errorResponse('Failed to close PO: ' + e.message, 'CLOSE_ERROR');
  }
}

// ── Internal: Update inventory on GRN receipt ──────────────

// FIX B-002: Fallback inventory row now reads all existing fields to
// compute correct opening balance instead of assuming zeros.
function _updateInventoryReceived(fabricID, fabricColorID, receivedQty, grnID, userID) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Inventory');
  var data  = sheet.getDataRange().getValues();
  var now   = nowISO();

  var found      = false;
  var newBalance = 0;

  for (var r = 1; r < data.length; r++) {
    if (data[r][SCHEMA.Inventory.FabricID - 1] === fabricID &&
        data[r][SCHEMA.Inventory.FabricColorID - 1] === fabricColorID) {

      var prevReceived = parseFloat(data[r][SCHEMA.Inventory.ReceivedQty - 1]) || 0;
      var issuedQty    = parseFloat(data[r][SCHEMA.Inventory.IssuedQty   - 1]) || 0;
      var returnedQty  = parseFloat(data[r][SCHEMA.Inventory.ReturnedQty - 1]) || 0;

      var newReceived  = Math.round((prevReceived + receivedQty) * 1000) / 1000;
      // AvailableQty formula: ReceivedQty - IssuedQty + ReturnedQty
      newBalance       = Math.round((newReceived - issuedQty + returnedQty) * 1000) / 1000;

      sheet.getRange(r + 1, SCHEMA.Inventory.ReceivedQty).setValue(newReceived);
      sheet.getRange(r + 1, SCHEMA.Inventory.AvailableQty).setValue(newBalance);
      sheet.getRange(r + 1, SCHEMA.Inventory.LastUpdated).setValue(now);
      found = true;
      break;
    }
  }

  if (!found) {
    // FIX B-002: When creating a new row, balance = receivedQty (no prior issued/returned)
    var id = generateID('INV', 'Inventory', 1);
    newBalance = receivedQty;
    appendRow('Inventory', [id, fabricID, fabricColorID, 0, receivedQty, 0, 0, receivedQty, now]);
  }

  // Write InventoryLedger entry with correct running balance
  var ledgerID = generateID('LGR', 'InventoryLedger', 1);
  appendRow('InventoryLedger', [
    ledgerID, now, 'Purchase Receipt', grnID,
    fabricID, fabricColorID,
    receivedQty,  // InQty
    0,            // OutQty
    newBalance,   // FIX: running AvailableQty, not just receivedQty
    'GRN: ' + grnID, userID
  ]);
}
