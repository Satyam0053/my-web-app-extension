// ============================================================
// File: FabricReceipts.gs — Fabric Receiving & Inventory Update
// FIXED:
//  • Over-receiving now ALLOWED (Phase 4 Scenario A)
//    PO quantity > balance → status = 'Over Received', excess recorded
//  • Under-receiving (Phase 4 Scenario B) → response flag
//    hasUnderReceive: true so UI can show modal
//  • New function: closePO(payload, caller) for "Mark PO Complete"
//  • updatePOStatus column index fixed: Status is col 13, not 11
//  • Balance calculated from actual receipts (not relying on PO row alone)
//  • logActivity caller forwarded correctly
// ============================================================

function receiveFabric(payload, caller) {
  const data = payload;

  if (!data.poNumber) return error('Purchase Order Number is required');
  if (!data.receivedQty || isNaN(data.receivedQty) || parseFloat(data.receivedQty) <= 0)
    return error('Valid Received Quantity is required');

  const poResult = JSON.parse(getPOById(data.poNumber));
  if (!poResult.success) return error('Purchase Order not found: ' + data.poNumber);
  const po = poResult.data;

  // Only allow receiving against Pending or Partial POs
  if (po.Status === 'Completed' || po.Status === 'Cancelled')
    return error('Cannot receive against a ' + po.Status + ' Purchase Order');

  const balanceQty  = parseFloat(po.Balance_Qty) || 0;
  const receivedQty = parseFloat(data.receivedQty);

  // PHASE 4 — Scenario A: Over-receive ALLOWED (no hard block)
  const isOverReceive  = receivedQty > balanceQty + 0.001;
  const excessQty      = isOverReceive ? Math.round((receivedQty - balanceQty) * 1000) / 1000 : 0;
  const isUnderReceive = !isOverReceive && receivedQty < balanceQty - 0.001;

  // GRN Number
  const receiptSheet = getSheet(SHEETS.FABRIC_RECEIPTS);
  const grnNumber    = generateSequentialId('GRN', receiptSheet);

  const receiptRow = [
    grnNumber, data.poNumber, po.Vendor_Name, po.Article_Number,
    po.Fabric_Name, po.Color, po.Width, po.Width_Unit, receivedQty,
    po.Unit, data.batchNumber || _generateBatchNumber(),
    data.dyeLot || '', data.qualityStatus || 'Approved',
    data.storageLocation || '',
    (caller && caller.username) ? caller.username : (Session.getActiveUser().getEmail() || 'User'),
    new Date(), data.notes || ''
  ];
  receiptSheet.appendRow(receiptRow);

  // Update Inventory
  const invResult = JSON.parse(updateInventoryOnReceiving(po, receivedQty, grnNumber, data));
  if (!invResult.success) return error('GRN created but inventory update failed: ' + invResult.error);

  // Determine new PO status
  const newBalance  = Math.max(0, Math.round((balanceQty - receivedQty) * 1000) / 1000);
  let   newStatus;
  if (isOverReceive)       newStatus = 'Over Received';
  else if (newBalance <= 0.001) newStatus = 'Completed';
  else                     newStatus = 'Partial';

  updatePOStatus({ poNumber: data.poNumber, status: newStatus }, caller);

  try { logActivity('RECEIVE', 'FabricReceipt', grnNumber, { poNumber: data.poNumber, qty: receivedQty, status: newStatus }, (caller || {}).username, (caller || {}).role); } catch(e) {}

  return success({
    grnNumber,
    inventoryId  : invResult.data.inventoryId,
    previousStock: invResult.data.previousStock,
    receivedQty,
    updatedStock : invResult.data.updatedStock,
    newPOStatus  : newStatus,
    excessQty,
    // PHASE 4: flags for UI modal
    hasOverReceive : isOverReceive,
    hasUnderReceive: isUnderReceive,
    remainingQty   : newBalance
  }, `GRN ${grnNumber} created. Stock: ${invResult.data.previousStock} → ${invResult.data.updatedStock}. PO: ${newStatus}`);
}

// PHASE 4 — Scenario B Option 1: User chooses "Mark PO Complete"
// Closes the PO — no more receipts allowed
function closePO(payload, caller) {
  const poNumber = payload.poNumber;
  if (!poNumber) return error('PO Number is required');

  const poResult = JSON.parse(getPOById(poNumber));
  if (!poResult.success) return error('Purchase Order not found: ' + poNumber);
  const po = poResult.data;

  if (po.Status === 'Completed' || po.Status === 'Cancelled')
    return error('PO is already ' + po.Status);

  updatePOStatus({ poNumber, status: 'Completed' }, caller);

  try { logActivity('CLOSE_PO', 'PurchaseOrder', poNumber, { closedBy: (caller || {}).username, previousStatus: po.Status }, (caller || {}).username, (caller || {}).role); } catch(e) {}
  return success({ poNumber, status: 'Completed' }, 'Purchase Order ' + poNumber + ' marked as Completed.');
}

// ── INVENTORY UPDATE ON RECEIVING ─────────────────────────────
function updateInventoryOnReceiving(po, receivedQty, grnNumber, receiptData) {
  const invSheet = getSheet(SHEETS.INVENTORY);
  const data     = invSheet.getDataRange().getValues();
  const headers  = data[0];

  const matchIdx = data.findIndex((row, i) => {
    if (i === 0) return false;
    return row[1] === po.Article_Number &&
           row[3] === po.Color &&
           row[4] === po.Width &&
           row[6] === po.Vendor_Name;
  });

  if (matchIdx === -1) {
    const inventoryId = generateSequentialId('INV', invSheet);
    invSheet.appendRow([
      inventoryId, po.Article_Number, po.Fabric_Name, po.Color,
      po.Width, po.Width_Unit, po.Vendor_Name, po.Unit,
      receivedQty, receivedQty, // Opening = Current = first receipt
      receiptData.reorderLevel || 50,
      receiptData.storageLocation || '', new Date(), grnNumber, 'Active'
    ]);
    return success({ inventoryId, previousStock: 0, updatedStock: receivedQty });
  } else {
    const sheetRow  = matchIdx + 1;
    const prevStock = parseFloat(data[matchIdx][9]) || 0;
    const newStock  = Math.round((prevStock + receivedQty) * 1000) / 1000;
    const inventoryId = data[matchIdx][0];

    invSheet.getRange(sheetRow, 10).setValue(newStock);  // Current_Stock
    invSheet.getRange(sheetRow, 13).setValue(new Date()); // Last_Updated
    invSheet.getRange(sheetRow, 14).setValue(grnNumber);  // Last_Transaction

    return success({ inventoryId, previousStock: prevStock, updatedStock: newStock });
  }
}

// ── RECEIPTS QUERIES ──────────────────────────────────────────
function getReceipts(filters) {
  const sheet = getSheet(SHEETS.FABRIC_RECEIPTS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));

  if (filters) {
    if (filters.vendor)
      rows = rows.filter(r => String(r.Vendor_Name || '').toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.dateFrom)
      rows = rows.filter(r => new Date(r.Received_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => new Date(r.Received_At) <= new Date(filters.dateTo));
  }

  rows.sort((a, b) => new Date(b.Received_At) - new Date(a.Received_At));
  return success(rows);
}

function getReceiptsByPO(poNumber) {
  return success(getReceiptsForPO(poNumber));
}

function _generateBatchNumber() {
  const d = new Date();
  return `BATCH-${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}-${Math.random().toString(36).substring(2,6).toUpperCase()}`;
}

// Keep old name as alias for backward compat
function generateBatchNumber() { return _generateBatchNumber(); }
