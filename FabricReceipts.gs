// ============================================================
// File: FabricReceipts.gs — Fabric Receiving & Inventory Update
// Version: 2.4 | PO Completion Control
//
// NEW in v2.4:
//   • receiveFabric now accepts data.poAction = 'partial' | 'complete'
//     sent by the frontend confirmation step.
//   • When poAction === 'complete': PO status → 'Completed',
//     balance is forced to 0 by writing a closing note to the PO's
//     Notes column: "Closed short by X units on <date>".
//   • When poAction === 'partial': PO status → 'Partial', balance
//     remains open exactly as before.
//   • When receivedQty >= balanceQty (full receipt): auto-completes
//     exactly as before — no poAction needed.
//   • If poAction is missing/empty and qty < balance, defaults to
//     'partial' as a safe fallback (should not happen in normal UI
//     flow since frontend always sends it, but guards backend).
//   • updatePOStatusAndNote: new helper that writes both the Status
//     cell AND the Notes cell in a single sheet operation to avoid
//     two separate getRange calls.
// ============================================================

function receiveFabric(data, caller) {
  // ── INPUT VALIDATION ───────────────────────────────────────
  if (!data)          return error('No data received');
  if (!data.poNumber) return error('Purchase Order Number is required');
  if (!data.receivedQty || isNaN(data.receivedQty) || parseFloat(data.receivedQty) <= 0)
    return error('Valid Received Quantity is required');

  // ── FETCH THE PO ──────────────────────────────────────────
  const poResult = JSON.parse(getPOById(data.poNumber));
  if (!poResult.success) return error('Purchase Order not found: ' + data.poNumber);
  const po = poResult.data;

  // Defensive balance calculation (handles old/null rows)
  const balanceQty = (po.Balance_Qty !== undefined && po.Balance_Qty !== null && po.Balance_Qty !== '')
    ? parseFloat(po.Balance_Qty)
    : (parseFloat(po.Ordered_Qty) || 0) - (parseFloat(po.Total_Received) || 0);

  if (isNaN(balanceQty)) return error('Could not determine balance quantity for PO ' + data.poNumber);

  const receivedQty = parseFloat(data.receivedQty);

  // ── QUANTITY CHECK ────────────────────────────────────────
  if (receivedQty > balanceQty + 0.001)
    return error(`Cannot receive ${receivedQty}. Balance quantity is ${balanceQty.toFixed(3)}`);

  // ── GRN NUMBER ────────────────────────────────────────────
  const receiptSheet = getSheet(SHEETS.FABRIC_RECEIPTS);
  const grnNumber    = generateSequentialId('GRN', receiptSheet);

  // ── RECEIVED BY — safe fallback chain ────────────────────
  let receivedBy = '';
  try { receivedBy = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  if (!receivedBy && caller && caller.username) receivedBy = caller.username;
  if (!receivedBy) receivedBy = 'User';

  // ── SAVE GRN ROW ──────────────────────────────────────────
  const receiptRow = [
    grnNumber,
    data.poNumber,
    po.Vendor_Name    || '',
    po.Article_Number || '',
    po.Fabric_Name    || '',
    po.Color          || '',
    po.Width          || '',
    po.Width_Unit     || '',
    receivedQty,
    po.Unit           || 'Meter',
    data.batchNumber      || generateBatchNumber(),
    data.dyeLot           || '',
    data.qualityStatus    || 'Approved',
    data.storageLocation  || '',
    receivedBy,
    new Date(),
    data.notes            || ''
  ];
  receiptSheet.appendRow(receiptRow);

  // ── UPDATE INVENTORY ──────────────────────────────────────
  const invResult = JSON.parse(updateInventoryOnReceiving(po, receivedQty, grnNumber, data));
  if (!invResult.success)
    return error('GRN ' + grnNumber + ' created but inventory update failed: ' + invResult.error);

  // ── DETERMINE NEW PO STATUS ───────────────────────────────
  //
  // Three cases:
  //   A) receivedQty >= balanceQty  → auto-complete (full receipt)
  //   B) receivedQty <  balanceQty  AND poAction === 'complete'
  //                                 → user chose to close PO short
  //   C) receivedQty <  balanceQty  AND poAction === 'partial' (or missing)
  //                                 → keep open, status = Partial
  //
  const isFullReceipt = receivedQty >= balanceQty - 0.001;
  const poAction      = (data.poAction || 'partial').toLowerCase(); // 'complete' | 'partial'

  let newStatus    = 'Partial';
  let closingNote  = '';

  if (isFullReceipt) {
    // Case A — full or over-receipt: auto-complete, no dialog needed
    newStatus = 'Completed';
  } else if (poAction === 'complete') {
    // Case B — user explicitly closed PO short
    newStatus   = 'Completed';
    const shortBy = (balanceQty - receivedQty).toFixed(3);
    const today   = new Date().toLocaleDateString('en-IN');
    closingNote   = `Closed short by ${shortBy} ${po.Unit || 'units'} on ${today}`;
  }
  // Case C — poAction === 'partial': newStatus stays 'Partial', no note

  // ── UPDATE PO STATUS (+ optional closing note) ───────────
  updatePOStatusAndNote(data.poNumber, newStatus, closingNote, caller);

  try {
    logActivity('RECEIVE', 'FabricReceipt', grnNumber, {
      poNumber : data.poNumber,
      qty      : receivedQty,
      grnNumber,
      poStatus : newStatus,
      closedShort: closingNote || null
    }, caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({
    grnNumber,
    inventoryId  : invResult.data.inventoryId,
    previousStock: invResult.data.previousStock,
    receivedQty,
    updatedStock : invResult.data.updatedStock,
    poStatus     : newStatus,
    closingNote  : closingNote || null
  }, `GRN ${grnNumber} created. Stock: ${invResult.data.previousStock} → ${invResult.data.updatedStock}. PO: ${newStatus}.`);
}

// ── UPDATE PO STATUS AND OPTIONAL CLOSING NOTE ───────────────
// Writes Status (col 11) and, when closingNote is provided,
// appends it to the Notes column (col 14) of the PO row.
// Uses a single sheet.getDataRange() read to find the row index,
// then targeted setValues calls.
function updatePOStatusAndNote(poNumber, newStatus, closingNote, caller) {
  if (!poNumber || !newStatus) return error('PO Number and Status are required');

  const sheet  = getSheet(SHEETS.PURCHASE_ORDERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex(r => r[0] === poNumber);
  if (rowIdx === -1) return error('Purchase Order not found: ' + poNumber);

  const sheetRow = rowIdx + 1; // 1-based

  // Write Status — col 11
  sheet.getRange(sheetRow, 11).setValue(newStatus);

  // Status cell background colours
  const colors = {
    Pending  : '#FFF3CD', // yellow
    Partial  : '#CCE5FF', // blue
    Completed: '#D4EDDA', // green
    Cancelled: '#F8D7DA'  // red
  };
  sheet.getRange(sheetRow, 11).setBackground(colors[newStatus] || '#FFFFFF');

  // Append closing note to Notes column (col 14) if provided
  if (closingNote) {
    const existingNote = String(values[rowIdx][13] || '').trim(); // col 14, 0-based index 13
    const updatedNote  = existingNote
      ? existingNote + ' | ' + closingNote
      : closingNote;
    sheet.getRange(sheetRow, 14).setValue(updatedNote);
  }

  try {
    logActivity('UPDATE_STATUS', 'PurchaseOrder', poNumber,
      { status: newStatus, note: closingNote || null },
      caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({ poNumber, status: newStatus });
}

// ── INVENTORY UPDATE ON RECEIVING ────────────────────────────
function updateInventoryOnReceiving(po, receivedQty, grnNumber, receiptData) {
  const invSheet  = getSheet(SHEETS.INVENTORY);
  const sheetData = invSheet.getDataRange().getValues();

  if (sheetData.length === 0)
    return _createNewInventoryRecord(invSheet, po, receivedQty, grnNumber, receiptData);

  const matchIdx = sheetData.findIndex((row, i) => {
    if (i === 0) return false;
    return String(row[1]).trim() === String(po.Article_Number || '').trim() &&
           String(row[3]).trim() === String(po.Color          || '').trim() &&
           String(row[4]).trim() === String(po.Width          || '').trim() &&
           String(row[6]).trim() === String(po.Vendor_Name    || '').trim();
  });

  if (matchIdx === -1)
    return _createNewInventoryRecord(invSheet, po, receivedQty, grnNumber, receiptData);

  const sheetRow    = matchIdx + 1;
  const prevStock   = parseFloat(sheetData[matchIdx][9]) || 0;
  const newStock    = prevStock + receivedQty;
  const inventoryId = sheetData[matchIdx][0];

  invSheet.getRange(sheetRow, 10).setValue(newStock);
  invSheet.getRange(sheetRow, 13).setValue(new Date());
  invSheet.getRange(sheetRow, 14).setValue(grnNumber);

  return success({ inventoryId, previousStock: prevStock, updatedStock: newStock });
}

function _createNewInventoryRecord(invSheet, po, receivedQty, grnNumber, receiptData) {
  const inventoryId  = generateSequentialId('INV', invSheet);
  const reorderLevel = parseFloat((receiptData && receiptData.reorderLevel) || 0) || 50;
  const newRow = [
    inventoryId,
    po.Article_Number || '',
    po.Fabric_Name    || '',
    po.Color          || '',
    po.Width          || '',
    po.Width_Unit     || '',
    po.Vendor_Name    || '',
    po.Unit           || 'Meter',
    receivedQty,
    receivedQty,
    reorderLevel,
    (receiptData && receiptData.storageLocation) || '',
    new Date(),
    grnNumber,
    'Active'
  ];
  invSheet.appendRow(newRow);
  return success({ inventoryId, previousStock: 0, updatedStock: receivedQty });
}

// ── RECEIPTS QUERIES ─────────────────────────────────────────
function getReceipts(filters) {
  const sheet   = getSheet(SHEETS.FABRIC_RECEIPTS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));

  if (filters) {
    if (filters.vendor)
      rows = rows.filter(r =>
        String(r.Vendor_Name || '').toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.dateFrom)
      rows = rows.filter(r => r.Received_At && new Date(r.Received_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => r.Received_At && new Date(r.Received_At) <= new Date(filters.dateTo));
  }

  rows.sort((a, b) => {
    const da = a.Received_At ? new Date(a.Received_At) : 0;
    const db = b.Received_At ? new Date(b.Received_At) : 0;
    return db - da;
  });

  return success(rows);
}

function getReceiptsByPO(poNumber) {
  if (!poNumber) return success([]);
  return success(getReceiptsForPO(poNumber));
}

function generateBatchNumber() {
  const d = new Date();
  return `BATCH-${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}-${Math.random().toString(36).substring(2,6).toUpperCase()}`;
}