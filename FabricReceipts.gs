// ============================================================
// File: FabricReceipts.gs — Fabric Receiving & Inventory Update
// ============================================================

function receiveFabric(data) {
  // ── VALIDATION ─────────────────────────────────────────────
  if (!data.poNumber)       return error('Purchase Order Number is required');
  if (!data.receivedQty || isNaN(data.receivedQty) || parseFloat(data.receivedQty) <= 0)
    return error('Valid Received Quantity is required');

  // Fetch the PO
  const poResult = JSON.parse(getPOById(data.poNumber));
  if (!poResult.success) return error('Purchase Order not found: ' + data.poNumber);
  const po = poResult.data;

  // Balance check
  const balanceQty = po.Balance_Qty;
  const receivedQty = parseFloat(data.receivedQty);
  if (receivedQty > balanceQty + 0.001)
    return error(`Cannot receive ${receivedQty}. Balance quantity is ${balanceQty}`);

  // ── GRN NUMBER ──────────────────────────────────────────────
  const receiptSheet = getSheet(SHEETS.FABRIC_RECEIPTS);
  const grnNumber = generateSequentialId('GRN', receiptSheet);

  // ── SAVE GRN ROW ────────────────────────────────────────────
  const receiptRow = [
    grnNumber,
    data.poNumber,
    po.Vendor_Name,
    po.Article_Number,
    po.Fabric_Name,
    po.Color,
    po.Width,
    po.Width_Unit,
    receivedQty,
    po.Unit,
    data.batchNumber || generateBatchNumber(),
    data.dyeLot        || '',
    data.qualityStatus || 'Approved',
    data.storageLocation || '',
    Session.getActiveUser().getEmail() || 'User',
    new Date(),
    data.notes || ''
  ];
  receiptSheet.appendRow(receiptRow);

  // ── UPDATE INVENTORY ────────────────────────────────────────
  const invResult = JSON.parse(updateInventoryOnReceiving(po, receivedQty, grnNumber, data));
  if (!invResult.success) return error('GRN created but inventory update failed: ' + invResult.error);

  // ── UPDATE PO STATUS ────────────────────────────────────────
  const newBalance = balanceQty - receivedQty;
  const newStatus  = newBalance <= 0.001 ? 'Completed' : 'Partial';
  updatePOStatus({ poNumber: data.poNumber, status: newStatus });

  logActivity('RECEIVE', 'FabricReceipt', grnNumber, { poNumber: data.poNumber, qty: receivedQty });

  return success({
    grnNumber,
    inventoryId  : invResult.data.inventoryId,
    previousStock: invResult.data.previousStock,
    receivedQty,
    updatedStock : invResult.data.updatedStock
  }, `GRN ${grnNumber} created. Stock updated from ${invResult.data.previousStock} → ${invResult.data.updatedStock}`);
}

// ── INVENTORY UPDATE ON RECEIVING ────────────────────────────
function updateInventoryOnReceiving(po, receivedQty, grnNumber, receiptData) {
  const invSheet = getSheet(SHEETS.INVENTORY);
  const data     = invSheet.getDataRange().getValues();
  const headers  = data[0];

  // Find existing inventory record (Article + Color + Width + Vendor)
  const matchIdx = data.findIndex((row, i) => {
    if (i === 0) return false;
    return row[1] === po.Article_Number &&
           row[3] === po.Color &&
           row[4] === po.Width &&
           row[6] === po.Vendor_Name;
  });

  if (matchIdx === -1) {
    // CREATE new inventory record
    const inventoryId = generateSequentialId('INV', invSheet);
    const newRow = [
      inventoryId,
      po.Article_Number,
      po.Fabric_Name,
      po.Color,
      po.Width,
      po.Width_Unit,
      po.Vendor_Name,
      po.Unit,
      receivedQty,     // Opening Stock (first receipt)
      receivedQty,     // Current Stock
      data.reorderLevel || 50,
      receiptData.storageLocation || '',
      new Date(),
      grnNumber,
      'Active'
    ];
    invSheet.appendRow(newRow);
    return success({ inventoryId, previousStock: 0, updatedStock: receivedQty });
  } else {
    // UPDATE existing record
    const sheetRow  = matchIdx + 1;
    const prevStock = parseFloat(data[matchIdx][9]) || 0;
    const newStock  = prevStock + receivedQty;
    const inventoryId = data[matchIdx][0];

    invSheet.getRange(sheetRow, 10).setValue(newStock);        // Current_Stock
    invSheet.getRange(sheetRow, 13).setValue(new Date());      // Last_Updated
    invSheet.getRange(sheetRow, 14).setValue(grnNumber);       // Last_Transaction

    return success({ inventoryId, previousStock: prevStock, updatedStock: newStock });
  }
}

// ── RECEIPTS QUERIES ────────────────────────────────────────
function getReceipts(filters) {
  const sheet   = getSheet(SHEETS.FABRIC_RECEIPTS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));

  if (filters) {
    if (filters.vendor)
      rows = rows.filter(r => r.Vendor_Name.toLowerCase().includes(filters.vendor.toLowerCase()));
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

function generateBatchNumber() {
  const d = new Date();
  return `BATCH-${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}-${Math.random().toString(36).substring(2,6).toUpperCase()}`;
}
