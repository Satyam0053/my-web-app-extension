// ============================================================
// File: PurchaseOrders.gs — Purchase Order Module
// FIXED:
//  • createPurchaseOrder: sheet.appendRow called only ONCE (was doubled in old commented code)
//  • updatePOStatus: Status is column 13, not column 11 — fixed consistently
//  • updatePOStatus: caller param forwarded to logActivity
//  • getPurchaseOrders: Balance_Qty computed from real receipts (correct)
//  • createPurchaseOrder: caller param accepted and forwarded
// ============================================================

function createPurchaseOrder(data, caller) {
  const sheet    = getSheet(SHEETS.PURCHASE_ORDERS);
  const poNumber = generateSequentialId('PO', sheet);

  if (!data.vendorName)  return error('Vendor Name is required');
  if (!data.fabricName)  return error('Fabric Name is required');
  if (!data.color)       return error('Color is required');
  if (!data.orderedQty || isNaN(data.orderedQty) || data.orderedQty <= 0)
    return error('Valid Ordered Quantity is required');
  if (!data.deliveryDate) return error('Delivery Date is required');

  const row = [
    poNumber,
    data.vendorName    || '',
    data.articleNumber || '',
    data.fabricName    || '',
    data.color         || '',
    data.width         || '',
    data.widthUnit     || 'Inch',
    parseFloat(data.orderedQty) || 0,
    data.unit          || 'Meter',
    parseFloat(data.rate) || 0,
    data.uom           || '',
    data.deliveryDate,
    'Pending',
    (caller && caller.username) ? caller.username : (Session.getActiveUser().getEmail() || 'User'),
    new Date(),
    data.notes || ''
  ];

  // FIX: appendRow called exactly ONCE
  sheet.appendRow(row);

  try { logActivity('CREATE', 'PurchaseOrder', poNumber, data, (caller||{}).username, (caller||{}).role); } catch(e) {}

  // FIX: Status is column 13
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 13).setBackground('#FFF3CD');

  return success({ poNumber }, `Purchase Order ${poNumber} created successfully`);
}

function getPurchaseOrders(filters) {
  const sheet = getSheet(SHEETS.PURCHASE_ORDERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(row => rowToObj(headers, row));

  if (filters) {
    if (filters.status && filters.status !== 'All')
      rows = rows.filter(r => r.Status === filters.status);
    if (filters.vendor)
      rows = rows.filter(r => String(r.Vendor_Name || '').toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.articleNumber)
      rows = rows.filter(r => String(r.Article_Number || '').toLowerCase().includes(filters.articleNumber.toLowerCase()));
    if (filters.dateFrom)
      rows = rows.filter(r => new Date(r.Created_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => new Date(r.Created_At) <= new Date(filters.dateTo));
  }

  rows.sort((a, b) => new Date(b.Created_At) - new Date(a.Created_At));

  rows = rows.map(po => {
    const receipts      = getReceiptsForPO(po.PO_Number);
    const totalReceived = receipts.reduce((s, r) => s + (parseFloat(r.Received_Qty) || 0), 0);
    return {
      ...po,
      Total_Received: Math.round(totalReceived * 1000) / 1000,
      Balance_Qty   : Math.max(0, Math.round(((parseFloat(po.Ordered_Qty) || 0) - totalReceived) * 1000) / 1000),
      Receipt_Count : receipts.length
    };
  });

  return success(rows);
}

function getPOById(poNumber) {
  const sheet   = getSheet(SHEETS.PURCHASE_ORDERS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const row     = data.slice(1).find(r => r[0] === poNumber);
  if (!row) return error('Purchase Order not found: ' + poNumber);

  const po            = rowToObj(headers, row);
  const receipts      = getReceiptsForPO(poNumber);
  const totalReceived = receipts.reduce((s, r) => s + (parseFloat(r.Received_Qty) || 0), 0);

  return success({
    ...po,
    receipts,
    Total_Received: Math.round(totalReceived * 1000) / 1000,
    Balance_Qty   : Math.max(0, Math.round(((parseFloat(po.Ordered_Qty) || 0) - totalReceived) * 1000) / 1000)
  });
}

// FIX: Status is column 13 everywhere
function updatePOStatus(data, caller) {
  if (!data.poNumber || !data.status) return error('PO Number and Status are required');

  const sheet  = getSheet(SHEETS.PURCHASE_ORDERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex(r => r[0] === data.poNumber);
  if (rowIdx === -1) return error('Purchase Order not found');

  // FIX: Status is column 13 (1-based)
  sheet.getRange(rowIdx + 1, 13).setValue(data.status);

  const colors = {
    Pending      : '#FFF3CD',
    Partial      : '#FFE082',
    Completed    : '#C8E6C9',
    Cancelled    : '#FFCDD2',
    'Over Received': '#E1BEE7'
  };
  sheet.getRange(rowIdx + 1, 13).setBackground(colors[data.status] || '#FFFFFF');

  try { logActivity('UPDATE_STATUS', 'PurchaseOrder', data.poNumber, { status: data.status }, (caller||{}).username, (caller||{}).role); } catch(e) {}
  return success({ poNumber: data.poNumber, status: data.status });
}

// ── HELPERS ───────────────────────────────────────────────────
function getReceiptsForPO(poNumber) {
  const sheet = getSheet(SHEETS.FABRIC_RECEIPTS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).filter(r => r[1] === poNumber).map(r => rowToObj(headers, r));
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}
