// ============================================================
// File: PurchaseOrders.gs — Purchase Order Module
// ============================================================
// function test_createPO() {
//   const dummyData = {
//     vendorName: "Test Vendor",
//     fabricName: "Cotton",
//     color: "Blue",
//     orderedQty: 100,
//     deliveryDate: "2024-01-01"
//   };
//   // Call the function with the dummy data
//   createPurchaseOrder(dummyData);
// }
// function createPurchaseOrder(data) {
//   const sheet = getSheet(SHEETS.PURCHASE_ORDERS);
//   const poNumber = generateSequentialId('PO', sheet);

//   // Validate required fields
//   if (!data.vendorName) return error('Vendor Name is required');
//   if (!data.fabricName) return error('Fabric Name is required');
//   if (!data.color)      return error('Color is required');
//   if (!data.orderedQty || isNaN(data.orderedQty) || data.orderedQty <= 0)
//     return error('Valid Ordered Quantity is required');
//   if (!data.deliveryDate) return error('Delivery Date is required');

// const row = [
//     poNumber,
//     data.vendorName || '',
//     data.articleNumber || '',
//     data.fabricName || '',
//     data.color || '',
//     data.width || '',
//     data.widthUnit || 'Inch',
//     parseFloat(data.orderedQty) || 0,
//     data.unit || 'Meter',
//     parseFloat(data.rate) || 0,        // Maps to 'Rate'
//     data.uom || '',                    // Maps to 'UOM'
//     data.deliveryDate,
//     'Pending',
//     Session.getActiveUser().getEmail() || 'User',
//     new Date(),
//     data.notes || ''
//   ];

//   sheet.appendRow(row);
//   sheet.appendRow(row);
//   logActivity('CREATE', 'PurchaseOrder', poNumber, data);

//   // Auto-format the new row
//   const lastRow = sheet.getLastRow();
//   sheet.getRange(lastRow, 11).setBackground('#FFF3CD'); // Status cell - yellow for Pending

//   return success({ poNumber }, `Purchase Order ${poNumber} created successfully`);
// }
function createPurchaseOrder(data) {
  const sheet = getSheet(SHEETS.PURCHASE_ORDERS);
  const poNumber = generateSequentialId('PO', sheet);

  // Validate required fields
  if (!data.vendorName) return error('Vendor Name is required');
  if (!data.fabricName) return error('Fabric Name is required');
  if (!data.color)      return error('Color is required');
  if (!data.orderedQty || isNaN(data.orderedQty) || data.orderedQty <= 0)
    return error('Valid Ordered Quantity is required');
  if (!data.deliveryDate) return error('Delivery Date is required');

  const row = [
    poNumber,
    data.vendorName || '',
    data.articleNumber || '',
    data.fabricName || '',
    data.color || '',
    data.width || '',
    data.widthUnit || 'Inch',
    parseFloat(data.orderedQty) || 0,
    data.unit || 'Meter',               // Receives selected value from frontend dropdown
    parseFloat(data.rate) || 0,        // Maps to 'Rate'
    data.uom || '',                    // Maps to 'UOM'
    data.deliveryDate,
    'Pending',
    Session.getActiveUser().getEmail() || 'User',
    new Date(),
    data.notes || ''
  ];

  // Append data only once
  sheet.appendRow(row);
  logActivity('CREATE', 'PurchaseOrder', poNumber, data);

  // Auto-format the correct cell (Status is Column 13)
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 13).setBackground('#FFF3CD'); // Fixed: Changed from 11 to 13

  return success({ poNumber }, `Purchase Order ${poNumber} created successfully`);
}
function getPurchaseOrders(filters) {
  const sheet = getSheet(SHEETS.PURCHASE_ORDERS);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);

  const headers = data[0];
  let rows = data.slice(1).map(row => rowToObj(headers, row));

  // Apply filters
  if (filters) {
    if (filters.status && filters.status !== 'All')
      rows = rows.filter(r => r.Status === filters.status);
    if (filters.vendor)
      rows = rows.filter(r => r.Vendor_Name.toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.articleNumber)
      rows = rows.filter(r => r.Article_Number.toLowerCase().includes(filters.articleNumber.toLowerCase()));
    if (filters.dateFrom)
      rows = rows.filter(r => new Date(r.Created_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => new Date(r.Created_At) <= new Date(filters.dateTo));
  }

  // Sort newest first
  rows.sort((a, b) => new Date(b.Created_At) - new Date(a.Created_At));

  // Attach receipt summary
  rows = rows.map(po => {
    const receipts = getReceiptsForPO(po.PO_Number);
    const totalReceived = receipts.reduce((s, r) => s + (parseFloat(r.Received_Qty) || 0), 0);
    return {
      ...po,
      Total_Received: totalReceived,
      Balance_Qty: (parseFloat(po.Ordered_Qty) || 0) - totalReceived,
      Receipt_Count: receipts.length
    };
  });

  return success(rows);
}

function getPOById(poNumber) {
  const sheet = getSheet(SHEETS.PURCHASE_ORDERS);
  const data   = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data.slice(1).find(r => r[0] === poNumber);
  if (!row) return error('Purchase Order not found: ' + poNumber);
  const po = rowToObj(headers, row);
  const receipts = getReceiptsForPO(poNumber);
  const totalReceived = receipts.reduce((s, r) => s + (parseFloat(r.Received_Qty) || 0), 0);
  return success({
    ...po,
    receipts,
    Total_Received: totalReceived,
    Balance_Qty: (parseFloat(po.Ordered_Qty) || 0) - totalReceived
  });
}

function updatePOStatus(data) {
  if (!data.poNumber || !data.status)
    return error('PO Number and Status are required');

  const sheet = getSheet(SHEETS.PURCHASE_ORDERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex(r => r[0] === data.poNumber);
  if (rowIdx === -1) return error('Purchase Order not found');

  sheet.getRange(rowIdx + 1, 11).setValue(data.status);

  const colors = { Pending:'#FFF3CD', Partial:'#676E00', Completed:'#00821B', Cancelled:'#F53144' };
  sheet.getRange(rowIdx + 1, 11).setBackground(colors[data.status] || '#FFFFFF');

  logActivity('UPDATE_STATUS', 'PurchaseOrder', data.poNumber, { status: data.status });
  return success({ poNumber: data.poNumber, status: data.status });
}

// Helper: used internally
function getReceiptsForPO(poNumber) {
  const sheet = getSheet(SHEETS.FABRIC_RECEIPTS);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(r => r[1] === poNumber)
    .map(r => rowToObj(headers, r));
}

// Helper: convert row array → object using headers
function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}
