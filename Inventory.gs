// ============================================================
// File: Inventory.gs — Inventory Master, Stock Issue & Returns
// ============================================================

// ── INVENTORY MASTER ─────────────────────────────────────────
function getInventory(filters) {
  const sheet   = getSheet(SHEETS.INVENTORY);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));

  if (filters) {
    if (filters.articleNumber)
      rows = rows.filter(r => String(r.Article_Number).toLowerCase().includes(filters.articleNumber.toLowerCase()));
    if (filters.color)
      rows = rows.filter(r => r.Color.toLowerCase().includes(filters.color.toLowerCase()));
    if (filters.vendor)
      rows = rows.filter(r => r.Vendor_Name.toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.lowStock)
      rows = rows.filter(r => parseFloat(r.Current_Stock) <= parseFloat(r.Reorder_Level));
    if (filters.status)
      rows = rows.filter(r => r.Status === filters.status);
  }

  // Add computed fields
  rows = rows.map(r => ({
    ...r,
    Is_Low_Stock: parseFloat(r.Current_Stock) <= parseFloat(r.Reorder_Level),
    Stock_Value: parseFloat(r.Current_Stock)
  }));

  return success(rows);
}

function getInventoryById(id) {
  const sheet  = getSheet(SHEETS.INVENTORY);
  const data   = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data.slice(1).find(r => r[0] === id);
  if (!row) return error('Inventory item not found: ' + id);
  return success(rowToObj(headers, row));
}

function addInventoryItem(data) {
  if (!data.fabricName)      return error('Fabric Name is required');
  if (!data.color)           return error('Color is required');
  if (!data.vendorName)      return error('Vendor Name is required');

  const invSheet = getSheet(SHEETS.INVENTORY);
  // Duplicate check
  const existing = invSheet.getDataRange().getValues().slice(1);
  const dup = existing.find(r =>
    r[1] === (data.articleNumber || '') &&
    r[3] === data.color &&
    r[4] === (data.width || '') &&
    r[6] === data.vendorName
  );
  if (dup) return error('Inventory record already exists for this Article + Color + Width + Vendor combination');

  const inventoryId = generateSequentialId('INV', invSheet);
  const row = [
    inventoryId,
    data.articleNumber || '',
    data.fabricName,
    data.color,
    data.width || '',
    data.widthUnit || 'Inch',
    data.vendorName,
    data.unit || 'Meter',
    parseFloat(data.openingStock) || 0,
    parseFloat(data.openingStock) || 0,
    parseFloat(data.reorderLevel) || 50,
    data.storageLocation || '',
    new Date(),
    'MANUAL',
    'Active'
  ];
  invSheet.appendRow(row);
  logActivity('CREATE', 'Inventory', inventoryId, data);
  return success({ inventoryId }, `Inventory item ${inventoryId} created`);
}

function getLowStockAlerts() {
  const sheet  = getSheet(SHEETS.INVENTORY);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  const lowStock = data.slice(1)
    .map(r => rowToObj(headers, r))
    .filter(r => parseFloat(r.Current_Stock) <= parseFloat(r.Reorder_Level) && r.Status === 'Active')
    .map(r => ({
      ...r,
      Shortage: parseFloat(r.Reorder_Level) - parseFloat(r.Current_Stock)
    }))
    .sort((a, b) => a.Current_Stock - b.Current_Stock);
  return success(lowStock);
}

// ── STOCK ISSUE ──────────────────────────────────────────────
function issueStock(data) {
  if (!data.inventoryId)  return error('Inventory ID is required');
  if (!data.issueQty || isNaN(data.issueQty) || parseFloat(data.issueQty) <= 0)
    return error('Valid Issue Quantity is required');
  if (!data.department)   return error('Department is required');

  const invSheet = getSheet(SHEETS.INVENTORY);
  const invData  = invSheet.getDataRange().getValues();
  const headers  = invData[0];
  const rowIdx   = invData.findIndex((r, i) => i > 0 && r[0] === data.inventoryId);
  if (rowIdx === -1) return error('Inventory item not found: ' + data.inventoryId);

  const inv      = rowToObj(headers, invData[rowIdx]);
  const prevStock = parseFloat(inv.Current_Stock) || 0;
  const issueQty  = parseFloat(data.issueQty);

  // ── PREVENT NEGATIVE STOCK ──────────────────────────────────
  if (issueQty > prevStock)
    return error(`Insufficient stock. Available: ${prevStock} ${inv.Unit}, Requested: ${issueQty} ${inv.Unit}`);

  const newStock = prevStock - issueQty;

  // Save issue record
  const issueSheet = getSheet(SHEETS.STOCK_ISSUE);
  const issueId = generateSequentialId('ISS', issueSheet);
  issueSheet.appendRow([
    issueId,
    data.inventoryId,
    inv.Article_Number,
    inv.Fabric_Name,
    inv.Color,
    inv.Width,
    inv.Width_Unit,
    issueQty,
    inv.Unit,
    data.productionOrder || '',
    data.department,
    prevStock,
    newStock,
    Session.getActiveUser().getEmail() || 'User',
    new Date(),
    data.notes || '',
    data.batchNumber || ''
  ]);

  // Update inventory
  const sheetRow = rowIdx + 1;
  invSheet.getRange(sheetRow, 10).setValue(newStock);
  invSheet.getRange(sheetRow, 13).setValue(new Date());
  invSheet.getRange(sheetRow, 14).setValue(issueId);

  // Low stock check
  const reorderLevel = parseFloat(inv.Reorder_Level) || 0;
  const lowStockAlert = newStock <= reorderLevel;

  logActivity('ISSUE', 'StockIssue', issueId, { inventoryId: data.inventoryId, qty: issueQty });

  return success({
    issueId,
    previousStock: prevStock,
    issuedQty: issueQty,
    currentStock: newStock,
    lowStockAlert,
    reorderLevel
  }, `Stock issued: ${issueQty} ${inv.Unit}. New stock: ${newStock} ${inv.Unit}${lowStockAlert ? ' ⚠️ Low stock alert!' : ''}`);
}

function getIssues(filters) {
  const sheet  = getSheet(SHEETS.STOCK_ISSUE);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));

  if (filters) {
    if (filters.department)
      rows = rows.filter(r => r.Department === filters.department);
    if (filters.dateFrom)
      rows = rows.filter(r => new Date(r.Issued_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => new Date(r.Issued_At) <= new Date(filters.dateTo));
  }
  rows.sort((a, b) => new Date(b.Issued_At) - new Date(a.Issued_At));
  return success(rows);
}

function getIssuesByInventory(inventoryId) {
  const sheet  = getSheet(SHEETS.STOCK_ISSUE);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  const rows = data.slice(1)
    .filter(r => r[1] === inventoryId)
    .map(r => rowToObj(headers, r));
  return success(rows);
}

// ── STOCK RETURNS ────────────────────────────────────────────
function returnStock(data) {
  if (!data.issueId)      return error('Issue ID is required');
  if (!data.returnQty || isNaN(data.returnQty) || parseFloat(data.returnQty) <= 0)
    return error('Valid Return Quantity is required');

  // Get original issue
  const issueSheet = getSheet(SHEETS.STOCK_ISSUE);
  const issueData  = issueSheet.getDataRange().getValues();
  const issueHeaders = issueData[0];
  const issueRow   = issueData.slice(1).find(r => r[0] === data.issueId);
  if (!issueRow) return error('Issue record not found: ' + data.issueId);
  const issue = rowToObj(issueHeaders, issueRow);

  const returnQty = parseFloat(data.returnQty);
  const issuedQty = parseFloat(issue.Issue_Qty);

  // Get total already returned for this issue
  const returnSheet = getSheet(SHEETS.STOCK_RETURNS);
  const returnData  = returnSheet.getDataRange().getValues();
  const alreadyReturned = returnData.slice(1)
    .filter(r => r[1] === data.issueId)
    .reduce((s, r) => s + (parseFloat(r[8]) || 0), 0);

  const maxReturn = issuedQty - alreadyReturned;
  if (returnQty > maxReturn + 0.001)
    return error(`Cannot return ${returnQty}. Max returnable quantity is ${maxReturn}`);

  // Get current inventory stock
  const invSheet = getSheet(SHEETS.INVENTORY);
  const invData  = invSheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRowIdx = invData.findIndex((r, i) => i > 0 && r[0] === issue.Inventory_ID);
  if (invRowIdx === -1) return error('Inventory item not found');

  const prevStock = parseFloat(invData[invRowIdx][9]) || 0;
  const newStock  = prevStock + returnQty;

  // Save return record
  const returnId = generateSequentialId('RET', returnSheet);
  returnSheet.appendRow([
    returnId,
    data.issueId,
    issue.Inventory_ID,
    issue.Article_Number,
    issue.Fabric_Name,
    issue.Color,
    issue.Width,
    issue.Width_Unit,
    returnQty,
    issue.Unit,
    prevStock,
    newStock,
    data.reason || '',
    Session.getActiveUser().getEmail() || 'User',
    new Date(),
    data.notes || ''
  ]);

  // Update inventory
  const sheetRow = invRowIdx + 1;
  invSheet.getRange(sheetRow, 10).setValue(newStock);
  invSheet.getRange(sheetRow, 13).setValue(new Date());
  invSheet.getRange(sheetRow, 14).setValue(returnId);

  logActivity('RETURN', 'StockReturn', returnId, { issueId: data.issueId, qty: returnQty });

  return success({
    returnId,
    previousStock: prevStock,
    returnedQty: returnQty,
    currentStock: newStock
  }, `Stock returned: ${returnQty} ${issue.Unit}. New stock: ${newStock} ${issue.Unit}`);
}

function getReturns(filters) {
  const sheet   = getSheet(SHEETS.STOCK_RETURNS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));
  if (filters && filters.dateFrom)
    rows = rows.filter(r => new Date(r.Returned_At) >= new Date(filters.dateFrom));
  if (filters && filters.dateTo)
    rows = rows.filter(r => new Date(r.Returned_At) <= new Date(filters.dateTo));
  rows.sort((a, b) => new Date(b.Returned_At) - new Date(a.Returned_At));
  return success(rows);
}

// ── VENDORS ──────────────────────────────────────────────────
function getVendors() {
  const sheet  = getSheet(SHEETS.VENDORS);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  return success(data.slice(1).map(r => rowToObj(headers, r)));
}

function addVendor(data) {
  if (!data.vendorName) return error('Vendor Name is required');
  const sheet = getSheet(SHEETS.VENDORS);
  const vendorId = generateSequentialId('VEN', sheet);
  sheet.appendRow([
    vendorId, data.vendorName, data.contactPerson || '',
    data.phone || '', data.email || '', data.address || '',
    data.gstNumber || '', data.paymentTerms || '30 days',
    'Active', new Date()
  ]);
  return success({ vendorId }, 'Vendor added');
}

// ── PRODUCTION ORDERS ────────────────────────────────────────
function getProductionOrders() {
  const sheet  = getSheet(SHEETS.PRODUCTION_ORDERS);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  return success(data.slice(1).map(r => rowToObj(headers, r)));
}

function addProductionOrder(data) {
  if (!data.poNumber || !data.productName) return error('PO Number and Product Name are required');
  const sheet = getSheet(SHEETS.PRODUCTION_ORDERS);
  const poId  = generateSequentialId('PROD', sheet);
  sheet.appendRow([
    poId, data.poNumber, data.productName, data.styleNumber || '',
    parseFloat(data.quantity) || 0, data.startDate || '',
    data.endDate || '', data.department || '',
    'Active', new Date(), data.notes || ''
  ]);
  return success({ poId }, 'Production Order added');
}
