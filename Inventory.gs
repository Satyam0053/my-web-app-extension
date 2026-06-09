// ============================================================
// File: Inventory.gs — Inventory Master, Stock Issue & Returns
// Version: 2.3 | Audit-Fixed
//
// FIXES:
//   • issueStock: function signature was (data) — dispatcher calls
//     issueStock(payload, caller). Added caller parameter and used
//     it for audit logging and Received_By field.
//   • issueStock: Session.getActiveUser().getEmail() wrapped in
//     try/catch with fallback to caller.username (same issue as
//     FabricReceipts — can throw in web app context).
//   • returnStock: same Session.getActiveUser() fix.
//   • addInventoryItem: same Session issues; also added caller param.
//   • addVendor: same; added caller param.
//   • addProductionOrder: same; added caller param.
//   • getInventory filters: null-safe String() wrapping on all
//     filter fields to avoid crash when a cell is empty/null.
//   • getInventory computed fields: parseFloat on null returns NaN;
//     Is_Low_Stock comparison now handles NaN gracefully.
//   • returnStock: alreadyReturned was reading column index 8
//     (0-based) for Return_Qty which is correct per the header.
//     Confirmed correct — no change needed there.
//   • addInventoryItem: caller arg was ignored — now used.
//   • addVendor / addProductionOrder: caller arg was ignored — now used.
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
      rows = rows.filter(r =>
        String(r.Article_Number || '').toLowerCase().includes(filters.articleNumber.toLowerCase()));
    if (filters.color)
      rows = rows.filter(r =>
        String(r.Color || '').toLowerCase().includes(filters.color.toLowerCase()));
    if (filters.vendor)
      rows = rows.filter(r =>
        String(r.Vendor_Name || '').toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.lowStock)
      rows = rows.filter(r => {
        const stock   = parseFloat(r.Current_Stock)  || 0;
        const reorder = parseFloat(r.Reorder_Level) || 0;
        return stock <= reorder;
      });
    if (filters.status)
      rows = rows.filter(r => r.Status === filters.status);
  }

  // Add computed fields — handle NaN from empty cells
  rows = rows.map(r => {
    const stock   = parseFloat(r.Current_Stock)  || 0;
    const reorder = parseFloat(r.Reorder_Level) || 0;
    return {
      ...r,
      Is_Low_Stock: stock <= reorder,
      Stock_Value : stock
    };
  });

  return success(rows);
}

function getInventoryById(id) {
  if (!id) return error('Inventory ID is required');
  const sheet   = getSheet(SHEETS.INVENTORY);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const row     = data.slice(1).find(r => r[0] === id);
  if (!row) return error('Inventory item not found: ' + id);
  return success(rowToObj(headers, row));
}

function addInventoryItem(data, caller) {
  if (!data.fabricName)  return error('Fabric Name is required');
  if (!data.color)       return error('Color is required');
  if (!data.vendorName)  return error('Vendor Name is required');

  const invSheet = getSheet(SHEETS.INVENTORY);

  // Duplicate check
  const existing = invSheet.getDataRange().getValues().slice(1);
  const dup = existing.find(r =>
    String(r[1] || '').trim() === String(data.articleNumber || '').trim() &&
    String(r[3] || '').trim() === String(data.color         || '').trim() &&
    String(r[4] || '').trim() === String(data.width         || '').trim() &&
    String(r[6] || '').trim() === String(data.vendorName    || '').trim()
  );
  if (dup) return error('Inventory record already exists for this Article + Color + Width + Vendor combination');

  const inventoryId   = generateSequentialId('INV', invSheet);
  const openingStock  = parseFloat(data.openingStock) || 0;
  const row = [
    inventoryId,
    data.articleNumber     || '',
    data.fabricName,
    data.color,
    data.width             || '',
    data.widthUnit         || 'Inch',
    data.vendorName,
    data.unit              || 'Meter',
    openingStock,                         // Opening_Stock
    openingStock,                         // Current_Stock (= opening on creation)
    parseFloat(data.reorderLevel) || 50,
    data.storageLocation   || '',
    new Date(),
    'MANUAL',
    'Active'
  ];
  invSheet.appendRow(row);

  try {
    logActivity('CREATE', 'Inventory', inventoryId, {
      article: data.articleNumber, fabric: data.fabricName, vendor: data.vendorName
    }, caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({ inventoryId }, `Inventory item ${inventoryId} created`);
}

function getLowStockAlerts() {
  const sheet   = getSheet(SHEETS.INVENTORY);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  const lowStock = data.slice(1)
    .map(r => rowToObj(headers, r))
    .filter(r => {
      const stock   = parseFloat(r.Current_Stock)  || 0;
      const reorder = parseFloat(r.Reorder_Level) || 0;
      return stock <= reorder && r.Status === 'Active';
    })
    .map(r => ({
      ...r,
      Shortage: (parseFloat(r.Reorder_Level) || 0) - (parseFloat(r.Current_Stock) || 0)
    }))
    .sort((a, b) => (parseFloat(a.Current_Stock) || 0) - (parseFloat(b.Current_Stock) || 0));

  return success(lowStock);
}

// ── STOCK ISSUE ──────────────────────────────────────────────
function issueStock(data, caller) {
  if (!data.inventoryId) return error('Inventory ID is required');
  if (!data.issueQty || isNaN(data.issueQty) || parseFloat(data.issueQty) <= 0)
    return error('Valid Issue Quantity is required');
  if (!data.department) return error('Department is required');

  const invSheet = getSheet(SHEETS.INVENTORY);
  const invData  = invSheet.getDataRange().getValues();
  const headers  = invData[0];
  const rowIdx   = invData.findIndex((r, i) => i > 0 && r[0] === data.inventoryId);
  if (rowIdx === -1) return error('Inventory item not found: ' + data.inventoryId);

  const inv       = rowToObj(headers, invData[rowIdx]);
  const prevStock = parseFloat(inv.Current_Stock) || 0;
  const issueQty  = parseFloat(data.issueQty);

  if (issueQty > prevStock)
    return error(`Insufficient stock. Available: ${prevStock} ${inv.Unit}, Requested: ${issueQty} ${inv.Unit}`);

  const newStock = prevStock - issueQty;

  // ── Safe "issued by" ──────────────────────────────────────
  let issuedBy = '';
  try { issuedBy = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  if (!issuedBy && caller && caller.username) issuedBy = caller.username;
  if (!issuedBy) issuedBy = 'User';

  // Save issue record
  const issueSheet = getSheet(SHEETS.STOCK_ISSUE);
  const issueId    = generateSequentialId('ISS', issueSheet);
  issueSheet.appendRow([
    issueId,
    data.inventoryId,
    inv.Article_Number || '',
    inv.Fabric_Name    || '',
    inv.Color          || '',
    inv.Width          || '',
    inv.Width_Unit     || '',
    issueQty,
    inv.Unit           || 'Meter',
    data.productionOrder || '',
    data.department,
    prevStock,
    newStock,
    issuedBy,
    new Date(),
    data.notes       || '',
    data.batchNumber || ''
  ]);

  // Update inventory
  const sheetRow = rowIdx + 1;
  invSheet.getRange(sheetRow, 10).setValue(newStock);   // Current_Stock
  invSheet.getRange(sheetRow, 13).setValue(new Date()); // Last_Updated
  invSheet.getRange(sheetRow, 14).setValue(issueId);    // Last_Transaction

  const reorderLevel  = parseFloat(inv.Reorder_Level) || 0;
  const lowStockAlert = newStock <= reorderLevel;

  try {
    logActivity('ISSUE', 'StockIssue', issueId,
      { inventoryId: data.inventoryId, qty: issueQty, dept: data.department },
      caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({
    issueId,
    previousStock: prevStock,
    issuedQty    : issueQty,
    currentStock : newStock,
    lowStockAlert,
    reorderLevel
  }, `Stock issued: ${issueQty} ${inv.Unit}. New stock: ${newStock} ${inv.Unit}${lowStockAlert ? ' ⚠️ Low stock alert!' : ''}`);
}

function getIssues(filters) {
  const sheet   = getSheet(SHEETS.STOCK_ISSUE);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));

  if (filters) {
    if (filters.department)
      rows = rows.filter(r => r.Department === filters.department);
    if (filters.dateFrom)
      rows = rows.filter(r => r.Issued_At && new Date(r.Issued_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => r.Issued_At && new Date(r.Issued_At) <= new Date(filters.dateTo));
  }

  rows.sort((a, b) => {
    const da = a.Issued_At ? new Date(a.Issued_At) : 0;
    const db = b.Issued_At ? new Date(b.Issued_At) : 0;
    return db - da;
  });

  return success(rows);
}

function getIssuesByInventory(inventoryId) {
  if (!inventoryId) return success([]);
  const sheet   = getSheet(SHEETS.STOCK_ISSUE);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  const rows    = data.slice(1)
    .filter(r => r[1] === inventoryId)
    .map(r => rowToObj(headers, r));
  return success(rows);
}

// ── STOCK RETURNS ────────────────────────────────────────────
function returnStock(data, caller) {
  if (!data.issueId)  return error('Issue ID is required');
  if (!data.returnQty || isNaN(data.returnQty) || parseFloat(data.returnQty) <= 0)
    return error('Valid Return Quantity is required');

  // Get original issue
  const issueSheet   = getSheet(SHEETS.STOCK_ISSUE);
  const issueData    = issueSheet.getDataRange().getValues();
  if (issueData.length <= 1) return error('Issue record not found: ' + data.issueId);
  const issueHeaders = issueData[0];
  const issueRow     = issueData.slice(1).find(r => r[0] === data.issueId);
  if (!issueRow) return error('Issue record not found: ' + data.issueId);
  const issue = rowToObj(issueHeaders, issueRow);

  const returnQty = parseFloat(data.returnQty);
  const issuedQty = parseFloat(issue.Issue_Qty) || 0;

  // Total already returned for this issue (Return_Qty is col index 8)
  const returnSheet    = getSheet(SHEETS.STOCK_RETURNS);
  const returnData     = returnSheet.getDataRange().getValues();
  const alreadyReturned = returnData.length > 1
    ? returnData.slice(1)
        .filter(r => r[1] === data.issueId)
        .reduce((s, r) => s + (parseFloat(r[8]) || 0), 0)
    : 0;

  const maxReturn = issuedQty - alreadyReturned;
  if (returnQty > maxReturn + 0.001)
    return error(`Cannot return ${returnQty}. Max returnable quantity is ${maxReturn.toFixed(3)}`);

  // Get current inventory stock
  const invSheet  = getSheet(SHEETS.INVENTORY);
  const invData   = invSheet.getDataRange().getValues();
  const invRowIdx = invData.findIndex((r, i) => i > 0 && r[0] === issue.Inventory_ID);
  if (invRowIdx === -1) return error('Inventory item not found for Inventory_ID: ' + issue.Inventory_ID);

  const prevStock = parseFloat(invData[invRowIdx][9]) || 0;
  const newStock  = prevStock + returnQty;

  // Safe "returned by"
  let returnedBy = '';
  try { returnedBy = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  if (!returnedBy && caller && caller.username) returnedBy = caller.username;
  if (!returnedBy) returnedBy = 'User';

  // Save return record
  const returnId = generateSequentialId('RET', returnSheet);
  returnSheet.appendRow([
    returnId,
    data.issueId,
    issue.Inventory_ID    || '',
    issue.Article_Number  || '',
    issue.Fabric_Name     || '',
    issue.Color           || '',
    issue.Width           || '',
    issue.Width_Unit      || '',
    returnQty,
    issue.Unit            || 'Meter',
    prevStock,
    newStock,
    data.reason           || '',
    returnedBy,
    new Date(),
    data.notes            || ''
  ]);

  // Update inventory
  const sheetRow = invRowIdx + 1;
  invSheet.getRange(sheetRow, 10).setValue(newStock);   // Current_Stock
  invSheet.getRange(sheetRow, 13).setValue(new Date()); // Last_Updated
  invSheet.getRange(sheetRow, 14).setValue(returnId);   // Last_Transaction

  try {
    logActivity('RETURN', 'StockReturn', returnId,
      { issueId: data.issueId, qty: returnQty },
      caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({
    returnId,
    previousStock: prevStock,
    returnedQty  : returnQty,
    currentStock : newStock
  }, `Stock returned: ${returnQty} ${issue.Unit}. New stock: ${newStock} ${issue.Unit}`);
}

function getReturns(filters) {
  const sheet   = getSheet(SHEETS.STOCK_RETURNS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1).map(r => rowToObj(headers, r));

  if (filters) {
    if (filters.dateFrom)
      rows = rows.filter(r => r.Returned_At && new Date(r.Returned_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => r.Returned_At && new Date(r.Returned_At) <= new Date(filters.dateTo));
  }

  rows.sort((a, b) => {
    const da = a.Returned_At ? new Date(a.Returned_At) : 0;
    const db = b.Returned_At ? new Date(b.Returned_At) : 0;
    return db - da;
  });

  return success(rows);
}

// ── VENDORS ──────────────────────────────────────────────────
function getVendors() {
  const sheet   = getSheet(SHEETS.VENDORS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  return success(data.slice(1).map(r => rowToObj(headers, r)));
}

function addVendor(data, caller) {
  if (!data.vendorName) return error('Vendor Name is required');

  const sheet    = getSheet(SHEETS.VENDORS);
  const vendorId = generateSequentialId('VEN', sheet);
  sheet.appendRow([
    vendorId,
    data.vendorName,
    data.contactPerson  || '',
    data.phone          || '',
    data.email          || '',
    data.address        || '',
    data.gstNumber      || '',
    data.paymentTerms   || '30 days',
    'Active',
    new Date()
  ]);

  try {
    logActivity('CREATE', 'Vendor', vendorId,
      { vendor: data.vendorName },
      caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({ vendorId }, 'Vendor added: ' + data.vendorName);
}

// ── PRODUCTION ORDERS ────────────────────────────────────────
function getProductionOrders() {
  const sheet   = getSheet(SHEETS.PRODUCTION_ORDERS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  return success(data.slice(1).map(r => rowToObj(headers, r)));
}

function addProductionOrder(data, caller) {
  if (!data.poNumber || !data.productName)
    return error('PO Number and Product Name are required');

  const sheet = getSheet(SHEETS.PRODUCTION_ORDERS);
  const poId  = generateSequentialId('PROD', sheet);
  sheet.appendRow([
    poId,
    data.poNumber,
    data.productName,
    data.styleNumber  || '',
    parseFloat(data.quantity) || 0,
    data.startDate    || '',
    data.endDate      || '',
    data.department   || '',
    'Active',
    new Date(),
    data.notes        || ''
  ]);

  try {
    logActivity('CREATE', 'ProductionOrder', poId,
      { po: data.poNumber, product: data.productName },
      caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({ poId }, 'Production Order added: ' + poId);
}