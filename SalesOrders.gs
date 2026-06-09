// ============================================================
// File: SalesOrders.gs
// Modules: Customer Master | Article Master | Sales Orders
// Version: 1.0 | Production Ready
//
// Sheet structure:
//   Customers       → CUST0001, CUST0002 …
//   Articles        → ART0001, ART0002 …  (colors stored as JSON)
//   Sales_Orders    → SO0001, SO0002 …    (line items stored as JSON)
// ============================================================

// ── SEQUENTIAL ID HELPERS ────────────────────────────────────
// Format: PREFIX + zero-padded 4-digit counter (CUST0001, ART0001, SO0001)
function _nextSeqId(prefix, sheet) {
  const data = sheet.getDataRange().getValues();
  // Count data rows (excluding header)
  const count = Math.max(data.length - 1, 0);
  return prefix + String(count + 1).padStart(4, '0');
}

// ── EMAIL VALIDATOR ──────────────────────────────────────────
function _isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// ============================================================
// SECTION 1 — CUSTOMER MASTER
// ============================================================

function getCustomers() {
  const sheet = getSheet(SHEETS.CUSTOMERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  return success(data.slice(1).map(r => rowToObj(headers, r)));
}

function getCustomerById(customerId) {
  const sheet   = getSheet(SHEETS.CUSTOMERS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const row     = data.slice(1).find(r => r[0] === customerId);
  if (!row) return error('Customer not found: ' + customerId);
  return success(rowToObj(headers, row));
}

function createCustomer(data, caller) {
  if (!data.customerName || !data.customerName.trim())
    return error('Customer Name is required');

  if (data.email && data.email.trim() && !_isValidEmail(data.email.trim()))
    return error('Invalid email format');

  const sheet = getSheet(SHEETS.CUSTOMERS);

  // Duplicate name check (case-insensitive)
  const existing = sheet.getDataRange().getValues().slice(1);
  const dup = existing.find(r =>
    String(r[1]).trim().toLowerCase() === data.customerName.trim().toLowerCase()
  );
  if (dup) return error('Customer "' + data.customerName.trim() + '" already exists');

  const customerId = _nextSeqId('CUST', sheet);

  sheet.appendRow([
    customerId,
    data.customerName.trim(),
    data.mobile   ? data.mobile.trim()   : '',
    data.email    ? data.email.trim()    : '',
    data.location ? data.location.trim() : '',
    new Date()
  ]);

  logActivity('CREATE', 'Customer', customerId,
    { name: data.customerName },
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ customerId }, 'Customer ' + customerId + ' created successfully');
}

function updateCustomer(data, caller) {
  if (!data.customerId) return error('Customer ID is required');
  if (!data.customerName || !data.customerName.trim())
    return error('Customer Name is required');
  if (data.email && data.email.trim() && !_isValidEmail(data.email.trim()))
    return error('Invalid email format');

  const sheet  = getSheet(SHEETS.CUSTOMERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex((r, i) => i > 0 && r[0] === data.customerId);
  if (rowIdx === -1) return error('Customer not found: ' + data.customerId);

  // Duplicate name check (excluding self)
  const dup = values.slice(1).find((r, i) =>
    i + 1 !== rowIdx &&
    String(r[1]).trim().toLowerCase() === data.customerName.trim().toLowerCase()
  );
  if (dup) return error('Another customer with this name already exists');

  const sheetRow = rowIdx + 1;
  sheet.getRange(sheetRow, 2).setValue(data.customerName.trim());
  sheet.getRange(sheetRow, 3).setValue(data.mobile   ? data.mobile.trim()   : '');
  sheet.getRange(sheetRow, 4).setValue(data.email    ? data.email.trim()    : '');
  sheet.getRange(sheetRow, 5).setValue(data.location ? data.location.trim() : '');

  logActivity('UPDATE', 'Customer', data.customerId,
    { name: data.customerName },
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ customerId: data.customerId }, 'Customer updated successfully');
}

function deleteCustomer(data, caller) {
  if (!data.customerId) return error('Customer ID is required');

  // Guard: do not delete if customer has SOs
  const soSheet = getSheet(SHEETS.SALES_ORDERS);
  const soData  = soSheet.getDataRange().getValues();
  const linked  = soData.slice(1).find(r => r[1] === data.customerId);
  if (linked) return error('Cannot delete — this customer has Sales Orders linked to them');

  const sheet  = getSheet(SHEETS.CUSTOMERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex((r, i) => i > 0 && r[0] === data.customerId);
  if (rowIdx === -1) return error('Customer not found: ' + data.customerId);

  sheet.deleteRow(rowIdx + 1);

  logActivity('DELETE', 'Customer', data.customerId, {},
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ customerId: data.customerId }, 'Customer deleted successfully');
}

// ============================================================
// SECTION 2 — ARTICLE MASTER
// ============================================================
// Colors are stored as a JSON string in column 5 (index 4).
// Schema: [ { color: "Red", qty: 100 }, … ]

function getArticles() {
  const sheet = getSheet(SHEETS.ARTICLES);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  return success(data.slice(1).map(r => {
    const obj = rowToObj(headers, r);
    // Parse colors JSON safely
    try { obj.Colors = JSON.parse(obj.Colors || '[]'); } catch(e) { obj.Colors = []; }
    return obj;
  }));
}

function getArticleById(articleId) {
  const sheet   = getSheet(SHEETS.ARTICLES);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const row     = data.slice(1).find(r => r[0] === articleId);
  if (!row) return error('Article not found: ' + articleId);
  const obj = rowToObj(headers, row);
  try { obj.Colors = JSON.parse(obj.Colors || '[]'); } catch(e) { obj.Colors = []; }
  return success(obj);
}

function _validateColors(colors) {
  if (!Array.isArray(colors) || colors.length === 0)
    return 'At least one color is required';
  for (let i = 0; i < colors.length; i++) {
    const c = colors[i];
    if (!c.color || !String(c.color).trim())
      return 'Color name cannot be blank (row ' + (i + 1) + ')';
    if (c.qty === undefined || c.qty === null || c.qty === '')
      return 'Quantity is required for color "' + c.color + '"';
    if (isNaN(parseFloat(c.qty)))
      return 'Quantity must be numeric for color "' + c.color + '"';
    if (parseFloat(c.qty) < 0)
      return 'Quantity cannot be negative for color "' + c.color + '"';
  }
  return null; // no error
}

function createArticle(data, caller) {
  if (!data.articleNumber || !data.articleNumber.trim())
    return error('Article Number is required');

  const colorErr = _validateColors(data.colors);
  if (colorErr) return error(colorErr);

  const sheet = getSheet(SHEETS.ARTICLES);

  // Duplicate article number check (case-insensitive)
  const existing = sheet.getDataRange().getValues().slice(1);
  const dup = existing.find(r =>
    String(r[2]).trim().toLowerCase() === data.articleNumber.trim().toLowerCase()
  );
  if (dup) return error('Article Number "' + data.articleNumber.trim() + '" already exists');

  const articleId = _nextSeqId('ART', sheet);

  // Normalise colors — trim whitespace, ensure numeric qty
  const colorsClean = data.colors.map(c => ({
    color: String(c.color).trim(),
    qty  : parseFloat(c.qty)
  }));

  sheet.appendRow([
    articleId,
    data.jobCardNumber ? data.jobCardNumber.trim() : '',
    data.articleNumber.trim(),
    JSON.stringify(colorsClean),
    new Date()
  ]);

  logActivity('CREATE', 'Article', articleId,
    { articleNumber: data.articleNumber },
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ articleId }, 'Article ' + articleId + ' created successfully');
}

function updateArticle(data, caller) {
  if (!data.articleId) return error('Article ID is required');
  if (!data.articleNumber || !data.articleNumber.trim())
    return error('Article Number is required');

  const colorErr = _validateColors(data.colors);
  if (colorErr) return error(colorErr);

  const sheet  = getSheet(SHEETS.ARTICLES);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex((r, i) => i > 0 && r[0] === data.articleId);
  if (rowIdx === -1) return error('Article not found: ' + data.articleId);

  // Duplicate article number check (excluding self)
  const dup = values.slice(1).find((r, i) =>
    i + 1 !== rowIdx &&
    String(r[2]).trim().toLowerCase() === data.articleNumber.trim().toLowerCase()
  );
  if (dup) return error('Another article with this Article Number already exists');

  const colorsClean = data.colors.map(c => ({
    color: String(c.color).trim(),
    qty  : parseFloat(c.qty)
  }));

  const sheetRow = rowIdx + 1;
  sheet.getRange(sheetRow, 2).setValue(data.jobCardNumber ? data.jobCardNumber.trim() : '');
  sheet.getRange(sheetRow, 3).setValue(data.articleNumber.trim());
  sheet.getRange(sheetRow, 4).setValue(JSON.stringify(colorsClean));

  logActivity('UPDATE', 'Article', data.articleId,
    { articleNumber: data.articleNumber },
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ articleId: data.articleId }, 'Article updated successfully');
}

function deleteArticle(data, caller) {
  if (!data.articleId) return error('Article ID is required');

  // Guard: do not delete if article is referenced in any SO
  const soSheet = getSheet(SHEETS.SALES_ORDERS);
  const soData  = soSheet.getDataRange().getValues();
  const linked  = soData.slice(1).find(r => r[3] === data.articleId);
  if (linked) return error('Cannot delete — this article is used in one or more Sales Orders');

  const sheet  = getSheet(SHEETS.ARTICLES);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex((r, i) => i > 0 && r[0] === data.articleId);
  if (rowIdx === -1) return error('Article not found: ' + data.articleId);

  sheet.deleteRow(rowIdx + 1);

  logActivity('DELETE', 'Article', data.articleId, {},
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ articleId: data.articleId }, 'Article deleted successfully');
}

// ============================================================
// SECTION 3 — SALES ORDERS
// ============================================================
// One SO = one header row.  Line items stored as JSON in col 5.
// Line item schema:
//   { articleId, articleNumber, color, size, qty, ratePerPiece,
//     discountPct, grossAmount, discountAmount, netAmount }
//
// Status: Pending | Dispatched

const SO_SIZES = ['S', 'M', 'L', 'XL', 'XXL', 'XXXL', 'XXXXL'];

function getSalesOrders() {
  const sheet = getSheet(SHEETS.SALES_ORDERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  return success(data.slice(1).map(r => {
    const obj = rowToObj(headers, r);
    try { obj.Line_Items = JSON.parse(obj.Line_Items || '[]'); } catch(e) { obj.Line_Items = []; }
    return obj;
  }));
}

function getSalesOrderById(soId) {
  const sheet   = getSheet(SHEETS.SALES_ORDERS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const row     = data.slice(1).find(r => r[0] === soId);
  if (!row) return error('Sales Order not found: ' + soId);
  const obj = rowToObj(headers, row);
  try { obj.Line_Items = JSON.parse(obj.Line_Items || '[]'); } catch(e) { obj.Line_Items = []; }
  return success(obj);
}

function _validateSOLineItems(items) {
  if (!Array.isArray(items) || items.length === 0)
    return 'At least one line item is required';
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!it.articleId)          return 'Article is required on line ' + (i + 1);
    if (!it.color)              return 'Color is required on line ' + (i + 1);
    if (!it.size || !SO_SIZES.includes(it.size))
      return 'Valid Size is required on line ' + (i + 1);
    if (!it.qty || isNaN(parseFloat(it.qty)) || parseFloat(it.qty) <= 0)
      return 'Valid Qty is required on line ' + (i + 1);
    if (it.ratePerPiece === undefined || isNaN(parseFloat(it.ratePerPiece)) || parseFloat(it.ratePerPiece) < 0)
      return 'Valid Rate Per Piece is required on line ' + (i + 1);
    if (it.discountPct === undefined || isNaN(parseFloat(it.discountPct)) ||
        parseFloat(it.discountPct) < 0 || parseFloat(it.discountPct) > 100)
      return 'Discount % must be between 0 and 100 on line ' + (i + 1);
  }
  return null;
}

function _calcLineItem(it) {
  const qty          = parseFloat(it.qty)          || 0;
  const rate         = parseFloat(it.ratePerPiece) || 0;
  const discountPct  = parseFloat(it.discountPct)  || 0;
  const grossAmount  = qty * rate;
  const discountAmt  = grossAmount * discountPct / 100;
  const netAmount    = grossAmount - discountAmt;
  return {
    articleId    : it.articleId,
    articleNumber: it.articleNumber || '',
    color        : it.color,
    size         : it.size,
    qty,
    ratePerPiece : rate,
    discountPct,
    grossAmount  : Math.round(grossAmount  * 100) / 100,
    discountAmount: Math.round(discountAmt * 100) / 100,
    netAmount    : Math.round(netAmount    * 100) / 100
  };
}

function createSalesOrder(data, caller) {
  if (!data.customerId) return error('Customer is required');

  const lineErr = _validateSOLineItems(data.lineItems);
  if (lineErr) return error(lineErr);

  // Validate customer exists
  const custRes = JSON.parse(getCustomerById(data.customerId));
  if (!custRes.success) return error('Customer not found');
  const customer = custRes.data;

  const sheet = getSheet(SHEETS.SALES_ORDERS);
  const soId  = _nextSeqId('SO', sheet);

  // Calculate line items + SO totals
  const lineItems    = data.lineItems.map(_calcLineItem);
  const totalQty     = lineItems.reduce((s, l) => s + l.qty,          0);
  const totalGross   = lineItems.reduce((s, l) => s + l.grossAmount,   0);
  const totalDiscount= lineItems.reduce((s, l) => s + l.discountAmount, 0);
  const totalNet     = lineItems.reduce((s, l) => s + l.netAmount,     0);

  sheet.appendRow([
    soId,
    data.customerId,
    customer.Customer_Name,
    lineItems.length > 0 ? lineItems[0].articleId : '',       // first article (for quick reference)
    JSON.stringify(lineItems),
    Math.round(totalQty      * 100) / 100,
    Math.round(totalGross    * 100) / 100,
    Math.round(totalDiscount * 100) / 100,
    Math.round(totalNet      * 100) / 100,
    'Pending',
    data.notes ? data.notes.trim() : '',
    caller ? (caller.username || '') : '',
    new Date()
  ]);

  logActivity('CREATE', 'SalesOrder', soId,
    { customerId: data.customerId, totalNet },
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ soId }, 'Sales Order ' + soId + ' created successfully');
}

function updateSOStatus(data, caller) {
  if (!data.soId)   return error('SO ID is required');
  if (!data.status) return error('Status is required');
  if (!['Pending', 'Dispatched'].includes(data.status))
    return error('Status must be Pending or Dispatched');

  const sheet  = getSheet(SHEETS.SALES_ORDERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex((r, i) => i > 0 && r[0] === data.soId);
  if (rowIdx === -1) return error('Sales Order not found: ' + data.soId);

  sheet.getRange(rowIdx + 1, 10).setValue(data.status);
  const colors = { Pending: '#FFF3CD', Dispatched: '#D1FAE5' };
  sheet.getRange(rowIdx + 1, 10).setBackground(colors[data.status] || '#FFFFFF');

  logActivity('UPDATE_STATUS', 'SalesOrder', data.soId,
    { status: data.status },
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ soId: data.soId, status: data.status }, 'Status updated to ' + data.status);
}

function deleteSalesOrder(data, caller) {
  if (!data.soId) return error('SO ID is required');

  const sheet  = getSheet(SHEETS.SALES_ORDERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex((r, i) => i > 0 && r[0] === data.soId);
  if (rowIdx === -1) return error('Sales Order not found: ' + data.soId);

  sheet.deleteRow(rowIdx + 1);

  logActivity('DELETE', 'SalesOrder', data.soId, {},
    caller ? caller.username : 'System',
    caller ? caller.role     : 'System'
  );

  return success({ soId: data.soId }, 'Sales Order deleted successfully');
}

// Dashboard helper: SO stats for getDashboardData()
function getSalesOrderStats() {
  const sheet = getSheet(SHEETS.SALES_ORDERS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { totalSOs: 0, pendingSOs: 0, dispatchedSOs: 0 };
  const rows = data.slice(1);
  return {
    totalSOs      : rows.length,
    pendingSOs    : rows.filter(r => r[9] === 'Pending').length,
    dispatchedSOs : rows.filter(r => r[9] === 'Dispatched').length
  };
}