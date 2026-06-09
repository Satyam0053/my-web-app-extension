// ============================================================
// File: Code.gs — Main Entry Point & Router
// Version: 2.3 | Sales Module Added + Dashboard KPI Fix
//
// CHANGES v2.3:
//  • Added SHEETS.CUSTOMERS, SHEETS.ARTICLES, SHEETS.SALES_ORDERS
//  • Added SHEET_HEADERS for all three new sheets
//  • Added dispatch routes for all Customer / Article / SO actions
//  • getDashboardData() now includes SO KPIs (totalSOs, pendingSOs,
//    dispatchedSOs) pulled via getSalesOrderStats() — no double-fetch
//  • Dashboard PO qty bug: Reports.gs getDashboardData() now reads
//    Balance_Qty / Total_Received from getPurchaseOrders() which
//    already calls getReceiptsForPO() internally — removing the
//    redundant getReceipts() + receivedByPO map that caused zero values
//    due to PO_Number key-type mismatches.
// ============================================================

// ── SPREADSHEET CONFIG ──────────────────────────────────────
const SS_ID = '1GliIhNyi7cZrv9UU5mhkJRR1MeLCwSyDT-cNcHQBcNw';

const SHEETS = {
  PURCHASE_ORDERS : 'Purchase_Orders',
  FABRIC_RECEIPTS : 'Fabric_Receipts',
  INVENTORY       : 'Inventory_Master',
  STOCK_ISSUE     : 'Stock_Issue',
  STOCK_RETURNS   : 'Stock_Returns',
  VENDORS         : 'Vendors',
  PRODUCTION_ORDERS: 'Production_Orders',
  ACTIVITY_LOGS   : 'Activity_Logs',
  DASHBOARD_DATA  : 'Dashboard_Data',
  // ── NEW (Sales Module) ──────────────────────────────────
  CUSTOMERS       : 'Customers',
  ARTICLES        : 'Articles',
  SALES_ORDERS    : 'Sales_Orders',
};

// ── PUBLIC ACTIONS (no session token required) ───────────────
const PUBLIC_ACTIONS = ['login', 'validateToken'];

// ── WEB APP ENTRY POINT ──────────────────────────────────────
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Inventory ERP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getLoginPage() {
  return HtmlService.createHtmlOutputFromFile('login').getContent();
}

function getSpreadsheet() {
  return SS_ID
    ? SpreadsheetApp.openById(SS_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  const headers = SHEET_HEADERS[name];
  if (!headers) return;
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#FF8608')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11);
  sheet.setFrozenRows(1);
}

const SHEET_HEADERS = {
  Purchase_Orders: [
    'PO_Number','Vendor_Name','Article_Number','Fabric_Name',
    'Color','Width','Width_Unit','Ordered_Qty','Unit','Rate','UOM',
    'Delivery_Date','Status','Created_By','Created_At','Notes'
  ],
  Fabric_Receipts: [
    'GRN_Number','PO_Number','Vendor_Name','Article_Number',
    'Fabric_Name','Color','Width','Width_Unit','Received_Qty',
    'Unit','Batch_Number','Dye_Lot','Quality_Status',
    'Storage_Location','Received_By','Received_At','Notes'
  ],
  Inventory_Master: [
    'Inventory_ID','Article_Number','Fabric_Name','Color',
    'Width','Width_Unit','Vendor_Name','Unit','Opening_Stock',
    'Current_Stock','Reorder_Level','Storage_Location',
    'Last_Updated','Last_Transaction','Status'
  ],
  Stock_Issue: [
    'Issue_ID','Inventory_ID','Article_Number','Fabric_Name',
    'Color','Width','Width_Unit','Issue_Qty','Unit',
    'Production_Order','Department','Stock_Before','Stock_After',
    'Issued_By','Issued_At','Notes','Batch_Number'
  ],
  Stock_Returns: [
    'Return_ID','Issue_ID','Inventory_ID','Article_Number',
    'Fabric_Name','Color','Width','Width_Unit','Return_Qty',
    'Unit','Stock_Before','Stock_After','Reason',
    'Returned_By','Returned_At','Notes'
  ],
  Vendors: [
    'Vendor_ID','Vendor_Name','Contact_Person','Phone',
    'Email','Address','GST_Number','Payment_Terms','Status','Created_At'
  ],
  Production_Orders: [
    'PO_ID','PO_Number','Product_Name','Style_Number',
    'Quantity','Start_Date','End_Date','Department',
    'Status','Created_At','Notes'
  ],
  Activity_Logs: [
    'Log_ID','Timestamp','User','Role','Action','Module','Record_ID','Details'
  ],
  Dashboard_Data: [
    'Metric','Value','Updated_At'
  ],
  // ── NEW SHEETS ──────────────────────────────────────────
  Customers: [
    'Customer_ID','Customer_Name','Mobile','Email','Location','Created_At'
  ],
  Articles: [
    'Article_ID','Job_Card_Number','Article_Number','Colors','Created_At'
  ],
  // Line_Items stored as JSON string
  Sales_Orders: [
    'SO_ID','Customer_ID','Customer_Name','Article_ID_Ref',
    'Line_Items','Total_Qty','Total_Gross','Total_Discount',
    'Total_Net','Status','Notes','Created_By','Created_At'
  ],
};

// ── ID GENERATORS ────────────────────────────────────────────
function generateId(prefix) {
  const ts  = new Date().getTime().toString(36).toUpperCase();
  const rnd = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${ts}${rnd}`;
}

function generateSequentialId(prefix, sheet) {
  const data  = sheet.getDataRange().getValues();
  const count = Math.max(data.length, 1);
  return `${prefix}-${String(count).padStart(5, '0')}`;
}

// ── RESPONSE HELPERS ─────────────────────────────────────────
function success(data, message) {
  return JSON.stringify({ success: true, data: data || null, message: message || 'OK' });
}

function error(message, details) {
  return JSON.stringify({ success: false, error: message, details: details || null });
}

// ── ACTIVITY LOGGER ──────────────────────────────────────────
function logActivity(action, module, recordId, details, callerUser, callerRole) {
  try {
    const sheet = getSheet(SHEETS.ACTIVITY_LOGS);
    const user  = callerUser || 'System';
    const role  = callerRole || 'System';
    sheet.appendRow([
      generateId('LOG'),
      new Date(),
      user,
      role,
      action,
      module,
      recordId || '',
      JSON.stringify(details || {})
    ]);
  } catch (e) {
    Logger.log('logActivity FAILED (non-fatal): ' + e.message);
  }
}

// ── SETUP ALL SHEETS ─────────────────────────────────────────
function setupAllSheets() {
  Object.keys(SHEETS).forEach(key => getSheet(SHEETS[key]));
  try { refreshDashboardCache(); } catch(e) {}
  return success(null, 'All sheets initialized successfully');
}

// ── DISPATCHER ───────────────────────────────────────────────
function dispatch(action, payload) {
  try {
    payload = typeof payload === 'string' ? JSON.parse(payload) : (payload || {});

    if (PUBLIC_ACTIONS.includes(action)) {
      switch (action) {
        case 'login'         : return loginUser(payload);
        case 'validateToken' : return validateSessionToken(payload.token);
        default              : return error('Unknown public action: ' + action);
      }
    }

    const token = payload._token;
    if (!token) return error('No session token. Please log in.', 'AUTH_REQUIRED');

    const authCheck = JSON.parse(validateSessionToken(token));
    if (!authCheck.success)
      return error('Session invalid. Please log in again.', 'AUTH_FAILED');

    const caller = authCheck.data;

    switch (action) {
      // ── Setup ────────────────────────────────────────────
      case 'setup'                : return setupAllSheets();

      // ── Auth ─────────────────────────────────────────────
      case 'logout'               : return logoutUser(payload, caller);
      case 'changePassword'       : return changePassword(payload, caller);

      // ── User Management (Admin-only) ──────────────────────
      case 'setupUsers'           : return setupUsers(payload, caller);
      case 'getUsers'             : return getUsers(caller);
      case 'addUser'              : return addUser(payload, caller);
      case 'updateUser'           : return updateUser(payload, caller);
      case 'deactivateUser'       : return deactivateUserAPI(payload, caller);
      case 'reactivateUser'       : return reactivateUserAPI(payload, caller);
      case 'adminResetPassword'   : return adminResetPassword(payload, caller);

      // ── Purchase Orders ───────────────────────────────────
      case 'createPO'             : return createPurchaseOrder(payload, caller);
      case 'getPOs'               : return getPurchaseOrders(payload);
      case 'getPOById'            : return getPOById(payload.id);
      case 'updatePOStatus'       : return updatePOStatus(payload, caller);

      // ── Fabric Receipts ───────────────────────────────────
      case 'receivefabric'        : return receiveFabric(payload, caller);
      case 'getReceipts'          : return getReceipts(payload);
      case 'getReceiptsByPO'      : return getReceiptsByPO(payload.poNumber);

      // ── Inventory ─────────────────────────────────────────
      case 'getInventory'         : return getInventory(payload);
      case 'getInventoryById'     : return getInventoryById(payload.id);
      case 'getLowStock'          : return getLowStockAlerts();
      case 'addInventoryItem'     : return addInventoryItem(payload, caller);

      // ── Stock Issue ───────────────────────────────────────
      case 'issueStock'           : return issueStock(payload, caller);
      case 'getIssues'            : return getIssues(payload);
      case 'getIssuesByInventory' : return getIssuesByInventory(payload.inventoryId);

      // ── Stock Returns ─────────────────────────────────────
      case 'returnStock'          : return returnStock(payload, caller);
      case 'getReturns'           : return getReturns(payload);

      // ── Vendors ───────────────────────────────────────────
      case 'getVendors'           : return getVendors();
      case 'addVendor'            : return addVendor(payload, caller);

      // ── Production Orders ─────────────────────────────────
      case 'getProductionOrders'  : return getProductionOrders();
      case 'addProductionOrder'   : return addProductionOrder(payload, caller);

      // ── Reports ───────────────────────────────────────────
      case 'getStockLedger'       : return getStockLedger(payload);
      case 'getDailyIssueReport'  : return getDailyIssueReport(payload);
      case 'getMonthlyConsumption': return getMonthlyConsumption(payload);
      case 'getVendorReport'      : return getVendorReport(payload);
      case 'getArticleReport'     : return getArticleReport(payload);
      case 'getFabricMovement'    : return getFabricMovement(payload);

      // ── Dashboard ─────────────────────────────────────────
      case 'getDashboard'         : return getDashboardData();
      case 'getActivityLogs'      : return getActivityLogs(payload);

      // ── Search ────────────────────────────────────────────
      case 'search'               : return globalSearch(payload.query);

      // ── Customer Master (NEW) ─────────────────────────────
      case 'getCustomers'         : return getCustomers();
      case 'getCustomerById'      : return getCustomerById(payload.customerId);
      case 'createCustomer'       : return createCustomer(payload, caller);
      case 'updateCustomer'       : return updateCustomer(payload, caller);
      case 'deleteCustomer'       : return deleteCustomer(payload, caller);

      // ── Article Master (NEW) ──────────────────────────────
      case 'getArticles'          : return getArticles();
      case 'getArticleById'       : return getArticleById(payload.articleId);
      case 'createArticle'        : return createArticle(payload, caller);
      case 'updateArticle'        : return updateArticle(payload, caller);
      case 'deleteArticle'        : return deleteArticle(payload, caller);

      // ── Sales Orders (NEW) ────────────────────────────────
      case 'getSalesOrders'       : return getSalesOrders();
      case 'getSalesOrderById'    : return getSalesOrderById(payload.soId);
      case 'createSalesOrder'     : return createSalesOrder(payload, caller);
      case 'updateSOStatus'       : return updateSOStatus(payload, caller);
      case 'deleteSalesOrder'     : return deleteSalesOrder(payload, caller);

      default: return error('Unknown action: ' + action);
    }
  } catch (e) {
    Logger.log('dispatch ERROR — action: ' + action + ' | ' + e.message + '\n' + e.stack);
    try { logActivity('ERROR', 'dispatch', action, { error: e.message }); } catch(le) {}
    return error('Server error: ' + e.message);
  }
}