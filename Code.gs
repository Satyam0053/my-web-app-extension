// ============================================================
// INVENTORY ERP — GOOGLE APPS SCRIPT BACKEND
// File: Code.gs — Main Entry Point & Router
// Version: 2.2 | Bug-Fixed
//
// FIXES v2.2:
//   • error() helper: first arg is always the user-facing message.
//     PERMISSION_DENIED was previously passed as first arg — now
//     kept consistent: error(message, code) so frontend always
//     reads res.error for the display message.
//   • logActivity: null-safe fallback when Sheets is unreachable
//   • dispatch: catches sheet init errors that previously surfaced
//     as "Server error" even on login
// ============================================================

// ── SPREADSHEET CONFIG ──────────────────────────────────────
const SS_ID = '1GliIhNyi7cZrv9UU5mhkJRR1MeLCwSyDT-cNcHQBcNw';

const SHEETS = {
  PURCHASE_ORDERS  : 'Purchase_Orders',
  FABRIC_RECEIPTS  : 'Fabric_Receipts',
  INVENTORY        : 'Inventory_Master',
  STOCK_ISSUE      : 'Stock_Issue',
  STOCK_RETURNS    : 'Stock_Returns',
  VENDORS          : 'Vendors',
  PRODUCTION_ORDERS: 'Production_Orders',
  ACTIVITY_LOGS    : 'Activity_Logs',
  DASHBOARD_DATA   : 'Dashboard_Data',
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

// ── LOGIN PAGE HELPER ────────────────────────────────────────
function getLoginPage() {
  return HtmlService.createHtmlOutputFromFile('login').getContent();
}

// ── SPREADSHEET ACCESSOR ─────────────────────────────────────
function getSpreadsheet() {
  return SS_ID
    ? SpreadsheetApp.openById(SS_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss    = getSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

// ── SHEET INITIALIZER ────────────────────────────────────────
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
// success(data, message) → JSON string
// error(message, details) → JSON string
//
// IMPORTANT: error() first arg is ALWAYS the human-readable message.
// The frontend reads res.error to display it. Keep this consistent
// everywhere — never swap the argument order.
function success(data, message) {
  return JSON.stringify({ success: true, data: data || null, message: message || 'OK' });
}

function error(message, details) {
  return JSON.stringify({ success: false, error: message, details: details || null });
}

// ── ACTIVITY LOGGER ──────────────────────────────────────────
// FIX v2.2: Full try/catch so sheet errors NEVER propagate up
// and silently break transactions or logins.
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
    // Never break a transaction — log to Apps Script logger only
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
// All frontend → backend calls route through here.
// Step 1: allow public actions (login, validateToken) without a token.
// Step 2: validate session token for all other actions.
// Step 3: Admin-only actions are further guarded inside their functions.
function dispatch(action, payload) {
  try {
    payload = typeof payload === 'string' ? JSON.parse(payload) : (payload || {});

    // ── STEP 1: PUBLIC ACTIONS ────────────────────────────
    if (PUBLIC_ACTIONS.includes(action)) {
      switch (action) {
        case 'login'         : return loginUser(payload);
        case 'validateToken' : return validateSessionToken(payload.token);
        default              : return error('Unknown public action: ' + action);
      }
    }

    // ── STEP 2: AUTHENTICATE ──────────────────────────────
    const token = payload._token;
    if (!token)
      return error('No session token. Please log in.', 'AUTH_REQUIRED');

    const authCheck = JSON.parse(validateSessionToken(token));
    if (!authCheck.success)
      return error('Session invalid. Please log in again.', 'AUTH_FAILED');

    // caller = { username, displayName, role }
    const caller = authCheck.data;

    // ── STEP 3: ROUTE ─────────────────────────────────────
    switch (action) {

      // ── Setup ────────────────────────────────────────────
      case 'setup'               : return setupAllSheets();

      // ── Auth ─────────────────────────────────────────────
      case 'logout'              : return logoutUser(payload, caller);
      case 'changePassword'      : return changePassword(payload, caller);

      // ── User Management (Admin-only) ──────────────────────
      case 'setupUsers'          : return setupUsers(payload, caller);
      case 'getUsers'            : return getUsers(caller);
      case 'addUser'             : return addUser(payload, caller);
      case 'updateUser'          : return updateUser(payload, caller);
      case 'deactivateUser'      : return deactivateUserAPI(payload, caller);
      case 'reactivateUser'      : return reactivateUserAPI(payload, caller);
      case 'adminResetPassword'  : return adminResetPassword(payload, caller);

      // ── Purchase Orders ───────────────────────────────────
      case 'createPO'            : return createPurchaseOrder(payload, caller);
      case 'getPOs'              : return getPurchaseOrders(payload);
      case 'getPOById'           : return getPOById(payload.id);
      case 'updatePOStatus'      : return updatePOStatus(payload, caller);

      // ── Fabric Receipts ───────────────────────────────────
      case 'receivefabric'       : return receiveFabric(payload, caller);
      case 'getReceipts'         : return getReceipts(payload);
      case 'getReceiptsByPO'     : return getReceiptsByPO(payload.poNumber);

      // ── Inventory ─────────────────────────────────────────
      case 'getInventory'        : return getInventory(payload);
      case 'getInventoryById'    : return getInventoryById(payload.id);
      case 'getLowStock'         : return getLowStockAlerts();
      case 'addInventoryItem'    : return addInventoryItem(payload, caller);

      // ── Stock Issue ───────────────────────────────────────
      case 'issueStock'          : return issueStock(payload, caller);
      case 'getIssues'           : return getIssues(payload);
      case 'getIssuesByInventory': return getIssuesByInventory(payload.inventoryId);

      // ── Stock Returns ─────────────────────────────────────
      case 'returnStock'         : return returnStock(payload, caller);
      case 'getReturns'          : return getReturns(payload);

      // ── Vendors ───────────────────────────────────────────
      case 'getVendors'          : return getVendors();
      case 'addVendor'           : return addVendor(payload, caller);

      // ── Production Orders ─────────────────────────────────
      case 'getProductionOrders' : return getProductionOrders();
      case 'addProductionOrder'  : return addProductionOrder(payload, caller);

      // ── Reports ───────────────────────────────────────────
      case 'getStockLedger'       : return getStockLedger(payload);
      case 'getDailyIssueReport'  : return getDailyIssueReport(payload);
      case 'getMonthlyConsumption': return getMonthlyConsumption(payload);
      case 'getVendorReport'      : return getVendorReport(payload);
      case 'getArticleReport'     : return getArticleReport(payload);
      case 'getFabricMovement'    : return getFabricMovement(payload);

      // ── Dashboard ─────────────────────────────────────────
      case 'getDashboard'        : return getDashboardData();
      case 'getActivityLogs'     : return getActivityLogs(payload);

      // ── Search ────────────────────────────────────────────
      case 'search'              : return globalSearch(payload.query);

      default: return error('Unknown action: ' + action);
    }

  } catch (e) {
    Logger.log('dispatch ERROR — action: ' + action + ' | ' + e.message + '\n' + e.stack);
    try { logActivity('ERROR', 'dispatch', action, { error: e.message }); } catch(le) {}
    return error('Server error: ' + e.message);
  }
}