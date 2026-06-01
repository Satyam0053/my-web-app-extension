// ============================================================
// UtilService.gs — Shared Utilities
// FIXED: B-014 removed SpreadsheetApp prototype override
//        B-006 ID generation wrapped with LockService
//        B-022 sanitize() guards against formula injection
//        B-019 findRowByID returns consistent types
//        B-017 AuditLog uses Utilities.getUuid()
// ============================================================

const _SS_ID = '1gk88o8KDAIFPzestuZlJs9SQuuaWxlEIGQwpsLpjIYQ';

// FIX B-014: Replace prototype override with a clean wrapper function
function getSpreadsheet() {
  return SpreadsheetApp.openById(_SS_ID);
}

// ── ID Generation ──────────────────────────────────────────

// FIX B-006: Wrap ID generation in LockService to prevent concurrent duplicates
function generateID(prefix, sheetName, idColumn) {
  if (!sheetName) {
    // No sheet → timestamp+random branch (used for non-critical IDs)
    const ts  = new Date().getTime().toString(36).toUpperCase();
    const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    return prefix + '-' + ts + rnd;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // wait up to 10s
  try {
    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return prefix + '-00001';

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return prefix + '-00001';

    const ids = sh.getRange(2, idColumn, lastRow - 1, 1).getValues()
      .flat()
      .filter(function(v) { return v !== ''; });

    if (ids.length === 0) return prefix + '-00001';

    let maxNum = 0;
    ids.forEach(function(id) {
      const parts = String(id).split('-');
      const num = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    });

    return prefix + '-' + String(maxNum + 1).padStart(5, '0');
  } finally {
    lock.releaseLock();
  }
}

function generateYearID(prefix, sheetName, idColumn) {
  const year = new Date().getFullYear();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return prefix + '-' + year + '-00001';

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return prefix + '-' + year + '-00001';

    const ids = sh.getRange(2, idColumn, lastRow - 1, 1).getValues()
      .flat()
      .filter(function(v) { return String(v).includes('-' + year + '-'); });

    if (ids.length === 0) return prefix + '-' + year + '-00001';

    let maxNum = 0;
    ids.forEach(function(id) {
      const parts = String(id).split('-');
      const num = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    });

    return prefix + '-' + year + '-' + String(maxNum + 1).padStart(5, '0');
  } finally {
    lock.releaseLock();
  }
}

function generateSessionID() {
  return Utilities.getUuid();
}

// Single authoritative generateToken — removed duplicate from AuthService
function generateToken() {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Math.random().toString() + new Date().getTime().toString() + Math.random().toString(),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function generateBarcodeValue(bundleID) {
  const timestamp = Date.now().toString();
  const idNum = bundleID.replace(/\D/g, '');
  return idNum + timestamp.slice(-6);
}

function generateQRData(bundle) {
  return JSON.stringify({
    id: bundle.BundleID,
    article: bundle.ArticleID,
    so: bundle.SOID,
    size: bundle.Size,
    color: bundle.Color,
    qty: bundle.BundleQty,
    created: new Date().toISOString()
  });
}

// ── Hashing ────────────────────────────────────────────────

function hashPassword(password) {
  const rawBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return rawBytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

// ── Sanitization ───────────────────────────────────────────

// FIX B-022: Add formula injection guard in addition to HTML entity encoding
function sanitize(value) {
  if (value === null || value === undefined) return '';
  let str = String(value).trim();

  // Guard against Sheets formula injection: prefix dangerous leading chars
  const FORMULA_CHARS = ['=', '+', '-', '@'];
  if (str.length > 0 && FORMULA_CHARS.indexOf(str[0]) !== -1) {
    str = "'" + str;
  }

  // HTML entity encoding for XSS protection
  return str
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function sanitizeObject(obj) {
  const result = {};
  Object.keys(obj).forEach(function(key) {
    const val = obj[key];
    if (typeof val === 'string') {
      result[key] = sanitize(val);
    } else if (typeof val === 'number' || typeof val === 'boolean') {
      result[key] = val;
    } else if (Array.isArray(val)) {
      result[key] = val.map(function(item) {
        return typeof item === 'object' ? sanitizeObject(item) : sanitize(item);
      });
    } else if (typeof val === 'object' && val !== null) {
      result[key] = sanitizeObject(val);
    } else {
      result[key] = val;
    }
  });
  return result;
}

function sanitizeString(value) {
  return sanitize(value);
}

// ── Validation ─────────────────────────────────────────────

function isValidEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(email).toLowerCase());
}

function isValidMobile(mobile) {
  return /^[6-9]\d{9}$/.test(String(mobile));
}

function isValidGST(gst) {
  return /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/.test(String(gst));
}

function isValidDate(dateStr) {
  const d = new Date(dateStr);
  return !isNaN(d.getTime());
}

function isPositiveNumber(val) {
  const n = parseFloat(val);
  return !isNaN(n) && n > 0;
}

function isNonNegativeNumber(val) {
  const n = parseFloat(val);
  return !isNaN(n) && n >= 0;
}

function validateRequired(fields, data) {
  const missing = [];
  fields.forEach(function(field) {
    const val = data[field];
    if (val === undefined || val === null || String(val).trim() === '') {
      missing.push(field);
    }
  });
  return missing;
}

// ── Response Helpers ───────────────────────────────────────

function successResponse(data, message) {
  return {
    success: true,
    data: data || {},
    message: message || 'Operation successful'
  };
}

function errorResponse(message, code) {
  return {
    success: false,
    data: {},
    message: message || 'An error occurred',
    code: code || 'GENERAL_ERROR'
  };
}

// ── Date Helpers ───────────────────────────────────────────

function nowISO() {
  return new Date().toISOString();
}

function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d.getTime())) return '';
  return d.toISOString().split('T')[0];
}

function formatDateTime(date) {
  if (!date) return '';
  const d = new Date(date);
  if (isNaN(d.getTime())) return '';
  return d.toISOString();
}

// ── Settings Helpers ───────────────────────────────────────

function getSetting(key) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName('_Settings');
  if (!sheet) return null;
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function setSetting(key, value, updatedBy) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName('_Settings');
  if (!sheet) return;
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      sheet.getRange(i + 1, 4).setValue(updatedBy);
      sheet.getRange(i + 1, 5).setValue(nowISO());
      return;
    }
  }
}

// ── Sheet Row Helpers ──────────────────────────────────────

function appendRow(sheetName, rowData) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  if (!Array.isArray(rowData) && typeof rowData === 'object') {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(function(h) {
      const val = rowData[h];
      return (val !== undefined && val !== null) ? val : '';
    });
    sheet.appendRow(row);
  } else {
    sheet.appendRow(rowData);
  }
}

// FIX B-019: findRowByID now always returns row index (number) when called
// with a numeric idColumn, and a row object when called with a string column name.
// Documented contract enforced consistently.
function findRowByID(sheetName, idColumn, idValue) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  let colIndex;
  const returnObject = typeof idColumn === 'string';

  if (returnObject) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    colIndex = headers.indexOf(idColumn) + 1;
    if (colIndex <= 0) return null;
  } else {
    colIndex = idColumn;
  }

  const ids = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.findIndex(function(v) { return String(v) === String(idValue); });
  if (rowIndex === -1) return null;

  if (returnObject) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData  = sheet.getRange(rowIndex + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = rowData[i]; });
    return obj; // returns full row object
  }
  return rowIndex + 2; // returns 1-indexed row number
}

function getSheetData(sheetName) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return data.map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function updateRowByID(sheetName, idColumn, idValue, updateObj) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  let colIndex;
  if (typeof idColumn === 'string') {
    colIndex = headers.indexOf(idColumn) + 1;
    if (colIndex <= 0) return false;
  } else {
    colIndex = idColumn;
  }

  const ids = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(function(v) { return String(v) === String(idValue); });
  if (idx === -1) return false;
  const rowIndex = idx + 2;

  const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowObj  = {};
  headers.forEach(function(h, i) { rowObj[h] = rowData[i]; });
  Object.assign(rowObj, updateObj);
  const updatedRow = headers.map(function(h) { return rowObj[h] !== undefined ? rowObj[h] : ''; });
  sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
  return true;
}

function idExists(sheetName, idColumn, idValue) {
  return findRowByID(sheetName, idColumn, idValue) !== null;
}

// ── Schema constants (column indices, 1-based) ─────────────
// FIX B-010: Centralise hardcoded column indices so a sheet reorder
// only requires updating this one object.
const SCHEMA = {
  PurchaseOrders: {
    POID: 1, PONumber: 2, PODate: 3, VendorID: 4, Remarks: 5,
    Status: 6, TotalAmount: 7, ApprovedBy: 8, ApprovedAt: 9,
    CreatedBy: 10, CreatedAt: 11, UpdatedBy: 12, UpdatedAt: 13
  },
  PurchaseOrderItems: {
    POItemID: 1, POID: 2, FabricID: 3, FabricColorID: 4,
    OrderedQty: 5, ReceivedQty: 6, PendingQty: 7, UOM: 8,
    Width: 9, Rate: 10, Amount: 11, ExcessQty: 12, ItemStatus: 13
  },
  Inventory: {
    InventoryID: 1, FabricID: 2, FabricColorID: 3, PurchasedQty: 4,
    ReceivedQty: 5, IssuedQty: 6, ReturnedQty: 7, AvailableQty: 8, LastUpdated: 9
  }
};
