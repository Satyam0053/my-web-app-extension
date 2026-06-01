// ============================================================
// SalesOrderService.gs — Sales Orders (Sprint S3)
// SO Header + Multi-item (Article × Size × Color × Qty)
// Status flow: Draft → Confirmed → In Production → Dispatched → Closed
//              Draft → Cancelled
// ============================================================

// ── SO Status constants ───────────────────────────────────
var SO_STATUSES  = ['Draft', 'Confirmed', 'In Production', 'Dispatched', 'Closed', 'Cancelled'];
var VALID_SIZES_ = ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'Custom'];

// ── Allowed status transitions ─────────────────────────────
var SO_TRANSITIONS = {
  'Draft':         ['Confirmed', 'Cancelled'],
  'Confirmed':     ['In Production', 'Cancelled'],
  'In Production': ['Dispatched'],
  'Dispatched':    ['Closed'],
  'Closed':        [],
  'Cancelled':     []
};

// ══════════════════════════════════════════════════════════════
// GET SALES ORDERS
// ══════════════════════════════════════════════════════════════

function getSalesOrders(token, filters) {
  try {
    var auth = requireReadAccess(token, 'SalesOrders');
    if (!auth.valid) return auth.error;

    var orders = getSheetData('SalesOrders');

    if (filters) {
      if (filters.buyerID) orders = orders.filter(function(r) { return r.BuyerID === filters.buyerID; });
      if (filters.status)  orders = orders.filter(function(r) { return r.Status  === filters.status; });
      if (filters.from)    orders = orders.filter(function(r) { return r.SODate  >= filters.from; });
      if (filters.to)      orders = orders.filter(function(r) { return r.SODate  <= filters.to; });
    }

    // Sort newest first
    orders.sort(function(a, b) { return String(b.SOID).localeCompare(String(a.SOID)); });

    return successResponse({ orders: orders, count: orders.length });
  } catch (e) {
    return errorResponse('Failed to fetch sales orders: ' + e.message, 'FETCH_ERROR');
  }
}

function getSalesOrderByID(token, soID) {
  try {
    var auth = requireReadAccess(token, 'SalesOrders');
    if (!auth.valid) return auth.error;

    if (!soID) return errorResponse('SOID is required', 'VALIDATION_ERROR');

    var orders = getSheetData('SalesOrders').filter(function(r) { return r.SOID === soID; });
    if (orders.length === 0) return errorResponse('Sales Order not found: ' + soID, 'NOT_FOUND');

    var so    = orders[0];
    var items = getSheetData('SalesOrderItems').filter(function(r) { return r.SOID === soID; });

    return successResponse({ so: so, items: items });
  } catch (e) {
    return errorResponse('Failed to fetch sales order: ' + e.message, 'FETCH_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// CREATE SALES ORDER
// ══════════════════════════════════════════════════════════════

/**
 * Creates a new Sales Order with one or more line items.
 * data: {
 *   BuyerID, SODate, DeliveryDate, Season, Remarks,
 *   items: [{ ArticleID, Size, Color, OrderedQty, Rate }]
 * }
 */
function createSalesOrder(token, data) {
  try {
    var auth = requireWriteAccess(token, 'SalesOrders');
    if (!auth.valid) return auth.error;

    // Header validation
    var missing = validateRequired(['BuyerID', 'SODate', 'DeliveryDate'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    if (!idExists('BuyerMaster', 'BuyerID', data.BuyerID)) {
      return errorResponse('Invalid BuyerID: ' + data.BuyerID, 'VALIDATION_ERROR');
    }
    if (data.DeliveryDate < data.SODate) {
      return errorResponse('Delivery date cannot be before SO date', 'VALIDATION_ERROR');
    }

    // Items validation
    if (!Array.isArray(data.items) || data.items.length === 0) {
      return errorResponse('At least one item line is required', 'VALIDATION_ERROR');
    }

    var seenLines = {};
    for (var i = 0; i < data.items.length; i++) {
      var item    = data.items[i];
      var lineNum = i + 1;

      var missingItem = validateRequired(['ArticleID', 'Size', 'Color', 'OrderedQty'], item);
      if (missingItem.length) {
        return errorResponse('Item line ' + lineNum + ': Missing ' + missingItem.join(', '), 'VALIDATION_ERROR');
      }
      if (!idExists('ArticleMaster', 'ArticleID', item.ArticleID)) {
        return errorResponse('Item line ' + lineNum + ': Invalid ArticleID: ' + item.ArticleID, 'VALIDATION_ERROR');
      }
      if (VALID_SIZES_.indexOf(item.Size) === -1) {
        return errorResponse('Item line ' + lineNum + ': Invalid size: ' + item.Size, 'VALIDATION_ERROR');
      }
      if (!isPositiveNumber(item.OrderedQty) || parseInt(item.OrderedQty) < 1) {
        return errorResponse('Item line ' + lineNum + ': Ordered qty must be a positive integer', 'VALIDATION_ERROR');
      }
      if (item.Rate !== undefined && item.Rate !== '' && parseFloat(item.Rate) < 0) {
        return errorResponse('Item line ' + lineNum + ': Rate cannot be negative', 'VALIDATION_ERROR');
      }

      // Check same article+size+color not duplicated in submission
      var lineKey = item.ArticleID + '|' + item.Size + '|' + String(item.Color).toLowerCase();
      if (seenLines[lineKey]) {
        return errorResponse('Duplicate item at line ' + lineNum + ': same Article, Size, and Color', 'VALIDATION_ERROR');
      }
      seenLines[lineKey] = true;
    }

    var now  = nowISO();
    var user = auth.session;
    var soID = generateYearID('SO', 'SalesOrders', 1);

    // Calculate total pieces
    var totalPieces = data.items.reduce(function(sum, item) {
      return sum + (parseInt(item.OrderedQty) || 0);
    }, 0);

    // Write SO header — columns: SOID,SONumber,BuyerID,SODate,DeliveryDate,Season,Remarks,Status,TotalPieces,CreatedBy,CreatedAt,UpdatedBy,UpdatedAt
    var soNumber = 'SO-' + soID.replace('SO-','');
    var soRow = {
      SOID:         soID,
      SONumber:     soNumber,
      BuyerID:      data.BuyerID,
      SODate:       data.SODate,
      DeliveryDate: data.DeliveryDate,
      Season:       sanitize(data.Season   || ''),
      Remarks:      sanitize(data.Remarks  || ''),
      Status:       'Draft',
      TotalPieces:  totalPieces,
      CreatedBy:    user.userID,
      CreatedAt:    now,
      UpdatedBy:    '',
      UpdatedAt:    ''
    };
    appendRow('SalesOrders', soRow);

    // Write SO items
    data.items.forEach(function(item) {
      var qty    = parseInt(item.OrderedQty) || 0;
      var rate   = parseFloat(item.Rate)     || 0;
      var amount = Math.round(qty * rate * 100) / 100;
      var itemID = generateID('SOI', 'SalesOrderItems', 1);

      var itemRow = {
        SOItemID:   itemID,
        SOID:       soID,
        ArticleID:  item.ArticleID,
        Size:       item.Size,
        Color:      sanitize(item.Color),
        OrderedQty: qty,
        Rate:       rate,
        Amount:     amount
      };
      appendRow('SalesOrderItems', itemRow);
    });

    writeAuditLog({
      userID:    user.userID,
      userName:  user.userName,
      action:    'CREATE',
      module:    'SalesOrders',
      recordID:  soID,
      oldValues: '',
      newValues: data
    });

    return successResponse({ SOID: soID, totalPieces: totalPieces }, 'Sales Order created: ' + soID);
  } catch (e) {
    return errorResponse('Failed to create sales order: ' + e.message, 'CREATE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// UPDATE SALES ORDER (header only — items managed separately)
// ══════════════════════════════════════════════════════════════

function updateSalesOrder(token, soID, data) {
  try {
    var auth = requireWriteAccess(token, 'SalesOrders');
    if (!auth.valid) return auth.error;

    if (!soID) return errorResponse('SOID required', 'VALIDATION_ERROR');

    var rowIndex = findRowByID('SalesOrders', 1, soID);
    if (!rowIndex) return errorResponse('Sales Order not found: ' + soID, 'NOT_FOUND');

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SalesOrders');
    var old   = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
    var currentStatus = old[7]; // col 8 = Status (0-indexed: 7)

    // Only allow edits when in Draft status
    if (currentStatus !== 'Draft') {
      return errorResponse('Only Draft orders can be edited. Current status: ' + currentStatus, 'INVALID_STATUS');
    }

    if (data.BuyerID && !idExists('BuyerMaster', 'BuyerID', data.BuyerID)) {
      return errorResponse('Invalid BuyerID', 'VALIDATION_ERROR');
    }
    if (data.DeliveryDate && data.SODate && data.DeliveryDate < data.SODate) {
      return errorResponse('Delivery date cannot be before SO date', 'VALIDATION_ERROR');
    }

    var now = nowISO();
    // Columns: SOID(1),SONumber(2),BuyerID(3),SODate(4),DeliveryDate(5),Season(6),Remarks(7),Status(8),TotalPieces(9),CreatedBy(10),CreatedAt(11),UpdatedBy(12),UpdatedAt(13)
    if (data.BuyerID)      sheet.getRange(rowIndex, 3).setValue(data.BuyerID);
    if (data.SODate)       sheet.getRange(rowIndex, 4).setValue(data.SODate);
    if (data.DeliveryDate) sheet.getRange(rowIndex, 5).setValue(data.DeliveryDate);
    if (data.Season !== undefined)  sheet.getRange(rowIndex, 6).setValue(sanitize(data.Season));
    if (data.Remarks !== undefined) sheet.getRange(rowIndex, 7).setValue(sanitize(data.Remarks));

    // Recalculate total pieces if items were also passed
    if (Array.isArray(data.items) && data.items.length > 0) {
      var total = data.items.reduce(function(s, it) { return s + (parseInt(it.OrderedQty) || 0); }, 0);
      sheet.getRange(rowIndex, 9).setValue(total);
    }

    sheet.getRange(rowIndex, 12).setValue(auth.session.userID);
    sheet.getRange(rowIndex, 13).setValue(now);

    writeAuditLog({
      userID:    auth.session.userID,
      userName:  auth.session.userName,
      action:    'UPDATE',
      module:    'SalesOrders',
      recordID:  soID,
      oldValues: old,
      newValues: data
    });

    return successResponse({ SOID: soID }, 'Sales Order updated');
  } catch (e) {
    return errorResponse('Failed to update sales order: ' + e.message, 'UPDATE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// UPDATE SO STATUS
// ══════════════════════════════════════════════════════════════

function updateSOStatus(token, soID, newStatus) {
  try {
    var auth = requireWriteAccess(token, 'SalesOrders');
    if (!auth.valid) return auth.error;

    if (!soID)      return errorResponse('SOID required', 'VALIDATION_ERROR');
    if (!newStatus) return errorResponse('New status required', 'VALIDATION_ERROR');

    var rowIndex = findRowByID('SalesOrders', 1, soID);
    if (!rowIndex) return errorResponse('Sales Order not found: ' + soID, 'NOT_FOUND');

    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName('SalesOrders');
    // Columns: SOID(1),SONumber(2),BuyerID(3),SODate(4),DeliveryDate(5),Season(6),Remarks(7),Status(8),TotalPieces(9),CreatedBy(10),CreatedAt(11),UpdatedBy(12),UpdatedAt(13)
    var current = sheet.getRange(rowIndex, 8).getValue();

    var allowed = SO_TRANSITIONS[current] || [];
    if (allowed.indexOf(newStatus) === -1) {
      return errorResponse(
        'Cannot move from "' + current + '" to "' + newStatus + '". ' +
        'Allowed: ' + (allowed.length ? allowed.join(', ') : 'none'),
        'INVALID_TRANSITION'
      );
    }

    sheet.getRange(rowIndex, 8).setValue(newStatus);
    sheet.getRange(rowIndex, 12).setValue(auth.session.userID);
    sheet.getRange(rowIndex, 13).setValue(nowISO());

    writeAuditLog({
      userID:    auth.session.userID,
      userName:  auth.session.userName,
      action:    'UPDATE',
      module:    'SalesOrders',
      recordID:  soID,
      oldValues: { Status: current },
      newValues: { Status: newStatus }
    });

    return successResponse({ SOID: soID, oldStatus: current, newStatus: newStatus }, 'Status updated to ' + newStatus);
  } catch (e) {
    return errorResponse('Failed to update SO status: ' + e.message, 'UPDATE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// SO ITEMS — Add / Update / Delete individual lines
// ══════════════════════════════════════════════════════════════

function addSOItem(token, soID, itemData) {
  try {
    var auth = requireWriteAccess(token, 'SalesOrders');
    if (!auth.valid) return auth.error;

    if (!soID) return errorResponse('SOID required', 'VALIDATION_ERROR');

    var soRow = findRowByID('SalesOrders', 1, soID);
    if (!soRow) return errorResponse('Sales Order not found: ' + soID, 'NOT_FOUND');

    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var soSheet = ss.getSheetByName('SalesOrders');
    var status  = soSheet.getRange(soRow, 8).getValue(); // col 8 = Status
    if (status !== 'Draft') {
      return errorResponse('Items can only be added to Draft orders', 'INVALID_STATUS');
    }

    var missing = validateRequired(['ArticleID', 'Size', 'Color', 'OrderedQty'], itemData);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    if (!idExists('ArticleMaster', 'ArticleID', itemData.ArticleID)) return errorResponse('Invalid ArticleID', 'VALIDATION_ERROR');
    if (VALID_SIZES_.indexOf(itemData.Size) === -1) return errorResponse('Invalid size', 'VALIDATION_ERROR');
    if (!isPositiveNumber(itemData.OrderedQty)) return errorResponse('OrderedQty must be positive', 'VALIDATION_ERROR');

    // Duplicate check within this SO
    var existing = getSheetData('SalesOrderItems').filter(function(r) {
      return r.SOID === soID
        && r.ArticleID === itemData.ArticleID
        && r.Size === itemData.Size
        && String(r.Color).toLowerCase() === String(itemData.Color).toLowerCase();
    });
    if (existing.length > 0) {
      return errorResponse('This Article + Size + Color already exists in the order', 'DUPLICATE');
    }

    var qty    = parseInt(itemData.OrderedQty) || 0;
    var rate   = parseFloat(itemData.Rate) || 0;
    var amount = Math.round(qty * rate * 100) / 100;
    var itemID = generateID('SOI', 'SalesOrderItems', 1);

    appendRow('SalesOrderItems', [itemID, soID, itemData.ArticleID, itemData.Size, sanitize(itemData.Color), qty, rate, amount]);

    // Update total pieces on header
    _recalcSOTotalPieces(soID, ss, soSheet, soRow);

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'CREATE', module: 'SalesOrders', recordID: soID, oldValues: '', newValues: itemData });
    return successResponse({ SOItemID: itemID }, 'Item added to order');
  } catch (e) {
    return errorResponse('Failed to add SO item: ' + e.message, 'CREATE_ERROR');
  }
}

function deleteSOItem(token, soItemID) {
  try {
    var auth = requireWriteAccess(token, 'SalesOrders');
    if (!auth.valid) return auth.error;

    if (!soItemID) return errorResponse('SOItemID required', 'VALIDATION_ERROR');

    var rowIndex = findRowByID('SalesOrderItems', 1, soItemID);
    if (!rowIndex) return errorResponse('SO Item not found: ' + soItemID, 'NOT_FOUND');

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var itemSheet = ss.getSheetByName('SalesOrderItems');
    var itemRow  = itemSheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    var soID     = itemRow[1];

    // Check SO is in Draft
    var soSheet = ss.getSheetByName('SalesOrders');
    var soRowIdx = findRowByID('SalesOrders', 1, soID);
    if (soRowIdx) {
      var status = soSheet.getRange(soRowIdx, 8).getValue(); // col 8 = Status
      if (status !== 'Draft') {
        return errorResponse('Items can only be deleted from Draft orders', 'INVALID_STATUS');
      }
    }

    itemSheet.deleteRow(rowIndex);
    if (soRowIdx) _recalcSOTotalPieces(soID, ss, soSheet, soRowIdx);

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'DELETE', module: 'SalesOrders', recordID: soItemID, oldValues: itemRow, newValues: '' });
    return successResponse({ SOItemID: soItemID }, 'Item removed from order');
  } catch (e) {
    return errorResponse('Failed to delete SO item: ' + e.message, 'DELETE_ERROR');
  }
}

// ── Internal: recalculate and update TotalPieces on SO header ─
function _recalcSOTotalPieces(soID, ss, soSheet, soRowIdx) {
  var items = getSheetData('SalesOrderItems').filter(function(r) { return r.SOID === soID; });
  var total = items.reduce(function(s, r) { return s + (parseInt(r.OrderedQty) || 0); }, 0);
  soSheet.getRange(soRowIdx, 9).setValue(total);  // col 9 = TotalPieces
}