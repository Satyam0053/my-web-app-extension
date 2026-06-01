// ============================================================
// FabricIssueService.gs — Fabric Issue
// FIXED: B-009 AvailableQty snapshot uses pre-validation value, not re-fetch
//        Performance: Pre-load master data maps before item loop
// ============================================================

var ISSUE_FOR_VALUES = ['In House Production', 'Fabricator', 'Sample'];

function getFabricIssues(token, filters) {
  try {
    var auth = requireReadAccess(token, 'FabricIssue');
    if (!auth.valid) return auth.error;

    var issues = getSheetData('FabricIssue');

    if (filters) {
      if (filters.articleID) issues = issues.filter(function(r) { return r.ArticleID === filters.articleID; });
      if (filters.soID)      issues = issues.filter(function(r) { return r.SOID      === filters.soID; });
      if (filters.status)    issues = issues.filter(function(r) { return r.Status    === filters.status; });
      if (filters.issueFor)  issues = issues.filter(function(r) { return r.IssueFor  === filters.issueFor; });
      if (filters.from)      issues = issues.filter(function(r) { return r.IssueDate >= filters.from; });
      if (filters.to)        issues = issues.filter(function(r) { return r.IssueDate <= filters.to; });
    }

    issues.sort(function(a, b) { return String(b.IssueID).localeCompare(String(a.IssueID)); });
    return successResponse({ issues: issues, count: issues.length });
  } catch (e) {
    return errorResponse('Failed to fetch fabric issues: ' + e.message, 'FETCH_ERROR');
  }
}

function getFabricIssueByID(token, issueID) {
  try {
    var auth = requireReadAccess(token, 'FabricIssue');
    if (!auth.valid) return auth.error;

    if (!issueID) return errorResponse('IssueID required', 'VALIDATION_ERROR');

    var issues = getSheetData('FabricIssue').filter(function(r) { return r.IssueID === issueID; });
    if (issues.length === 0) return errorResponse('Fabric Issue not found: ' + issueID, 'NOT_FOUND');

    var issue = issues[0];
    var items = getSheetData('FabricIssueItems').filter(function(r) { return r.IssueID === issueID; });

    return successResponse({ issue: issue, items: items });
  } catch (e) {
    return errorResponse('Failed to fetch fabric issue: ' + e.message, 'FETCH_ERROR');
  }
}

function getBOMSuggestion(token, articleID, soID) {
  try {
    var auth = requireReadAccess(token, 'FabricIssue');
    if (!auth.valid) return auth.error;

    if (!articleID) return errorResponse('ArticleID required', 'VALIDATION_ERROR');

    var bomLines = getSheetData('BOM').filter(function(r) { return r.ArticleID === articleID; });
    if (bomLines.length === 0) {
      return successResponse({ suggestions: [], message: 'No BOM defined for this article' });
    }

    // Build SO qty map using ArticleID+Size+Color (resolved to consistent lowercase string)
    var soItemQtyMap = {};
    if (soID) {
      getSheetData('SalesOrderItems')
        .filter(function(r) { return r.SOID === soID && r.ArticleID === articleID; })
        .forEach(function(r) {
          // FIX B-015: Key uses ArticleID+Size+Color — all compared as lowercase strings
          var key = String(r.ArticleID) + '|' + String(r.Size) + '|' + String(r.Color).toLowerCase();
          soItemQtyMap[key] = (soItemQtyMap[key] || 0) + (parseInt(r.OrderedQty) || 0);
        });
    }

    var inventory  = getSheetData('Inventory');
    var fabricReq  = {};

    bomLines.forEach(function(b) {
      // FIX B-015: Match key exactly as built in soItemQtyMap
      var soKey  = String(b.ArticleID) + '|' + String(b.Size) + '|' + String(b.Color).toLowerCase();
      var soQty  = soItemQtyMap[soKey] || 0;
      var reqKey = b.FabricID + '|' + b.FabricColorID;
      var eff    = parseFloat(b.EffectiveConsumption) || 0;
      var reqQty = soQty > 0 ? Math.round(soQty * eff * 1000) / 1000 : 0;

      if (!fabricReq[reqKey]) {
        fabricReq[reqKey] = { FabricID: b.FabricID, FabricColorID: b.FabricColorID, UOM: b.UOM, SuggestedQty: 0, AvailableQty: 0 };
        var invRow = inventory.filter(function(r) {
          return r.FabricID === b.FabricID && r.FabricColorID === b.FabricColorID;
        });
        fabricReq[reqKey].AvailableQty = invRow.length > 0 ? (parseFloat(invRow[0].AvailableQty) || 0) : 0;
      }
      fabricReq[reqKey].SuggestedQty = Math.round((fabricReq[reqKey].SuggestedQty + reqQty) * 1000) / 1000;
    });

    var suggestions = Object.keys(fabricReq).map(function(k) {
      var s = fabricReq[k];
      s.InsufficientStock = s.SuggestedQty > s.AvailableQty;
      return s;
    });

    return successResponse({ suggestions: suggestions });
  } catch (e) {
    return errorResponse('Failed to get BOM suggestion: ' + e.message, 'FETCH_ERROR');
  }
}

function createFabricIssue(token, data) {
  try {
    var auth = requireWriteAccess(token, 'FabricIssue');
    if (!auth.valid) return auth.error;

    var missing = validateRequired(['IssueDate', 'ArticleID', 'IssueFor'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    if (ISSUE_FOR_VALUES.indexOf(data.IssueFor) === -1) {
      return errorResponse('Invalid IssueFor value. Must be: ' + ISSUE_FOR_VALUES.join(', '), 'VALIDATION_ERROR');
    }
    if (!idExists('ArticleMaster', 'ArticleID', data.ArticleID)) {
      return errorResponse('Invalid ArticleID', 'VALIDATION_ERROR');
    }
    if (data.SOID && !idExists('SalesOrders', 'SOID', data.SOID)) {
      return errorResponse('Invalid SOID', 'VALIDATION_ERROR');
    }

    if (data.IssueFor === 'Fabricator') {
      if (!data.FabricatorID) return errorResponse('FabricatorID required for Fabricator issues', 'VALIDATION_ERROR');
      if (!idExists('VendorMaster', 'VendorID', data.FabricatorID)) return errorResponse('Invalid FabricatorID', 'VALIDATION_ERROR');
      var fabricators = getSheetData('VendorMaster').filter(function(r) {
        return r.VendorID === data.FabricatorID && (r.VendorType === 'Fabricator' || r.VendorType === 'Both');
      });
      if (fabricators.length === 0) return errorResponse('Selected vendor is not a Fabricator', 'VALIDATION_ERROR');
    }

    if (!Array.isArray(data.items) || data.items.length === 0) {
      return errorResponse('At least one fabric item is required', 'VALIDATION_ERROR');
    }

    // Pre-load master data maps to avoid repeated getSheetData in the loop
    var fabricMasterMap = {};
    getSheetData('FabricMaster').forEach(function(f) { fabricMasterMap[f.FabricID] = f; });
    var fabricColorMap = {};
    getSheetData('FabricColors').forEach(function(c) { fabricColorMap[c.FabricColorID] = c; });
    var allInventory = getSheetData('Inventory');

    // Item validation — build inventory snapshot ONCE before loop
    var inventorySnapshots = {};
    allInventory.forEach(function(r) {
      var key = r.FabricID + '|' + r.FabricColorID;
      inventorySnapshots[key] = parseFloat(r.AvailableQty) || 0;
    });

    for (var i = 0; i < data.items.length; i++) {
      var item    = data.items[i];
      var lineNum = i + 1;

      var missingItem = validateRequired(['FabricID', 'FabricColorID', 'IssueQty', 'UOM'], item);
      if (missingItem.length) return errorResponse('Line ' + lineNum + ': Missing ' + missingItem.join(', '), 'VALIDATION_ERROR');

      if (['Meters', 'Yards'].indexOf(item.UOM) === -1) {
        return errorResponse('Line ' + lineNum + ': Invalid UOM', 'VALIDATION_ERROR');
      }
      if (!isPositiveNumber(item.IssueQty)) {
        return errorResponse('Line ' + lineNum + ': IssueQty must be positive', 'VALIDATION_ERROR');
      }
      if (!fabricMasterMap[item.FabricID]) {
        return errorResponse('Line ' + lineNum + ': Invalid FabricID', 'VALIDATION_ERROR');
      }
      if (!fabricColorMap[item.FabricColorID]) {
        return errorResponse('Line ' + lineNum + ': Invalid FabricColorID', 'VALIDATION_ERROR');
      }

      var invKey    = item.FabricID + '|' + item.FabricColorID;
      var available = inventorySnapshots[invKey] !== undefined ? inventorySnapshots[invKey] : 0;
      var issueQty  = parseFloat(item.IssueQty) || 0;

      if (issueQty > available) {
        return errorResponse(
          'Line ' + lineNum + ': INSUFFICIENT STOCK. ' +
          'Requested: ' + issueQty + ', Available: ' + available +
          ' (Fabric: ' + item.FabricID + ', Color: ' + item.FabricColorID + ')',
          'INSUFFICIENT_STOCK'
        );
      }

      // Reduce snapshot for subsequent lines using same fabric+color
      inventorySnapshots[invKey] = Math.round((available - issueQty) * 1000) / 1000;
      // FIX B-009: Store the snapshot value at validation time on the item
      item._availableAtValidation = available;
    }

    // All validations passed — write
    var now     = nowISO();
    var user    = auth.session;
    var issueID = generateYearID('ISS', 'FabricIssue', 1);

    appendRow('FabricIssue', [
      issueID, issueID, data.IssueDate, data.ArticleID,
      data.SOID || '', data.IssueFor, data.FabricatorID || '',
      sanitize(data.Remarks || ''), 'Issued', user.userID, now, '', ''
    ]);

    data.items.forEach(function(item) {
      var issueQty = parseFloat(item.IssueQty) || 0;
      var itemID   = generateID('ISI', 'FabricIssueItems', 1);

      // FIX B-009: Use pre-validation snapshot, NOT a new getSheetData() call
      var availableSnapshot = item._availableAtValidation !== undefined ? item._availableAtValidation : 0;

      appendRow('FabricIssueItems', [
        itemID, issueID, item.FabricID, item.FabricColorID,
        availableSnapshot,  // Correct snapshot at time of validation
        issueQty, item.UOM,
        0  // ReturnedQty starts at 0
      ]);

      _deductInventory(item.FabricID, item.FabricColorID, issueQty, issueID, user.userID, data.IssueFor);
    });

    if (data.IssueFor === 'Fabricator' && data.FabricatorID) {
      _createFabricatorChallanFromIssue(issueID, data, user, now);
    }

    writeAuditLog({ userID: user.userID, userName: user.userName, action: 'CREATE', module: 'FabricIssue', recordID: issueID, oldValues: '', newValues: data });
    return successResponse({ IssueID: issueID }, 'Fabric Issue created: ' + issueID);
  } catch (e) {
    return errorResponse('Failed to create fabric issue: ' + e.message, 'CREATE_ERROR');
  }
}

function _deductInventory(fabricID, fabricColorID, issueQty, issueID, userID, issueFor) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Inventory');
  var data  = sheet.getDataRange().getValues();
  var now   = nowISO();
  var newBalance = 0;

  for (var r = 1; r < data.length; r++) {
    if (data[r][SCHEMA.Inventory.FabricID - 1] === fabricID &&
        data[r][SCHEMA.Inventory.FabricColorID - 1] === fabricColorID) {

      var issued    = parseFloat(data[r][SCHEMA.Inventory.IssuedQty - 1])   || 0;
      var available = parseFloat(data[r][SCHEMA.Inventory.AvailableQty - 1]) || 0;
      var newIssued = Math.round((issued + issueQty) * 1000) / 1000;
      newBalance    = Math.round((available - issueQty) * 1000) / 1000;

      sheet.getRange(r + 1, SCHEMA.Inventory.IssuedQty).setValue(newIssued);
      sheet.getRange(r + 1, SCHEMA.Inventory.AvailableQty).setValue(newBalance);
      sheet.getRange(r + 1, SCHEMA.Inventory.LastUpdated).setValue(now);
      break;
    }
  }

  var transType = issueFor === 'Fabricator' ? 'Fabricator Issue' : 'Fabric Issue';
  var ledgerID  = generateID('LGR', 'InventoryLedger', 1);
  appendRow('InventoryLedger', [
    ledgerID, now, transType, issueID,
    fabricID, fabricColorID,
    0, issueQty, newBalance,
    issueFor + ': ' + issueID, userID
  ]);
}

function _createFabricatorChallanFromIssue(issueID, data, user, now) {
  try {
    var challanID = generateYearID('FJC', 'FabricatorChallan', 1);

    appendRow('FabricatorChallan', [
      challanID, data.IssueDate, 'Out', data.FabricatorID, issueID,
      'Auto-created from Fabric Issue: ' + issueID, 'Open',
      user.userID, now, '', ''
    ]);

    data.items.forEach(function(item) {
      var cItemID = generateID('FCI', 'FabricatorChallanItems', 1);
      appendRow('FabricatorChallanItems', [
        cItemID, challanID, item.FabricID, item.FabricColorID,
        parseFloat(item.IssueQty) || 0, item.UOM, ''
      ]);
    });

    writeAuditLog({ userID: user.userID, userName: user.userName, action: 'CREATE', module: 'FabricatorChallan', recordID: challanID, oldValues: '', newValues: { issueID: issueID } });
  } catch (e) {
    writeAuditLog({ userID: user.userID, userName: user.userName, action: 'ERROR', module: 'FabricatorChallan', recordID: issueID, oldValues: '', newValues: { error: e.message } });
  }
}
