// ============================================================
// BOMService.gs — Bill of Materials (Sprint S2)
// BOM is defined at Article × Size × Color × Fabric level
// Each article can have multiple BOM lines (one per size/color/fabric combo)
// ============================================================

// ── Column index map (1-based) for BOM sheet ──────────────
// BOMID(1) ArticleID(2) Size(3) Color(4) FabricID(5) FabricColorID(6)
// ConsumptionPerPc(7) UOM(8) WastagePercent(9) EffectiveConsumption(10)
// Remarks(11) CreatedBy(12) CreatedAt(13) UpdatedBy(14) UpdatedAt(15)

var VALID_SIZES = ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'Custom'];
var VALID_UOMS  = ['Meters', 'Yards','Pcs','Rolls'];

// ══════════════════════════════════════════════════════════════
// GET BOM
// ══════════════════════════════════════════════════════════════

/**
 * Returns all BOM lines for a given article.
 * Optionally filter by size or color.
 */
function getBOM(token, articleID, filters) {
  try {
    var auth = requireReadAccess(token, 'BOM');
    if (!auth.valid) return auth.error;

    if (!articleID) return errorResponse('ArticleID is required', 'VALIDATION_ERROR');

    // Verify article exists
    if (!idExists('ArticleMaster', 1, articleID)) {
      return errorResponse('Article not found: ' + articleID, 'NOT_FOUND');
    }

    var rows = getSheetData('BOM').filter(function(r) {
      return r.ArticleID === articleID;
    });

    if (filters) {
      if (filters.size)  rows = rows.filter(function(r) { return r.Size === filters.size; });
      if (filters.color) rows = rows.filter(function(r) { return r.Color === filters.color; });
    }

    // Sort: Size order, then Color alphabetically
    var sizeOrder = { XS: 1, S: 2, M: 3, L: 4, XL: 5, XXL: 6, Custom: 7 };
    rows.sort(function(a, b) {
      var sA = sizeOrder[a.Size] || 99;
      var sB = sizeOrder[b.Size] || 99;
      if (sA !== sB) return sA - sB;
      return String(a.Color).localeCompare(String(b.Color));
    });

    return successResponse({ bom: rows, count: rows.length });
  } catch (e) {
    return errorResponse('Failed to fetch BOM: ' + e.message, 'FETCH_ERROR');
  }
}

/**
 * Returns a summary of unique sizes and colors defined in the BOM for an article.
 * Used by production and sales order modules.
 */
function getBOMSummary(token, articleID) {
  try {
    var auth = requireReadAccess(token, 'BOM');
    if (!auth.valid) return auth.error;

    if (!articleID) return errorResponse('ArticleID required', 'VALIDATION_ERROR');

    var rows = getSheetData('BOM').filter(function(r) {
      return r.ArticleID === articleID;
    });

    var sizes  = [];
    var colors = [];
    rows.forEach(function(r) {
      if (r.Size  && sizes.indexOf(r.Size)   === -1) sizes.push(r.Size);
      if (r.Color && colors.indexOf(r.Color) === -1) colors.push(r.Color);
    });

    var sizeOrder = { XS: 1, S: 2, M: 3, L: 4, XL: 5, XXL: 6, Custom: 7 };
    sizes.sort(function(a, b) { return (sizeOrder[a] || 99) - (sizeOrder[b] || 99); });
    colors.sort();

    return successResponse({ sizes: sizes, colors: colors, totalLines: rows.length });
  } catch (e) {
    return errorResponse('Failed to fetch BOM summary: ' + e.message, 'FETCH_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// SAVE BOM (Upsert — full replace for an article)
// ══════════════════════════════════════════════════════════════

/**
 * Saves BOM for an article. Replaces all existing lines with the new set.
 * items: Array of { Size, Color, FabricID, FabricColorID, ConsumptionPerPc, UOM, WastagePercent, Remarks }
 *
 * Business rule: Each Size+Color+FabricID combination must be unique per article.
 */
function saveBOM(token, articleID, items) {
  try {
    var auth = requireWriteAccess(token, 'BOM');
    if (!auth.valid) return auth.error;

    if (!articleID) return errorResponse('ArticleID is required', 'VALIDATION_ERROR');
    if (!Array.isArray(items) || items.length === 0) {
      return errorResponse('At least one BOM line is required', 'VALIDATION_ERROR');
    }
    if (!idExists('ArticleMaster', 1, articleID)) {
      return errorResponse('Article not found: ' + articleID, 'NOT_FOUND');
    }

    // Validate each line
    var seenKeys = {};
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var lineNum = i + 1;

      var missing = validateRequired(['Size', 'Color', 'FabricID', 'FabricColorID', 'ConsumptionPerPc', 'UOM'], item);
      if (missing.length) {
        return errorResponse('Line ' + lineNum + ': Missing required fields: ' + missing.join(', '), 'VALIDATION_ERROR');
      }
      if (VALID_SIZES.indexOf(item.Size) === -1) {
        return errorResponse('Line ' + lineNum + ': Invalid size: ' + item.Size, 'VALIDATION_ERROR');
      }
      if (VALID_UOMS.indexOf(item.UOM) === -1) {
        return errorResponse('Line ' + lineNum + ': Invalid UOM: ' + item.UOM + '. Must be Meters or Yards', 'VALIDATION_ERROR');
      }
      if (!isPositiveNumber(item.ConsumptionPerPc)) {
        return errorResponse('Line ' + lineNum + ': Consumption must be a positive number', 'VALIDATION_ERROR');
      }
      if (!idExists('FabricMaster', 1, item.FabricID)) {
        return errorResponse('Line ' + lineNum + ': Invalid FabricID: ' + item.FabricID, 'VALIDATION_ERROR');
      }
      if (!idExists('FabricColors', 1, item.FabricColorID)) {
        return errorResponse('Line ' + lineNum + ': Invalid FabricColorID: ' + item.FabricColorID, 'VALIDATION_ERROR');
      }
      // Check FabricColor belongs to FabricMaster
      var fcRows = getSheetData('FabricColors').filter(function(r) {
        return r.FabricColorID === item.FabricColorID && r.FabricID === item.FabricID;
      });
      if (fcRows.length === 0) {
        return errorResponse('Line ' + lineNum + ': FabricColor does not belong to the selected Fabric', 'VALIDATION_ERROR');
      }

      var wastage = parseFloat(item.WastagePercent) || 0;
      if (wastage < 0 || wastage > 100) {
        return errorResponse('Line ' + lineNum + ': Wastage % must be between 0 and 100', 'VALIDATION_ERROR');
      }

      // Uniqueness check within this submission
      var key = item.Size + '|' + String(item.Color).toLowerCase() + '|' + item.FabricID + '|' + item.FabricColorID;
      if (seenKeys[key]) {
        return errorResponse('Duplicate BOM line at line ' + lineNum + ': same Size, Color, and Fabric combination', 'VALIDATION_ERROR');
      }
      seenKeys[key] = true;
    }

    var now  = nowISO();
    var user = auth.session;

    // Read old BOM for audit trail
    var oldBOM = getSheetData('BOM').filter(function(r) { return r.ArticleID === articleID; });

    // Delete all existing BOM lines for this article
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('BOM');
    var allData = sheet.getDataRange().getValues();
    // Collect row indices to delete (bottom to top to preserve indices)
    var toDelete = [];
    for (var r = allData.length - 1; r >= 1; r--) {
      if (allData[r][1] === articleID) { // column index 1 = ArticleID
        toDelete.push(r + 1); // 1-based row number
      }
    }
    toDelete.forEach(function(rowNum) {
      sheet.deleteRow(rowNum);
    });

    // Insert new BOM lines
    var newIDs = [];
    items.forEach(function(item) {
      var consumption = parseFloat(item.ConsumptionPerPc);
      var wastage     = parseFloat(item.WastagePercent) || 0;
      var effective   = Math.round((consumption * (1 + wastage / 100)) * 10000) / 10000;
      var id          = generateID('BOM', 'BOM', 1);

      var row = [
        id,
        articleID,
        item.Size,
        sanitize(item.Color),
        item.FabricID,
        item.FabricColorID,
        consumption,
        item.UOM,
        wastage,
        effective,
        sanitize(item.Remarks || ''),
        user.userID,
        now,
        '',
        ''
      ];
      appendRow('BOM', row);
      newIDs.push(id);
    });

    writeAuditLog({
      userID:    user.userID,
      userName:  user.userName,
      action:    'UPDATE',
      module:    'BOM',
      recordID:  articleID,
      oldValues: oldBOM,
      newValues: items
    });

    return successResponse({ savedCount: newIDs.length, bomIDs: newIDs }, 'BOM saved successfully (' + newIDs.length + ' lines)');
  } catch (e) {
    return errorResponse('Failed to save BOM: ' + e.message, 'SAVE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// DELETE SINGLE BOM LINE
// ══════════════════════════════════════════════════════════════

function deleteBOMItem(token, bomID) {
  try {
    var auth = requireWriteAccess(token, 'BOM');
    if (!auth.valid) return auth.error;

    if (!bomID) return errorResponse('BOMID required', 'VALIDATION_ERROR');

    var rowIndex = findRowByID('BOM', 1, bomID);
    if (!rowIndex) return errorResponse('BOM line not found: ' + bomID, 'NOT_FOUND');

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('BOM');
    var old   = sheet.getRange(rowIndex, 1, 1, 15).getValues()[0];

    sheet.deleteRow(rowIndex);

    writeAuditLog({
      userID:    auth.session.userID,
      userName:  auth.session.userName,
      action:    'DELETE',
      module:    'BOM',
      recordID:  bomID,
      oldValues: old,
      newValues: ''
    });

    return successResponse({ BOMID: bomID }, 'BOM line deleted');
  } catch (e) {
    return errorResponse('Failed to delete BOM line: ' + e.message, 'DELETE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// GET BOM FOR FABRIC REQUIREMENT CALCULATION
// ══════════════════════════════════════════════════════════════

/**
 * Given a Sales Order ID, calculates total fabric requirement by fabric + color.
 * Used by Purchase Planning and Fabric Issue modules.
 */
function getBOMFabricRequirement(token, soID) {
  try {
    var auth = requireReadAccess(token, 'BOM');
    if (!auth.valid) return auth.error;

    if (!soID) return errorResponse('SOID required', 'VALIDATION_ERROR');

    var soItems = getSheetData('SalesOrderItems').filter(function(r) {
      return r.SOID === soID;
    });

    if (soItems.length === 0) {
      return errorResponse('No items found for Sales Order: ' + soID, 'NOT_FOUND');
    }

    var allBOM = getSheetData('BOM');

    // Aggregate fabric requirement: FabricID + FabricColorID → total meters
    var requirement = {};

    soItems.forEach(function(soItem) {
      var bomLines = allBOM.filter(function(b) {
        return b.ArticleID === soItem.ArticleID
          && b.Size  === soItem.Size
          && String(b.Color).toLowerCase() === String(soItem.Color).toLowerCase();
      });

      bomLines.forEach(function(b) {
        var key = b.FabricID + '|' + b.FabricColorID;
        var qty = parseFloat(soItem.OrderedQty) || 0;
        var eff = parseFloat(b.EffectiveConsumption) || 0;

        if (!requirement[key]) {
          requirement[key] = {
            FabricID:      b.FabricID,
            FabricColorID: b.FabricColorID,
            UOM:           b.UOM,
            TotalQty:      0
          };
        }
        requirement[key].TotalQty = Math.round((requirement[key].TotalQty + (qty * eff)) * 1000) / 1000;
      });
    });

    var result = Object.keys(requirement).map(function(k) { return requirement[k]; });

    return successResponse({ requirements: result, soID: soID });
  } catch (e) {
    return errorResponse('Failed to calculate BOM requirement: ' + e.message, 'CALC_ERROR');
  }
}
