// ============================================================
// MasterService.gs — Master Data CRUD
// Buyer, Article, Fabric, FabricColors, Vendor
// ============================================================

// ══════════════════════════════════════════════════════════════
// BUYER MASTER
// ══════════════════════════════════════════════════════════════

function getBuyers(token) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'BuyerMaster');
    if (!valid) return error;
    const buyers = getSheetData('BuyerMaster');
    return successResponse({ buyers: buyers });
  } catch (e) {
    return errorResponse('Failed to fetch buyers: ' + e.message, 'FETCH_ERROR');
  }
}

function createBuyer(token, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'BuyerMaster');
    if (!valid) return error;

    const missing = validateRequired(['BuyerName'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    // Duplicate name check
    const existing = getSheetData('BuyerMaster');
    const dup = existing.find(function(r) {
      return String(r.BuyerName).toLowerCase() === String(data.BuyerName).toLowerCase();
    });
    if (dup) return errorResponse('Buyer name already exists: ' + data.BuyerName, 'DUPLICATE');

    if (data.Mobile && !isValidMobile(data.Mobile)) {
      return errorResponse('Invalid mobile number (10 digits required)', 'VALIDATION_ERROR');
    }
    if (data.Email && !isValidEmail(data.Email)) {
      return errorResponse('Invalid email address', 'VALIDATION_ERROR');
    }

    const now = nowISO();
    const id = generateID('BYR', 'BuyerMaster', 1);
    const row = [
      id,
      sanitize(data.BuyerName),
      sanitize(data.BuyerCode || ''),
      sanitize(data.ContactPerson || ''),
      sanitize(data.Mobile || ''),
      sanitize(data.Email || '').toLowerCase(),
      sanitize(data.Address || ''),
      sanitize(data.Country || ''),
      data.Status || 'Active',
      session.userID, now, '', ''
    ];
    appendRow('BuyerMaster', row);
    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'CREATE', module: 'BuyerMaster', recordID: id, oldValues: '', newValues: data });
    return successResponse({ BuyerID: id }, 'Buyer created successfully');
  } catch (e) {
    return errorResponse('Failed to create buyer: ' + e.message, 'CREATE_ERROR');
  }
}

function updateBuyer(token, id, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'BuyerMaster');
    if (!valid) return error;
    if (!id) return errorResponse('BuyerID required', 'VALIDATION_ERROR');

    const rowIndex = findRowByID('BuyerMaster', 1, id);
    if (!rowIndex) return errorResponse('Buyer not found: ' + id, 'NOT_FOUND');

    if (data.Mobile && !isValidMobile(data.Mobile)) return errorResponse('Invalid mobile number', 'VALIDATION_ERROR');
    if (data.Email && !isValidEmail(data.Email)) return errorResponse('Invalid email address', 'VALIDATION_ERROR');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BuyerMaster');
    const old = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];

    if (data.BuyerName) sheet.getRange(rowIndex, 2).setValue(sanitize(data.BuyerName));
    if (data.BuyerCode !== undefined) sheet.getRange(rowIndex, 3).setValue(sanitize(data.BuyerCode));
    if (data.ContactPerson !== undefined) sheet.getRange(rowIndex, 4).setValue(sanitize(data.ContactPerson));
    if (data.Mobile !== undefined) sheet.getRange(rowIndex, 5).setValue(sanitize(data.Mobile));
    if (data.Email !== undefined) sheet.getRange(rowIndex, 6).setValue(sanitize(data.Email).toLowerCase());
    if (data.Address !== undefined) sheet.getRange(rowIndex, 7).setValue(sanitize(data.Address));
    if (data.Country !== undefined) sheet.getRange(rowIndex, 8).setValue(sanitize(data.Country));
    if (data.Status) sheet.getRange(rowIndex, 9).setValue(data.Status);
    sheet.getRange(rowIndex, 12).setValue(session.userID);
    sheet.getRange(rowIndex, 13).setValue(nowISO());

    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'UPDATE', module: 'BuyerMaster', recordID: id, oldValues: old, newValues: data });
    return successResponse({ BuyerID: id }, 'Buyer updated successfully');
  } catch (e) {
    return errorResponse('Failed to update buyer: ' + e.message, 'UPDATE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// ARTICLE MASTER
// ══════════════════════════════════════════════════════════════

function getArticles(token, filters) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'ArticleMaster');
    if (!valid) return error;
    let articles = getSheetData('ArticleMaster');
    if (filters && filters.buyerID) {
      articles = articles.filter(function(r) { return r.BuyerID === filters.buyerID; });
    }
    if (filters && filters.status) {
      articles = articles.filter(function(r) { return r.Status === filters.status; });
    }
    return successResponse({ articles: articles });
  } catch (e) {
    return errorResponse('Failed to fetch articles: ' + e.message, 'FETCH_ERROR');
  }
}

function createArticle(token, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'ArticleMaster');
    if (!valid) return error;

    const missing = validateRequired(['ArticleCode', 'ArticleName', 'BuyerID'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    if (!idExists('BuyerMaster', 'BuyerID', data.BuyerID)) return errorResponse('Invalid BuyerID', 'VALIDATION_ERROR');

    const existing = getSheetData('ArticleMaster');
    const dup = existing.find(function(r) {
      return String(r.ArticleCode).toLowerCase() === String(data.ArticleCode).toLowerCase();
    });
    if (dup) return errorResponse('Article code already exists: ' + data.ArticleCode, 'DUPLICATE');

    const now = nowISO();
    const id = generateID('ART', 'ArticleMaster', 1);
    const row = [
      id,
      sanitize(data.ArticleCode),
      sanitize(data.ArticleName),
      data.BuyerID,
      sanitize(data.Category || ''),
      sanitize(data.Description || ''),
      sanitize(data.Season || ''),
      data.Status || 'Active',
      session.userID, now, '', ''
    ];
    appendRow('ArticleMaster', row);
    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'CREATE', module: 'ArticleMaster', recordID: id, oldValues: '', newValues: data });
    return successResponse({ ArticleID: id }, 'Article created successfully');
  } catch (e) {
    return errorResponse('Failed to create article: ' + e.message, 'CREATE_ERROR');
  }
}

function updateArticle(token, id, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'ArticleMaster');
    if (!valid) return error;
    if (!id) return errorResponse('ArticleID required', 'VALIDATION_ERROR');
    const rowIndex = findRowByID('ArticleMaster', 1, id);
    if (!rowIndex) return errorResponse('Article not found', 'NOT_FOUND');
    if (data.BuyerID && !idExists('BuyerMaster', 'BuyerID', data.BuyerID)) return errorResponse('Invalid BuyerID', 'VALIDATION_ERROR');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ArticleMaster');
    const old = sheet.getRange(rowIndex, 1, 1, 12).getValues()[0];

    if (data.ArticleCode) sheet.getRange(rowIndex, 2).setValue(sanitize(data.ArticleCode));
    if (data.ArticleName) sheet.getRange(rowIndex, 3).setValue(sanitize(data.ArticleName));
    if (data.BuyerID) sheet.getRange(rowIndex, 4).setValue(data.BuyerID);
    if (data.Category !== undefined) sheet.getRange(rowIndex, 5).setValue(sanitize(data.Category));
    if (data.Description !== undefined) sheet.getRange(rowIndex, 6).setValue(sanitize(data.Description));
    if (data.Season !== undefined) sheet.getRange(rowIndex, 7).setValue(sanitize(data.Season));
    if (data.Status) sheet.getRange(rowIndex, 8).setValue(data.Status);
    sheet.getRange(rowIndex, 11).setValue(session.userID);
    sheet.getRange(rowIndex, 12).setValue(nowISO());

    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'UPDATE', module: 'ArticleMaster', recordID: id, oldValues: old, newValues: data });
    return successResponse({ ArticleID: id }, 'Article updated successfully');
  } catch (e) {
    return errorResponse('Failed to update article: ' + e.message, 'UPDATE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// FABRIC MASTER
// ══════════════════════════════════════════════════════════════

function getFabrics(token) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'FabricMaster');
    if (!valid) return error;
    const fabrics = getSheetData('FabricMaster');
    return successResponse({ fabrics: fabrics });
  } catch (e) {
    return errorResponse('Failed to fetch fabrics: ' + e.message, 'FETCH_ERROR');
  }
}

function createFabric(token, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'FabricMaster');
    if (!valid) return error;

    const missing = validateRequired(['DesignName', 'DesignCode'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    const existing = getSheetData('FabricMaster');
    const dup = existing.find(function(r) {
      return String(r.DesignCode).toLowerCase() === String(data.DesignCode).toLowerCase();
    });
    if (dup) return errorResponse('Design code already exists: ' + data.DesignCode, 'DUPLICATE');

    const now = nowISO();
    const id = generateID('FAB', 'FabricMaster', 1);
    const row = [
      id,
      sanitize(data.DesignName),
      sanitize(data.DesignCode),
      sanitize(data.Category || ''),
      parseFloat(data.Width) || 0,
      data.UOM || 'Meters',
      sanitize(data.Remarks || ''),
      data.Status || 'Active',
      session.userID, now, '', ''
    ];
    appendRow('FabricMaster', row);
    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'CREATE', module: 'FabricMaster', recordID: id, oldValues: '', newValues: data });
    return successResponse({ FabricID: id }, 'Fabric created successfully');
  } catch (e) {
    return errorResponse('Failed to create fabric: ' + e.message, 'CREATE_ERROR');
  }
}

function updateFabric(token, id, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'FabricMaster');
    if (!valid) return error;
    if (!id) return errorResponse('FabricID required', 'VALIDATION_ERROR');
    const rowIndex = findRowByID('FabricMaster', 1, id);
    if (!rowIndex) return errorResponse('Fabric not found', 'NOT_FOUND');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('FabricMaster');
    const old = sheet.getRange(rowIndex, 1, 1, 12).getValues()[0];

    if (data.DesignName) sheet.getRange(rowIndex, 2).setValue(sanitize(data.DesignName));
    if (data.DesignCode) sheet.getRange(rowIndex, 3).setValue(sanitize(data.DesignCode));
    if (data.Category !== undefined) sheet.getRange(rowIndex, 4).setValue(sanitize(data.Category));
    if (data.Width !== undefined) sheet.getRange(rowIndex, 5).setValue(parseFloat(data.Width) || 0);
    if (data.UOM) sheet.getRange(rowIndex, 6).setValue(data.UOM);
    if (data.Remarks !== undefined) sheet.getRange(rowIndex, 7).setValue(sanitize(data.Remarks));
    if (data.Status) sheet.getRange(rowIndex, 8).setValue(data.Status);
    sheet.getRange(rowIndex, 11).setValue(session.userID);
    sheet.getRange(rowIndex, 12).setValue(nowISO());

    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'UPDATE', module: 'FabricMaster', recordID: id, oldValues: old, newValues: data });
    return successResponse({ FabricID: id }, 'Fabric updated successfully');
  } catch (e) {
    return errorResponse('Failed to update fabric: ' + e.message, 'UPDATE_ERROR');
  }
}

// ── Fabric Colors ────────────────────────────────────────────

function getFabricColors(token, fabricID) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'FabricMaster');
    if (!valid) return error;
    let colors = getSheetData('FabricColors');
    if (fabricID) colors = colors.filter(function(r) { return r.FabricID === fabricID; });
    return successResponse({ colors: colors });
  } catch (e) {
    return errorResponse('Failed to fetch fabric colors: ' + e.message, 'FETCH_ERROR');
  }
}

function addFabricColor(token, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'FabricMaster');
    if (!valid) return error;

    const missing = validateRequired(['FabricID', 'ColorName'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');
    if (!idExists('FabricMaster', 'FabricID', data.FabricID)) return errorResponse('Invalid FabricID', 'VALIDATION_ERROR');

    const existing = getSheetData('FabricColors').filter(function(r) { return r.FabricID === data.FabricID; });
    const dup = existing.find(function(r) {
      return String(r.ColorName).toLowerCase() === String(data.ColorName).toLowerCase();
    });
    if (dup) return errorResponse('Color already exists for this fabric', 'DUPLICATE');

    const id = generateID('FCOL', 'FabricColors', 1);
    const row = [id, data.FabricID, sanitize(data.ColorName), sanitize(data.ColorCode || ''), data.Status || 'Active'];
    appendRow('FabricColors', row);
    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'CREATE', module: 'FabricMaster', recordID: id, oldValues: '', newValues: data });
    return successResponse({ FabricColorID: id }, 'Color added successfully');
  } catch (e) {
    return errorResponse('Failed to add color: ' + e.message, 'CREATE_ERROR');
  }
}

function updateFabricColor(token, id, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'FabricMaster');
    if (!valid) return error;
    const rowIndex = findRowByID('FabricColors', 1, id);
    if (!rowIndex) return errorResponse('Color not found', 'NOT_FOUND');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('FabricColors');
    if (data.ColorName) sheet.getRange(rowIndex, 3).setValue(sanitize(data.ColorName));
    if (data.ColorCode !== undefined) sheet.getRange(rowIndex, 4).setValue(sanitize(data.ColorCode));
    if (data.Status) sheet.getRange(rowIndex, 5).setValue(data.Status);
    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'UPDATE', module: 'FabricMaster', recordID: id, oldValues: '', newValues: data });
    return successResponse({ FabricColorID: id }, 'Color updated');
  } catch (e) {
    return errorResponse('Failed to update color: ' + e.message, 'UPDATE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// VENDOR MASTER
// ══════════════════════════════════════════════════════════════

function getVendors(token, type) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'VendorMaster');
    if (!valid) return error;
    let vendors = getSheetData('VendorMaster');
    if (type) vendors = vendors.filter(function(r) { return r.VendorType === type || r.VendorType === 'Both'; });
    return successResponse({ vendors: vendors });
  } catch (e) {
    return errorResponse('Failed to fetch vendors: ' + e.message, 'FETCH_ERROR');
  }
}

function createVendor(token, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'VendorMaster');
    if (!valid) return error;

    const missing = validateRequired(['VendorName', 'VendorType'], data);
    if (missing.length) return errorResponse('Required: ' + missing.join(', '), 'VALIDATION_ERROR');

    const existing = getSheetData('VendorMaster');
    const dup = existing.find(function(r) {
      return String(r.VendorName).toLowerCase() === String(data.VendorName).toLowerCase();
    });
    if (dup) return errorResponse('Vendor name already exists', 'DUPLICATE');

    if (data.Mobile && !isValidMobile(data.Mobile)) return errorResponse('Invalid mobile number', 'VALIDATION_ERROR');
    if (data.Email && !isValidEmail(data.Email)) return errorResponse('Invalid email address', 'VALIDATION_ERROR');
    if (data.GSTNumber && !isValidGST(data.GSTNumber)) return errorResponse('Invalid GST number format', 'VALIDATION_ERROR');

    const validTypes = ['Fabric Vendor', 'Fabricator', 'Both'];
    if (validTypes.indexOf(data.VendorType) === -1) return errorResponse('Invalid vendor type', 'VALIDATION_ERROR');

    const now = nowISO();
    const id = generateID('VND', 'VendorMaster', 1);
    const row = [
      id,
      sanitize(data.VendorName),
      sanitize(data.VendorCode || ''),
      sanitize(data.Address || ''),
      sanitize(data.GSTNumber || ''),
      sanitize(data.ContactPerson || ''),
      sanitize(data.Mobile || ''),
      sanitize(data.Email || '').toLowerCase(),
      data.VendorType,
      data.Status || 'Active',
      session.userID, now, '', ''
    ];
    appendRow('VendorMaster', row);
    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'CREATE', module: 'VendorMaster', recordID: id, oldValues: '', newValues: data });
    return successResponse({ VendorID: id }, 'Vendor created successfully');
  } catch (e) {
    return errorResponse('Failed to create vendor: ' + e.message, 'CREATE_ERROR');
  }
}

function updateVendor(token, id, data) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'VendorMaster');
    if (!valid) return error;
    if (!id) return errorResponse('VendorID required', 'VALIDATION_ERROR');
    const rowIndex = findRowByID('VendorMaster', 1, id);
    if (!rowIndex) return errorResponse('Vendor not found', 'NOT_FOUND');

    if (data.Mobile && !isValidMobile(data.Mobile)) return errorResponse('Invalid mobile number', 'VALIDATION_ERROR');
    if (data.Email && !isValidEmail(data.Email)) return errorResponse('Invalid email', 'VALIDATION_ERROR');
    if (data.GSTNumber && !isValidGST(data.GSTNumber)) return errorResponse('Invalid GST number', 'VALIDATION_ERROR');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('VendorMaster');
    const old = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];

    if (data.VendorName) sheet.getRange(rowIndex, 2).setValue(sanitize(data.VendorName));
    if (data.VendorCode !== undefined) sheet.getRange(rowIndex, 3).setValue(sanitize(data.VendorCode));
    if (data.Address !== undefined) sheet.getRange(rowIndex, 4).setValue(sanitize(data.Address));
    if (data.GSTNumber !== undefined) sheet.getRange(rowIndex, 5).setValue(sanitize(data.GSTNumber));
    if (data.ContactPerson !== undefined) sheet.getRange(rowIndex, 6).setValue(sanitize(data.ContactPerson));
    if (data.Mobile !== undefined) sheet.getRange(rowIndex, 7).setValue(sanitize(data.Mobile));
    if (data.Email !== undefined) sheet.getRange(rowIndex, 8).setValue(sanitize(data.Email).toLowerCase());
    if (data.VendorType) sheet.getRange(rowIndex, 9).setValue(data.VendorType);
    if (data.Status) sheet.getRange(rowIndex, 10).setValue(data.Status);
    sheet.getRange(rowIndex, 13).setValue(session.userID);
    sheet.getRange(rowIndex, 14).setValue(nowISO());

    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'UPDATE', module: 'VendorMaster', recordID: id, oldValues: old, newValues: data });
    return successResponse({ VendorID: id }, 'Vendor updated successfully');
  } catch (e) {
    return errorResponse('Failed to update vendor: ' + e.message, 'UPDATE_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// HOME DASHBOARD DATA
// ══════════════════════════════════════════════════════════════

function getHomeDashboard(token) {
  try {
    const { valid, session, error } = requireSession(token);
    if (!valid) return error;

    const soData = getSheetData('SalesOrders');
    const poData = getSheetData('PurchaseOrders');
    const invData = getSheetData('Inventory');
    const bundleData = getSheetData('Bundles');
    const dispatchData = getSheetData('Dispatch');
    const returnsData = getSheetData('Returns');

    const now = new Date();
    const thisMonth = now.getMonth();
    const thisYear = now.getFullYear();

    const activeSO = soData.filter(function(r) { return r.Status === 'Confirmed' || r.Status === 'In Production'; }).length;
    const pendingPO = poData.filter(function(r) { return r.Status === 'Approved' || r.Status === 'Partially Received'; }).length;
    const threshold = parseFloat(getSetting('LOW_STOCK_THRESHOLD') || '50');
    const lowStock = invData.filter(function(r) { return parseFloat(r.AvailableQty) < threshold; }).length;
    const activeBundles = bundleData.filter(function(r) { return r.CurrentProcess !== 'Dispatched'; }).length;
    const packedBundles = bundleData.filter(function(r) { return r.PackingStatus === 'Done' && r.CurrentProcess !== 'Dispatched'; }).length;
    const pendingReturns = returnsData.filter(function(r) { return r.Status === 'Open'; }).length;

    const monthDispatch = dispatchData.filter(function(r) {
      const d = new Date(r.DispatchDate || r.CreatedAt);
      return d.getMonth() === thisMonth && d.getFullYear() === thisYear && r.Status === 'Dispatched';
    });
    const totalPieces = monthDispatch.reduce(function(sum, r) { return sum + (parseInt(r.TotalPieces) || 0); }, 0);

    // Production status for dashboard
    const processes = ['Cutting', 'Stitching', 'Finishing', 'Packing'];
    const colorMap = { Cutting: 'orange', Stitching: 'blue', Finishing: 'green', Packing: 'purple' };
    const productionStatus = processes.map(function(p) {
      const statusKey = p + 'Status';
      const total = bundleData.filter(function(r) { return r.CurrentProcess === p || r[statusKey] === 'Done'; }).length;
      const done = bundleData.filter(function(r) { return r[statusKey] === 'Done'; }).length;
      return { label: p, total: Math.max(total, 1), completed: done, color: colorMap[p] };
    }).filter(function(s) { return s.total > 0; });

    return successResponse({
      totalSO: soData.length,
      activeSO: activeSO,
      totalPO: poData.length,
      pendingPO: pendingPO,
      stockItems: invData.length,
      lowStock: lowStock,
      activeBundles: activeBundles,
      packedBundles: packedBundles,
      thisMonthDispatch: monthDispatch.length,
      totalPieces: totalPieces,
      pendingReturns: pendingReturns,
      productionStatus: productionStatus
    });
  } catch (e) {
    return errorResponse('Dashboard error: ' + e.message, 'FETCH_ERROR');
  }
}

// ══════════════════════════════════════════════════════════════
// SETTINGS
// ══════════════════════════════════════════════════════════════

function getSettingsAll(token) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'Settings');
    if (!valid) return error;
    const settings = getSheetData('_Settings');
    return successResponse({ settings: settings });
  } catch (e) {
    return errorResponse('Failed to load settings: ' + e.message, 'FETCH_ERROR');
  }
}

function updateSettingValue(token, key, value) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'Settings');
    if (!valid) return error;
    if (!key) return errorResponse('Setting key required', 'VALIDATION_ERROR');
    setSetting(key, value, session.userID);
    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'UPDATE', module: 'Settings', recordID: key, oldValues: '', newValues: { key: key, value: value } });
    return successResponse({}, 'Setting updated');
  } catch (e) {
    return errorResponse('Failed to update setting: ' + e.message, 'UPDATE_ERROR');
  }
}