// ============================================================
// Code.gs — Main Entry Point & Router
// FIXED: B-020 XFrameOptionsMode changed to SAMEORIGIN
//        B-011 getDropdowns() excludes Fully Dispatched SOs
//        Added serverClosePO route for Phase 4 under-receive workflow
// ============================================================

function doGet(e) {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('Garment ERP')
    .setFaviconUrl('https://www.google.com/images/favicon.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    // FIX B-020: SAMEORIGIN prevents clickjacking
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.SAMEORIGIN);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Auth Routes ────────────────────────────────────────────
function serverLogin(email, password)               { return login(email, password); }
function serverLogout(token)                         { return logout(token); }
function serverGetCurrentUser(token)                 { return getCurrentUser(token); }
function serverChangePassword(token, cur, np)        { return changePassword(token, cur, np); }

// ── User Management Routes ─────────────────────────────────
function serverGetUsers(token)                       { return getUsers(token); }
function serverCreateUser(token, data)               { return createUser(token, data); }
function serverUpdateUser(token, id, data)           { return updateUser(token, id, data); }
function serverResetPassword(token, id, np)          { return resetUserPassword(token, id, np); }
function serverGetRoles()                            { return getRoles(); }

// ── Master Routes ──────────────────────────────────────────
function serverGetBuyers(token)                      { return getBuyers(token); }
function serverCreateBuyer(token, data)              { return createBuyer(token, data); }
function serverUpdateBuyer(token, id, data)          { return updateBuyer(token, id, data); }

function serverGetArticles(token, filters)           { return getArticles(token, filters); }
function serverCreateArticle(token, data)            { return createArticle(token, data); }
function serverUpdateArticle(token, id, data)        { return updateArticle(token, id, data); }

function serverGetFabrics(token)                     { return getFabrics(token); }
function serverCreateFabric(token, data)             { return createFabric(token, data); }
function serverUpdateFabric(token, id, data)         { return updateFabric(token, id, data); }
function serverGetFabricColors(token, fabricID)      { return getFabricColors(token, fabricID); }
function serverAddFabricColor(token, data)           { return addFabricColor(token, data); }
function serverUpdateFabricColor(token, id, data)    { return updateFabricColor(token, id, data); }

function serverGetVendors(token, type)               { return getVendors(token, type); }
function serverCreateVendor(token, data)             { return createVendor(token, data); }
function serverUpdateVendor(token, id, data)         { return updateVendor(token, id, data); }

// ── BOM Routes ─────────────────────────────────────────────
function serverGetBOM(token, articleID)              { return getBOM(token, articleID); }
function serverSaveBOM(token, articleID, items)      { return saveBOM(token, articleID, items); }
function serverDeleteBOMItem(token, bomID)           { return deleteBOMItem(token, bomID); }

// ── Sales Order Routes ─────────────────────────────────────
function serverGetSalesOrders(token, filters)        { return getSalesOrders(token, filters); }
function serverGetSalesOrderByID(token, id)          { return getSalesOrderByID(token, id); }
function serverCreateSalesOrder(token, data)         { return createSalesOrder(token, data); }
function serverUpdateSalesOrder(token, id, data)     { return updateSalesOrder(token, id, data); }
function serverUpdateSOStatus(token, id, status)     { return updateSOStatus(token, id, status); }

// ── Purchase Order Routes ──────────────────────────────────
function serverGetPurchaseOrders(token, filters)     { return getPurchaseOrders(token, filters); }
function serverGetPOByID(token, id)                  { return getPOByID(token, id); }
function serverCreatePO(token, data)                 { return createPO(token, data); }
function serverUpdatePO(token, id, data)             { return updatePO(token, id, data); }
function serverApprovePO(token, id)                  { return approvePO(token, id); }
function serverCancelPO(token, id, reason)           { return cancelPO(token, id, reason); }
// Phase 4 — Under Receive: close PO at user's choice
function serverClosePO(token, id)                    { return closePO(token, id); }

// ── GRN Routes ─────────────────────────────────────────────
function serverGetGRNs(token, filters)               { return getGRNs(token, filters); }
function serverGetGRNByID(token, id)                 { return getGRNByID(token, id); }
function serverCreateGRN(token, data)                { return createGRN(token, data); }
function serverGetPOItemsForGRN(token, poID)         { return getPOItemsForGRN(token, poID); }

// ── Inventory Routes ───────────────────────────────────────
function serverGetInventory(token, filters)          { return getInventory(token, filters); }
function serverGetInventoryLedger(token, filters)    { return getInventoryLedger(token, filters); }
function serverGetAvailableStock(token, fid, cid)    { return getAvailableStock(token, fid, cid); }
function serverStockAdjustment(token, data)          { return stockAdjustment(token, data); }

// ── Fabric Issue Routes ────────────────────────────────────
function serverGetFabricIssues(token, filters)       { return getFabricIssues(token, filters); }
function serverGetFabricIssueByID(token, id)         { return getFabricIssueByID(token, id); }
function serverCreateFabricIssue(token, data)        { return createFabricIssue(token, data); }
function serverGetBOMSuggestion(token, aid, sid)     { return getBOMSuggestion(token, aid, sid); }

// ── Production Plan Routes ─────────────────────────────────
function serverGetProductionPlans(token, filters)          { return getProductionPlans(token, filters); }
function serverGetProductionPlanByID(token, id)            { return getProductionPlanByID(token, id); }
function serverCreateProductionPlan(token, data)           { return createProductionPlan(token, data); }
function serverUpdateProductionPlan(token, id, data)       { return updateProductionPlan(token, id, data); }
function serverUpdatePlanStatus(token, id, status)         { return updatePlanStatus(token, id, status); }
function serverGenerateBundles(token, planID)              { return generateBundles(token, planID); }
function serverDeletePlanBundles(token, planID)            { return deletePlanBundles(token, planID); }
function serverGetSOArticleBreakdown(token, soID)          { return getSOArticleBreakdown(token, soID); }
function serverGetBundleProcessSummary(token, planID)      { return getBundleProcessSummary(token, planID); }
function serverGetBundleLabelData(token, planID)           { return getBundleLabelData(token, planID); }

// ── Bundle Routes ──────────────────────────────────────────
function serverGetBundles(token, filters)                  { return getBundles(token, filters); }
function serverGetBundleByID(token, id)                    { return getBundleByID(token, id); }
function serverGetBundleByBarcode(token, barcode)          { return getBundleByBarcode(token, barcode); }
function serverUpdateBundleProcess(token, id, p, s, pl)   { return updateBundleProcess(token, id, p, s, pl); }
function serverBulkUpdateBundleProcess(token, ids, p, s, pl) { return bulkUpdateBundleProcess(token, ids, p, s, pl); }
function serverGetReadyBundles(token, soID)                { return getReadyBundles(token, soID); }

// ── Fabricator Challan Routes ──────────────────────────────
function serverGetFabricatorChallans(token, filters)       { return getFabricatorChallans(token, filters); }
function serverGetFabricatorChallanByID(token, id)         { return getFabricatorChallanByID(token, id); }
function serverCreateFabricatorChallan(token, data)        { return createFabricatorChallan(token, data); }
function serverReceiveFabricatorChallan(token, id, data)   { return receiveFabricatorChallan(token, id, data); }
function serverGetFabricatorSummary(token)                 { return getFabricatorSummary(token); }

// ── Dispatch Routes ────────────────────────────────────────
function serverGetDispatches(token, filters)               { return getDispatches(token, filters); }
function serverGetDispatchByID(token, id)                  { return getDispatchByID(token, id); }
function serverCreateDispatch(token, data)                 { return createDispatch(token, data); }
function serverGetReadyBundlesForDispatch(token, soID)     { return getReadyBundlesForDispatch(token, soID); }
function serverGetDispatchSummary(token)                   { return getDispatchSummary(token); }

// ── Returns Routes ─────────────────────────────────────────
function serverGetReturns(token, filters)                  { return getReturns(token, filters); }
function serverGetReturnByID(token, id)                    { return getReturnByID(token, id); }
function serverCreateReturn(token, data)                   { return createReturn(token, data); }
function serverUpdateReturnStatus(token, id, status)       { return updateReturnStatus(token, id, status); }

// ── Report Routes ──────────────────────────────────────────
function serverGetPurchaseDashboard(token)                 { return getPurchaseDashboard(token); }
function serverGetInventoryDashboard(token)                { return getInventoryDashboard(token); }
function serverGetProductionDashboard(token)               { return getProductionDashboard(token); }
function serverGetDispatchDashboard(token)                 { return getDispatchDashboard(token); }
function serverGetSalesReport(token, filters)              { return getSalesReport(token, filters); }
function serverGetHomeDashboard(token)                     { return getHomeDashboard(token); }

// ── Settings Routes ────────────────────────────────────────
function serverGetSettings(token)                          { return getSettingsAll(token); }
function serverUpdateSetting(token, key, value)            { return updateSettingValue(token, key, value); }

// ── Audit Routes ───────────────────────────────────────────
function serverGetAuditLogs(token, filters)                { return getAuditLogs(token, filters); }

// ── Lookup Routes (dropdown data) ─────────────────────────
function serverGetDropdowns(token) { return getDropdowns(token); }

function getDropdowns(token) {
  try {
    const { valid, session, error } = requireSession(token);
    if (!valid) return error;

    const buyers = getSheetData('BuyerMaster')
      .filter(function(r) { return r.Status === 'Active'; })
      .map(function(r) { return { id: r.BuyerID, name: r.BuyerName, code: r.BuyerCode }; });

    const articles = getSheetData('ArticleMaster')
      .filter(function(r) { return r.Status === 'Active'; })
      .map(function(r) { return { id: r.ArticleID, name: r.ArticleName, code: r.ArticleCode, buyerID: r.BuyerID }; });

    const fabrics = getSheetData('FabricMaster')
      .filter(function(r) { return r.Status === 'Active'; })
      .map(function(r) { return { id: r.FabricID, name: r.DesignName, code: r.DesignCode }; });

    const fabricColors = getSheetData('FabricColors')
      .filter(function(r) { return r.Status === 'Active'; })
      .map(function(r) { return { id: r.FabricColorID, fabricID: r.FabricID, name: r.ColorName, code: r.ColorCode }; });

    const vendors = getSheetData('VendorMaster')
      .filter(function(r) { return r.Status === 'Active'; })
      .map(function(r) { return { id: r.VendorID, name: r.VendorName, code: r.VendorCode, type: r.VendorType }; });

    // FIX B-011: Exclude Cancelled, Closed, AND Dispatched SOs from active dropdowns
    const INACTIVE_SO_STATUSES = ['Cancelled', 'Closed', 'Dispatched'];
    const soRows    = getSheetData('SalesOrders').filter(function(r) {
      return INACTIVE_SO_STATUSES.indexOf(r.Status) === -1;
    });
    const buyerMap2 = {};
    getSheetData('BuyerMaster').forEach(function(r) { buyerMap2[r.BuyerID] = r.BuyerName; });
    const salesOrders = soRows.map(function(r) {
      return {
        value: r.SOID, label: (r.SONumber || r.SOID) + ' — ' + (buyerMap2[r.BuyerID] || ''),
        id: r.SOID, buyerID: r.BuyerID, status: r.Status, deliveryDate: r.DeliveryDate
      };
    });

    return successResponse({
      buyers, articles, fabrics, fabricColors, vendors, salesOrders,
      sizes: ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'Custom'],
      uoms: ['Meters', 'Yards'],
      fabricCategories: ['Cotton', 'Polyester', 'Blend', 'Linen', 'Silk', 'Denim', 'Knit', 'Other'],
      articleCategories: ['Formal Shirt', 'Casual Shirt', 'Kids Shirt', 'Ladies Shirt', 'T-Shirt', 'Other'],
      vendorTypes: ['Fabric Vendor', 'Fabricator', 'Both']
    });
  } catch (e) {
    return errorResponse('Failed to load dropdowns: ' + e.message, 'FETCH_ERROR');
  }
}
