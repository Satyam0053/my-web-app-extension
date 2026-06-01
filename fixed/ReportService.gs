// ============================================================
// ReportService.gs — Reports & Dashboards
// FIXED: B-012 recentGRNs sorted by GRNDate before slicing
//        Performance: each function reads sheets once into local vars
// ============================================================

function getPurchaseDashboard(token) {
  var auth = requireReadAccess(token, 'Reports');
  if (!auth.valid) return auth.error;
  try {
    var pos     = getSheetData('PurchaseOrders');
    var grns    = getSheetData('GRN');
    var vendors = getSheetData('VendorMaster');

    var vendorMap = {};
    vendors.forEach(function(v) { vendorMap[v.VendorID] = v.VendorName; });

    var statusCounts = {};
    pos.forEach(function(p) { statusCounts[p.Status] = (statusCounts[p.Status] || 0) + 1; });

    var totalValue  = pos.reduce(function(s, p) { return s + (Number(p.TotalAmount) || 0); }, 0);
    var pendingRecv = pos.filter(function(p) { return p.Status === 'Approved' || p.Status === 'Partially Received'; }).length;

    var byVendor = {};
    pos.forEach(function(p) {
      var name = vendorMap[p.VendorID] || p.VendorID;
      byVendor[name] = (byVendor[name] || 0) + 1;
    });

    // FIX B-012: Sort by GRNDate descending before taking last 10
    var sortedGRNs = grns.slice().sort(function(a, b) {
      return String(b.GRNDate).localeCompare(String(a.GRNDate));
    });
    var recentGRNs = sortedGRNs.slice(0, 10).map(function(g) {
      return { GRNID: g.GRNID, GRNDate: g.GRNDate, VendorName: vendorMap[g.VendorID] || g.VendorID };
    });

    return successResponse({ totalPOs: pos.length, totalValue: totalValue, pendingRecv: pendingRecv, statusCounts: statusCounts, byVendor: byVendor, recentGRNs: recentGRNs });
  } catch (e) {
    return errorResponse('Failed to get purchase dashboard: ' + e.message);
  }
}

function getInventoryDashboard(token) {
  var auth = requireReadAccess(token, 'Reports');
  if (!auth.valid) return auth.error;
  try {
    var inventory = getSheetData('Inventory');
    var threshold = Number(getSetting('LOW_STOCK_THRESHOLD')) || 50;

    var fabricMap = {}; getSheetData('FabricMaster').forEach(function(f) { fabricMap[f.FabricID] = f.DesignName; });
    var colorMap  = {}; getSheetData('FabricColors').forEach(function(c) { colorMap[c.FabricColorID] = c.ColorName; });

    var totalItems    = inventory.length;
    var lowStock      = inventory.filter(function(i) { return Number(i.AvailableQty) < threshold; });
    var zeroStock     = inventory.filter(function(i) { return Number(i.AvailableQty) <= 0; });
    var totalReceived = inventory.reduce(function(s, i) { return s + (Number(i.ReceivedQty) || 0); }, 0);
    var totalIssued   = inventory.reduce(function(s, i) { return s + (Number(i.IssuedQty)   || 0); }, 0);

    var enrichedInv = inventory.map(function(i) {
      return {
        FabricName:   fabricMap[i.FabricID]     || i.FabricID,
        ColorName:    colorMap[i.FabricColorID] || i.FabricColorID,
        AvailableQty: Number(i.AvailableQty) || 0,
        ReceivedQty:  Number(i.ReceivedQty)  || 0,
        IssuedQty:    Number(i.IssuedQty)    || 0,
        IsLow:        Number(i.AvailableQty) < threshold
      };
    }).sort(function(a, b) { return a.AvailableQty - b.AvailableQty; });

    return successResponse({ totalItems: totalItems, lowStockCount: lowStock.length, zeroStockCount: zeroStock.length, totalReceived: totalReceived, totalIssued: totalIssued, inventory: enrichedInv });
  } catch (e) {
    return errorResponse('Failed to get inventory dashboard: ' + e.message);
  }
}

function getProductionDashboard(token) {
  var auth = requireReadAccess(token, 'Reports');
  if (!auth.valid) return auth.error;
  try {
    var plans   = getSheetData('ProductionPlan');
    var bundles = getSheetData('Bundles');

    var active = plans.filter(function(p) { return p.Status === 'In Progress'; });
    var done   = plans.filter(function(p) { return p.Status === 'Completed'; });

    var processCounts = {
      'Pre-Cutting': 0, 'Cutting': 0, 'Pre-Stitching': 0, 'Stitching': 0,
      'Pre-Finishing': 0, 'Finishing': 0, 'Pre-Packing': 0, 'Packing': 0, 'Completed': 0
    };
    bundles.forEach(function(b) {
      var stage = b.CurrentProcess || 'Pre-Cutting';
      if (processCounts[stage] !== undefined) processCounts[stage]++;
    });

    var articleMap = {}; getSheetData('ArticleMaster').forEach(function(a) { articleMap[a.ArticleID] = a.ArticleName; });
    var soMap      = {}; getSheetData('SalesOrders').forEach(function(s)   { soMap[s.SOID] = s.SONumber || s.SOID; });

    var planProgress = active.map(function(p) {
      var planBundles = bundles.filter(function(b) { return b.PlanID === p.PlanID; });
      var packed      = planBundles.filter(function(b) { return b.PackingStatus === 'Done'; }).length;
      return {
        PlanID:       p.PlanID,
        SONumber:     soMap[p.SOID]          || p.SOID,
        ArticleName:  articleMap[p.ArticleID] || p.ArticleID,
        TotalBundles: planBundles.length,
        Packed:       packed,
        PctDone:      planBundles.length ? Math.round((packed / planBundles.length) * 100) : 0
      };
    });

    return successResponse({ totalPlans: plans.length, activePlans: active.length, completedPlans: done.length, totalBundles: bundles.length, processCounts: processCounts, planProgress: planProgress });
  } catch (e) {
    return errorResponse('Failed to get production dashboard: ' + e.message);
  }
}

function getDispatchDashboard(token) {
  var auth = requireReadAccess(token, 'Reports');
  if (!auth.valid) return auth.error;
  try {
    var dispatches = getSheetData('Dispatch');
    var soMap      = {}; getSheetData('SalesOrders').forEach(function(s) { soMap[s.SOID]   = s.SONumber || s.SOID; });
    var buyerMap   = {}; getSheetData('BuyerMaster').forEach(function(b) { buyerMap[b.BuyerID] = b.BuyerName; });

    var thisMonth  = new Date().toISOString().slice(0, 7);
    var monthDisp  = dispatches.filter(function(d) { return String(d.DispatchDate).slice(0, 7) === thisMonth; });
    var totalPcs   = dispatches.reduce(function(s, d) { return s + (Number(d.TotalPieces) || 0); }, 0);
    var monthPcs   = monthDisp.reduce(function(s, d)  { return s + (Number(d.TotalPieces) || 0); }, 0);

    var dispatchedIDs = {};
    getSheetData('DispatchItems').forEach(function(d) { if (d.BundleID) dispatchedIDs[d.BundleID] = true; });
    var pendingBundles = getSheetData('Bundles').filter(function(b) {
      return b.PackingStatus === 'Done' && !dispatchedIDs[b.BundleID];
    }).length;

    var byBuyer = {};
    dispatches.forEach(function(d) {
      var name = buyerMap[d.BuyerID] || d.BuyerID;
      byBuyer[name] = (byBuyer[name] || 0) + (Number(d.TotalPieces) || 0);
    });

    // FIX B-012 pattern: sort by DispatchDate before taking recent
    var recentDisp = dispatches.slice().sort(function(a, b) {
      return String(b.DispatchDate).localeCompare(String(a.DispatchDate));
    }).slice(0, 10).map(function(d) {
      return { DispatchID: d.DispatchID, DispatchDate: d.DispatchDate, SONumber: soMap[d.SOID] || d.SOID, BuyerName: buyerMap[d.BuyerID] || d.BuyerID, TotalPieces: d.TotalPieces };
    });

    return successResponse({ totalDispatches: dispatches.length, totalPieces: totalPcs, thisMonthCount: monthDisp.length, thisMonthPieces: monthPcs, pendingBundles: pendingBundles, byBuyer: byBuyer, recentDispatches: recentDisp });
  } catch (e) {
    return errorResponse('Failed to get dispatch dashboard: ' + e.message);
  }
}

function getSalesReport(token, filters) {
  var auth = requireReadAccess(token, 'Reports');
  if (!auth.valid) return auth.error;
  try {
    filters = filters || {};
    var orders = getSheetData('SalesOrders');
    var items  = getSheetData('SalesOrderItems');

    if (filters.from) orders = orders.filter(function(o) { return o.SODate >= filters.from; });
    if (filters.to)   orders = orders.filter(function(o) { return o.SODate <= filters.to; });

    var buyerMap   = {}; getSheetData('BuyerMaster').forEach(function(b)   { buyerMap[b.BuyerID]     = b.BuyerName; });
    var articleMap = {}; getSheetData('ArticleMaster').forEach(function(a) { articleMap[a.ArticleID] = a.ArticleName; });

    var soIDs         = {};
    orders.forEach(function(o) { soIDs[o.SOID] = true; });
    var filteredItems = items.filter(function(i) { return soIDs[i.SOID]; });

    var byBuyer = {};
    orders.forEach(function(o) {
      var name = buyerMap[o.BuyerID] || o.BuyerID;
      if (!byBuyer[name]) byBuyer[name] = { orders: 0, pieces: 0 };
      byBuyer[name].orders++;
      byBuyer[name].pieces += Number(o.TotalPieces) || 0;
    });

    var byArticle = {};
    filteredItems.forEach(function(i) {
      var name = articleMap[i.ArticleID] || i.ArticleID;
      byArticle[name] = (byArticle[name] || 0) + (Number(i.OrderedQty) || 0);
    });

    var statusCounts = {};
    orders.forEach(function(o) { statusCounts[o.Status] = (statusCounts[o.Status] || 0) + 1; });

    return successResponse({
      totalOrders:  orders.length,
      totalPieces:  orders.reduce(function(s, o) { return s + (Number(o.TotalPieces) || 0); }, 0),
      statusCounts: statusCounts,
      byBuyer:      byBuyer,
      byArticle:    byArticle
    });
  } catch (e) {
    return errorResponse('Failed to get sales report: ' + e.message);
  }
}
