// ============================================================
// DispatchService.gs — Dispatch Management (Sprint S9)
// ============================================================

// ── Get all dispatches ───────────────────────────────────────
function getDispatches(token, filters) {
  var auth = requireReadAccess(token, 'Dispatch');
  if (!auth.valid) return auth.error;
  try {
    filters = filters || {};
    var rows = getSheetData('Dispatch');

    if (filters.soID)   rows = rows.filter(function(r) { return r.SOID    === filters.soID; });
    if (filters.status) rows = rows.filter(function(r) { return r.Status  === filters.status; });

    var buyerMap = {};
    getSheetData('BuyerMaster').forEach(function(b) { buyerMap[b.BuyerID] = b.BuyerName; });
    var soMap = {};
    getSheetData('SalesOrders').forEach(function(s) { soMap[s.SOID] = s.SONumber || s.SOID; });

    var enriched = rows.map(function(r) {
      return Object.assign({}, r, {
        BuyerName: buyerMap[r.BuyerID] || r.BuyerID,
        SONumber:  soMap[r.SOID]       || r.SOID
      });
    });

    enriched.sort(function(a, b) { return String(b.DispatchID).localeCompare(String(a.DispatchID)); });
    return successResponse(enriched);
  } catch (e) {
    return errorResponse('Failed to get dispatches: ' + e.message);
  }
}

// ── Get single dispatch with items ───────────────────────────
function getDispatchByID(token, dispatchID) {
  var auth = requireReadAccess(token, 'Dispatch');
  if (!auth.valid) return auth.error;
  try {
    var dispatch = findRowByID('Dispatch', 'DispatchID', dispatchID);
    if (!dispatch) return errorResponse('Dispatch not found');

    var buyerMap   = {}; getSheetData('BuyerMaster').forEach(function(b) { buyerMap[b.BuyerID] = b.BuyerName; });
    var soMap      = {}; getSheetData('SalesOrders').forEach(function(s) { soMap[s.SOID] = s.SONumber || s.SOID; });
    var articleMap = {}; getSheetData('ArticleMaster').forEach(function(a) { articleMap[a.ArticleID] = a.ArticleName; });

    dispatch.BuyerName = buyerMap[dispatch.BuyerID] || dispatch.BuyerID;
    dispatch.SONumber  = soMap[dispatch.SOID]       || dispatch.SOID;

    var items = getSheetData('DispatchItems')
      .filter(function(i) { return i.DispatchID === dispatchID; })
      .map(function(i) {
        return Object.assign({}, i, { ArticleName: articleMap[i.ArticleID] || i.ArticleID });
      });

    dispatch.Items      = items;
    dispatch.TotalItems = items.length;
    return successResponse(dispatch);
  } catch (e) {
    return errorResponse('Failed to get dispatch: ' + e.message);
  }
}

// ── Get bundles ready for dispatch ───────────────────────────
function getReadyBundlesForDispatch(token, soID) {
  var auth = requireReadAccess(token, 'Dispatch');
  if (!auth.valid) return auth.error;
  try {
    if (!soID) return errorResponse('Sales Order is required');

    // All packing-done bundles for this SO
    var bundles = getSheetData('Bundles').filter(function(b) {
      return b.SOID === soID && b.PackingStatus === 'Done';
    });

    // Exclude already dispatched bundle IDs
    var dispatchedIDs = new Set();
    getSheetData('DispatchItems').forEach(function(d) {
      if (d.BundleID) dispatchedIDs.add(d.BundleID);
    });

    var available = bundles.filter(function(b) { return !dispatchedIDs.has(b.BundleID); });

    var articleMap = {};
    getSheetData('ArticleMaster').forEach(function(a) { articleMap[a.ArticleID] = a.ArticleName; });

    var enriched = available.map(function(b) {
      return Object.assign({}, b, { ArticleName: articleMap[b.ArticleID] || b.ArticleID });
    });

    return successResponse(enriched);
  } catch (e) {
    return errorResponse('Failed to get ready bundles: ' + e.message);
  }
}

// ── Create Dispatch ──────────────────────────────────────────
function createDispatch(token, data) {
  var auth = requireWriteAccess(token, 'Dispatch');
  if (!auth.valid) return auth.error;
  try {
    if (!data.SOID)         return errorResponse('Sales Order is required');
    if (!data.DispatchDate) return errorResponse('Dispatch date is required');
    if (!Array.isArray(data.bundleIDs) || data.bundleIDs.length === 0) {
      return errorResponse('Select at least one bundle to dispatch');
    }

    var so = findRowByID('SalesOrders', 'SOID', data.SOID);
    if (!so) return errorResponse('Sales Order not found');
    if (!['Confirmed','In Production'].includes(so.Status)) {
      return errorResponse('Sales Order must be Confirmed or In Production to dispatch');
    }

    // Validate all bundles are packing done and not already dispatched
    var dispatchedIDs = new Set();
    getSheetData('DispatchItems').forEach(function(d) { if (d.BundleID) dispatchedIDs.add(d.BundleID); });

    var bundles = [];
    for (var i = 0; i < data.bundleIDs.length; i++) {
      var bid    = data.bundleIDs[i];
      var bundle = findRowByID('Bundles', 'BundleID', bid);
      if (!bundle)                          return errorResponse('Bundle not found: ' + bid);
      if (bundle.PackingStatus !== 'Done')  return errorResponse('Bundle ' + bid + ' is not packing-done');
      if (dispatchedIDs.has(bid))           return errorResponse('Bundle ' + bid + ' already dispatched');
      bundles.push(bundle);
    }

    var now        = nowISO();
    var dispatchID = generateYearID('DSP', 'Dispatch', 1);
    var dispNumber = dispatchID;
    var totalPcs   = bundles.reduce(function(s, b) { return s + (Number(b.Pieces) || 0); }, 0);

    appendRow('Dispatch', {
      DispatchID:     dispatchID,
      DispatchNumber: dispNumber,
      DispatchDate:   data.DispatchDate,
      SOID:           data.SOID,
      BuyerID:        so.BuyerID,
      VehicleNumber:  sanitizeString(data.VehicleNumber || ''),
      DriverName:     sanitizeString(data.DriverName    || ''),
      DriverMobile:   sanitizeString(data.DriverMobile  || ''),
      TotalPieces:    totalPcs,
      Remarks:        sanitizeString(data.Remarks       || ''),
      Status:         'Dispatched',
      CreatedBy:      auth.session.userID,
      CreatedAt:      now
    });

    // Write dispatch items
    bundles.forEach(function(b) {
      appendRow('DispatchItems', {
        DispatchItemID: generateID('DSPI'),
        DispatchID:     dispatchID,
        BundleID:       b.BundleID,
        ArticleID:      b.ArticleID,
        Size:           b.Size,
        Color:          b.Color,
        Qty:            Number(b.Pieces) || 0
      });
    });

    // Update SO status to Dispatched
    updateRowByID('SalesOrders', 'SOID', data.SOID, {
      Status:    'Dispatched',
      UpdatedAt: now
    });

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'CREATE', module: 'Dispatch', recordID: dispatchID, oldValues: '', newValues: { soID: data.SOID, bundles: data.bundleIDs.length, totalPcs: totalPcs } });
    return successResponse({ dispatchID: dispatchID, totalPieces: totalPcs }, 'Dispatched successfully: ' + dispatchID);
  } catch (e) {
    return errorResponse('Failed to create dispatch: ' + e.message);
  }
}

// ── Get dispatch summary for dashboard ───────────────────────
function getDispatchSummary(token) {
  var auth = requireReadAccess(token, 'Dispatch');
  if (!auth.valid) return auth.error;
  try {
    var dispatches = getSheetData('Dispatch');
    var thisMonth  = new Date().toISOString().slice(0, 7); // YYYY-MM
    var monthDisp  = dispatches.filter(function(d) { return String(d.DispatchDate).slice(0, 7) === thisMonth; });
    var totalPcs   = monthDisp.reduce(function(s, d) { return s + (Number(d.TotalPieces) || 0); }, 0);

    // Pending (SOs confirmed/in-production with packing-done bundles not dispatched)
    var dispatchedBundleIDs = new Set();
    getSheetData('DispatchItems').forEach(function(d) { if (d.BundleID) dispatchedBundleIDs.add(d.BundleID); });
    var pendingBundles = getSheetData('Bundles').filter(function(b) {
      return b.PackingStatus === 'Done' && !dispatchedBundleIDs.has(b.BundleID);
    }).length;

    return successResponse({
      totalDispatches: dispatches.length,
      thisMonthCount:  monthDisp.length,
      thisMonthPieces: totalPcs,
      pendingBundles:  pendingBundles
    });
  } catch (e) {
    return errorResponse('Failed to get dispatch summary: ' + e.message);
  }
}