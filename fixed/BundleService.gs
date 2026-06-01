// ============================================================
// BundleService.gs — Bundle Process Tracking
// Sprint S6
// ============================================================

// ── Process list in order ────────────────────────────────────
const PROCESSES = ['Cutting', 'Stitching', 'Finishing', 'Packing'];
const PROCESS_STATUS_FIELD = {
  'Cutting':   'CuttingStatus',
  'Stitching': 'StitchingStatus',
  'Finishing': 'FinishingStatus',
  'Packing':   'PackingStatus'
};
const PROCESS_OPERATOR_FIELD = {
  'Cutting':   'CuttingOperator',
  'Stitching': 'StitchingOperator',
  'Finishing': 'FinishingOperator',
  'Packing':   'PackingOperator'
};
const PROCESS_DATE_FIELD = {
  'Cutting':   'CuttingDate',
  'Stitching': 'StitchingDate',
  'Finishing': 'FinishingDate',
  'Packing':   'PackingDate'
};

// ── Get bundles with filters ─────────────────────────────────
function getBundles(token, filters) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    filters = filters || {};
    let rows = getSheetData('Bundles');

    if (filters.planID)    rows = rows.filter(r => r.PlanID    === filters.planID);
    if (filters.soID)      rows = rows.filter(r => r.SOID      === filters.soID);
    if (filters.articleID) rows = rows.filter(r => r.ArticleID === filters.articleID);
    if (filters.size)      rows = rows.filter(r => r.Size      === filters.size);
    if (filters.process && filters.processStatus) {
      const field = PROCESS_STATUS_FIELD[filters.process];
      if (field) rows = rows.filter(r => r[field] === filters.processStatus);
    }
    if (filters.currentProcess) rows = rows.filter(r => r.CurrentProcess === filters.currentProcess);
    if (filters.packingStatus)  rows = rows.filter(r => r.PackingStatus  === filters.packingStatus);
    if (filters.fabricatorOut)  rows = rows.filter(r => r.FabricatorOut  === filters.fabricatorOut);

    // Enrich with article + plan info
    const planMap    = {};
    const articleMap = {};
    getSheetData('ProductionPlan').forEach(p => { planMap[p.PlanID] = p; });
    getSheetData('ArticleMaster').forEach(a => { articleMap[a.ArticleID] = a; });

    const enriched = rows.map(r => ({
      ...r,
      ArticleName: (articleMap[r.ArticleID] || {}).ArticleName || '',
      StyleCode:   (articleMap[r.ArticleID] || {}).StyleCode   || '',
      PlanStatus:  (planMap[r.PlanID]       || {}).Status       || ''
    }));

    return successResponse(enriched);
  } catch (e) {
    return errorResponse('Failed to get bundles: ' + e.message);
  }
}

// ── Get single bundle by ID ───────────────────────────────────
function getBundleByID(token, bundleID) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    const bundle = findRowByID('Bundles', 'BundleID', bundleID);
    if (!bundle) return errorResponse('Bundle not found');

    // Enrich
    const plan    = findRowByID('ProductionPlan', 'PlanID',     bundle.PlanID)    || {};
    const article = findRowByID('ArticleMaster',  'ArticleID',  bundle.ArticleID) || {};
    const so      = findRowByID('SalesOrders',     'SOID',       bundle.SOID)      || {};
    const buyer   = findRowByID('BuyerMaster',     'BuyerID',    so.BuyerID)       || {};

    bundle.ArticleName = article.ArticleName || '';
    bundle.StyleCode   = article.StyleCode   || '';
    bundle.SONumber    = so.SONumber         || '';
    bundle.BuyerName   = buyer.BuyerName     || '';
    bundle.PlanStatus  = plan.Status         || '';

    // Get tracking history
    const tracking = getSheetData('BundleTracking')
      .filter(t => t.BundleID === bundleID)
      .sort((a, b) => new Date(a.TrackedAt) - new Date(b.TrackedAt));
    bundle.Tracking = tracking;

    return successResponse(bundle);
  } catch (e) {
    return errorResponse('Failed to get bundle: ' + e.message);
  }
}

// ── Scan / lookup bundle by barcode value ─────────────────────
function getBundleByBarcode(token, barcodeValue) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    const bundles = getSheetData('Bundles').filter(b => b.BarcodeValue === barcodeValue);
    if (bundles.length === 0) return errorResponse('Bundle not found for barcode: ' + barcodeValue);

    return getBundleByID(token, bundles[0].BundleID);
  } catch (e) {
    return errorResponse('Barcode lookup failed: ' + e.message);
  }
}

// ── Update bundle process status ─────────────────────────────
// process: 'Cutting' | 'Stitching' | 'Finishing' | 'Packing'
// status:  'In Progress' | 'Done'
function updateBundleProcess(token, bundleID, process, status, payload) {
  const { valid: _wv, session, error: _we } = requireWriteAccess(token, 'production');
  if (!_wv) return _we;

  try {
    if (!PROCESSES.includes(process)) {
      return errorResponse('Invalid process. Must be one of: ' + PROCESSES.join(', '));
    }
    if (!['In Progress', 'Done'].includes(status)) {
      return errorResponse('Invalid status. Must be In Progress or Done');
    }

    const bundle = findRowByID('Bundles', 'BundleID', bundleID);
    if (!bundle) return errorResponse('Bundle not found');

    // Validate plan is active
    const plan = findRowByID('ProductionPlan', 'PlanID', bundle.PlanID);
    if (!plan || !['Confirmed', 'In Progress'].includes(plan.Status)) {
      return errorResponse('Production plan is not active');
    }

    // Enforce process order — previous must be Done
    const procIdx = PROCESSES.indexOf(process);
    if (procIdx > 0) {
      const prevProcess = PROCESSES[procIdx - 1];
      const prevStatus  = bundle[PROCESS_STATUS_FIELD[prevProcess]];
      if (prevStatus !== 'Done') {
        return errorResponse(`${prevProcess} must be completed before ${process} can be updated`);
      }
    }

    // Prevent going backwards: if already Done, cannot set In Progress
    const currentStatus = bundle[PROCESS_STATUS_FIELD[process]];
    if (currentStatus === 'Done' && status === 'In Progress') {
      return errorResponse(`${process} is already marked Done`);
    }

    payload = payload || {};
    const before = { ...bundle };
    const now    = nowISO();

    // Update status field
    bundle[PROCESS_STATUS_FIELD[process]] = status;
    bundle[PROCESS_DATE_FIELD[process]]   = now;
    if (payload.operator) {
      bundle[PROCESS_OPERATOR_FIELD[process]] = sanitizeString(payload.operator);
    }

    // Update CurrentProcess
    bundle.CurrentProcess = _computeCurrentProcess(bundle);
    bundle.UpdatedAt = now;

    updateRowByID('Bundles', 'BundleID', bundleID, bundle);

    // Write tracking entry
    const trackID = generateID('TRK');
    const tracking = {
      TrackID:    trackID,
      BundleID:   bundleID,
      PlanID:     bundle.PlanID,
      Process:    process,
      Status:     status,
      Operator:   sanitizeString(payload.operator || session.userName),
      Notes:      sanitizeString(payload.notes || ''),
      TrackedBy:  session.userName,
      TrackedAt:  now
    };
    appendRow('BundleTracking', tracking);

    // Update plan status to In Progress if first bundle starts cutting
    if (plan.Status === 'Confirmed' && process === 'Cutting' && status === 'In Progress') {
      plan.Status    = 'In Progress';
      plan.UpdatedAt = now;
      updateRowByID('ProductionPlan', 'PlanID', bundle.PlanID, plan);
    }

    // Check if all bundles for plan are Packing Done → auto-complete plan
    _checkPlanCompletion(bundle.PlanID, session.userName);

    writeAuditLog({ userID: '', userName: session.userName, action: 'BUNDLE_PROCESS_UPDATE', module: 'Bundles', recordID: bundleID, oldValues: before, newValues: {
      process, status, operator: payload.operator
    } });

    return successResponse({ bundleID, process, status }, `${process} status updated to ${status}`);
  } catch (e) {
    return errorResponse('Failed to update bundle process: ' + e.message);
  }
}

// ── Bulk update bundles by process ───────────────────────────
function bulkUpdateBundleProcess(token, bundleIDs, process, status, payload) {
  const { valid: _wv, session, error: _we } = requireWriteAccess(token, 'production');
  if (!_wv) return _we;

  if (!Array.isArray(bundleIDs) || bundleIDs.length === 0) {
    return errorResponse('No bundle IDs provided');
  }

  const results = { success: [], failed: [] };
  for (const bundleID of bundleIDs) {
    const res = updateBundleProcess(token, bundleID, process, status, payload);
    if (res.success) {
      results.success.push(bundleID);
    } else {
      results.failed.push({ bundleID, error: res.message });
    }
  }

  return successResponse(results,
    `Updated ${results.success.length} bundles. ${results.failed.length} failed.`
  );
}

// ── Get bundle process summary for a plan ────────────────────
function getBundleProcessSummary(token, planID) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    const bundles = getSheetData('Bundles').filter(b => b.PlanID === planID);
    if (bundles.length === 0) return successResponse({ planID, bundles: 0, summary: {} });

    const summary = {};
    PROCESSES.forEach(proc => {
      const field = PROCESS_STATUS_FIELD[proc];
      const counts = { Pending: 0, 'In Progress': 0, Done: 0 };
      bundles.forEach(b => {
        const s = b[field] || 'Pending';
        counts[s] = (counts[s] || 0) + 1;
      });
      summary[proc] = {
        ...counts,
        total:    bundles.length,
        pctDone:  Math.round((counts['Done'] / bundles.length) * 100)
      };
    });

    // Pieces summary
    const totalPieces = bundles.reduce((s, b) => s + (Number(b.Pieces) || 0), 0);
    const donePackingPieces = bundles
      .filter(b => b.PackingStatus === 'Done')
      .reduce((s, b) => s + (Number(b.Pieces) || 0), 0);

    return successResponse({
      planID,
      totalBundles: bundles.length,
      totalPieces,
      donePackingPieces,
      summary
    });
  } catch (e) {
    return errorResponse('Failed to get bundle summary: ' + e.message);
  }
}

// ── Get bundles ready for dispatch (PackingStatus=Done, not dispatched) ──
function getReadyBundles(token, soID) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'dispatch');
  if (!_rv) return _re;

  try {
    let bundles = getSheetData('Bundles').filter(b =>
      b.SOID           === soID &&
      b.PackingStatus  === 'Done' &&
      b.FabricatorOut  !== 'Yes'
    );

    // Exclude already-dispatched bundles
    const dispatchedBundleIDs = new Set();
    getSheetData('DispatchItems').forEach(d => {
      if (d.BundleID) dispatchedBundleIDs.add(d.BundleID);
    });
    bundles = bundles.filter(b => !dispatchedBundleIDs.has(b.BundleID));

    const articleMap = {};
    getSheetData('ArticleMaster').forEach(a => { articleMap[a.ArticleID] = a; });

    const enriched = bundles.map(b => ({
      ...b,
      ArticleName: (articleMap[b.ArticleID] || {}).ArticleName || '',
      StyleCode:   (articleMap[b.ArticleID] || {}).StyleCode   || ''
    }));

    return successResponse(enriched);
  } catch (e) {
    return errorResponse('Failed to get ready bundles: ' + e.message);
  }
}

// ── Print bundle labels (returns label data for all plan bundles) ──
function getBundleLabelData(token, planID) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    const plan    = findRowByID('ProductionPlan', 'PlanID', planID);
    if (!plan) return errorResponse('Plan not found');

    const bundles = getSheetData('Bundles').filter(b => b.PlanID === planID);
    const article = findRowByID('ArticleMaster', 'ArticleID', plan.ArticleID) || {};
    const so      = findRowByID('SalesOrders',   'SOID',      plan.SOID)      || {};
    const buyer   = findRowByID('BuyerMaster',   'BuyerID',   so.BuyerID)    || {};

    const labels = bundles.map(b => ({
      BundleID:    b.BundleID,
      BarcodeValue:b.BarcodeValue,
      QRData:      b.QRData,
      BundleNumber:b.BundleNumber,
      Size:        b.Size,
      Color:       b.Color,
      Pieces:      b.Pieces,
      ArticleName: article.ArticleName || '',
      StyleCode:   article.StyleCode   || '',
      SONumber:    so.SONumber         || '',
      BuyerName:   buyer.BuyerName     || ''
    }));

    return successResponse({ plan, labels });
  } catch (e) {
    return errorResponse('Failed to get label data: ' + e.message);
  }
}

// ── Internal: Compute CurrentProcess from statuses ────────────
function _computeCurrentProcess(bundle) {
  if (bundle.PackingStatus   === 'Done') return 'Completed';
  if (bundle.PackingStatus   === 'In Progress') return 'Packing';
  if (bundle.FinishingStatus === 'Done') return 'Pre-Packing';
  if (bundle.FinishingStatus === 'In Progress') return 'Finishing';
  if (bundle.StitchingStatus === 'Done') return 'Pre-Finishing';
  if (bundle.StitchingStatus === 'In Progress') return 'Stitching';
  if (bundle.CuttingStatus   === 'Done') return 'Pre-Stitching';
  if (bundle.CuttingStatus   === 'In Progress') return 'Cutting';
  return 'Pre-Cutting';
}

// ── Internal: Auto-complete plan when all bundles packed ──────
function _checkPlanCompletion(planID, username) {
  try {
    const plan    = findRowByID('ProductionPlan', 'PlanID', planID);
    if (!plan || plan.Status !== 'In Progress') return;

    const bundles = getSheetData('Bundles').filter(b => b.PlanID === planID);
    if (bundles.length === 0) return;

    const allDone = bundles.every(b => b.PackingStatus === 'Done');
    if (allDone) {
      const before   = { ...plan };
      plan.Status    = 'Completed';
      plan.UpdatedAt = nowISO();
      updateRowByID('ProductionPlan', 'PlanID', planID, plan);
      writeAuditLog({ userID: '', userName: username, action: 'AUTO_COMPLETE', module: 'ProductionPlan', recordID: planID, oldValues: before, newValues: plan });
    }
  } catch (e) {
    // Non-fatal — don't block the caller
    Logger.log('_checkPlanCompletion error: ' + e.message);
  }
}