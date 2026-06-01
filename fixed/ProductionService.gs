// ============================================================
// ProductionService.gs — Production Plan + Bundle Management
// Sprint S6
// ============================================================

// ── Status transition matrices ───────────────────────────────
const PLAN_TRANSITIONS = {
  'Draft':       ['Confirmed', 'Cancelled'],
  'Confirmed':   ['In Progress', 'Cancelled'],
  'In Progress': ['Completed'],
  'Completed':   [],
  'Cancelled':   []
};

// ── Get all production plans ─────────────────────────────────
function getProductionPlans(token, filters) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    filters = filters || {};
    let rows = getSheetData('ProductionPlan');

    if (filters.soID)     rows = rows.filter(r => r.SOID === filters.soID);
    if (filters.status)   rows = rows.filter(r => r.Status === filters.status);
    if (filters.articleID) rows = rows.filter(r => r.ArticleID === filters.articleID);

    // Enrich with SO + Article info
    const soRows      = getSheetData('SalesOrders');
    const articleRows = getSheetData('ArticleMaster');
    const buyerRows   = getSheetData('BuyerMaster');

    const soMap      = {};
    soRows.forEach(r => { soMap[r.SOID] = r; });
    const articleMap = {};
    articleRows.forEach(r => { articleMap[r.ArticleID] = r; });
    const buyerMap   = {};
    buyerRows.forEach(r => { buyerMap[r.BuyerID] = r; });

    const enriched = rows.map(r => ({
      ...r,
      SONumber:     soMap[r.SOID]      ? soMap[r.SOID].SONumber      : '',
      BuyerName:    soMap[r.SOID]      ? (buyerMap[soMap[r.SOID].BuyerID] || {}).BuyerName || '' : '',
      ArticleName:  articleMap[r.ArticleID] ? articleMap[r.ArticleID].ArticleName : '',
      StyleCode:    articleMap[r.ArticleID] ? articleMap[r.ArticleID].StyleCode   : ''
    }));

    return successResponse(enriched);
  } catch (e) {
    return errorResponse('Failed to get production plans: ' + e.message);
  }
}

// ── Get single plan with bundles ─────────────────────────────
function getProductionPlanByID(token, planID) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    const plan = findRowByID('ProductionPlan', 'PlanID', planID);
    if (!plan) return errorResponse('Production plan not found');

    // Enrich
    const soRow      = findRowByID('SalesOrders', 'SOID', plan.SOID) || {};
    const articleRow = findRowByID('ArticleMaster', 'ArticleID', plan.ArticleID) || {};
    const buyerRow   = findRowByID('BuyerMaster', 'BuyerID', soRow.BuyerID) || {};

    plan.SONumber    = soRow.SONumber    || '';
    plan.BuyerName   = buyerRow.BuyerName || '';
    plan.ArticleName = articleRow.ArticleName || '';
    plan.StyleCode   = articleRow.StyleCode   || '';

    // Get bundles for this plan
    const bundles = getSheetData('Bundles').filter(b => b.PlanID === planID);
    plan.Bundles  = bundles;
    plan.TotalBundles = bundles.length;
    plan.BundlesByStatus = _countBundleStatuses(bundles);

    return successResponse(plan);
  } catch (e) {
    return errorResponse('Failed to get production plan: ' + e.message);
  }
}

// ── Create Production Plan ────────────────────────────────────
function createProductionPlan(token, payload) {
  const { valid: _wv, session, error: _we } = requireWriteAccess(token, 'production');
  if (!_wv) return _we;

  try {
    // Validate required fields
    if (!payload.SOID)      return errorResponse('Sales Order is required');
    if (!payload.ArticleID) return errorResponse('Article is required');
    if (!payload.PlannedStartDate) return errorResponse('Planned Start Date is required');
    if (!payload.PlannedEndDate)   return errorResponse('Planned End Date is required');
    if (!payload.BundleSize || payload.BundleSize < 1) return errorResponse('Bundle Size must be at least 1');

    // Validate SO exists and is Confirmed or In Production
    const so = findRowByID('SalesOrders', 'SOID', payload.SOID);
    if (!so) return errorResponse('Sales Order not found');
    if (!['Confirmed', 'In Production'].includes(so.Status)) {
      return errorResponse('Sales Order must be Confirmed or In Production to create a plan');
    }

    // Validate Article belongs to SO
    const soItems = getSheetData('SalesOrderItems').filter(i => i.SOID === payload.SOID && i.ArticleID === payload.ArticleID);
    if (soItems.length === 0) return errorResponse('Article not found in this Sales Order');

    // Check for duplicate active plan for same SO+Article
    const existing = getSheetData('ProductionPlan').filter(p =>
      p.SOID === payload.SOID &&
      p.ArticleID === payload.ArticleID &&
      !['Cancelled', 'Completed'].includes(p.Status)
    );
    if (existing.length > 0) return errorResponse('An active production plan already exists for this SO + Article');

    // Calculate total pieces for this article in SO
    const totalPieces = soItems.reduce((sum, i) => sum + (Number(i.Quantity) || 0), 0);

    const planID = generateID('PLN');
    const plan = {
      PlanID:           planID,
      SOID:             payload.SOID,
      ArticleID:        payload.ArticleID,
      PlannedStartDate: payload.PlannedStartDate,
      PlannedEndDate:   payload.PlannedEndDate,
      BundleSize:       Number(payload.BundleSize),
      TotalPieces:      totalPieces,
      Status:           'Draft',
      Notes:            sanitizeString(payload.Notes || ''),
      CreatedBy:        session.userName,
      CreatedAt:        nowISO(),
      UpdatedAt:        nowISO()
    };

    appendRow('ProductionPlan', plan);
    writeAuditLog({ userID: '', userName: session.userName, action: 'CREATE', module: 'ProductionPlan', recordID: planID, oldValues: null, newValues: plan });

    return successResponse({ planID }, 'Production plan created successfully');
  } catch (e) {
    return errorResponse('Failed to create production plan: ' + e.message);
  }
}

// ── Update Production Plan (Draft only) ──────────────────────
function updateProductionPlan(token, planID, payload) {
  const { valid: _wv, session, error: _we } = requireWriteAccess(token, 'production');
  if (!_wv) return _we;

  try {
    const plan = findRowByID('ProductionPlan', 'PlanID', planID);
    if (!plan) return errorResponse('Production plan not found');
    if (plan.Status !== 'Draft') return errorResponse('Only Draft plans can be edited');

    const before = { ...plan };

    if (payload.PlannedStartDate) plan.PlannedStartDate = payload.PlannedStartDate;
    if (payload.PlannedEndDate)   plan.PlannedEndDate   = payload.PlannedEndDate;
    if (payload.BundleSize && Number(payload.BundleSize) >= 1) plan.BundleSize = Number(payload.BundleSize);
    if (payload.Notes !== undefined) plan.Notes = sanitizeString(payload.Notes);
    plan.UpdatedAt = nowISO();

    updateRowByID('ProductionPlan', 'PlanID', planID, plan);
    writeAuditLog({ userID: '', userName: session.userName, action: 'UPDATE', module: 'ProductionPlan', recordID: planID, oldValues: before, newValues: plan });

    return successResponse({ planID }, 'Production plan updated successfully');
  } catch (e) {
    return errorResponse('Failed to update production plan: ' + e.message);
  }
}

// ── Update Plan Status ────────────────────────────────────────
function updatePlanStatus(token, planID, newStatus) {
  const { valid: _wv, session, error: _we } = requireWriteAccess(token, 'production');
  if (!_wv) return _we;

  try {
    const plan = findRowByID('ProductionPlan', 'PlanID', planID);
    if (!plan) return errorResponse('Production plan not found');

    const allowed = PLAN_TRANSITIONS[plan.Status] || [];
    if (!allowed.includes(newStatus)) {
      return errorResponse(`Cannot transition from ${plan.Status} to ${newStatus}`);
    }

    const before = { ...plan };
    plan.Status    = newStatus;
    plan.UpdatedAt = nowISO();

    updateRowByID('ProductionPlan', 'PlanID', planID, plan);
    writeAuditLog({ userID: '', userName: session.userName, action: 'STATUS_CHANGE', module: 'ProductionPlan', recordID: planID, oldValues: before, newValues: plan });

    return successResponse({ planID, status: newStatus }, `Plan status updated to ${newStatus}`);
  } catch (e) {
    return errorResponse('Failed to update plan status: ' + e.message);
  }
}

// ── Generate Bundles for a Plan ───────────────────────────────
// Creates bundle records from BOM size/color breakdown
function generateBundles(token, planID) {
  const { valid: _wv, session, error: _we } = requireWriteAccess(token, 'production');
  if (!_wv) return _we;

  try {
    const plan = findRowByID('ProductionPlan', 'PlanID', planID);
    if (!plan) return errorResponse('Production plan not found');
    if (!['Draft', 'Confirmed'].includes(plan.Status)) {
      return errorResponse('Bundles can only be generated for Draft or Confirmed plans');
    }

    // Check if bundles already generated
    const existingBundles = getSheetData('Bundles').filter(b => b.PlanID === planID);
    if (existingBundles.length > 0) {
      return errorResponse('Bundles already generated for this plan. Delete existing bundles first.');
    }

    // Get SO items for this article
    const soItems = getSheetData('SalesOrderItems').filter(i =>
      i.SOID === plan.SOID && i.ArticleID === plan.ArticleID
    );
    if (soItems.length === 0) return errorResponse('No SO items found for this article');

    // Get BOM for this article
    const bomRows = getSheetData('BOM').filter(b => b.ArticleID === plan.ArticleID);

    // Build size→color quantities from SO items
    // soItems have Size + Quantity (total pieces for that size across colors)
    // We'll generate bundles per size with BundleSize pieces each

    const bundleSize = Number(plan.BundleSize);
    const bundlesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bundles');
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let bundleNumber = 1;
    const createdBundles = [];
    const now = nowISO();
    const user = session.userName;

    for (const soItem of soItems) {
      const size        = soItem.Size;
      const totalPieces = Number(soItem.Quantity) || 0;
      const color       = soItem.Color || '';

      if (totalPieces === 0) continue;

      let remaining = totalPieces;
      while (remaining > 0) {
        const piecesInBundle = Math.min(remaining, bundleSize);
        remaining -= piecesInBundle;

        const bundleID  = generateID('BND');
        const barcodeVal = _generateBarcodeValue(planID, bundleNumber);
        const qrData     = _generateQRData(bundleID, planID, plan.SOID, plan.ArticleID, size, color, piecesInBundle);

        const bundle = {
          BundleID:       bundleID,
          PlanID:         planID,
          SOID:           plan.SOID,
          ArticleID:      plan.ArticleID,
          Size:           size,
          Color:          color,
          BundleNumber:   bundleNumber,
          Pieces:         piecesInBundle,
          BarcodeValue:   barcodeVal,
          QRData:         qrData,
          CuttingStatus:  'Pending',
          StitchingStatus:'Pending',
          FinishingStatus:'Pending',
          PackingStatus:  'Pending',
          CurrentProcess: 'Pre-Cutting',
          FabricatorOut:  'No',
          CreatedBy:      user,
          CreatedAt:      now,
          UpdatedAt:      now
        };

        appendRow('Bundles', bundle);
        createdBundles.push(bundle);
        bundleNumber++;
      }
    }

    // Update plan status to Confirmed if still Draft
    if (plan.Status === 'Draft') {
      plan.Status    = 'Confirmed';
      plan.UpdatedAt = now;
      updateRowByID('ProductionPlan', 'PlanID', planID, plan);
    }

    writeAuditLog({ userID: '', userName: user, action: 'GENERATE_BUNDLES', module: 'ProductionPlan', recordID: planID, oldValues: null, newValues: {
      bundleCount: createdBundles.length,
      planID
    } });

    return successResponse({
      planID,
      bundleCount: createdBundles.length,
      bundles: createdBundles
    }, `${createdBundles.length} bundles generated successfully`);
  } catch (e) {
    return errorResponse('Failed to generate bundles: ' + e.message);
  }
}

// ── Delete all bundles for a plan (reset) ────────────────────
function deletePlanBundles(token, planID) {
  const { valid: _wv, session, error: _we } = requireWriteAccess(token, 'production');
  if (!_wv) return _we;

  try {
    const plan = findRowByID('ProductionPlan', 'PlanID', planID);
    if (!plan) return errorResponse('Plan not found');
    if (!['Draft', 'Confirmed'].includes(plan.Status)) {
      return errorResponse('Cannot delete bundles for In Progress or Completed plans');
    }

    // Check none are processed
    const bundles = getSheetData('Bundles').filter(b => b.PlanID === planID);
    const processed = bundles.filter(b => b.CuttingStatus !== 'Pending');
    if (processed.length > 0) {
      return errorResponse('Cannot delete — some bundles already have cutting records');
    }

    // Delete from sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Bundles');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const bundleIDCol = headers.indexOf('BundleID') + 1;

    const data = sheet.getDataRange().getValues();
    const planIDCol = headers.indexOf('PlanID') + 1;
    const rowsToDelete = [];

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][planIDCol - 1] === planID) {
        rowsToDelete.push(i + 1);
      }
    }
    rowsToDelete.forEach(r => sheet.deleteRow(r));

    // Reset plan to Draft
    plan.Status    = 'Draft';
    plan.UpdatedAt = nowISO();
    updateRowByID('ProductionPlan', 'PlanID', planID, plan);

    writeAuditLog({ userID: '', userName: session.userName, action: 'DELETE_BUNDLES', module: 'ProductionPlan', recordID: planID, oldValues: null, newValues: { deleted: rowsToDelete.length } });

    return successResponse({}, `${rowsToDelete.length} bundles deleted`);
  } catch (e) {
    return errorResponse('Failed to delete bundles: ' + e.message);
  }
}

// ── Get SO items for plan creation (article breakdown) ───────
function getSOArticleBreakdown(token, soID) {
  const { valid: _rv, session, error: _re } = requireReadAccess(token, 'production');
  if (!_rv) return _re;

  try {
    const so = findRowByID('SalesOrders', 'SOID', soID);
    if (!so) return errorResponse('Sales Order not found');

    const soItems = getSheetData('SalesOrderItems').filter(i => i.SOID === soID);
    const articleMap = {};
    getSheetData('ArticleMaster').forEach(r => { articleMap[r.ArticleID] = r; });

    // Group by article
    const byArticle = {};
    soItems.forEach(item => {
      if (!byArticle[item.ArticleID]) {
        byArticle[item.ArticleID] = {
          ArticleID:   item.ArticleID,
          ArticleName: (articleMap[item.ArticleID] || {}).ArticleName || '',
          StyleCode:   (articleMap[item.ArticleID] || {}).StyleCode   || '',
          TotalPieces: 0,
          Sizes: []
        };
      }
      byArticle[item.ArticleID].TotalPieces += Number(item.Quantity) || 0;
      byArticle[item.ArticleID].Sizes.push({
        Size:     item.Size,
        Color:    item.Color || '',
        Quantity: Number(item.Quantity) || 0
      });
    });

    return successResponse({
      so,
      articles: Object.values(byArticle)
    });
  } catch (e) {
    return errorResponse('Failed to get SO breakdown: ' + e.message);
  }
}

// ── Internal: Generate barcode value ─────────────────────────
function _generateBarcodeValue(planID, bundleNumber) {
  // Format: PLN-XXXXX-BNNN  (short, scannable)
  const planShort = planID.replace('PLN-', '').substring(0, 6);
  const bn = String(bundleNumber).padStart(4, '0');
  return `${planShort}-B${bn}`;
}

// ── Internal: Generate QR data (JSON string) ─────────────────
function _generateQRData(bundleID, planID, soID, articleID, size, color, pieces) {
  return JSON.stringify({
    bid: bundleID,
    pid: planID,
    so:  soID,
    art: articleID,
    sz:  size,
    col: color,
    pcs: pieces,
    ts:  new Date().getTime()
  });
}

// ── Internal: Count bundle statuses ──────────────────────────
function _countBundleStatuses(bundles) {
  const counts = { Pending: 0, 'In Progress': 0, Done: 0 };
  const processes = ['CuttingStatus', 'StitchingStatus', 'FinishingStatus', 'PackingStatus'];
  processes.forEach(proc => {
    bundles.forEach(b => {
      const s = b[proc] || 'Pending';
      counts[s] = (counts[s] || 0) + 1;
    });
  });
  return counts;
}