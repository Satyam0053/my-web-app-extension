// ============================================================
// FabricatorService.gs — Fabricator Job Work (Sprint S8)
// Manage challans sent to and received from fabricators
// ============================================================

// ── Get all challans ─────────────────────────────────────────
function getFabricatorChallans(token, filters) {
  var auth = requireReadAccess(token, 'FabricatorChallan');
  if (!auth.valid) return auth.error;
  try {
    filters = filters || {};
    var rows = getSheetData('FabricatorChallan');

    if (filters.type)          rows = rows.filter(function(r) { return r.ChallanType === filters.type; });
    if (filters.fabricatorID)  rows = rows.filter(function(r) { return r.FabricatorID === filters.fabricatorID; });
    if (filters.status)        rows = rows.filter(function(r) { return r.Status === filters.status; });

    // Enrich with vendor names
    var vendorMap = {};
    getSheetData('VendorMaster').forEach(function(v) { vendorMap[v.VendorID] = v.VendorName; });

    var enriched = rows.map(function(r) {
      return Object.assign({}, r, {
        FabricatorName: vendorMap[r.FabricatorID] || r.FabricatorID
      });
    });

    // Sort newest first
    enriched.sort(function(a, b) { return String(b.ChallanID).localeCompare(String(a.ChallanID)); });

    return successResponse(enriched);
  } catch (e) {
    return errorResponse('Failed to get challans: ' + e.message);
  }
}

// ── Get single challan with items ────────────────────────────
function getFabricatorChallanByID(token, challanID) {
  var auth = requireReadAccess(token, 'FabricatorChallan');
  if (!auth.valid) return auth.error;
  try {
    var challan = findRowByID('FabricatorChallan', 'ChallanID', challanID);
    if (!challan) return errorResponse('Challan not found');

    var vendorMap = {};
    getSheetData('VendorMaster').forEach(function(v) { vendorMap[v.VendorID] = v.VendorName; });
    challan.FabricatorName = vendorMap[challan.FabricatorID] || challan.FabricatorID;

    // Get items with fabric info
    var fabricMap = {};
    getSheetData('FabricMaster').forEach(function(f) { fabricMap[f.FabricID] = f.DesignName; });
    var colorMap = {};
    getSheetData('FabricColors').forEach(function(c) { colorMap[c.FabricColorID] = c.ColorName; });

    var items = getSheetData('FabricatorChallanItems')
      .filter(function(i) { return i.ChallanID === challanID; })
      .map(function(i) {
        return Object.assign({}, i, {
          FabricName: fabricMap[i.FabricID] || i.FabricID,
          ColorName:  colorMap[i.FabricColorID] || i.FabricColorID
        });
      });

    challan.Items = items;
    return successResponse(challan);
  } catch (e) {
    return errorResponse('Failed to get challan: ' + e.message);
  }
}

// ── Create manual Challan Out ─────────────────────────────────
function createFabricatorChallan(token, data) {
  var auth = requireWriteAccess(token, 'FabricatorChallan');
  if (!auth.valid) return auth.error;
  try {
    if (!data.FabricatorID)  return errorResponse('Fabricator is required');
    if (!data.ChallanDate)   return errorResponse('Challan date is required');
    if (!Array.isArray(data.items) || data.items.length === 0) return errorResponse('At least one item is required');

    // Validate fabricator exists and is a fabricator
    var vendor = findRowByID('VendorMaster', 'VendorID', data.FabricatorID);
    if (!vendor) return errorResponse('Fabricator not found');
    if (vendor.VendorType !== 'Fabricator' && vendor.VendorType !== 'Both') {
      return errorResponse('Selected vendor is not a Fabricator');
    }

    var now       = nowISO();
    var challanID = generateYearID('FJC', 'FabricatorChallan', 1);

    appendRow('FabricatorChallan', {
      ChallanID:    challanID,
      ChallanDate:  data.ChallanDate,
      ChallanType:  'Out',
      FabricatorID: data.FabricatorID,
      IssueID:      data.IssueID || '',
      Remarks:      sanitizeString(data.Remarks || ''),
      Status:       'Open',
      CreatedBy:    auth.session.userID,
      CreatedAt:    now,
      UpdatedBy:    '',
      UpdatedAt:    ''
    });

    data.items.forEach(function(item) {
      var cItemID = generateID('FCI');
      appendRow('FabricatorChallanItems', {
        ChallanItemID: cItemID,
        ChallanID:     challanID,
        FabricID:      item.FabricID,
        FabricColorID: item.FabricColorID,
        Qty:           parseFloat(item.Qty) || 0,
        UOM:           item.UOM || 'Meters',
        Remarks:       sanitizeString(item.Remarks || '')
      });
    });

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'CREATE', module: 'FabricatorChallan', recordID: challanID, oldValues: '', newValues: data });
    return successResponse({ challanID: challanID }, 'Challan created: ' + challanID);
  } catch (e) {
    return errorResponse('Failed to create challan: ' + e.message);
  }
}

// ── Receive goods back from fabricator ───────────────────────
function receiveFabricatorChallan(token, challanID, data) {
  var auth = requireWriteAccess(token, 'FabricatorChallan');
  if (!auth.valid) return auth.error;
  try {
    var challan = findRowByID('FabricatorChallan', 'ChallanID', challanID);
    if (!challan) return errorResponse('Challan not found');
    if (challan.Status !== 'Open') return errorResponse('Only Open challans can be received');

    var now    = nowISO();
    var before = Object.assign({}, challan);

    challan.Status    = 'Received';
    challan.UpdatedBy = auth.session.userID;
    challan.UpdatedAt = now;
    if (data.Remarks) challan.Remarks = sanitizeString(data.Remarks);

    updateRowByID('FabricatorChallan', 'ChallanID', challanID, challan);

    // Create a Challan In record
    var inChallanID = generateYearID('FJC', 'FabricatorChallan', 1);
    appendRow('FabricatorChallan', {
      ChallanID:    inChallanID,
      ChallanDate:  data.ReceiveDate || now.slice(0, 10),
      ChallanType:  'In',
      FabricatorID: challan.FabricatorID,
      IssueID:      challan.IssueID,
      Remarks:      'Received against: ' + challanID + (data.Remarks ? ' — ' + data.Remarks : ''),
      Status:       'Closed',
      CreatedBy:    auth.session.userID,
      CreatedAt:    now,
      UpdatedBy:    '',
      UpdatedAt:    ''
    });

    // Copy items to Challan In with received quantities
    var outItems = getSheetData('FabricatorChallanItems').filter(function(i) { return i.ChallanID === challanID; });
    outItems.forEach(function(item) {
      var inItemID = generateID('FCI');
      var receivedQty = (data.items && data.items[item.ChallanItemID]) || item.Qty;
      appendRow('FabricatorChallanItems', {
        ChallanItemID: inItemID,
        ChallanID:     inChallanID,
        FabricID:      item.FabricID,
        FabricColorID: item.FabricColorID,
        Qty:           parseFloat(receivedQty) || 0,
        UOM:           item.UOM,
        Remarks:       'In against ' + item.ChallanItemID
      });
    });

    writeAuditLog({ userID: auth.session.userID, userName: auth.session.userName, action: 'RECEIVE', module: 'FabricatorChallan', recordID: challanID, oldValues: before, newValues: challan });
    return successResponse({ challanID: challanID, inChallanID: inChallanID }, 'Challan received successfully');
  } catch (e) {
    return errorResponse('Failed to receive challan: ' + e.message);
  }
}

// ── Get pending challans summary ─────────────────────────────
function getFabricatorSummary(token) {
  var auth = requireReadAccess(token, 'FabricatorChallan');
  if (!auth.valid) return auth.error;
  try {
    var challans = getSheetData('FabricatorChallan');
    var open     = challans.filter(function(c) { return c.Status === 'Open' && c.ChallanType === 'Out'; });
    var received = challans.filter(function(c) { return c.Status === 'Received'; });

    var vendorMap = {};
    getSheetData('VendorMaster').forEach(function(v) { vendorMap[v.VendorID] = v.VendorName; });

    // Group open by fabricator
    var byFabricator = {};
    open.forEach(function(c) {
      var name = vendorMap[c.FabricatorID] || c.FabricatorID;
      if (!byFabricator[name]) byFabricator[name] = 0;
      byFabricator[name]++;
    });

    return successResponse({
      totalOpen:     open.length,
      totalReceived: received.length,
      byFabricator:  byFabricator
    });
  } catch (e) {
    return errorResponse('Failed to get summary: ' + e.message);
  }
}