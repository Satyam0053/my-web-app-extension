// ============================================================
// File: Reports.gs — Reports, Analytics & Dashboard
// ============================================================

// ── DASHBOARD DATA ───────────────────────────────────────────
function getDashboardData() {
  const inv     = JSON.parse(getInventory({}));
  const issues  = JSON.parse(getIssues({}));
  const pos     = JSON.parse(getPurchaseOrders({}));
  const returns = JSON.parse(getReturns({}));

  const invItems   = inv.data    || [];
  const issueItems = issues.data || [];
  const poItems    = pos.data    || [];
  const retItems   = returns.data || [];

  // Total stock value (sum of all current stocks)
  const totalStock = invItems.reduce((s, i) => s + (parseFloat(i.Current_Stock) || 0), 0);
  const totalItems = invItems.length;
  const lowStockCount = invItems.filter(i =>
    parseFloat(i.Current_Stock) <= parseFloat(i.Reorder_Level)).length;

  // Today's issues
  const today = new Date(); today.setHours(0,0,0,0);
  const todayIssues = issueItems.filter(i => new Date(i.Issued_At) >= today);
  const todayIssuedQty = todayIssues.reduce((s, i) => s + (parseFloat(i.Issue_Qty) || 0), 0);

  // This month's issues
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const monthIssues = issueItems.filter(i => new Date(i.Issued_At) >= monthStart);
  const monthIssuedQty = monthIssues.reduce((s, i) => s + (parseFloat(i.Issue_Qty) || 0), 0);

  // PO stats
// PO stats
  const pendingPOs   = poItems.filter(p => p.Status === 'Pending').length;
  const completedPOs = poItems.filter(p => p.Status === 'Completed').length;

  // Total Order Qty & Total Pending Qty (across ALL POs, where received < ordered)
  const receiptDataForQty = JSON.parse(getReceipts({})).data || [];
  const receivedByPO = {};
  receiptDataForQty.forEach(r => {
    receivedByPO[r.PO_Number] = (receivedByPO[r.PO_Number] || 0) + (parseFloat(r.Received_Qty) || 0);
  });

  let totalOrderQty   = 0;
  let totalPendingQty = 0;
  const vendorPendingMap = {};

  poItems.forEach(po => {
    const ordered  = parseFloat(po.Ordered_Qty) || 0;
    const received = receivedByPO[po.PO_Number]  || 0;
    const pending  = Math.max(0, ordered - received);
    totalOrderQty   += ordered;
    totalPendingQty += pending;
    if (pending > 0) {
      const vendor = po.Vendor_Name || 'Unknown';
      vendorPendingMap[vendor] = (vendorPendingMap[vendor] || 0) + pending;
    }
  });

  const vendorPendingQty = Object.entries(vendorPendingMap)
    .sort((a, b) => b[1] - a[1])
    .map(([vendor, qty]) => ({ vendor, qty: Math.round(qty * 100) / 100 }));

  // Top 5 articles by consumption
  const articleConsumption = {};
  issueItems.forEach(i => {
    const key = i.Article_Number || i.Fabric_Name;
    articleConsumption[key] = (articleConsumption[key] || 0) + (parseFloat(i.Issue_Qty) || 0);
  });
  const topArticles = Object.entries(articleConsumption)
    .sort((a, b) => b[1] - a[1]).slice(0, 5)
    .map(([name, qty]) => ({ name, qty }));

  // Vendor-wise receipt quantities
  const vendorInward = {};
  const receiptData = JSON.parse(getReceipts({})).data || [];
  receiptData.forEach(r => {
    vendorInward[r.Vendor_Name] = (vendorInward[r.Vendor_Name] || 0) + (parseFloat(r.Received_Qty) || 0);
  });
  const topVendors = Object.entries(vendorInward)
    .sort((a, b) => b[1] - a[1]).slice(0, 5)
    .map(([name, qty]) => ({ name, qty }));

  // Color-wise consumption
  const colorConsumption = {};
  issueItems.forEach(i => {
    colorConsumption[i.Color] = (colorConsumption[i.Color] || 0) + (parseFloat(i.Issue_Qty) || 0);
  });
  const colorData = Object.entries(colorConsumption)
    .sort((a, b) => b[1] - a[1]).slice(0, 8)
    .map(([color, qty]) => ({ color, qty }));

  // Daily trend (last 7 days)
  const dailyTrend = [];
  for (let d = 6; d >= 0; d--) {
    const day = new Date(); day.setHours(0,0,0,0); day.setDate(day.getDate() - d);
    const nextDay = new Date(day); nextDay.setDate(nextDay.getDate() + 1);
    const dayIssues = issueItems.filter(i => {
      const t = new Date(i.Issued_At);
      return t >= day && t < nextDay;
    });
    dailyTrend.push({
      date: day.toLocaleDateString('en-IN', { month: 'short', day: 'numeric' }),
      qty: dayIssues.reduce((s, i) => s + (parseFloat(i.Issue_Qty) || 0), 0)
    });
  }

  // Monthly consumption trend (last 6 months)
  const monthlyTrend = [];
  for (let m = 5; m >= 0; m--) {
    const mStart = new Date(today.getFullYear(), today.getMonth() - m, 1);
    const mEnd   = new Date(today.getFullYear(), today.getMonth() - m + 1, 0);
    const mIssues = issueItems.filter(i => {
      const t = new Date(i.Issued_At);
      return t >= mStart && t <= mEnd;
    });
    monthlyTrend.push({
      month: mStart.toLocaleDateString('en-IN', { month: 'short', year: '2-digit' }),
      qty: mIssues.reduce((s, i) => s + (parseFloat(i.Issue_Qty) || 0), 0)
    });
  }

  return success({
    stats: {
      totalStock: Math.round(totalStock * 100) / 100,
      totalItems,
      lowStockCount,
      todayIssuedQty: Math.round(todayIssuedQty * 100) / 100,
      todayIssueCount: todayIssues.length,
      monthIssuedQty: Math.round(monthIssuedQty * 100) / 100,
      pendingPOs,
      completedPOs,
      totalPOs: poItems.length,
      totalReturns: retItems.length,
      totalOrderQty: Math.round(totalOrderQty * 100) / 100,
      totalPendingQty: Math.round(totalPendingQty * 100) / 100
    },
    topArticles,
    topVendors,
    colorData,
    dailyTrend,
    monthlyTrend,
    vendorPendingQty,
    lowStockItems: invItems.filter(i => parseFloat(i.Current_Stock) <= parseFloat(i.Reorder_Level)).slice(0,5)
  });
}

// ── STOCK LEDGER ─────────────────────────────────────────────
function getStockLedger(filters) {
  const invId = filters && filters.inventoryId;
  const txns  = [];

  // GRN receipts
  const receipts = JSON.parse(getReceipts(filters || {})).data || [];
  receipts.forEach(r => {
    if (!invId || r.Inventory_ID === invId)
      txns.push({ date: r.Received_At, type: 'GRN', ref: r.GRN_Number, article: r.Article_Number, fabric: r.Fabric_Name, color: r.Color, qty: parseFloat(r.Received_Qty) || 0, in: parseFloat(r.Received_Qty) || 0, out: 0, vendor: r.Vendor_Name, dept: '' });
  });

  // Issues
  const issues = JSON.parse(getIssues(filters || {})).data || [];
  issues.forEach(i => {
    if (!invId || i.Inventory_ID === invId)
      txns.push({ date: i.Issued_At, type: 'ISSUE', ref: i.Issue_ID, article: i.Article_Number, fabric: i.Fabric_Name, color: i.Color, qty: parseFloat(i.Issue_Qty) || 0, in: 0, out: parseFloat(i.Issue_Qty) || 0, vendor: '', dept: i.Department });
  });

  // Returns
  const returns = JSON.parse(getReturns(filters || {})).data || [];
  returns.forEach(r => {
    if (!invId || r.Inventory_ID === invId)
      txns.push({ date: r.Returned_At, type: 'RETURN', ref: r.Return_ID, article: r.Article_Number, fabric: r.Fabric_Name, color: r.Color, qty: parseFloat(r.Return_Qty) || 0, in: parseFloat(r.Return_Qty) || 0, out: 0, vendor: '', dept: '' });
  });

  // Sort by date
  txns.sort((a, b) => new Date(a.date) - new Date(b.date));

  // Running balance
  let balance = 0;
  txns.forEach(t => { balance += t.in - t.out; t.balance = balance; });

  return success(txns.reverse()); // newest first for display
}

// ── DAILY ISSUE REPORT ───────────────────────────────────────
function getDailyIssueReport(filters) {
  const date  = filters && filters.date ? new Date(filters.date) : new Date();
  const start = new Date(date); start.setHours(0,0,0,0);
  const end   = new Date(date); end.setHours(23,59,59,999);

  const issues = JSON.parse(getIssues({})).data || [];
  const dayIssues = issues.filter(i => {
    const t = new Date(i.Issued_At);
    return t >= start && t <= end;
  });

  // Group by department
  const byDept = {};
  dayIssues.forEach(i => {
    if (!byDept[i.Department]) byDept[i.Department] = { department: i.Department, items: [], totalQty: 0 };
    byDept[i.Department].items.push(i);
    byDept[i.Department].totalQty += parseFloat(i.Issue_Qty) || 0;
  });

  return success({
    date: start.toLocaleDateString(),
    totalIssues: dayIssues.length,
    totalQty: dayIssues.reduce((s, i) => s + (parseFloat(i.Issue_Qty) || 0), 0),
    byDepartment: Object.values(byDept),
    transactions: dayIssues
  });
}

// ── MONTHLY CONSUMPTION ──────────────────────────────────────
function getMonthlyConsumption(filters) {
  const now   = new Date();
  const month = filters && filters.month ? parseInt(filters.month) : now.getMonth();
  const year  = filters && filters.year  ? parseInt(filters.year)  : now.getFullYear();
  const start = new Date(year, month, 1);
  const end   = new Date(year, month + 1, 0, 23, 59, 59);

  const issues = JSON.parse(getIssues({})).data || [];
  const monthIssues = issues.filter(i => {
    const t = new Date(i.Issued_At);
    return t >= start && t <= end;
  });

  // Article-wise summary
  const byArticle = {};
  monthIssues.forEach(i => {
    const key = `${i.Article_Number}|${i.Fabric_Name}|${i.Color}`;
    if (!byArticle[key]) byArticle[key] = { article: i.Article_Number, fabric: i.Fabric_Name, color: i.Color, totalQty: 0, txnCount: 0 };
    byArticle[key].totalQty += parseFloat(i.Issue_Qty) || 0;
    byArticle[key].txnCount++;
  });

  return success({
    month: start.toLocaleDateString('en-IN', { month: 'long', year: 'numeric' }),
    totalQty: monthIssues.reduce((s, i) => s + (parseFloat(i.Issue_Qty) || 0), 0),
    byArticle: Object.values(byArticle).sort((a, b) => b.totalQty - a.totalQty),
    transactions: monthIssues
  });
}

// ── VENDOR REPORT ────────────────────────────────────────────
function getVendorReport(filters) {
  const receipts = JSON.parse(getReceipts(filters || {})).data || [];
  const byVendor = {};
  receipts.forEach(r => {
    if (!byVendor[r.Vendor_Name])
      byVendor[r.Vendor_Name] = { vendor: r.Vendor_Name, totalReceived: 0, grnCount: 0, articles: new Set(), fabrics: [] };
    byVendor[r.Vendor_Name].totalReceived += parseFloat(r.Received_Qty) || 0;
    byVendor[r.Vendor_Name].grnCount++;
    byVendor[r.Vendor_Name].articles.add(r.Article_Number);
    byVendor[r.Vendor_Name].fabrics.push(`${r.Fabric_Name} ${r.Color}`);
  });

  const result = Object.values(byVendor).map(v => ({
    ...v, articles: Array.from(v.articles), fabrics: [...new Set(v.fabrics)]
  })).sort((a, b) => b.totalReceived - a.totalReceived);

  return success(result);
}

// ── ARTICLE REPORT ───────────────────────────────────────────
function getArticleReport(filters) {
  const invItems = JSON.parse(getInventory(filters || {})).data || [];
  const issues   = JSON.parse(getIssues({})).data || [];

  return success(invItems.map(inv => {
    const artIssues = issues.filter(i => i.Article_Number === inv.Article_Number && i.Color === inv.Color);
    const totalIssued = artIssues.reduce((s, i) => s + (parseFloat(i.Issue_Qty) || 0), 0);
    return {
      ...inv,
      Total_Issued: totalIssued,
      Issue_Count: artIssues.length
    };
  }).sort((a, b) => b.Total_Issued - a.Total_Issued));
}

// ── FABRIC MOVEMENT REPORT ───────────────────────────────────
function getFabricMovement(filters) {
  return getStockLedger(filters);
}

// ── ACTIVITY LOGS ────────────────────────────────────────────
function getActivityLogs(filters) {
  const sheet  = getSheet(SHEETS.ACTIVITY_LOGS);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);
  const headers = data[0];
  let rows = data.slice(1)
    .map(r => rowToObj(headers, r))
    .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))
    .slice(0, 100);
  return success(rows);
}

// ── GLOBAL SEARCH ─────────────────────────────────────────────
function globalSearch(query) {
  if (!query || query.length < 2) return success([]);
  const q = query.toLowerCase();
  const results = [];

  const invItems = JSON.parse(getInventory({})).data || [];
  invItems.filter(i =>
    String(i.Article_Number).toLowerCase().includes(q) ||
    i.Fabric_Name.toLowerCase().includes(q) ||
    i.Color.toLowerCase().includes(q) ||
    i.Vendor_Name.toLowerCase().includes(q)
  ).slice(0, 10).forEach(i => results.push({ type: 'Inventory', id: i.Inventory_ID, label: `${i.Article_Number} | ${i.Fabric_Name} | ${i.Color}`, meta: `Stock: ${i.Current_Stock} ${i.Unit}` }));

  const poItems = JSON.parse(getPurchaseOrders({})).data || [];
  poItems.filter(p =>
    p.PO_Number.toLowerCase().includes(q) ||
    p.Vendor_Name.toLowerCase().includes(q) ||
    p.Fabric_Name.toLowerCase().includes(q)
  ).slice(0, 5).forEach(p => results.push({ type: 'PurchaseOrder', id: p.PO_Number, label: `${p.PO_Number} | ${p.Vendor_Name}`, meta: `Status: ${p.Status}` }));

  return success(results);
}

// ── DASHBOARD CACHE REFRESH ──────────────────────────────────
function refreshDashboardCache() {
  // Called on setup and can be scheduled via time-trigger
  try {
    const dashData = JSON.parse(getDashboardData());
    if (!dashData.success) return;
    const sheet = getSheet(SHEETS.DASHBOARD_DATA);
    sheet.clearContents();
    sheet.appendRow(['Metric', 'Value', 'Updated_At']);
    const stats = dashData.data.stats;
    Object.entries(stats).forEach(([k, v]) => {
      sheet.appendRow([k, v, new Date()]);
    });
  } catch(e) { /* silent */ }
}
