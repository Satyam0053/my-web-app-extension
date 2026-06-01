// ============================================================
// AuditService.gs — Audit Trail
// FIXED: B-017 Log IDs use Utilities.getUuid() (guaranteed unique)
// ============================================================

function writeAuditLog(params) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('_AuditLogs');
    if (!sheet) return;

    // FIX B-017: Use UUID instead of Date.now()+random (eliminates duplicate risk)
    const logID = 'LOG-' + Utilities.getUuid();
    const row = [
      logID,
      new Date().toISOString(),
      params.userID   || '',
      params.userName || '',
      params.action   || 'UNKNOWN',
      params.module   || '',
      params.recordID || '',
      typeof params.oldValues === 'object' ? JSON.stringify(params.oldValues) : (params.oldValues || ''),
      typeof params.newValues === 'object' ? JSON.stringify(params.newValues) : (params.newValues || ''),
      params.ipAddress || ''
    ];
    sheet.appendRow(row);
  } catch (e) {
    Logger.log('AuditLog write error: ' + e.toString());
    // Never throw — audit failure must not block business operations
  }
}

function getAuditLogs(token, filters) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'AuditLogs');
    if (!valid) return error;

    const data = getSheetData('_AuditLogs');
    let filtered = data;

    if (filters) {
      if (filters.module)   filtered = filtered.filter(function(r) { return r.Module === filters.module; });
      if (filters.action)   filtered = filtered.filter(function(r) { return r.Action === filters.action; });
      if (filters.userID)   filtered = filtered.filter(function(r) { return r.UserID === filters.userID; });
      if (filters.fromDate) {
        const from = new Date(filters.fromDate);
        filtered = filtered.filter(function(r) { return new Date(r.Timestamp) >= from; });
      }
      if (filters.toDate) {
        const to = new Date(filters.toDate);
        to.setHours(23, 59, 59, 999);
        filtered = filtered.filter(function(r) { return new Date(r.Timestamp) <= to; });
      }
    }

    filtered.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });

    const page     = (filters && filters.page)     ? parseInt(filters.page)     : 1;
    const pageSize = (filters && filters.pageSize) ? parseInt(filters.pageSize) : 50;
    const total    = filtered.length;
    const start    = (page - 1) * pageSize;
    const paged    = filtered.slice(start, start + pageSize);

    return successResponse({ logs: paged, total: total, page: page, pageSize: pageSize, totalPages: Math.ceil(total / pageSize) });
  } catch (e) {
    return errorResponse('Failed to fetch audit logs: ' + e.message, 'FETCH_ERROR');
  }
}
