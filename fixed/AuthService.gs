// ============================================================
// AuthService.gs — Authentication & Session Management
// FIXED: B-007 session boolean check handles Sheets "TRUE"
//        B-013 CacheService session lookup + inactive session purge
//        B-020 XFrameOptionsMode note added (must be set in Code.gs)
//        B-021 removed attempt count from error message
// ============================================================

const ALL_ROLES   = ['Super Admin', 'Admin', 'Purchase', 'Store', 'Production', 'Dispatch', 'Viewer'];
const ADMIN_ROLES = ['Super Admin', 'Admin'];

const PERMISSIONS = {
  'UserManagement':    ADMIN_ROLES,
  'AuditLogs':         ADMIN_ROLES,
  'Settings':          ADMIN_ROLES,
  'BuyerMaster':       ['Super Admin', 'Admin', 'Purchase', 'Production', 'Viewer'],
  'ArticleMaster':     ['Super Admin', 'Admin', 'Purchase', 'Production', 'Viewer'],
  'BOM':               ALL_ROLES,
  'FabricMaster':      ALL_ROLES,
  'VendorMaster':      ['Super Admin', 'Admin', 'Purchase', 'Store', 'Viewer'],
  'SalesOrder':        ALL_ROLES,
  'SalesOrders':       ALL_ROLES,
  'PurchaseOrder':     ['Super Admin', 'Admin', 'Purchase', 'Store', 'Viewer'],
  'PurchaseOrders':    ['Super Admin', 'Admin', 'Purchase', 'Store', 'Viewer'],
  // FIX: ApprovePO also needs a READ permission entry so UI can check it
  'ApprovePO':         ADMIN_ROLES,
  'GRN':               ['Super Admin', 'Admin', 'Purchase', 'Store', 'Viewer'],
  'Inventory':         ALL_ROLES,
  'FabricIssue':       ['Super Admin', 'Admin', 'Store', 'Production', 'Viewer'],
  'InventoryLedger':   ALL_ROLES,
  'ProductionPlan':    ['Super Admin', 'Admin', 'Production', 'Viewer'],
  'production':        ['Super Admin', 'Admin', 'Production', 'Viewer'],
  'BundleManagement':  ['Super Admin', 'Admin', 'Production', 'Viewer'],
  'Dispatch':          ['Super Admin', 'Admin', 'Dispatch', 'Viewer'],
  'dispatch':          ['Super Admin', 'Admin', 'Dispatch', 'Viewer'],
  'FabricatorChallan': ['Super Admin', 'Admin', 'Store', 'Production', 'Viewer'],
  'Returns':           ALL_ROLES,
  'Reports':           ALL_ROLES
};

const WRITE_PERMISSIONS = {
  'UserManagement':       ADMIN_ROLES,
  'Settings':             ADMIN_ROLES,
  'BuyerMaster':          ADMIN_ROLES,
  'ArticleMaster':        ADMIN_ROLES,
  'BOM':                  ADMIN_ROLES,
  'FabricMaster':         ['Super Admin', 'Admin', 'Purchase'],
  'VendorMaster':         ['Super Admin', 'Admin', 'Purchase'],
  'SalesOrder':           ADMIN_ROLES,
  'SalesOrders':          ADMIN_ROLES,
  'PurchaseOrder':        ['Super Admin', 'Admin', 'Purchase'],
  'PurchaseOrders':       ['Super Admin', 'Admin', 'Purchase'],
  'ApprovePO':            ADMIN_ROLES,
  'PurchaseOrderApprove': ADMIN_ROLES,
  'GRN':                  ['Super Admin', 'Admin', 'Store'],
  'Inventory':            ['Super Admin', 'Admin', 'Store'],
  'FabricIssue':          ['Super Admin', 'Admin', 'Store'],
  'ProductionPlan':       ['Super Admin', 'Admin', 'Production'],
  'production':           ['Super Admin', 'Admin', 'Production'],
  'BundleManagement':     ['Super Admin', 'Admin', 'Production'],
  'FabricatorChallan':    ['Super Admin', 'Admin', 'Store', 'Production'],
  'Returns':              ['Super Admin', 'Admin', 'Store', 'Production', 'Dispatch'],
  'Dispatch':             ['Super Admin', 'Admin', 'Dispatch'],
  'dispatch':             ['Super Admin', 'Admin', 'Dispatch']
};

// Session cache TTL: 1 hour (in seconds)
const SESSION_CACHE_TTL = 3600;

// ── Login ──────────────────────────────────────────────────

function login(email, password) {
  try {
    if (!email || !password) {
      return errorResponse('Email and password are required', 'MISSING_CREDENTIALS');
    }

    email = String(email).toLowerCase().trim();
    const passwordHash = hashPassword(password);

    const ss         = getSpreadsheet();
    const usersSheet = ss.getSheetByName('_Users');
    if (!usersSheet) return errorResponse('System error: Users sheet not found', 'SYSTEM_ERROR');

    const lastRow = usersSheet.getLastRow();
    if (lastRow < 2) return errorResponse('Invalid email or password', 'AUTH_FAILED');

    const userData = usersSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    let userRow      = null;
    let userRowIndex = -1;

    for (let i = 0; i < userData.length; i++) {
      if (String(userData[i][2]).toLowerCase() === email) {
        userRow      = userData[i];
        userRowIndex = i + 2;
        break;
      }
    }

    if (!userRow) return errorResponse('Invalid email or password', 'AUTH_FAILED');

    if (userRow[5] === 'Locked')   return errorResponse('Account is locked. Contact your administrator.', 'ACCOUNT_LOCKED');
    if (userRow[5] === 'Inactive') return errorResponse('Account is inactive. Contact your administrator.', 'ACCOUNT_INACTIVE');

    if (userRow[3] !== passwordHash) {
      const failedAttempts = (parseInt(userRow[7]) || 0) + 1;
      usersSheet.getRange(userRowIndex, 8).setValue(failedAttempts);
      if (failedAttempts >= 5) {
        usersSheet.getRange(userRowIndex, 6).setValue('Locked');
        return errorResponse('Account locked after too many failed attempts. Contact administrator.', 'ACCOUNT_LOCKED');
      }
      // FIX B-021: Do NOT reveal remaining attempt count
      return errorResponse('Invalid email or password', 'AUTH_FAILED');
    }

    // Reset failed attempts
    usersSheet.getRange(userRowIndex, 8).setValue(0);
    usersSheet.getRange(userRowIndex, 7).setValue(nowISO());

    const sessionToken = generateToken();
    const sessionID    = generateSessionID();
    const now          = nowISO();

    const sessionsSheet = ss.getSheetByName('_Sessions');
    if (sessionsSheet) {
      sessionsSheet.appendRow([sessionID, userRow[0], sessionToken, now, now, true, '']);
    }

    writeAuditLog({
      userID: userRow[0], userName: userRow[1], action: 'LOGIN',
      module: 'Authentication', recordID: userRow[0],
      oldValues: '', newValues: JSON.stringify({ email: email, timestamp: now })
    });

    return successResponse({
      token: sessionToken,
      sessionID: sessionID,
      user: { userID: userRow[0], fullName: userRow[1], email: userRow[2], role: userRow[4] }
    }, 'Login successful');

  } catch (e) {
    Logger.log('Login error: ' + e.toString());
    return errorResponse('Login failed: ' + e.message, 'SYSTEM_ERROR');
  }
}

// ── Logout ─────────────────────────────────────────────────

function logout(token) {
  try {
    if (!token) return errorResponse('No session token', 'NO_TOKEN');

    const session = getSession(token);
    if (!session) return errorResponse('Session not found', 'SESSION_NOT_FOUND');

    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('_Sessions');
    if (!sheet) return errorResponse('System error', 'SYSTEM_ERROR');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return errorResponse('Session not found', 'SESSION_NOT_FOUND');

    const tokens   = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
    const rowIndex = tokens.indexOf(token);
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 2, 6).setValue(false);
    }

    // FIX B-013: Evict from cache on logout
    try {
      CacheService.getScriptCache().remove('sess_' + token);
    } catch (cacheErr) { /* non-fatal */ }

    writeAuditLog({
      userID: session.userID, userName: session.userName, action: 'LOGOUT',
      module: 'Authentication', recordID: session.userID,
      oldValues: '', newValues: JSON.stringify({ timestamp: nowISO() })
    });

    return successResponse({}, 'Logged out successfully');
  } catch (e) {
    Logger.log('Logout error: ' + e.toString());
    return errorResponse('Logout failed', 'SYSTEM_ERROR');
  }
}

// ── Session Validation ─────────────────────────────────────

function validateSession(token) {
  if (!token) return null;
  const session = getSession(token);
  if (!session) return null;
  updateSessionActivity(token);
  return session;
}

// FIX B-007 + B-013: Fix boolean check + add CacheService
function getSession(token) {
  try {
    if (!token) return null;

    // FIX B-013: Check cache first — avoids full sheet scan on every request
    const cache  = CacheService.getScriptCache();
    const cached = cache.get('sess_' + token);
    if (cached) {
      const parsed = JSON.parse(cached);
      // Still verify user is not locked/inactive
      const user = getUserByID(parsed.userID);
      if (!user || user.Status === 'Inactive' || user.Status === 'Locked') {
        cache.remove('sess_' + token);
        return null;
      }
      return parsed;
    }

    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('_Sessions');
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][2] === token) {
        // FIX B-007: Handles true (boolean), "true" (lowercase), "TRUE" (Sheets uppercase)
        const isActive = String(data[i][5]).toLowerCase() === 'true';
        if (!isActive) return null;

        const userID = data[i][1];
        const user   = getUserByID(userID);
        if (!user) return null;
        if (user.Status === 'Inactive' || user.Status === 'Locked') return null;

        const session = {
          sessionID: data[i][0],
          userID:    data[i][1],
          token:     data[i][2],
          userName:  user.FullName,
          role:      user.Role,
          email:     user.Email
        };

        // FIX B-013: Store in cache to skip sheet scan next time
        cache.put('sess_' + token, JSON.stringify(session), SESSION_CACHE_TTL);
        return session;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getSession error: ' + e.toString());
    return null;
  }
}

function updateSessionActivity(token) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('_Sessions');
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const tokens   = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
    const rowIndex = tokens.indexOf(token);
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 2, 5).setValue(nowISO());
    }
  } catch (e) {
    Logger.log('updateSessionActivity error: ' + e.toString());
  }
}

// FIX B-013: Purge old inactive sessions — call via a weekly time-based trigger
function purgeOldSessions() {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('_Sessions');
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const cutoff = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000); // 30 days ago

    // Collect rows to delete (bottom-up to preserve indices)
    const toDelete = [];
    for (let i = data.length - 1; i >= 0; i--) {
      const isActive    = String(data[i][5]).toLowerCase() === 'true';
      const lastActivity = new Date(data[i][4]);
      if (!isActive && (!isNaN(lastActivity.getTime()) && lastActivity < cutoff)) {
        toDelete.push(i + 2); // 1-indexed
      }
    }
    toDelete.forEach(function(row) { sheet.deleteRow(row); });
    Logger.log('purgeOldSessions: removed ' + toDelete.length + ' stale sessions');
  } catch (e) {
    Logger.log('purgeOldSessions error: ' + e.toString());
  }
}

// ── Permission Check ───────────────────────────────────────

function hasReadPermission(session, module) {
  if (!session) return false;
  const allowed = PERMISSIONS[module];
  if (!allowed) return false;
  return allowed.indexOf(session.role) !== -1;
}

function hasWritePermission(session, module) {
  if (!session) return false;
  const allowed = WRITE_PERMISSIONS[module];
  if (!allowed) return false;
  return allowed.indexOf(session.role) !== -1;
}

function requireSession(token) {
  const session = validateSession(token);
  if (!session) return { valid: false, error: errorResponse('Session expired or invalid. Please login again.', 'SESSION_INVALID') };
  return { valid: true, session: session };
}

function requireReadAccess(token, module) {
  const { valid, session, error } = requireSession(token);
  if (!valid) return { valid: false, error };
  if (!hasReadPermission(session, module)) {
    return { valid: false, error: errorResponse('Access denied: Insufficient permissions for ' + module, 'ACCESS_DENIED') };
  }
  return { valid: true, session };
}

function requireWriteAccess(token, module) {
  const { valid, session, error } = requireSession(token);
  if (!valid) return { valid: false, error };
  if (!hasWritePermission(session, module)) {
    return { valid: false, error: errorResponse('Access denied: You do not have write permission for ' + module, 'ACCESS_DENIED') };
  }
  return { valid: true, session };
}

// ── User Helper ────────────────────────────────────────────

function getUserByID(userID) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName('_Users');
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === userID) {
      return { UserID: data[i][0], FullName: data[i][1], Email: data[i][2], Role: data[i][4], Status: data[i][5] };
    }
  }
  return null;
}

// ── Change Password ────────────────────────────────────────

function changePassword(token, currentPassword, newPassword) {
  try {
    const { valid, session, error } = requireSession(token);
    if (!valid) return error;

    if (!currentPassword || !newPassword) return errorResponse('Current and new passwords are required', 'MISSING_FIELDS');
    if (newPassword.length < 8) return errorResponse('New password must be at least 8 characters', 'WEAK_PASSWORD');

    const ss      = getSpreadsheet();
    const sheet   = ss.getSheetByName('_Users');
    const lastRow = sheet.getLastRow();
    const data    = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    let userRowIndex = -1;
    let currentHash  = '';

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === session.userID) {
        userRowIndex = i + 2;
        currentHash  = data[i][3];
        break;
      }
    }

    if (userRowIndex === -1) return errorResponse('User not found', 'USER_NOT_FOUND');
    if (currentHash !== hashPassword(currentPassword)) return errorResponse('Current password is incorrect', 'WRONG_PASSWORD');

    sheet.getRange(userRowIndex, 4).setValue(hashPassword(newPassword));
    sheet.getRange(userRowIndex, 12).setValue(nowISO());

    writeAuditLog({ userID: session.userID, userName: session.userName, action: 'UPDATE', module: 'UserManagement', recordID: session.userID, oldValues: '', newValues: JSON.stringify({ action: 'Password changed' }) });
    return successResponse({}, 'Password changed successfully');
  } catch (e) {
    return errorResponse('Change password failed: ' + e.message, 'SYSTEM_ERROR');
  }
}

// ── Get Current User Info ──────────────────────────────────

function getCurrentUser(token) {
  try {
    const { valid, session, error } = requireSession(token);
    if (!valid) return error;
    return successResponse({ userID: session.userID, fullName: session.userName, role: session.role, email: session.email });
  } catch (e) {
    return errorResponse('Failed to get user info', 'SYSTEM_ERROR');
  }
}
