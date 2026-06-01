// ============================================================
// INVENTORY ERP — GOOGLE APPS SCRIPT BACKEND
// File: Auth.gs — Authentication & User Management
// Version: 2.2 | Bug-Fixed
//
// ARCHITECTURE:
//   • Credentials stored in Script Properties as JSON
//   • Passwords hashed with SHA-256 before storing
//   • Session tokens stored in Script Properties (server-side)
//   • Sessions never expire — persist until manual logout
//   • Admin-only actions enforced server-side via requireAdmin()
//
// ROLES: Admin | Manager | EntryOP | User | PC
//
// FIXES v2.2:
//   • loginUser: wrapped logActivity in try/catch so a sheet error
//     never blocks login and never returns a false "Invalid password"
//   • loginUser: removed .trim() from password before hashing so
//     passwords with intentional leading/trailing spaces work correctly
//   • setupUsers_seed: same — password not trimmed before sha256
//   • Added verifyUser() debug helper (run from editor)
// ============================================================

// ── CONSTANTS ────────────────────────────────────────────────
const PROP_USERS_KEY    = 'ERP_USERS';
const PROP_SESSIONS_KEY = 'ERP_SESSIONS';
const VALID_ROLES       = ['Admin', 'Manager', 'EntryOP', 'User', 'PC'];

// ── ROLE GUARD ───────────────────────────────────────────────
function requireAdmin(caller) {
  if (!caller || caller.role !== 'Admin') {
    return error('PERMISSION_DENIED', 'This action requires Admin role.');
  }
  return null;
}

// ── SHA-256 HASH ─────────────────────────────────────────────
function sha256(input) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    input,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ── TOKEN GENERATOR ──────────────────────────────────────────
function generateToken() {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Math.random().toString() + new Date().getTime().toString() + Math.random().toString(),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ── USER STORE HELPERS ───────────────────────────────────────
function _getUsers() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(PROP_USERS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    Logger.log('_getUsers ERROR: ' + e.message);
    return {};
  }
}

function _saveUsers(usersObj) {
  PropertiesService.getScriptProperties().setProperty(
    PROP_USERS_KEY, JSON.stringify(usersObj)
  );
}

// ── SESSION STORE HELPERS ────────────────────────────────────
function _getSessions() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(PROP_SESSIONS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    Logger.log('_getSessions ERROR: ' + e.message);
    return {};
  }
}

function _saveSessions(sessionsObj) {
  PropertiesService.getScriptProperties().setProperty(
    PROP_SESSIONS_KEY, JSON.stringify(sessionsObj)
  );
}

// ── KILL ALL SESSIONS FOR A USER ─────────────────────────────
function _killUserSessions(username) {
  const sessions = _getSessions();
  let changed = false;
  Object.keys(sessions).forEach(token => {
    if (sessions[token].username === username) {
      delete sessions[token];
      changed = true;
    }
  });
  if (changed) _saveSessions(sessions);
}

// ── SAFE USER OBJECT (no password hash) ──────────────────────
function _safeUser(u) {
  return {
    username   : u.username,
    displayName: u.displayName,
    role       : u.role,
    active     : u.active,
    createdAt  : u.createdAt,
    updatedAt  : u.updatedAt || null
  };
}

// ============================================================
// AUTH FUNCTIONS
// ============================================================

// ── LOGIN ────────────────────────────────────────────────────
// FIX v2.2:
//   1. Password is NOT trimmed — trimming before hashing would
//      cause mismatch if the stored hash was made without trimming.
//      The frontend should trim UI whitespace; the backend hashes exactly
//      what it receives.
//   2. logActivity is wrapped in its own try/catch so a Sheets
//      error never causes a false "Invalid username or password".
//   3. Detailed Logger.log added for server-side debugging.
function loginUser(payload) {
  try {
    // username: lowercase + trim (safe — usernames are never intentionally spaced)
    // password: DO NOT trim — hash must match exactly what was stored at creation
    const username = (payload.username || '').trim().toLowerCase();
    const password = (payload.password || ''); // ← no .trim() here

    Logger.log('LOGIN ATTEMPT — username: "' + username + '"');

    if (!username || !password) {
      return error('Username and password are required');
    }

    const users = _getUsers();
    Logger.log('Users loaded — count: ' + Object.keys(users).length);

    const user = users[username];

    if (!user) {
      Logger.log('LOGIN FAIL — user not found: "' + username + '"');
      return error('Invalid username or password');
    }

    if (!user.active) {
      Logger.log('LOGIN FAIL — account deactivated: "' + username + '"');
      return error('This account has been deactivated. Contact your administrator.');
    }

    const incomingHash = sha256(password);
    Logger.log('Hash check — stored: ' + user.passwordHash + ' | incoming: ' + incomingHash);

    if (incomingHash !== user.passwordHash) {
      Logger.log('LOGIN FAIL — wrong password for: "' + username + '"');
      // FIX: wrapped logActivity so sheet errors don't propagate
      try {
        logActivity('LOGIN_FAIL', 'Auth', username, { reason: 'Wrong password' });
      } catch (logErr) {
        Logger.log('logActivity error (non-fatal): ' + logErr.message);
      }
      return error('Invalid username or password');
    }

    // ── PASSWORD MATCHED — create session ─────────────────────
    const token     = generateToken();
    const sessions  = _getSessions();
    sessions[token] = {
      username   : user.username,
      displayName: user.displayName,
      role       : user.role,
      createdAt  : new Date().toISOString()
    };
    _saveSessions(sessions);

    // FIX: wrapped logActivity so sheet errors don't prevent login
    try {
      logActivity('LOGIN', 'Auth', username, { role: user.role }, user.username, user.role);
    } catch (logErr) {
      Logger.log('logActivity error (non-fatal): ' + logErr.message);
    }

    Logger.log('LOGIN SUCCESS — username: "' + username + '" role: "' + user.role + '"');

    return success({
      token      : token,
      username   : user.username,
      displayName: user.displayName,
      role       : user.role
    }, 'Welcome, ' + user.displayName + '!');

  } catch (e) {
    Logger.log('loginUser UNEXPECTED ERROR: ' + e.message + '\n' + e.stack);
    return error('Login failed due to a server error. Please try again.');
  }
}

// ── VALIDATE SESSION TOKEN ───────────────────────────────────
function validateSessionToken(token) {
  if (!token) return error('No token provided');

  const sessions = _getSessions();
  const session  = sessions[token];
  if (!session) return error('Session not found. Please log in again.');

  const users = _getUsers();
  const user  = users[session.username];
  if (!user || !user.active) {
    delete sessions[token];
    _saveSessions(sessions);
    return error('Account not found or deactivated.');
  }

  return success({
    username   : user.username,
    displayName: user.displayName,
    role       : user.role
  }, 'Token valid');
}

// ── LOGOUT ───────────────────────────────────────────────────
function logoutUser(payload, caller) {
  const token    = payload._token;
  const sessions = _getSessions();
  if (token && sessions[token]) {
    delete sessions[token];
    _saveSessions(sessions);
  }
  const u = caller || {};
  try {
    logActivity('LOGOUT', 'Auth', u.username || 'unknown', {}, u.username, u.role);
  } catch(e) {}
  return success(null, 'Logged out successfully');
}

// ── CHANGE OWN PASSWORD ──────────────────────────────────────
function changePassword(payload, caller) {
  const currentPassword = (payload.currentPassword || '');
  const newPassword     = (payload.newPassword     || '');
  const confirmPassword = (payload.confirmPassword || '');

  if (!currentPassword || !newPassword || !confirmPassword)
    return error('All password fields are required');
  if (newPassword !== confirmPassword)
    return error('New password and confirmation do not match');
  if (newPassword.length < 6)
    return error('Password must be at least 6 characters');

  const users = _getUsers();
  const user  = users[caller.username];
  if (!user) return error('User not found');

  if (sha256(currentPassword) !== user.passwordHash)
    return error('Current password is incorrect');

  users[caller.username].passwordHash = sha256(newPassword);
  users[caller.username].updatedAt    = new Date().toISOString();
  _saveUsers(users);

  try {
    logActivity('CHANGE_PASSWORD', 'Auth', caller.username, {}, caller.username, caller.role);
  } catch(e) {}
  return success(null, 'Password changed successfully');
}

// ============================================================
// USER MANAGEMENT — ADMIN ONLY
// ============================================================

function getUsers(caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const users = _getUsers();
  const list  = Object.values(users)
    .map(_safeUser)
    .sort((a, b) => a.username.localeCompare(b.username));
  return success(list);
}

function addUser(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username    = (payload.username    || '').trim().toLowerCase();
  const displayName = (payload.displayName || '').trim();
  const role        = (payload.role        || '').trim();
  const password    = (payload.password    || '');

  if (!username)
    return error('Username is required');
  if (!/^[a-z0-9_]{3,30}$/.test(username))
    return error('Username must be 3-30 characters: lowercase letters, numbers, underscore only');
  if (!displayName)
    return error('Display name is required');
  if (!VALID_ROLES.includes(role))
    return error('Invalid role. Must be one of: ' + VALID_ROLES.join(', '));
  if (!password || password.length < 6)
    return error('Password must be at least 6 characters');

  const users = _getUsers();
  if (users[username])
    return error('Username "' + username + '" already exists. Choose a different username.');

  users[username] = {
    username    : username,
    displayName : displayName,
    role        : role,
    passwordHash: sha256(password),
    active      : true,
    createdAt   : new Date().toISOString(),
    createdBy   : caller.username,
    updatedAt   : null
  };
  _saveUsers(users);

  try {
    logActivity('ADD_USER', 'UserManagement', username,
      { role, displayName, createdBy: caller.username },
      caller.username, caller.role);
  } catch(e) {}

  return success(_safeUser(users[username]),
    'User "' + username + '" created successfully'
  );
}

function updateUser(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username = (payload.username || '').trim().toLowerCase();
  if (!username) return error('Username is required');

  const users = _getUsers();
  const user  = users[username];
  if (!user) return error('User not found: ' + username);

  if (username === caller.username && payload.role && payload.role !== 'Admin')
    return error('You cannot change your own role. Ask another Admin.');

  const oldRole = user.role;
  if (payload.displayName && payload.displayName.trim())
    users[username].displayName = payload.displayName.trim();
  if (payload.role && VALID_ROLES.includes(payload.role))
    users[username].role = payload.role;
  users[username].updatedAt = new Date().toISOString();

  _saveUsers(users);

  if (payload.role && payload.role !== oldRole) {
    _killUserSessions(username);
  }

  try {
    logActivity('UPDATE_USER', 'UserManagement', username,
      { oldRole, newRole: users[username].role, displayName: users[username].displayName },
      caller.username, caller.role);
  } catch(e) {}

  return success(_safeUser(users[username]), 'User updated successfully');
}

function deactivateUserAPI(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username = (payload.username || '').trim().toLowerCase();
  if (!username) return error('Username is required');
  if (username === caller.username)
    return error('You cannot deactivate your own account.');

  const users = _getUsers();
  const user  = users[username];
  if (!user) return error('User not found: ' + username);
  if (!user.active) return error('User is already deactivated.');

  users[username].active    = false;
  users[username].updatedAt = new Date().toISOString();
  _saveUsers(users);

  _killUserSessions(username);

  try {
    logActivity('DEACTIVATE_USER', 'UserManagement', username,
      { deactivatedBy: caller.username },
      caller.username, caller.role);
  } catch(e) {}

  return success({ username, active: false },
    'User "' + username + '" deactivated. All sessions terminated.'
  );
}

function reactivateUserAPI(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username = (payload.username || '').trim().toLowerCase();
  if (!username) return error('Username is required');

  const users = _getUsers();
  const user  = users[username];
  if (!user) return error('User not found: ' + username);
  if (user.active) return error('User is already active.');

  users[username].active    = true;
  users[username].updatedAt = new Date().toISOString();
  _saveUsers(users);

  try {
    logActivity('REACTIVATE_USER', 'UserManagement', username,
      { reactivatedBy: caller.username },
      caller.username, caller.role);
  } catch(e) {}

  return success({ username, active: true },
    'User "' + username + '" reactivated. They can now log in.'
  );
}

function adminResetPassword(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username        = (payload.username        || '').trim().toLowerCase();
  const newPassword     = (payload.newPassword     || '');
  const confirmPassword = (payload.confirmPassword || '');

  if (!username)
    return error('Username is required');
  if (!newPassword || newPassword.length < 6)
    return error('New password must be at least 6 characters');
  if (newPassword !== confirmPassword)
    return error('Passwords do not match');

  const users = _getUsers();
  const user  = users[username];
  if (!user) return error('User not found: ' + username);

  users[username].passwordHash = sha256(newPassword);
  users[username].updatedAt    = new Date().toISOString();
  _saveUsers(users);

  _killUserSessions(username);

  try {
    logActivity('ADMIN_RESET_PASSWORD', 'UserManagement', username,
      { resetBy: caller.username },
      caller.username, caller.role);
  } catch(e) {}

  return success({ username },
    'Password reset for "' + username + '". All their sessions have been terminated.'
  );
}

// ============================================================
// SETUP UTILITIES — Run from Apps Script Editor
// ============================================================

// ── SETUP USERS (ADDITIVE) ───────────────────────────────────
// FIX v2.2: Password NOT trimmed before hashing — must be consistent
// with how loginUser hashes the incoming password.
function setupUsers(payload, caller) {
  const incoming = payload.users || [];
  if (!Array.isArray(incoming) || incoming.length === 0)
    return error('No users provided in payload.users array');

  const users  = _getUsers();
  const report = { added: [], skipped: [], updated: [], errors: [] };

  incoming.forEach(entry => {
    const username = (entry.username || '').trim().toLowerCase();
    if (!username) { report.errors.push('Blank username skipped'); return; }

    if (!VALID_ROLES.includes(entry.role)) {
      report.errors.push(username + ': invalid role "' + entry.role + '"');
      return;
    }

    if (users[username] && !entry.forceUpdate) {
      report.skipped.push(username);
      return;
    }

    if (users[username] && entry.forceUpdate) {
      users[username].role        = entry.role;
      users[username].displayName = entry.displayName || users[username].displayName;
      users[username].updatedAt   = new Date().toISOString();
      // FIX: no .trim() on password before hashing
      if (entry.password && entry.password.length >= 6)
        users[username].passwordHash = sha256(entry.password);
      report.updated.push(username);
      return;
    }

    if (!entry.password || entry.password.length < 6) {
      report.errors.push(username + ': password must be at least 6 characters');
      return;
    }

    users[username] = {
      username    : username,
      displayName : entry.displayName || username,
      role        : entry.role,
      // FIX: no .trim() on password — hash the raw string
      passwordHash: sha256(entry.password),
      active      : true,
      createdAt   : new Date().toISOString(),
      createdBy   : caller ? caller.username : 'seed',
      updatedAt   : null
    };
    report.added.push(username);
  });

  _saveUsers(users);

  const callerName = caller ? caller.username : 'script';
  const callerRole = caller ? caller.role     : 'System';
  try {
    logActivity('SETUP_USERS', 'Auth', 'batch', report, callerName, callerRole);
  } catch(e) {}

  return success(report,
    'Done — Added: '   + report.added.length +
    ', Updated: '      + report.updated.length +
    ', Skipped: '      + report.skipped.length +
    ', Errors: '       + report.errors.length
  );
}

// ── SEED — Run once from Apps Script Editor ──────────────────
function setupUsers_seed() {
  const result = JSON.parse(setupUsers({
    users: [
      { username:'admin',    displayName:'Administrator',    role:'Admin',   password:'Admin@1234'   },
      { username:'manager1', displayName:'Manager',          role:'Manager', password:'Manager@1234' },
      { username:'entryop1', displayName:'Entry Operator',   role:'EntryOP', password:'Entry@1234'   },
      { username:'user1',    displayName:'General User',     role:'User',    password:'User@1234'    },
      { username:'pc1',      displayName:'Process Coordinator', role:'PC',   password:'PC@1234'      }
    ]
  }, null));
  Logger.log('setupUsers_seed result:\n' + JSON.stringify(result, null, 2));
}

// ── VERIFY USER (Debug helper — run from editor) ─────────────
// Run this from the editor to confirm a user exists and their
// password hash is correct, without needing to go through the UI.
function verifyUser_debug() {
  const username = 'admin';       // ← change as needed
  const password = 'Admin@1234';  // ← change as needed

  const users = _getUsers();
  Logger.log('Total users stored: ' + Object.keys(users).length);
  Logger.log('Stored usernames: ' + Object.keys(users).join(', '));

  const user = users[username];
  if (!user) {
    Logger.log('❌ User "' + username + '" NOT FOUND. Run setupUsers_seed() first.');
    return;
  }

  Logger.log('User found: ' + JSON.stringify({
    username   : user.username,
    displayName: user.displayName,
    role       : user.role,
    active     : user.active,
    storedHash : user.passwordHash
  }));

  const incomingHash = sha256(password);
  Logger.log('Hash of test password "' + password + '": ' + incomingHash);

  if (incomingHash === user.passwordHash) {
    Logger.log('✅ Password MATCHES — login should work.');
  } else {
    Logger.log('❌ Password MISMATCH — run setupUsers_seed() with forceUpdate:true to reset.');
  }
}

// ── RE-SEED WITH FORCE UPDATE (run if hashes are corrupt) ───
// Use this if verifyUser_debug() shows a hash mismatch.
// It will overwrite ALL user passwords back to defaults.
function resetAllUsers_seed() {
  const result = JSON.parse(setupUsers({
    users: [
      { username:'admin',    displayName:'Administrator',    role:'Admin',   password:'Admin@1234',   forceUpdate:true },
      { username:'manager1', displayName:'Manager',          role:'Manager', password:'Manager@1234', forceUpdate:true },
      { username:'entryop1', displayName:'Entry Operator',   role:'EntryOP', password:'Entry@1234',   forceUpdate:true },
      { username:'user1',    displayName:'General User',     role:'User',    password:'User@1234',    forceUpdate:true },
      { username:'pc1',      displayName:'Process Coordinator', role:'PC',   password:'PC@1234',      forceUpdate:true }
    ]
  }, null));
  Logger.log('resetAllUsers_seed result:\n' + JSON.stringify(result, null, 2));
}

// ── LIST USERS (Editor utility) ──────────────────────────────
function listUsers_log() {
  const users = _getUsers();
  Logger.log(JSON.stringify(Object.values(users).map(_safeUser), null, 2));
}

// ── CLEAR ALL SESSIONS (Emergency, Editor only) ──────────────
function clearAllSessions_admin() {
  PropertiesService.getScriptProperties().setProperty(PROP_SESSIONS_KEY, '{}');
  Logger.log('All sessions cleared.');
}

// ── DISPATCH LOGIN DEBUG (Editor only) ───────────────────────
function debugDispatchLogin() {
  try {
    const result = dispatch('login', JSON.stringify({
      username: 'admin',
      password: 'Admin@1234'
    }));
    Logger.log('Result: ' + result);
  } catch(e) {
    Logger.log('Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
}