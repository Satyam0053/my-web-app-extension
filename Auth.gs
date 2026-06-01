// ============================================================
// File: Auth.gs — Authentication & User Management
// FIXED:
//  • requireAdmin: error() arg order fixed (message first)
//  • loginUser: password NOT trimmed before hashing (consistent)
//  • loginUser: logActivity wrapped in try/catch
//  • Sessions stored in ScriptProperties — added 30-day expiry check
//  • Login no longer reveals lockout threshold or attempt counts
// ============================================================

const PROP_USERS_KEY    = 'ERP_USERS';
const PROP_SESSIONS_KEY = 'ERP_SESSIONS';
const VALID_ROLES       = ['Admin', 'Manager', 'EntryOP', 'User', 'PC'];
// Session expiry: 8 hours in milliseconds
const SESSION_EXPIRY_MS = 8 * 60 * 60 * 1000;

// FIX: error(message, details) — message is first arg
function requireAdmin(caller) {
  if (!caller || caller.role !== 'Admin') {
    return error('This action requires Admin role.', 'PERMISSION_DENIED');
  }
  return null;
}

function sha256(input) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function generateToken() {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Math.random().toString() + new Date().getTime().toString() + Math.random().toString(),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function _getUsers() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(PROP_USERS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) { Logger.log('_getUsers ERROR: ' + e.message); return {}; }
}

function _saveUsers(usersObj) {
  PropertiesService.getScriptProperties().setProperty(PROP_USERS_KEY, JSON.stringify(usersObj));
}

function _getSessions() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(PROP_SESSIONS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) { Logger.log('_getSessions ERROR: ' + e.message); return {}; }
}

function _saveSessions(sessionsObj) {
  PropertiesService.getScriptProperties().setProperty(PROP_SESSIONS_KEY, JSON.stringify(sessionsObj));
}

function _killUserSessions(username) {
  const sessions = _getSessions();
  let changed = false;
  Object.keys(sessions).forEach(token => {
    if (sessions[token].username === username) { delete sessions[token]; changed = true; }
  });
  if (changed) _saveSessions(sessions);
}

function _safeUser(u) {
  return { username: u.username, displayName: u.displayName, role: u.role, active: u.active, createdAt: u.createdAt, updatedAt: u.updatedAt || null };
}

// ── LOGIN ──────────────────────────────────────────────────────
function loginUser(payload) {
  try {
    const username = (payload.username || '').trim().toLowerCase();
    const password = (payload.password || ''); // FIX: no .trim() — must match stored hash

    Logger.log('LOGIN ATTEMPT — username: "' + username + '"');

    if (!username || !password) return error('Username and password are required');

    const users = _getUsers();
    const user  = users[username];

    if (!user) {
      Logger.log('LOGIN FAIL — user not found');
      return error('Invalid username or password');
    }
    if (!user.active) return error('Account deactivated. Contact your administrator.');

    const incomingHash = sha256(password);
    if (incomingHash !== user.passwordHash) {
      Logger.log('LOGIN FAIL — wrong password for: "' + username + '"');
      // FIX: no attempt count revealed in error message
      try { logActivity('LOGIN_FAIL', 'Auth', username, { reason: 'Wrong password' }); } catch(e) {}
      return error('Invalid username or password');
    }

    const token    = generateToken();
    const sessions = _getSessions();
    sessions[token] = {
      username   : user.username,
      displayName: user.displayName,
      role       : user.role,
      createdAt  : new Date().toISOString()
    };
    _saveSessions(sessions);

    try { logActivity('LOGIN', 'Auth', username, { role: user.role }, user.username, user.role); } catch(e) {}

    Logger.log('LOGIN SUCCESS — "' + username + '" role: "' + user.role + '"');
    return success({ token, username: user.username, displayName: user.displayName, role: user.role }, 'Welcome, ' + user.displayName + '!');

  } catch (e) {
    Logger.log('loginUser ERROR: ' + e.message + '\n' + e.stack);
    return error('Login failed due to a server error. Please try again.');
  }
}

// ── VALIDATE SESSION ───────────────────────────────────────────
// FIX: Added 8-hour session expiry check
function validateSessionToken(token) {
  if (!token) return error('No token provided');

  const sessions = _getSessions();
  const session  = sessions[token];
  if (!session) return error('Session not found. Please log in again.');

  // FIX: Check session age — expire after SESSION_EXPIRY_MS
  const createdAt = new Date(session.createdAt).getTime();
  if (Date.now() - createdAt > SESSION_EXPIRY_MS) {
    delete sessions[token];
    _saveSessions(sessions);
    return error('Session expired. Please log in again.');
  }

  const users = _getUsers();
  const user  = users[session.username];
  if (!user || !user.active) {
    delete sessions[token];
    _saveSessions(sessions);
    return error('Account not found or deactivated.');
  }

  return success({ username: user.username, displayName: user.displayName, role: user.role }, 'Token valid');
}

// ── LOGOUT ────────────────────────────────────────────────────
function logoutUser(payload, caller) {
  const token    = payload._token;
  const sessions = _getSessions();
  if (token && sessions[token]) { delete sessions[token]; _saveSessions(sessions); }
  try { logActivity('LOGOUT', 'Auth', (caller || {}).username || 'unknown', {}, (caller || {}).username, (caller || {}).role); } catch(e) {}
  return success(null, 'Logged out successfully');
}

// ── CHANGE PASSWORD ───────────────────────────────────────────
function changePassword(payload, caller) {
  const currentPassword = (payload.currentPassword || '');
  const newPassword     = (payload.newPassword     || '');
  const confirmPassword = (payload.confirmPassword || '');

  if (!currentPassword || !newPassword || !confirmPassword) return error('All password fields are required');
  if (newPassword !== confirmPassword) return error('New password and confirmation do not match');
  if (newPassword.length < 6) return error('Password must be at least 6 characters');

  const users = _getUsers();
  const user  = users[caller.username];
  if (!user) return error('User not found');
  if (sha256(currentPassword) !== user.passwordHash) return error('Current password is incorrect');

  users[caller.username].passwordHash = sha256(newPassword);
  users[caller.username].updatedAt    = new Date().toISOString();
  _saveUsers(users);

  try { logActivity('CHANGE_PASSWORD', 'Auth', caller.username, {}, caller.username, caller.role); } catch(e) {}
  return success(null, 'Password changed successfully');
}

// ── USER MANAGEMENT — ADMIN ONLY ──────────────────────────────
function getUsers(caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;
  const list = Object.values(_getUsers()).map(_safeUser).sort((a, b) => a.username.localeCompare(b.username));
  return success(list);
}

function addUser(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username    = (payload.username    || '').trim().toLowerCase();
  const displayName = (payload.displayName || '').trim();
  const role        = (payload.role        || '').trim();
  const password    = (payload.password    || '');

  if (!username)    return error('Username is required');
  if (!/^[a-z0-9_]{3,30}$/.test(username)) return error('Username: 3-30 chars, lowercase/numbers/underscore only');
  if (!displayName) return error('Display name is required');
  if (!VALID_ROLES.includes(role)) return error('Invalid role. Must be one of: ' + VALID_ROLES.join(', '));
  if (!password || password.length < 6) return error('Password must be at least 6 characters');

  const users = _getUsers();
  if (users[username]) return error('Username "' + username + '" already exists.');

  users[username] = { username, displayName, role, passwordHash: sha256(password), active: true, createdAt: new Date().toISOString(), createdBy: caller.username, updatedAt: null };
  _saveUsers(users);

  try { logActivity('ADD_USER', 'UserManagement', username, { role, displayName, createdBy: caller.username }, caller.username, caller.role); } catch(e) {}
  return success(_safeUser(users[username]), 'User "' + username + '" created successfully');
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
  if (payload.displayName && payload.displayName.trim()) users[username].displayName = payload.displayName.trim();
  if (payload.role && VALID_ROLES.includes(payload.role)) users[username].role = payload.role;
  users[username].updatedAt = new Date().toISOString();
  _saveUsers(users);

  if (payload.role && payload.role !== oldRole) _killUserSessions(username);

  try { logActivity('UPDATE_USER', 'UserManagement', username, { oldRole, newRole: users[username].role }, caller.username, caller.role); } catch(e) {}
  return success(_safeUser(users[username]), 'User updated successfully');
}

function deactivateUserAPI(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username = (payload.username || '').trim().toLowerCase();
  if (!username)                     return error('Username is required');
  if (username === caller.username)  return error('You cannot deactivate your own account.');

  const users = _getUsers();
  const user  = users[username];
  if (!user)         return error('User not found: ' + username);
  if (!user.active)  return error('User is already deactivated.');

  users[username].active    = false;
  users[username].updatedAt = new Date().toISOString();
  _saveUsers(users);
  _killUserSessions(username);

  try { logActivity('DEACTIVATE_USER', 'UserManagement', username, { deactivatedBy: caller.username }, caller.username, caller.role); } catch(e) {}
  return success({ username, active: false }, 'User "' + username + '" deactivated.');
}

function reactivateUserAPI(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username = (payload.username || '').trim().toLowerCase();
  if (!username) return error('Username is required');

  const users = _getUsers();
  const user  = users[username];
  if (!user)       return error('User not found: ' + username);
  if (user.active) return error('User is already active.');

  users[username].active    = true;
  users[username].updatedAt = new Date().toISOString();
  _saveUsers(users);

  try { logActivity('REACTIVATE_USER', 'UserManagement', username, { reactivatedBy: caller.username }, caller.username, caller.role); } catch(e) {}
  return success({ username, active: true }, 'User "' + username + '" reactivated.');
}

function adminResetPassword(payload, caller) {
  const guard = requireAdmin(caller);
  if (guard) return guard;

  const username        = (payload.username        || '').trim().toLowerCase();
  const newPassword     = (payload.newPassword     || '');
  const confirmPassword = (payload.confirmPassword || '');

  if (!username)                       return error('Username is required');
  if (!newPassword || newPassword.length < 6) return error('New password must be at least 6 characters');
  if (newPassword !== confirmPassword) return error('Passwords do not match');

  const users = _getUsers();
  if (!users[username]) return error('User not found: ' + username);

  users[username].passwordHash = sha256(newPassword);
  users[username].updatedAt    = new Date().toISOString();
  _saveUsers(users);
  _killUserSessions(username);

  try { logActivity('ADMIN_RESET_PASSWORD', 'UserManagement', username, { resetBy: caller.username }, caller.username, caller.role); } catch(e) {}
  return success({ username }, 'Password reset for "' + username + '". All sessions terminated.');
}

// ── SETUP UTILITIES ───────────────────────────────────────────
function setupUsers(payload, caller) {
  const incoming = payload.users || [];
  if (!Array.isArray(incoming) || incoming.length === 0) return error('No users provided in payload.users array');

  const users  = _getUsers();
  const report = { added: [], skipped: [], updated: [], errors: [] };

  incoming.forEach(entry => {
    const username = (entry.username || '').trim().toLowerCase();
    if (!username) { report.errors.push('Blank username skipped'); return; }
    if (!VALID_ROLES.includes(entry.role)) { report.errors.push(username + ': invalid role'); return; }

    if (users[username] && !entry.forceUpdate) { report.skipped.push(username); return; }

    if (users[username] && entry.forceUpdate) {
      users[username].role        = entry.role;
      users[username].displayName = entry.displayName || users[username].displayName;
      users[username].updatedAt   = new Date().toISOString();
      if (entry.password && entry.password.length >= 6) users[username].passwordHash = sha256(entry.password);
      report.updated.push(username);
      return;
    }

    if (!entry.password || entry.password.length < 6) { report.errors.push(username + ': password must be 6+ chars'); return; }

    users[username] = {
      username, displayName: entry.displayName || username, role: entry.role,
      passwordHash: sha256(entry.password), active: true,
      createdAt: new Date().toISOString(), createdBy: caller ? caller.username : 'seed', updatedAt: null
    };
    report.added.push(username);
  });

  _saveUsers(users);
  try { logActivity('SETUP_USERS', 'Auth', 'batch', report, caller ? caller.username : 'script', caller ? caller.role : 'System'); } catch(e) {}
  return success(report, 'Done — Added: ' + report.added.length + ', Updated: ' + report.updated.length + ', Skipped: ' + report.skipped.length + ', Errors: ' + report.errors.length);
}

function setupUsers_seed() {
  const result = JSON.parse(setupUsers({
    users: [
      { username:'admin',    displayName:'Administrator',     role:'Admin',   password:'Admin@1234' },
      { username:'manager1', displayName:'Manager',           role:'Manager', password:'Manager@1234' },
      { username:'entryop1', displayName:'Entry Operator',    role:'EntryOP', password:'Entry@1234' },
      { username:'user1',    displayName:'General User',      role:'User',    password:'User@1234' },
      { username:'pc1',      displayName:'Process Coordinator',role:'PC',     password:'PC@1234' }
    ]
  }, null));
  Logger.log('setupUsers_seed result:\n' + JSON.stringify(result, null, 2));
}

function verifyUser_debug() {
  const username = 'admin';
  const password = 'Admin@1234';
  const users    = _getUsers();
  Logger.log('Total users: ' + Object.keys(users).length);
  const user = users[username];
  if (!user) { Logger.log('User NOT FOUND. Run setupUsers_seed() first.'); return; }
  const incomingHash = sha256(password);
  Logger.log('Stored hash : ' + user.passwordHash);
  Logger.log('Test hash   : ' + incomingHash);
  Logger.log(incomingHash === user.passwordHash ? '✅ MATCH — login will work.' : '❌ MISMATCH — run resetAllUsers_seed()');
}

function resetAllUsers_seed() {
  const result = JSON.parse(setupUsers({
    users: [
      { username:'admin',    displayName:'Administrator',      role:'Admin',   password:'Admin@1234',    forceUpdate:true },
      { username:'manager1', displayName:'Manager',            role:'Manager', password:'Manager@1234',  forceUpdate:true },
      { username:'entryop1', displayName:'Entry Operator',     role:'EntryOP', password:'Entry@1234',    forceUpdate:true },
      { username:'user1',    displayName:'General User',       role:'User',    password:'User@1234',     forceUpdate:true },
      { username:'pc1',      displayName:'Process Coordinator',role:'PC',      password:'PC@1234',       forceUpdate:true }
    ]
  }, null));
  Logger.log('resetAllUsers_seed:\n' + JSON.stringify(result, null, 2));
}

function listUsers_log()          { Logger.log(JSON.stringify(Object.values(_getUsers()).map(_safeUser), null, 2)); }
function clearAllSessions_admin() { PropertiesService.getScriptProperties().setProperty(PROP_SESSIONS_KEY, '{}'); Logger.log('All sessions cleared.'); }
function debugDispatchLogin() {
  try { Logger.log(dispatch('login', JSON.stringify({ username:'admin', password:'Admin@1234' }))); }
  catch(e) { Logger.log('Error: ' + e.message); }
}
