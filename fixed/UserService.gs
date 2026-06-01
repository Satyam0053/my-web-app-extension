// ============================================================
// UserService.gs — User Management
// CRUD operations for system users
// ============================================================

function getUsers(token) {
  try {
    const { valid, session, error } = requireReadAccess(token, 'UserManagement');
    if (!valid) return error;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('_Users');
    if (!sheet) return errorResponse('Users sheet not found', 'SYSTEM_ERROR');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return successResponse({ users: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const users = data
      .filter(function(r) { return r[0] !== ''; })
      .map(function(r) {
        return {
          UserID: r[0],
          FullName: r[1],
          Email: r[2],
          // Never send password hash to client
          Role: r[4],
          Status: r[5],
          LastLogin: r[6],
          FailedAttempts: r[7],
          CreatedBy: r[8],
          CreatedAt: r[9],
          UpdatedBy: r[10],
          UpdatedAt: r[11]
        };
      });

    return successResponse({ users: users });
  } catch (e) {
    return errorResponse('Failed to fetch users: ' + e.message, 'FETCH_ERROR');
  }
}

function createUser(token, userData) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'UserManagement');
    if (!valid) return error;

    // Validate required fields
    const required = ['FullName', 'Email', 'Password', 'Role'];
    const missing = validateRequired(required, userData);
    if (missing.length > 0) {
      return errorResponse('Missing required fields: ' + missing.join(', '), 'VALIDATION_ERROR');
    }

    // Validate email
    if (!isValidEmail(userData.Email)) {
      return errorResponse('Invalid email address', 'VALIDATION_ERROR');
    }

    // Validate role
    const validRoles = ['Super Admin', 'Admin', 'Purchase', 'Store', 'Production', 'Dispatch', 'Viewer'];
    if (validRoles.indexOf(userData.Role) === -1) {
      return errorResponse('Invalid role: ' + userData.Role, 'VALIDATION_ERROR');
    }

    // Only Super Admin can create Super Admin
    if (userData.Role === 'Super Admin' && session.role !== 'Super Admin') {
      return errorResponse('Only Super Admin can create Super Admin users', 'ACCESS_DENIED');
    }

    // Validate password strength
    if (!userData.Password || userData.Password.length < 8) {
      return errorResponse('Password must be at least 8 characters', 'WEAK_PASSWORD');
    }

    // Check duplicate email
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('_Users');
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
      if (emails.map(function(e) { return String(e).toLowerCase(); })
               .indexOf(userData.Email.toLowerCase()) !== -1) {
        return errorResponse('Email already exists: ' + userData.Email, 'DUPLICATE_EMAIL');
      }
    }

    const now = nowISO();
    const userID = generateID('USR', '_Users', 1);
    const passwordHash = hashPassword(userData.Password);

    const row = [
      userID,
      sanitize(userData.FullName),
      sanitize(userData.Email).toLowerCase(),
      passwordHash,
      userData.Role,
      userData.Status || 'Active',
      '',   // LastLogin
      0,    // FailedAttempts
      session.userID,
      now,
      '',   // UpdatedBy
      ''    // UpdatedAt
    ];

    sheet.appendRow(row);

    writeAuditLog({
      userID: session.userID,
      userName: session.userName,
      action: 'CREATE',
      module: 'UserManagement',
      recordID: userID,
      oldValues: '',
      newValues: JSON.stringify({ UserID: userID, Email: userData.Email, Role: userData.Role })
    });

    return successResponse({ UserID: userID }, 'User created successfully');
  } catch (e) {
    return errorResponse('Failed to create user: ' + e.message, 'CREATE_ERROR');
  }
}

function updateUser(token, userID, userData) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'UserManagement');
    if (!valid) return error;

    if (!userID) return errorResponse('UserID is required', 'VALIDATION_ERROR');

    // Cannot modify own role or status
    if (userID === session.userID && (userData.Role || userData.Status)) {
      return errorResponse('You cannot modify your own role or status', 'ACCESS_DENIED');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('_Users');
    const rowIndex = findRowByID('_Users', 1, userID);
    if (!rowIndex) return errorResponse('User not found: ' + userID, 'NOT_FOUND');

    const currentRow = sheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
    const oldValues = {
      FullName: currentRow[1], Email: currentRow[2], Role: currentRow[4], Status: currentRow[5]
    };

    // Build update
    const updates = {};
    if (userData.FullName) updates.FullName = sanitize(userData.FullName);
    if (userData.Role) {
      const validRoles = ['Super Admin', 'Admin', 'Purchase', 'Store', 'Production', 'Dispatch', 'Viewer'];
      if (validRoles.indexOf(userData.Role) === -1) {
        return errorResponse('Invalid role', 'VALIDATION_ERROR');
      }
      if (userData.Role === 'Super Admin' && session.role !== 'Super Admin') {
        return errorResponse('Only Super Admin can assign Super Admin role', 'ACCESS_DENIED');
      }
      updates.Role = userData.Role;
    }
    if (userData.Status) {
      const validStatuses = ['Active', 'Inactive', 'Locked'];
      if (validStatuses.indexOf(userData.Status) === -1) {
        return errorResponse('Invalid status', 'VALIDATION_ERROR');
      }
      updates.Status = userData.Status;
    }
    if (userData.NewPassword) {
      if (userData.NewPassword.length < 8) {
        return errorResponse('Password must be at least 8 characters', 'WEAK_PASSWORD');
      }
      updates.PasswordHash = hashPassword(userData.NewPassword);
    }

    // Apply updates to sheet
    const fieldMap = {
      'FullName': 2, 'PasswordHash': 4, 'Role': 5, 'Status': 6
    };
    // Column indices (1-based): UserID=1, FullName=2, Email=3, PasswordHash=4, Role=5, Status=6
    if (updates.FullName) sheet.getRange(rowIndex, 2).setValue(updates.FullName);
    if (updates.PasswordHash) sheet.getRange(rowIndex, 4).setValue(updates.PasswordHash);
    if (updates.Role) sheet.getRange(rowIndex, 5).setValue(updates.Role);
    if (updates.Status) {
      sheet.getRange(rowIndex, 6).setValue(updates.Status);
      // Reset failed attempts on unlock
      if (updates.Status === 'Active') {
        sheet.getRange(rowIndex, 8).setValue(0);
      }
    }
    sheet.getRange(rowIndex, 11).setValue(session.userID);
    sheet.getRange(rowIndex, 12).setValue(nowISO());

    writeAuditLog({
      userID: session.userID,
      userName: session.userName,
      action: 'UPDATE',
      module: 'UserManagement',
      recordID: userID,
      oldValues: oldValues,
      newValues: updates
    });

    return successResponse({ UserID: userID }, 'User updated successfully');
  } catch (e) {
    return errorResponse('Failed to update user: ' + e.message, 'UPDATE_ERROR');
  }
}

function resetUserPassword(token, userID, newPassword) {
  try {
    const { valid, session, error } = requireWriteAccess(token, 'UserManagement');
    if (!valid) return error;

    if (!userID || !newPassword) {
      return errorResponse('UserID and new password are required', 'VALIDATION_ERROR');
    }
    if (newPassword.length < 8) {
      return errorResponse('Password must be at least 8 characters', 'WEAK_PASSWORD');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('_Users');
    const rowIndex = findRowByID('_Users', 1, userID);
    if (!rowIndex) return errorResponse('User not found', 'NOT_FOUND');

    sheet.getRange(rowIndex, 4).setValue(hashPassword(newPassword));
    sheet.getRange(rowIndex, 6).setValue('Active'); // Unlock if locked
    sheet.getRange(rowIndex, 8).setValue(0); // Reset failed attempts
    sheet.getRange(rowIndex, 11).setValue(session.userID);
    sheet.getRange(rowIndex, 12).setValue(nowISO());

    writeAuditLog({
      userID: session.userID,
      userName: session.userName,
      action: 'UPDATE',
      module: 'UserManagement',
      recordID: userID,
      oldValues: '',
      newValues: JSON.stringify({ action: 'Password reset by admin' })
    });

    return successResponse({}, 'Password reset successfully');
  } catch (e) {
    return errorResponse('Failed to reset password: ' + e.message, 'RESET_ERROR');
  }
}

function getRoles() {
  return successResponse({
    roles: ['Super Admin', 'Admin', 'Purchase', 'Store', 'Production', 'Dispatch', 'Viewer']
  });
}