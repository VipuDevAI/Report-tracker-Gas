/************************************************
 MVM REPORT TRACKER - AUTHENTICATION & ACCESS
 File 3 of 7
 Google Email-Based Authentication
 
 ⚠️ IMPORTANT: OWNERSHIP vs APP ACCESS
 ────────────────────────────────────────────────
 SCRIPT OWNER (can edit code):
   → rishisans83@gmail.com ONLY
   
 APP ADMINS (can use admin features, NOT see code):
   → rishisans83@gmail.com
   → mvmseniors@gmail.com
   → Other school admins added below
   
 TEACHERS (filtered access via web app):
   → Added in Teachers sheet with their Google email
   
 🔒 DEPLOYMENT SETTINGS:
   1. Deploy → New deployment → Web app
   2. Execute as: "Me" (rishisans83@gmail.com)
   3. Who has access: "Anyone with Google account"
   4. Share ONLY the web app URL with school staff
   5. NEVER share script editor access
************************************************/

// APP ADMIN emails (can use admin features in the app)
// These users can manage students, teachers, exams, etc.
// They CANNOT see or edit the code (unless given editor access separately)
const ADMIN_EMAIL_LIST = [
  "rishisans83@gmail.com",    // Owner (can also edit code)
  "mvmseniors@gmail.com",     // Admin (app access only)
  "anithasivanesan4604@gmail.com"  // Admin (app access only)
];

/**
 * Get current logged in user's email
 * @returns {string} User email
 */
function getCurrentUser() {
  return Session.getActiveUser().getEmail();
}


/**
 * Get current user's effective user (for triggers)
 * @returns {string} Effective user email
 */
function getEffectiveUser() {
  return Session.getEffectiveUser().getEmail();
}


/**
 * Get the actual logged-in user email (checks both Session and stored)
 * @returns {string} User email
 */
function getActualUserEmail() {
  // Try session first
  const sessionEmail = getCurrentUser();
  if (sessionEmail) return sessionEmail;
  
  // Fall back to stored email
  return getLoggedInUser() || '';
}


/**
 * Check if user is registered (admin or teacher)
 * @returns {Object} { registered: boolean, role: string, message: string }
 */
function checkUserAccess() {
  const email = getCurrentUser();
  
  if (!email) {
    return { 
      registered: false, 
      role: null, 
      message: "Unable to identify user. Please ensure you are logged into Google." 
    };
  }
  
  // Check if admin
  if (ADMIN_EMAIL_LIST.includes(email)) {
    return { registered: true, role: "admin", message: "Welcome, Admin!" };
  }
  
  // Check if teacher exists in Teachers sheet
  const teacher = getTeacherByEmail(email);
  if (teacher) {
    return { registered: true, role: "teacher", message: `Welcome, ${teacher.name}!` };
  }
  
  // Not registered
  return { 
    registered: false, 
    role: null, 
    message: `Access Denied. Your email (${email}) is not registered in the system. Please contact the administrator.` 
  };
}


/**
 * Get current user's role
 * @returns {string} "admin", "teacher", or "unauthorized"
 */
function getCurrentUserRole() {
  const access = checkUserAccess();
  return access.registered ? access.role : "unauthorized";
}


/**
 * STEP 1 of login: verify email exists and check password setup status
 * @param {string} email
 * @returns {Object} { registered, role, requiresPasswordSetup, message }
 *   - requiresPasswordSetup=true → frontend should show "set new password" screen
 *   - requiresPasswordSetup=false → frontend should show password-entry screen
 */
function verifyUserEmail(email) {
  if (!email) {
    return { registered: false, role: null, requiresPasswordSetup: false, message: "Please enter an email address." };
  }
  email = email.trim().toLowerCase();
  
  // Determine role first
  let role = null;
  if (ADMIN_EMAIL_LIST.map(e => e.toLowerCase()).includes(email)) {
    role = "admin";
  } else {
    const teacher = getTeacherByEmail(email);
    if (teacher) role = "teacher";
  }
  if (!role) {
    return {
      registered: false, role: null, requiresPasswordSetup: false,
      message: `Email "${email}" is not registered. Please contact the administrator.`
    };
  }
  
  // Check Auth sheet for password presence
  const auth = _findAuthRow(email);
  const hasPassword = auth && auth.row[1]; // PasswordHash column
  const mustChange = auth && (auth.row[3] === true || String(auth.row[3]).toLowerCase() === "true");
  
  return {
    registered: true,
    role: role,
    requiresPasswordSetup: !hasPassword || mustChange,
    message: !hasPassword
      ? "First-time login: please set a password."
      : mustChange
        ? "Password reset required by admin. Please set a new password."
        : "Please enter your password."
  };
}


/**
 * STEP 2a: Set initial password (first-time user) or change password
 * @param {string} email
 * @param {string} newPassword - min 8 chars
 * @param {string} [oldPassword] - required if user already has a password (and not in mustChange state)
 * @returns {Object} { success, message, sessionToken?, role?, sessionExpiry? }
 */
function setUserPassword(email, newPassword, oldPassword) {
  if (!email || !newPassword) {
    return { success: false, message: "Email and new password are required." };
  }
  if (String(newPassword).length < 8) {
    return { success: false, message: "Password must be at least 8 characters long." };
  }
  email = email.trim().toLowerCase();
  
  // Confirm user is registered
  const adminEmails = ADMIN_EMAIL_LIST.map(e => e.toLowerCase());
  let role = null;
  if (adminEmails.includes(email)) role = "admin";
  else if (getTeacherByEmail(email)) role = "teacher";
  if (!role) {
    return { success: false, message: "Email not registered." };
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
    if (!sheet) return { success: false, message: "Auth sheet missing. Run Initialize App." };
    
    const auth = _findAuthRow(email);
    const hasPassword = auth && auth.row[1];
    const mustChange = auth && (auth.row[3] === true || String(auth.row[3]).toLowerCase() === "true");
    
    // If user already has password and is NOT in must-change state → require oldPassword
    if (hasPassword && !mustChange) {
      if (!oldPassword) {
        return { success: false, message: "Old password is required to change password." };
      }
      const oldHash = hashPassword(oldPassword, auth.row[2]);
      if (oldHash !== auth.row[1]) {
        try { writeAudit("PASSWORD_CHANGE_FAIL", "Auth", email, "PasswordHash", "", "", { reason: "wrong old password" }); } catch (e) {}
        return { success: false, message: "Old password is incorrect." };
      }
    }
    
    const salt = generateSalt();
    const hash = hashPassword(newPassword, salt);
    const sessionToken = Utilities.getUuid();
    const sessionDuration = parseInt(getSchoolSetting("SessionDurationHours") || "8", 10) || 8;
    const sessionExpiry = new Date(Date.now() + sessionDuration * 3600 * 1000);
    const now = new Date();
    
    const newRow = [email, hash, salt, false, sessionToken, sessionExpiry, 0, now, auth ? auth.row[8] : now];
    if (auth) {
      sheet.getRange(auth.rowNum, 1, 1, 9).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
    
    setLoggedInUser(email);
    writeAudit("PASSWORD_SET", "Auth", email, "PasswordHash", "", "(hashed)", { firstTime: !hasPassword });
    logAction("Password Set", `${email} set/changed password`);
    
    return {
      success: true,
      message: "Password set successfully. You are now logged in.",
      sessionToken: sessionToken,
      sessionExpiry: sessionExpiry.toISOString(),
      role: role
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * STEP 2b: Login with email + password
 * @param {string} email
 * @param {string} password
 * @returns {Object} { success, sessionToken, sessionExpiry, role, name, message }
 */
function loginWithPassword(email, password) {
  if (!email || !password) {
    return { success: false, message: "Email and password are required." };
  }
  email = email.trim().toLowerCase();
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    const auth = _findAuthRow(email);
    if (!auth || !auth.row[1]) {
      logAction("Login Fail", `No password set for ${email}`);
      writeAudit("LOGIN_FAIL", "Auth", email, "", "", "", { reason: "no password set" });
      return { success: false, message: "No password set. Please use 'First-time login' to set one.", requiresPasswordSetup: true };
    }
    
    const failedAttempts = parseInt(auth.row[6] || 0, 10);
    if (failedAttempts >= 5) {
      // Soft lockout
      const lastLogin = auth.row[7] ? new Date(auth.row[7]) : new Date(0);
      const lockoutMs = 15 * 60 * 1000;
      if (Date.now() - lastLogin.getTime() < lockoutMs) {
        writeAudit("LOGIN_LOCKED", "Auth", email, "", "", "", { failedAttempts });
        return { success: false, message: "Account temporarily locked due to too many failed attempts. Try again in 15 minutes." };
      }
    }
    
    const hash = hashPassword(password, auth.row[2]);
    if (hash !== auth.row[1]) {
      const newFailedCount = failedAttempts + 1;
      const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
      sheet.getRange(auth.rowNum, 7).setValue(newFailedCount);
      sheet.getRange(auth.rowNum, 8).setValue(new Date());
      logAction("Login Fail", `Wrong password for ${email} (attempt ${newFailedCount})`);
      writeAudit("LOGIN_FAIL", "Auth", email, "", "", "", { reason: "wrong password", failedAttempts: newFailedCount });
      return { success: false, message: `Invalid password. ${5 - newFailedCount} attempts remaining.` };
    }
    
    // Determine role (must still be valid — admin list or active teacher)
    let role = null;
    if (ADMIN_EMAIL_LIST.map(e => e.toLowerCase()).includes(email)) role = "admin";
    else {
      const teacher = getTeacherByEmail(email);
      if (teacher) role = "teacher";
    }
    if (!role) {
      writeAudit("LOGIN_FAIL", "Auth", email, "", "", "", { reason: "user no longer registered" });
      return { success: false, message: "Your account is no longer active. Contact admin." };
    }
    
    const sessionToken = Utilities.getUuid();
    const sessionDuration = parseInt(getSchoolSetting("SessionDurationHours") || "8", 10) || 8;
    const sessionExpiry = new Date(Date.now() + sessionDuration * 3600 * 1000);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
    sheet.getRange(auth.rowNum, 5).setValue(sessionToken);
    sheet.getRange(auth.rowNum, 6).setValue(sessionExpiry);
    sheet.getRange(auth.rowNum, 7).setValue(0); // reset failed attempts
    sheet.getRange(auth.rowNum, 8).setValue(new Date());
    
    setLoggedInUser(email);
    PropertiesService.getUserProperties().setProperty('sessionToken', sessionToken);
    
    const teacher = role === "teacher" ? getTeacherByEmail(email) : null;
    const name = teacher ? teacher.name : email.split("@")[0];
    
    logAction("Login", `${role} login: ${email}`);
    writeAudit("LOGIN_SUCCESS", "Auth", email, "", "", "", { role });
    
    return {
      success: true,
      sessionToken: sessionToken,
      sessionExpiry: sessionExpiry.toISOString(),
      role: role,
      name: name,
      email: email,
      message: `Welcome, ${name}!`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Validate a session token; returns user info or null
 * @param {string} sessionToken
 * @returns {Object|null}
 */
function validateSession(sessionToken) {
  if (!sessionToken) return null;
  const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][4] === sessionToken) {
      const expiry = data[i][5] ? new Date(data[i][5]) : null;
      if (!expiry || expiry.getTime() < Date.now()) return null;
      const email = String(data[i][0]).toLowerCase();
      const role = ADMIN_EMAIL_LIST.map(e => e.toLowerCase()).includes(email) ? "admin"
                 : getTeacherByEmail(email) ? "teacher" : null;
      if (!role) return null;
      setLoggedInUser(email);
      return { email, role, sessionExpiry: expiry.toISOString() };
    }
  }
  return null;
}


/**
 * Logout: invalidate session
 * @param {string} sessionToken - optional; if omitted, uses stored email's session
 * @returns {Object}
 */
function logoutUser(sessionToken) {
  try {
    const email = getLoggedInUser();
    if (email) {
      const auth = _findAuthRow(email);
      if (auth) {
        const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
        sheet.getRange(auth.rowNum, 5).setValue("");
        sheet.getRange(auth.rowNum, 6).setValue("");
      }
      writeAudit("LOGOUT", "Auth", email, "", "", "", {});
      logAction("Logout", email);
    }
  } catch (e) {}
  PropertiesService.getUserProperties().deleteProperty('loggedInEmail');
  PropertiesService.getUserProperties().deleteProperty('sessionToken');
  return { success: true, message: "Logged out." };
}


/**
 * Admin: reset another user's password (forces them to set new on next login)
 * @param {string} targetEmail
 * @returns {Object}
 */
function adminResetUserPassword(targetEmail) {
  if (!isAdmin()) return { success: false, message: "Admin access required." };
  if (!targetEmail) return { success: false, message: "Target email required." };
  targetEmail = targetEmail.trim().toLowerCase();
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
    if (!sheet) return { success: false, message: "Auth sheet missing." };
    const auth = _findAuthRow(targetEmail);
    if (!auth) {
      sheet.appendRow([targetEmail, "", "", true, "", "", 0, "", new Date()]);
    } else {
      sheet.getRange(auth.rowNum, 2).setValue(""); // clear hash
      sheet.getRange(auth.rowNum, 3).setValue(""); // clear salt
      sheet.getRange(auth.rowNum, 4).setValue(true); // mustChange
      sheet.getRange(auth.rowNum, 5).setValue(""); // invalidate session
      sheet.getRange(auth.rowNum, 6).setValue("");
      sheet.getRange(auth.rowNum, 7).setValue(0);
    }
    writeAudit("PASSWORD_RESET_BY_ADMIN", "Auth", targetEmail, "PasswordHash", "(was set)", "", { resetBy: getActualUserEmail() });
    logAction("Admin Reset Password", `${getActualUserEmail()} reset password for ${targetEmail}`);
    return { success: true, message: `Password for ${targetEmail} cleared. They must set a new password on next login.` };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * List Auth records (admin only) — for password management UI
 */
function listAuthUsers() {
  if (!isAdmin()) return [];
  const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  return data.map(r => ({
    email: r[0],
    hasPassword: !!r[1],
    mustChangePassword: r[3] === true || String(r[3]).toLowerCase() === "true",
    failedAttempts: r[6] || 0,
    lastLogin: r[7],
    createdAt: r[8]
  })).filter(u => u.email);
}


// Legacy compat — if old code calls this, treat as a simple email-only "session"
function verifyUserEmailLegacy(email) {
  return verifyUserEmail(email);
}


/**
 * Check if current user is an admin
 * Checks both Session email and stored login email
 * @param {string} email - Optional email to check
 * @returns {boolean} True if admin
 */
function isAdmin(email) {
  // If email provided, check it directly
  if (email) {
    return ADMIN_EMAIL_LIST.map(e => e.toLowerCase()).includes(email.toLowerCase());
  }
  
  // Try session first
  const sessionEmail = getCurrentUser();
  if (sessionEmail && ADMIN_EMAIL_LIST.map(e => e.toLowerCase()).includes(sessionEmail.toLowerCase())) {
    return true;
  }
  
  // Check stored login email from PropertiesService
  const storedEmail = PropertiesService.getUserProperties().getProperty('loggedInEmail');
  if (storedEmail && ADMIN_EMAIL_LIST.map(e => e.toLowerCase()).includes(storedEmail.toLowerCase())) {
    return true;
  }
  
  return false;
}


/**
 * Check if email is admin (direct check)
 * @param {string} email - Email to check
 * @returns {boolean} True if admin
 */
function isAdminEmail(email) {
  if (!email) return false;
  return ADMIN_EMAIL_LIST.map(e => e.toLowerCase()).includes(email.toLowerCase());
}


/**
 * Store logged in user email (called after successful login)
 * @param {string} email - User email
 */
function setLoggedInUser(email) {
  PropertiesService.getUserProperties().setProperty('loggedInEmail', email || '');
}


/**
 * Get logged in user email
 * @returns {string} Stored email or empty string
 */
function getLoggedInUser() {
  return PropertiesService.getUserProperties().getProperty('loggedInEmail') || '';
}


/**
 * Check if current user is a teacher
 * @param {string} email - Optional email to check
 * @returns {boolean} True if teacher
 */
function isTeacher(email) {
  // If email provided, check it directly
  if (email) {
    const teacher = getTeacherByEmail(email);
    return teacher !== null;
  }
  
  // Try session first
  const sessionEmail = getCurrentUser();
  if (sessionEmail) {
    const teacher = getTeacherByEmail(sessionEmail);
    if (teacher) return true;
  }
  
  // Check stored login email
  const storedEmail = getLoggedInUser();
  if (storedEmail) {
    const teacher = getTeacherByEmail(storedEmail);
    if (teacher) return true;
  }
  
  return false;
}


/**
 * Check if user has any valid access
 * @param {string} email - Optional email to check
 * @returns {boolean} True if admin or teacher
 */
function hasAccess(email) {
  return isAdmin(email) || isTeacher(email);
}


/**
 * Get teacher record by email
 * @param {string} email - Teacher email
 * @returns {Object|null} Teacher object or null
 */
function getTeacherByEmail(email) {
  if (!email) return null;
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return null;
  
  // Case-insensitive email comparison
  const emailLower = email.trim().toLowerCase();
  const row = data.find(r => {
    const teacherEmail = (r[5] || '').toString().trim().toLowerCase();
    const status = (r[8] || '').toString().trim();
    return teacherEmail === emailLower && status === "Active";
  });
  
  if (!row) return null;
  
  return {
    teacherId: row[0],
    name: row[1],
    subject: row[2],
    classes: row[3],
    sections: row[4],
    email: row[5],
    phone: row[6],
    joinDate: row[7],
    status: row[8]
  };
}


/**
 * Get teacher assignment details for filtering
 * @param {string} email - Teacher email (optional, uses current user if not provided)
 * @returns {Object|null} Teacher assignment object with parsed arrays
 */
function getTeacherAssignment(email) {
  // Get email from parameter, session, or stored login
  const teacherEmail = email || getActualUserEmail();
  const teacher = getTeacherByEmail(teacherEmail);
  
  if (!teacher) return null;
  
  // Parse comma/semicolon-separated values into arrays
  const classesRaw = String(teacher.classes || "");
  const sectionsRaw = String(teacher.sections || "");
  
  // Handle both comma and semicolon separators
  const classesSplit = classesRaw.includes(";") ? classesRaw.split(";") : classesRaw.split(",");
  const sectionsSplit = sectionsRaw.includes(";") ? sectionsRaw.split(";") : sectionsRaw.split(",");
  
  return {
    teacherId: teacher.teacherId,
    name: teacher.name,
    email: teacher.email,
    subject: teacher.subject,
    classes: classesSplit.map(c => c.trim()).filter(c => c),
    sections: sectionsSplit.map(s => s.trim()).filter(s => s),
    hasAllClasses: classesRaw.toLowerCase().includes("all"),
    hasAllSections: sectionsRaw.toLowerCase().includes("all")
  };
}


/**
 * Apply teacher filter to data array
 * @param {Array} data - Array of objects to filter
 * @param {Object} options - Filter options { filterBySubject: boolean, subjectField: string }
 * @returns {Array} Filtered data
 */
function applyTeacherFilter(data, options) {
  // Admin sees everything
  if (isAdmin()) {
    return data;
  }
  
  const assignment = getTeacherAssignment();
  if (!assignment) {
    return []; // No assignment = no access
  }
  
  const opts = options || {};
  const filterBySubject = opts.filterBySubject || false;
  const subjectField = opts.subjectField || "subject";
  
  return data.filter(item => {
    // Check class assignment
    const itemClass = String(item.class || item.Class || "");
    const classMatch = assignment.hasAllClasses || 
                       assignment.classes.includes(itemClass) ||
                       assignment.classes.some(c => itemClass.includes(c));
    
    if (!classMatch) return false;
    
    // Check section assignment
    const itemSection = String(item.section || item.Section || "");
    const sectionMatch = assignment.hasAllSections || 
                         assignment.sections.includes(itemSection) ||
                         itemSection === "";
    
    if (!sectionMatch) return false;
    
    // Check subject if required
    if (filterBySubject) {
      const itemSubject = String(item[subjectField] || "");
      const subjectMatch = assignment.subject === "All" || 
                           assignment.subject === itemSubject ||
                           itemSubject === "";
      if (!subjectMatch) return false;
    }
    
    return true;
  });
}


/**
 * Get current user info (admin/teacher details)
 * @returns {Object} User info object
 */
function getCurrentUserInfo() {
  const email = getCurrentUser();
  const access = checkUserAccess();
  
  if (!access.registered) {
    return {
      type: "unauthorized",
      role: "unauthorized",
      email: email,
      name: email ? email.split("@")[0] : "Unknown",
      message: access.message,
      hasAccess: false
    };
  }
  
  if (access.role === "admin") {
    return {
      type: "admin",
      role: "admin",
      email: email,
      name: email.split("@")[0],
      permissions: ["all"],
      canManageMasterData: true,
      canLockExams: true,
      canViewAllData: true,
      hasAccess: true
    };
  }
  
  const assignment = getTeacherAssignment(email);
  if (assignment) {
    return {
      type: "teacher",
      role: "teacher",
      email: email,
      name: assignment.name,
      teacherId: assignment.teacherId,
      subject: assignment.subject,
      classes: assignment.classes,
      sections: assignment.sections,
      permissions: ["view_own_data", "enter_marks", "view_students"],
      canManageMasterData: false,
      canLockExams: false,
      canViewAllData: false,
      hasAccess: true
    };
  }
  
  return {
    type: "unauthorized",
    role: "unauthorized",
    email: email,
    name: email ? email.split("@")[0] : "Unknown",
    message: "Access Denied",
    hasAccess: false
  };
}


/**
 * Get user info by email (for login-based auth)
 * @param {string} email - User email
 * @returns {Object} User info object
 */
function getUserInfoByEmail(email) {
  if (!email) {
    return {
      type: "unauthorized",
      role: "unauthorized",
      email: "",
      name: "Unknown",
      message: "No email provided",
      hasAccess: false
    };
  }
  
  email = email.trim().toLowerCase();
  
  // Check if admin
  const adminEmails = ADMIN_EMAIL_LIST.map(e => e.toLowerCase());
  if (adminEmails.includes(email)) {
    return {
      type: "admin",
      role: "admin",
      email: email,
      name: email.split("@")[0],
      permissions: ["all"],
      canManageMasterData: true,
      canLockExams: true,
      canViewAllData: true,
      hasAccess: true
    };
  }
  
  // Check if teacher
  const teacher = getTeacherByEmail(email);
  if (teacher) {
    const assignment = getTeacherAssignment(email);
    return {
      type: "teacher",
      role: "teacher",
      email: email,
      name: teacher.name || email.split("@")[0],
      teacherId: assignment ? assignment.teacherId : null,
      subject: assignment ? assignment.subject : null,
      classes: assignment ? assignment.classes : [],
      sections: assignment ? assignment.sections : [],
      permissions: ["view_own_data", "enter_marks", "view_students"],
      canManageMasterData: false,
      canLockExams: false,
      canViewAllData: false,
      hasAccess: true
    };
  }
  
  return {
    type: "unauthorized",
    role: "unauthorized",
    email: email,
    name: email.split("@")[0],
    message: "Access Denied - Email not registered",
    hasAccess: false
  };
}


/**
 * Check if user has specific permission
 * @param {string} permission - Permission to check
 * @returns {boolean} True if has permission
 */
function hasPermission(permission) {
  const userInfo = getCurrentUserInfo();
  
  if (!userInfo.hasAccess) return false;
  if (userInfo.role === "admin") return true;
  
  return userInfo.permissions && userInfo.permissions.includes(permission);
}


/**
 * Validate teacher can access specific class/section
 * @param {string} classNum - Class number
 * @param {string} section - Section
 * @returns {boolean} True if can access
 */
function canAccessClass(classNum, section) {
  if (isAdmin()) return true;
  
  const assignment = getTeacherAssignment();
  if (!assignment) return false;
  
  const classMatch = assignment.hasAllClasses || assignment.classes.includes(String(classNum));
  const sectionMatch = assignment.hasAllSections || assignment.sections.includes(section);
  
  return classMatch && sectionMatch;
}


/**
 * Validate teacher can edit marks for specific subject
 * @param {string} subject - Subject name
 * @returns {boolean} True if can edit
 */
function canEditSubject(subject) {
  if (isAdmin()) return true;
  
  const assignment = getTeacherAssignment();
  if (!assignment) return false;
  
  return assignment.subject === "All" || assignment.subject === subject;
}


/**
 * Validate teacher can access specific student
 * @param {Object} student - Student object with class and section
 * @returns {boolean} True if can access
 */
function canAccessStudent(student) {
  if (isAdmin()) return true;
  
  return canAccessClass(student.class, student.section);
}


/**
 * Get access control list for current user
 * @returns {Object} Access control object
 */
function getAccessControl() {
  const userInfo = getCurrentUserInfo();
  
  if (!userInfo.hasAccess) {
    return {
      user: userInfo,
      role: "unauthorized",
      hasAccess: false,
      message: userInfo.message
    };
  }
  
  const isUserAdmin = userInfo.role === "admin";
  
  return {
    user: userInfo,
    role: userInfo.role,
    hasAccess: true,
    // Master Data
    canUploadStudents: isUserAdmin,
    canUploadTeachers: isUserAdmin,
    canManageSubjects: isUserAdmin,
    canManageClasses: isUserAdmin,
    // Exams
    canCreateExam: isUserAdmin,
    canLockExam: isUserAdmin,
    canDeleteExam: isUserAdmin,
    // Marks
    canEnterMarks: true, // Both admin and teachers
    canViewMarks: true,
    canDeleteMarks: isUserAdmin,
    // Analytics & Reports
    canViewAnalytics: true,
    canViewReports: true,
    canExportData: true,
    canGenerateReportCards: isUserAdmin,
    // Settings
    canModifyGradeRanges: isUserAdmin,
    canModifySchoolSettings: isUserAdmin,
    canResetData: isUserAdmin,
    // Data scope
    viewAllData: isUserAdmin,
    filteredByAssignment: !isUserAdmin
  };
}


/**
 * Require valid access - throws error if not registered
 * @param {string} action - Action being attempted
 */
function requireAccess(action) {
  const access = checkUserAccess();
  if (!access.registered) {
    logAction("Access Denied", `${getCurrentUser()} attempted: ${action}`);
    throw new Error(access.message);
  }
}


/**
 * Require admin access - throws error if not admin
 * @param {string} action - Action being attempted
 */
function requireAdmin(action) {
  if (!isAdmin()) {
    logAction("Access Denied", `${getCurrentUser()} attempted: ${action}`);
    throw new Error(`Access denied. Admin privileges required for: ${action}`);
  }
}


/**
 * Require teacher or admin access
 * @param {string} action - Action being attempted
 */
function requireTeacherOrAdmin(action) {
  if (!isAdmin() && !isTeacher()) {
    logAction("Access Denied", `${getCurrentUser()} attempted: ${action}`);
    throw new Error(`Access denied. Teacher or Admin privileges required for: ${action}`);
  }
}


/**
 * Get current academic year from settings
 * @returns {string} Academic year (e.g., "2024-2025")
 */
function getCurrentAcademicYear() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
  const data = sheet.getDataRange().getValues();
  
  const yearRow = data.find(r => r[0] === "AcademicYear");
  return yearRow ? yearRow[1] : "2024-2025";
}


/**
 * Log unauthorized access attempt
 * @param {string} action - Attempted action
 */
function logUnauthorizedAccess(action) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Logs");
  
  sheet.appendRow([
    `LOG${Date.now()}`,
    "UNAUTHORIZED_ACCESS",
    getCurrentUser(),
    action,
    new Date()
  ]);
}
