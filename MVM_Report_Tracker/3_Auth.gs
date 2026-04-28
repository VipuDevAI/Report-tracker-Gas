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
  
  const role = getRole(email);
  if (role) {
    let displayName = email.split("@")[0];
    if (role !== "admin") {
      const t = getTeacherByEmail(email);
      if (t && t.name) displayName = t.name;
    }
    const roleLabel = role === "admin"      ? "Admin"
                    : role === "principal"  ? "Principal"
                    : role === "wing_admin" ? "Wing Admin"
                    : "Teacher";
    return { registered: true, role: role, message: `Welcome, ${displayName}! (${roleLabel})` };
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
  
  // Determine role first (supports all roles via Teachers sheet)
  const role = getRole(email);
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
  
  // Confirm user is registered (any role)
  const role = getRole(email);
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
    
    // Determine role (must still be valid — admin list or active teacher with a role)
    const role = getRole(email);
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
    
    const teacher = (role !== "admin") ? getTeacherByEmail(email) : null;
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
      const role = getRole(email);
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
  // If email provided, check directly via getRole (supports both ADMIN_EMAIL_LIST and Role column)
  if (email) {
    return getRole(email) === "admin";
  }
  
  // Try session first
  const sessionEmail = getCurrentUser();
  if (sessionEmail && getRole(sessionEmail) === "admin") {
    return true;
  }
  
  // Check stored login email from PropertiesService
  const storedEmail = PropertiesService.getUserProperties().getProperty('loggedInEmail');
  if (storedEmail && getRole(storedEmail) === "admin") {
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
    const isDeleted = (r[11] || '').toString().toLowerCase() === 'true';
    return teacherEmail === emailLower && status === "Active" && !isDeleted;
  });
  
  if (!row) return null;
  
  // Role column (index 12) — default to "TEACHER" if missing/empty (backward compat)
  const rawRole = (row[12] || '').toString().trim().toUpperCase();
  const role = rawRole || "TEACHER";
  
  return {
    teacherId: row[0],
    name: row[1],
    subject: row[2],
    classes: row[3],
    sections: row[4],
    email: row[5],
    phone: row[6],
    joinDate: row[7],
    status: row[8],
    isClassTeacher: row[9],
    classTeacherOf: row[10],
    role: role
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
  
  // Parse comma/semicolon-separated values into arrays (also handle spaces and pipes)
  const classesRaw = String(teacher.classes || "");
  const sectionsRaw = String(teacher.sections || "");
  
  // Standardize separators — accept comma, semicolon, pipe; trim each token
  const splitNorm = (s) => s.split(/[,;|]/).map(x => x.trim()).filter(x => x);
  const classesSplit = splitNorm(classesRaw);
  const sectionsSplit = splitNorm(sectionsRaw);
  
  return {
    teacherId: teacher.teacherId,
    name: teacher.name,
    email: teacher.email,
    subject: teacher.subject,
    role: teacher.role,
    classes: classesSplit,
    sections: sectionsSplit,
    hasAllClasses: classesRaw.toLowerCase().includes("all"),
    hasAllSections: sectionsRaw.toLowerCase().includes("all")
  };
}


/**
 * Get the role of a user (admin/principal/wing_admin/teacher/null)
 * Single source of truth for role determination.
 *
 * Resolution order:
 *   1. Hard-coded ADMIN_EMAIL_LIST → 'admin' (super-admin / script owner safety)
 *   2. Teachers sheet → uses Role column (col 13). Empty = 'teacher' (backward compat)
 *
 * @param {string} email - Optional email; falls back to current/stored user
 * @returns {string|null} 'admin' | 'principal' | 'wing_admin' | 'teacher' | null
 */
function getRole(email) {
  const e = (email || getActualUserEmail() || "").toString().trim().toLowerCase();
  if (!e) return null;
  
  // Super-admin via constant (cannot be locked out)
  if (ADMIN_EMAIL_LIST.map(x => x.toLowerCase()).includes(e)) return "admin";
  
  // Lookup in Teachers sheet (single source of truth for non-super-admin roles)
  const teacher = getTeacherByEmail(e);
  if (!teacher) return null;
  
  switch ((teacher.role || "TEACHER").toUpperCase()) {
    case "ADMIN":      return "admin";
    case "PRINCIPAL":  return "principal";
    case "WING_ADMIN":
    case "WINGADMIN":  return "wing_admin";
    case "TEACHER":
    default:           return "teacher";
  }
}


/**
 * Check if current/given user is a Principal (full read, no write).
 * @param {string} email - Optional email
 * @returns {boolean}
 */
function isPrincipal(email) {
  return getRole(email) === "principal";
}


/**
 * Check if current/given user is a Wing Admin (admin powers within wing scope).
 * @param {string} email - Optional email
 * @returns {boolean}
 */
function isWingAdmin(email) {
  return getRole(email) === "wing_admin";
}


/**
 * Check whether the user can perform a write action.
 * Roles:
 *   - admin       → all writes
 *   - wing_admin  → all writes within their wing (caller must additionally validate scope)
 *   - teacher     → only marks entry within assignment (caller validates scope)
 *   - principal   → no writes (read-only)
 *
 * @param {string} action - Optional action key: 'students' | 'marks' | 'exams' | 'reports' | 'settings'
 * @param {string} email - Optional email
 * @returns {boolean}
 */
function canWrite(action, email) {
  const role = getRole(email);
  if (!role) return false;
  if (role === "admin") return true;
  if (role === "principal") return false;
  
  if (role === "wing_admin") {
    // Wing admins can do all standard writes (admin's responsibilities) within scope.
    // System-wide settings (school config, year freeze, archive, grade ranges) remain
    // admin-only — those endpoints continue using isAdmin() directly.
    return action === "students" || action === "marks" ||
           action === "exams" || action === "reports";
  }
  
  if (role === "teacher") {
    // Teachers can only enter marks (existing behavior).
    return action === "marks";
  }
  
  return false;
}


/**
 * Check whether the user can read data. Anyone with a valid role can read;
 * scope is enforced separately by applyScopeFilter().
 * @param {string} email - Optional email
 * @returns {boolean}
 */
function canRead(email) {
  return getRole(email) !== null;
}


/**
 * Wing configuration helpers
 * Wings are configured in Settings_School: Wing_Primary, Wing_Secondary, Wing_Senior.
 * Defaults: Primary=6,7,8 | Secondary=9,10 | Senior=11,12.
 */
function _readWingConfig() {
  const cache = _readWingConfig._cache;
  if (cache && (Date.now() - cache.t) < 60000) return cache.v; // 60s cache
  
  const defaults = {
    PRIMARY: ["6", "7", "8"],
    SECONDARY: ["9", "10"],
    SENIOR: ["11", "12"]
  };
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
    if (!sheet) return defaults;
    const data = sheet.getDataRange().getValues();
    const map = {};
    data.forEach(r => { if (r[0]) map[String(r[0])] = String(r[1] || ""); });
    
    const parse = (s, fallback) => {
      const arr = String(s || "").split(/[,;|]/).map(x => x.trim()).filter(x => x);
      return arr.length ? arr : fallback;
    };
    
    const cfg = {
      PRIMARY:   parse(map["Wing_Primary"],   defaults.PRIMARY),
      SECONDARY: parse(map["Wing_Secondary"], defaults.SECONDARY),
      SENIOR:    parse(map["Wing_Senior"],    defaults.SENIOR)
    };
    _readWingConfig._cache = { t: Date.now(), v: cfg };
    return cfg;
  } catch (e) {
    return defaults;
  }
}


/**
 * Get wing name (PRIMARY/SECONDARY/SENIOR) for a given class.
 * @param {string|number} cls
 * @returns {string|null}
 */
function getWingForClass(cls) {
  const c = String(cls || "").trim();
  if (!c) return null;
  const cfg = _readWingConfig();
  if (cfg.PRIMARY.includes(c))   return "PRIMARY";
  if (cfg.SECONDARY.includes(c)) return "SECONDARY";
  if (cfg.SENIOR.includes(c))    return "SENIOR";
  return null;
}


/**
 * Get all class numbers belonging to a given wing.
 * @param {string} wingName - PRIMARY | SECONDARY | SENIOR
 * @returns {Array<string>}
 */
function getClassesForWing(wingName) {
  const cfg = _readWingConfig();
  return cfg[String(wingName || "").toUpperCase()] || [];
}


/**
 * Resolve the class scope for a Wing Admin user.
 *
 * Two configuration modes (auto-detected, in order):
 *   1. Their `Classes` column in Teachers sheet contains explicit class numbers
 *      (e.g. "9,10") — used as-is.
 *   2. Their `Classes` column contains a wing name (e.g. "PRIMARY", "SENIOR",
 *      "Wing_Senior") — expanded via getClassesForWing().
 *   3. Empty → empty scope (no access).
 *
 * @param {string} email
 * @returns {Object|null} { classes:[], wing:string|null }
 */
function getWingAdminAssignment(email) {
  const t = getTeacherByEmail(email || getActualUserEmail());
  if (!t) return null;
  
  const raw = String(t.classes || "").trim();
  if (!raw) return { classes: [], wing: null };
  
  const tokens = raw.split(/[,;|]/).map(x => x.trim()).filter(x => x);
  
  // If any token matches a wing name, expand it
  const wingNames = ["PRIMARY", "SECONDARY", "SENIOR",
                     "WING_PRIMARY", "WING_SECONDARY", "WING_SENIOR"];
  let classes = [];
  let firstWing = null;
  tokens.forEach(tok => {
    const upper = tok.toUpperCase();
    if (wingNames.includes(upper)) {
      const w = upper.replace("WING_", "");
      classes = classes.concat(getClassesForWing(w));
      if (!firstWing) firstWing = w;
    } else {
      classes.push(tok);
    }
  });
  
  // Dedupe
  classes = Array.from(new Set(classes));
  
  // If no explicit wing name was used but all classes happen to live in one wing, infer it
  if (!firstWing && classes.length) {
    const wings = Array.from(new Set(classes.map(c => getWingForClass(c)).filter(w => w)));
    if (wings.length === 1) firstWing = wings[0];
  }
  
  return { classes: classes, wing: firstWing };
}


/**
 * Apply role-based scope filter to data array.
 *
 * Rules:
 *   - admin       → returns full data
 *   - principal   → returns full data (read-only enforced by canWrite, not here)
 *   - wing_admin  → filters by wing class list
 *   - teacher     → existing assignment filter (class + section + optional subject)
 *   - unknown     → returns []
 *
 * @param {Array} data - Array of objects with .class/.section/.subject (or Class/Section)
 * @param {Object} options - { filterBySubject:boolean, subjectField:string }
 * @returns {Array}
 */
function applyScopeFilter(data, options) {
  const role = getRole();
  const opts = options || {};
  const filterBySubject = opts.filterBySubject || false;
  const subjectField = opts.subjectField || "subject";
  
  // Admin & Principal: full read access
  if (role === "admin" || role === "principal") {
    return data;
  }
  
  if (role === "wing_admin") {
    const wa = getWingAdminAssignment();
    if (!wa || !wa.classes.length) return [];
    const allowedClasses = wa.classes.map(String);
    return data.filter(item => {
      const itemClass = String(item.class || item.Class || "");
      return allowedClasses.includes(itemClass);
    });
  }
  
  if (role === "teacher") {
    const assignment = getTeacherAssignment();
    if (!assignment) return [];
    
    return data.filter(item => {
      const itemClass = String(item.class || item.Class || "");
      const classMatch = assignment.hasAllClasses ||
                         assignment.classes.includes(itemClass) ||
                         assignment.classes.some(c => itemClass.includes(c));
      if (!classMatch) return false;
      
      const itemSection = String(item.section || item.Section || "");
      const sectionMatch = assignment.hasAllSections ||
                           assignment.sections.includes(itemSection) ||
                           itemSection === "";
      if (!sectionMatch) return false;
      
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
  
  return [];
}


/**
 * Backward-compat alias. New code should use applyScopeFilter().
 * @deprecated Use applyScopeFilter
 */
function applyTeacherFilter(data, options) {
  return applyScopeFilter(data, options);
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
  
  return _buildUserInfo(email, access.role);
}


/**
 * Build a normalized user-info object for any role.
 * @private
 */
function _buildUserInfo(email, role) {
  email = (email || "").toLowerCase();
  const teacher = (role !== "admin") ? getTeacherByEmail(email) : null;
  const name = teacher ? teacher.name : (email ? email.split("@")[0] : "Unknown");
  
  if (role === "admin") {
    return {
      type: "admin",
      role: "admin",
      email: email,
      name: name,
      permissions: ["all"],
      canManageMasterData: true,
      canLockExams: true,
      canViewAllData: true,
      readOnly: false,
      hasAccess: true
    };
  }
  
  if (role === "principal") {
    return {
      type: "principal",
      role: "principal",
      email: email,
      name: name,
      teacherId: teacher ? teacher.teacherId : null,
      permissions: ["view_all", "view_analytics", "view_reports", "view_audit"],
      canManageMasterData: false,
      canLockExams: false,
      canViewAllData: true,
      readOnly: true,
      hasAccess: true
    };
  }
  
  if (role === "wing_admin") {
    const wa = getWingAdminAssignment(email);
    return {
      type: "wing_admin",
      role: "wing_admin",
      email: email,
      name: name,
      teacherId: teacher ? teacher.teacherId : null,
      wing: wa ? wa.wing : null,
      classes: wa ? wa.classes : [],
      permissions: ["manage_wing", "enter_marks", "view_students", "manage_exams", "generate_reports"],
      canManageMasterData: true, // within wing
      canLockExams: false,       // remains admin-only
      canViewAllData: false,
      readOnly: false,
      hasAccess: true
    };
  }
  
  // teacher (default)
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
      readOnly: false,
      hasAccess: true
    };
  }
  
  return {
    type: "unauthorized",
    role: "unauthorized",
    email: email,
    name: name,
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
  const role = getRole(email);
  if (!role) {
    return {
      type: "unauthorized",
      role: "unauthorized",
      email: email,
      name: email.split("@")[0],
      message: "Access Denied - Email not registered",
      hasAccess: false
    };
  }
  return _buildUserInfo(email, role);
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
  const role = getRole();
  if (role === "admin" || role === "principal") return true;
  
  if (role === "wing_admin") {
    const wa = getWingAdminAssignment();
    return wa && wa.classes.map(String).includes(String(classNum));
  }
  
  if (role === "teacher") {
    const assignment = getTeacherAssignment();
    if (!assignment) return false;
    const classMatch = assignment.hasAllClasses || assignment.classes.includes(String(classNum));
    const sectionMatch = assignment.hasAllSections || assignment.sections.includes(section);
    return classMatch && sectionMatch;
  }
  
  return false;
}


/**
 * Validate teacher can edit marks for specific subject
 * @param {string} subject - Subject name
 * @returns {boolean} True if can edit
 */
function canEditSubject(subject) {
  const role = getRole();
  if (role === "admin" || role === "wing_admin") return true;
  if (role === "principal") return false;
  
  const assignment = getTeacherAssignment();
  if (!assignment) return false;
  return assignment.subject === "All" || assignment.subject === subject;
}


/**
 * Validate user can access specific student (by class/section)
 * @param {Object} student - Student object with class and section
 * @returns {boolean}
 */
function canAccessStudent(student) {
  if (!student) return false;
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
  
  const role = userInfo.role;
  const isUserAdmin     = role === "admin";
  const isUserPrincipal = role === "principal";
  const isUserWing      = role === "wing_admin";
  const isUserTeacher   = role === "teacher";
  
  // Wing admins can do most admin work within their wing (server-side enforces scope)
  const canManage = isUserAdmin || isUserWing;
  
  return {
    user: userInfo,
    role: role,
    hasAccess: true,
    readOnly: isUserPrincipal,
    // Master Data
    canUploadStudents: canManage,
    canUploadTeachers: isUserAdmin,           // admin-only (teachers config is school-wide)
    canManageSubjects: isUserAdmin,
    canManageClasses:  isUserAdmin,
    // Exams
    canCreateExam: canManage,
    canLockExam:   isUserAdmin,                // remains admin-only
    canDeleteExam: canManage,
    // Marks
    canEnterMarks:  isUserAdmin || isUserWing || isUserTeacher,
    canViewMarks:   true,
    canDeleteMarks: canManage,
    // Analytics & Reports
    canViewAnalytics:       true,
    canViewReports:         true,
    canExportData:          true,
    canGenerateReportCards: canManage,
    canViewAuditTrail:      isUserAdmin || isUserPrincipal,
    // Settings (school-wide)
    canModifyGradeRanges:    isUserAdmin,
    canModifySchoolSettings: isUserAdmin,
    canResetData:            isUserAdmin,
    // Data scope
    viewAllData:           isUserAdmin || isUserPrincipal,
    filteredByAssignment:  isUserWing || isUserTeacher,
    wing:                  userInfo.wing || null
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
