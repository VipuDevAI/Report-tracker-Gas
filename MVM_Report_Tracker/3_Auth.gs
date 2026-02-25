/************************************************
 MVM REPORT TRACKER - AUTHENTICATION & ACCESS
 File 3 of 7
************************************************/

// Admin email whitelist (also defined in Init for reference)
const ADMIN_EMAIL_LIST = [
  "rishisans83@gmail.com",
  "mvmseniors26@gmail.com"
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
 * Check if current user is an admin
 * @returns {boolean} True if admin
 */
function isAdmin() {
  const email = getCurrentUser();
  return ADMIN_EMAIL_LIST.includes(email);
}


/**
 * Check if current user is a teacher
 * @returns {boolean} True if teacher
 */
function isTeacher() {
  const email = getCurrentUser();
  const teacher = getTeacherByEmail(email);
  return teacher !== null;
}


/**
 * Get teacher record by email
 * @param {string} email - Teacher email
 * @returns {Object|null} Teacher object or null
 */
function getTeacherByEmail(email) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return null;
  
  const row = data.find(r => r[5] === email && r[8] === "Active");
  
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
 * Get current user info (admin/teacher details)
 * @returns {Object} User info object
 */
function getCurrentUserInfo() {
  const email = getCurrentUser();
  
  if (ADMIN_EMAIL_LIST.includes(email)) {
    return {
      type: "admin",
      email: email,
      name: email.split("@")[0],
      permissions: ["all"]
    };
  }
  
  const teacher = getTeacherByEmail(email);
  if (teacher) {
    return {
      type: "teacher",
      email: email,
      name: teacher.name,
      teacherId: teacher.teacherId,
      subject: teacher.subject,
      classes: teacher.classes,
      sections: teacher.sections,
      permissions: ["view_own_marks", "enter_marks", "view_students"]
    };
  }
  
  return {
    type: "viewer",
    email: email,
    name: email.split("@")[0],
    permissions: ["view_only"]
  };
}


/**
 * Check if user has specific permission
 * @param {string} permission - Permission to check
 * @returns {boolean} True if has permission
 */
function hasPermission(permission) {
  const userInfo = getCurrentUserInfo();
  
  if (userInfo.type === "admin") return true;
  
  return userInfo.permissions.includes(permission);
}


/**
 * Validate teacher can access specific class/section
 * @param {string} classNum - Class number
 * @param {string} section - Section
 * @returns {boolean} True if can access
 */
function canAccessClass(classNum, section) {
  if (isAdmin()) return true;
  
  const teacher = getTeacherByEmail(getCurrentUser());
  if (!teacher) return false;
  
  const teacherClasses = teacher.classes.split(",").map(c => c.trim());
  const teacherSections = teacher.sections.split(",").map(s => s.trim());
  
  return teacherClasses.includes(String(classNum)) && 
         (teacherSections.includes(section) || teacherSections.includes("All"));
}


/**
 * Validate teacher can edit marks for specific subject
 * @param {string} subject - Subject name
 * @returns {boolean} True if can edit
 */
function canEditSubject(subject) {
  if (isAdmin()) return true;
  
  const teacher = getTeacherByEmail(getCurrentUser());
  if (!teacher) return false;
  
  return teacher.subject === subject || teacher.subject === "All";
}


/**
 * Get access control list for current user
 * @returns {Object} Access control object
 */
function getAccessControl() {
  const userInfo = getCurrentUserInfo();
  
  return {
    user: userInfo,
    canUploadStudents: userInfo.type === "admin",
    canUploadTeachers: userInfo.type === "admin",
    canCreateExam: userInfo.type === "admin",
    canLockExam: userInfo.type === "admin",
    canEnterMarks: userInfo.type === "admin" || userInfo.type === "teacher",
    canViewAllMarks: userInfo.type === "admin",
    canViewAnalytics: true,
    canViewReports: true,
    canModifySettings: userInfo.type === "admin",
    canResetData: userInfo.type === "admin"
  };
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
