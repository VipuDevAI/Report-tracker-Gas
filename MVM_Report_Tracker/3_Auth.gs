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
  "mvmseniors@gmail.com"      // Admin (app access only)
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
 * Check if user has any valid access
 * @returns {boolean} True if admin or teacher
 */
function hasAccess() {
  return isAdmin() || isTeacher();
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
 * Get teacher assignment details for filtering
 * @param {string} email - Teacher email (optional, uses current user if not provided)
 * @returns {Object|null} Teacher assignment object with parsed arrays
 */
function getTeacherAssignment(email) {
  const teacherEmail = email || getCurrentUser();
  const teacher = getTeacherByEmail(teacherEmail);
  
  if (!teacher) return null;
  
  // Parse comma-separated values into arrays
  const classesRaw = teacher.classes || "";
  const sectionsRaw = teacher.sections || "";
  
  return {
    teacherId: teacher.teacherId,
    name: teacher.name,
    email: teacher.email,
    subject: teacher.subject,
    classes: classesRaw.split(",").map(c => c.trim()).filter(c => c),
    sections: sectionsRaw.split(",").map(s => s.trim()).filter(s => s),
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
