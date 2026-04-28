/************************************************
 MVM REPORT TRACKER - INITIALIZATION & SETUP
 File 1 of 7
************************************************/

// Admin email whitelist
const ADMIN_EMAILS = [
  "rishisans83@gmail.com",
  "mvmseniors@gmail.com",
  "anithasivanesan4604@gmail.com"
];

// Grade ranges configuration
const GRADE_RANGES = [
  ["A+", 91, 100],
  ["A", 81, 90],
  ["B+", 71, 80],
  ["B", 61, 70],
  ["C", 51, 60],
  ["D", 41, 50],
  ["F", 0, 40]
];

// Streams configuration (only for Class 11-12; 6-10 has no stream)
const STREAMS = {
  "11": ["Science", "Computer Science", "Commerce"],
  "12": ["Science", "Computer Science", "Commerce"]
};

// Class sections - All non-12th use A1-A11, 11-12 uses A1-A12
const SECTIONS_6_10 = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11"];
const SECTIONS_9_10 = SECTIONS_6_10; // alias for backward compat
const SECTIONS_11_12 = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12"];

// Legacy constant for backward compatibility
const SECTIONS = SECTIONS_6_10;

// Performance safety thresholds
const MARKS_ROW_WARNING_THRESHOLD = 200000;

/**
 * Get sections for a specific class
 * @param {string|number} classNum - Class number
 * @returns {Array} Sections array
 */
function getSectionsForClass(classNum) {
  const cls = parseInt(classNum);
  if (cls === 11 || cls === 12) {
    return SECTIONS_11_12;
  }
  // 6, 7, 8, 9, 10 → A1-A11
  return SECTIONS_6_10;
}

// Weak student threshold
const WEAK_THRESHOLD = 40;

/**
 * Initialize the entire application
 * Creates all necessary sheets with proper structure
 */
function initializeApp() {
  const ss = SpreadsheetApp.getActive();
  
  const structure = {
    Students: [
      "StudentID", "Name", "Class", "Section", "Stream", 
      "RollNo", "ParentEmail", "Phone", "JoinDate", "Status",
      "ElectiveSubject", "AcademicYear",
      "LanguageL1", "LanguageL2", "LanguageL3",
      "IsDeleted"
    ],
    Teachers: [
      "TeacherID", "Name", "Subject", "Classes", "Sections", 
      "Email", "Phone", "JoinDate", "Status", "IsClassTeacher", "ClassTeacherOf",
      "IsDeleted", "Role"
    ],
    Subjects: [
      "SubjectID", "SubjectName", "SubjectCode", "Class", "Stream", 
      "MaxMarks", "PassingMarks", "IsActive", "LanguageGroup", "IsOptional"
    ],
    Classes: [
      "ClassID", "ClassName", "Sections", "Stream", "AcademicYear", "IsActive"
    ],
    Exams: [
      "ExamID", "ExamName", "ExamType", "Class", "MaxMarks", 
      "Weightage", "StartDate", "EndDate", "Locked", "CreatedBy", "CreatedAt", "AcademicYear",
      "HasInternals", "Internal1", "Internal2", "Internal3", "Internal4", "TotalMaxMarks",
      "IsDeleted"
    ],
    Marks_Master: [
      "EntryID", "StudentID", "StudentName", "Subject", "SubjectCode",
      "TeacherID", "TeacherName", "ExamID", "ExamName", "Class", "Section",
      "MaxMarks", "MarksObtained", "Percentage", "Grade", "UpdatedAt", "UpdatedBy", "AcademicYear",
      "Status", "IsDeleted"
    ],
    Auth: [
      "Email", "PasswordHash", "Salt", "MustChangePassword",
      "SessionToken", "SessionExpiry", "FailedAttempts", "LastLogin", "CreatedAt"
    ],
    Audit_Trail: [
      "Timestamp", "Action", "EntityType", "EntityID", "Field",
      "OldValue", "NewValue", "ChangedBy", "Context"
    ],
    Settings_Ranges: [
      "RangeName", "GradeLabel", "MinMarks", "MaxMarks", "Color"
    ],
    Settings_School: [
      "SettingKey", "SettingValue", "UpdatedAt"
    ],
    Aggregates: [
      "Type", "Key", "SubKey", "Value", "Count", "UpdatedAt"
    ],
    Alerts: [
      "AlertID", "AlertType", "StudentID", "StudentName", "Class", 
      "Subject", "Message", "Priority", "IsRead", "CreatedAt"
    ],
    Logs: [
      "LogID", "Action", "User", "Details", "Timestamp"
    ]
  };

  // Create sheets with headers
  Object.keys(structure).forEach(name => {
    let sheet = ss.getSheetByName(name);
    
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clearContents();
    }
    
    // Set headers with formatting
    const headerRange = sheet.getRange(1, 1, 1, structure[name].length);
    headerRange.setValues([structure[name]]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1a6b3a");
    headerRange.setFontColor("#ffffff");
    
    // Freeze header row
    sheet.setFrozenRows(1);
  });

  // Seed default data
  seedDefaultRanges();
  seedDefaultSchoolSettings();
  seedDefaultSubjects();
  seedDefaultClasses();
  
  // Idempotent migrations (safe on existing deployments)
  ensureTeachersRoleColumn();
  
  logAction("Initialize App", "Application initialized successfully");
  
  SpreadsheetApp.flush();
  
  return { success: true, message: "MVM Report Tracker initialized successfully!" };
}


/**
 * Seed default grade ranges (numeric only, no letter grades)
 */
function seedDefaultRanges() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_Ranges");
  
  const ranges = [
    ["91-100", "Excellent", 91, 100, "#22c55e"],
    ["81-90", "Very Good", 81, 90, "#16a34a"],
    ["71-80", "Good", 71, 80, "#3b82f6"],
    ["61-70", "Above Average", 61, 70, "#0ea5e9"],
    ["51-60", "Average", 51, 60, "#f59e0b"],
    ["41-50", "Below Average", 41, 50, "#f97316"],
    ["0-40", "Needs Improvement", 0, 40, "#ef4444"]
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, ranges.length, 5).setValues(ranges);
  }
}


/**
 * Seed default school settings
 */
function seedDefaultSchoolSettings() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
  
  const settings = [
    ["SchoolName", "MVM School", new Date()],
    ["SchoolCode", "MVM", new Date()],
    ["AcademicYear", "2024-2025", new Date()],
    ["WeakThreshold", "40", new Date()],
    ["PassingPercentage", "40", new Date()],
    ["LogoURL", "", new Date()],
    ["Address", "", new Date()],
    ["Phone", "", new Date()],
    ["Email", "", new Date()],
    ["IsYearFinalized", "false", new Date()],
    ["LastAggregatesUpdatedAt", "", new Date()],
    ["SessionDurationHours", "8", new Date()],
    ["Wing_Primary", "6,7,8", new Date()],
    ["Wing_Secondary", "9,10", new Date()],
    ["Wing_Senior", "11,12", new Date()]
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, settings.length, 3).setValues(settings);
    return;
  }
  
  // Idempotent: append any missing setting keys (e.g., Wing_* on existing deployments)
  const existing = sheet.getDataRange().getValues();
  const existingKeys = new Set(existing.slice(1).map(r => String(r[0] || "")));
  const toAppend = settings.filter(s => !existingKeys.has(s[0]));
  if (toAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAppend.length, 3).setValues(toAppend);
  }
}


/**
 * Idempotent: ensure the Teachers sheet has the new "Role" column.
 * Safe to call repeatedly. Returns true if a migration was applied.
 */
function ensureTeachersRoleColumn() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  if (!sheet) return false;
  const lastCol = sheet.getLastColumn();
  if (lastCol >= 13) {
    // Header may already be there; ensure label
    const header = sheet.getRange(1, 13).getValue();
    if (!header) sheet.getRange(1, 13).setValue("Role");
    return false;
  }
  // Add the Role header at column 13
  sheet.getRange(1, 13).setValue("Role");
  sheet.getRange(1, 13).setFontWeight("bold");
  sheet.getRange(1, 13).setBackground("#1a6b3a");
  sheet.getRange(1, 13).setFontColor("#ffffff");
  // Default existing rows to "TEACHER" (skip blank rows)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const emails = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    const roleCol = emails.map(r => [r[0] ? "TEACHER" : ""]);
    sheet.getRange(2, 13, lastRow - 1, 1).setValues(roleCol);
  }
  return true;
}


/**
 * Seed default subjects for all classes (6-12)
 * Schema: SubjectID, SubjectName, SubjectCode, Class, Stream, MaxMarks, PassingMarks, IsActive, LanguageGroup, IsOptional
 * - Class 6-10: stream blank (common)
 * - Class 11-12: stream-based (Science/Computer Science/Commerce/Elective)
 * - LanguageGroup: "L1" | "L2" | "L3" | "L2,L3" (comma-sep slots a language is allowed in) — empty for non-languages
 * - IsOptional: true if student can opt-in/swap (typically languages and electives)
 */
function seedDefaultSubjects() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Subjects");
  
  const subjects = [
    // ==================== CLASS 6-8 (no stream, 3-language system) ====================
    // Mandatory non-language subjects
    ["SUB101", "Mathematics",        "MATH",  "6,7,8", "", 100, 35, true, "",      false],
    ["SUB102", "Science",            "SCI",   "6,7,8", "", 100, 35, true, "",      false],
    ["SUB103", "Social Science",     "SST",   "6,7,8", "", 100, 35, true, "",      false],
    ["SUB104", "Computer Application","COMP", "6,7,8", "", 100, 35, true, "",      false],
    ["SUB105", "Art Education",      "ART",   "6,7,8", "", 100, 35, true, "",      false],
    // Languages — student picks one per slot (L1 fixed English; L2/L3 from Hindi/Sanskrit/Tamil)
    ["SUB110", "English",            "ENG",   "6,7,8", "", 100, 35, true, "L1",    false],
    ["SUB111", "Hindi",              "HIN",   "6,7,8", "", 100, 35, true, "L2,L3", true],
    ["SUB112", "Sanskrit",           "SANS",  "6,7,8", "", 100, 35, true, "L2,L3", true],
    ["SUB113", "Tamil",              "TAM",   "6,7,8", "", 100, 35, true, "L2,L3", true],

    // ==================== CLASS 9-10 (no stream, 2-language system) ====================
    ["SUB201", "Mathematics",        "MATH",  "9,10", "", 100, 35, true, "",      false],
    ["SUB202", "Science",            "SCI",   "9,10", "", 100, 35, true, "",      false],
    ["SUB203", "Social Science",     "SST",   "9,10", "", 100, 35, true, "",      false],
    ["SUB204", "Computer Application","COMP", "9,10", "", 100, 35, true, "",      false],
    ["SUB210", "English",            "ENG",   "9,10", "", 100, 35, true, "L1",    false],
    ["SUB211", "Hindi",              "HIN",   "9,10", "", 100, 35, true, "L2",    true],
    ["SUB212", "Sanskrit",           "SANS",  "9,10", "", 100, 35, true, "L2",    true],
    ["SUB213", "Tamil",              "TAM",   "9,10", "", 100, 35, true, "L2",    true],

    // ==================== CLASS 11-12 — Science Stream ====================
    ["SUB301", "Physics",            "PHY",   "11,12", "Science", 100, 35, true, "", false],
    ["SUB302", "Chemistry",          "CHEM",  "11,12", "Science", 100, 35, true, "", false],
    ["SUB303", "Biology",            "BIO",   "11,12", "Science", 100, 35, true, "", false],
    ["SUB304", "English",            "ENG",   "11,12", "Science", 100, 35, true, "L1", false],

    // ==================== CLASS 11-12 — Computer Science Stream ====================
    ["SUB311", "Physics",            "PHY",   "11,12", "Computer Science", 100, 35, true, "", false],
    ["SUB312", "Chemistry",          "CHEM",  "11,12", "Computer Science", 100, 35, true, "", false],
    ["SUB313", "Computer Science",   "CS",    "11,12", "Computer Science", 100, 35, true, "", false],
    ["SUB314", "English",            "ENG",   "11,12", "Computer Science", 100, 35, true, "L1", false],

    // ==================== CLASS 11-12 — Commerce Stream ====================
    ["SUB321", "Accountancy",        "ACC",   "11,12", "Commerce", 100, 35, true, "", false],
    ["SUB322", "Business Studies",   "BS",    "11,12", "Commerce", 100, 35, true, "", false],
    ["SUB323", "Economics",          "ECO",   "11,12", "Commerce", 100, 35, true, "", false],
    ["SUB324", "English",            "ENG",   "11,12", "Commerce", 100, 35, true, "L1", false],

    // ==================== CLASS 11-12 — ELECTIVES (student picks ONE, available to all streams) ====================
    ["SUB331", "Mathematics",        "MATH",  "11,12", "Elective", 100, 35, true, "", true],
    ["SUB332", "Applied Mathematics","AMATH", "11,12", "Elective", 100, 35, true, "", true],
    ["SUB333", "Hindi",              "HIN",   "11,12", "Elective", 100, 35, true, "", true],
    ["SUB334", "History",            "HIST",  "11,12", "Elective", 100, 35, true, "", true],
    ["SUB335", "Sanskrit",           "SANS",  "11,12", "Elective", 100, 35, true, "", true],
    ["SUB336", "Computer Science",   "CS",    "11,12", "Elective", 100, 35, true, "", true],
    ["SUB337", "Biology",            "BIO",   "11,12", "Elective", 100, 35, true, "", true]
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, subjects.length, 10).setValues(subjects);
  }
}


/**
 * Admin: Reset Subjects sheet to default (clears existing rows, reseeds defaults)
 * Called from "Reset Subjects to Default" button
 * @returns {Object} Result object
 */
function resetSubjectsToDefault() {
  if (typeof isAdmin === 'function' && !isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const sheet = SpreadsheetApp.getActive().getSheetByName("Subjects");
    if (!sheet) {
      return { success: false, message: "Subjects sheet not found. Run Initialize App first." };
    }

    // Clear all rows except header
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    // Re-seed defaults
    seedDefaultSubjects();

    // Invalidate cache
    _invalidateSubjectsCache();

    logAction("Reset Subjects", "Subjects sheet reset to defaults");

    return { success: true, message: "Subjects sheet reset to default values." };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Cached read of Subjects sheet (in-memory per execution)
 * Schema: SubjectID, SubjectName, SubjectCode, Class, Stream, MaxMarks, PassingMarks, IsActive, LanguageGroup, IsOptional
 * @returns {Array} Array of subject objects
 */
let _subjectsCache = null;
function _getSubjectsCache() {
  if (_subjectsCache !== null) return _subjectsCache;
  const sheet = SpreadsheetApp.getActive().getSheetByName("Subjects");
  if (!sheet || sheet.getLastRow() <= 1) {
    _subjectsCache = [];
    return _subjectsCache;
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  _subjectsCache = data
    .filter(row => row[0] && row[7] !== false)
    .map(row => ({
      subjectId: row[0],
      subjectName: String(row[1] || "").trim(),
      subjectCode: String(row[2] || "").trim(),
      classes: String(row[3] || "").split(",").map(c => c.trim()).filter(Boolean),
      stream: String(row[4] || "").trim(),
      maxMarks: row[5],
      passingMarks: row[6],
      isActive: row[7],
      languageGroups: String(row[8] || "").split(",").map(g => g.trim()).filter(Boolean),
      isOptional: row[9] === true
    }));
  return _subjectsCache;
}

function _invalidateSubjectsCache() {
  _subjectsCache = null;
}


/**
 * Get list of valid subject names for a given student
 * Logic:
 *   - All subjects matching (class, stream) where IsOptional=false AND no LanguageGroup → mandatory
 *   - English (LanguageGroup includes "L1") → mandatory L1
 *   - Student-chosen LanguageL2 / LanguageL3 → must match a subject with LanguageGroup containing "L2"/"L3"
 *   - Class 11-12 with stream != Elective: also include the chosen ElectiveSubject if it exists in Subjects (stream="Elective")
 * @param {Object} student - { class, stream, electiveSubject, languageL1, languageL2, languageL3 }
 * @returns {Array<string>} Valid subject names
 */
function getValidSubjectsForStudent(student) {
  if (!student) return [];
  const cls = String(student.class);
  const stream = String(student.stream || "").trim();
  const elective = String(student.electiveSubject || "").trim();
  const lang1 = String(student.languageL1 || "").trim();
  const lang2 = String(student.languageL2 || "").trim();
  const lang3 = String(student.languageL3 || "").trim();
  const cache = _getSubjectsCache();
  const result = new Set();

  cache.forEach(s => {
    if (!s.classes.includes(cls)) return;

    // For 6-10: stream is empty in Subjects (common). Match if subject stream is "" OR matches student stream.
    // For 11-12: stream must match exactly (or "Elective" handled separately).
    const isCommonClass = (cls === "6" || cls === "7" || cls === "8" || cls === "9" || cls === "10");
    const streamOk = isCommonClass
      ? (s.stream === "" || s.stream === stream)
      : (s.stream === stream);

    if (s.languageGroups.length > 0) {
      // Language subject → match by student's chosen language slot
      const isL1 = s.languageGroups.indexOf("L1") !== -1;
      const isL2 = s.languageGroups.indexOf("L2") !== -1;
      const isL3 = s.languageGroups.indexOf("L3") !== -1;
      if (streamOk) {
        if (isL1 && (!lang1 || s.subjectName === lang1)) result.add(s.subjectName);
        if (isL2 && lang2 && s.subjectName === lang2) result.add(s.subjectName);
        if (isL3 && lang3 && s.subjectName === lang3) result.add(s.subjectName);
      }
    } else if (s.stream === "Elective") {
      // Elective subjects (11-12 only) — only valid if student picked it
      if (elective && s.subjectName === elective) result.add(s.subjectName);
    } else if (streamOk && !s.isOptional) {
      // Mandatory subject for class+stream
      result.add(s.subjectName);
    } else if (streamOk && s.isOptional) {
      // Optional but stream matches — allow (e.g., optional language alternates)
      // Already handled by language block above; this catches non-language optionals
      result.add(s.subjectName);
    }
  });

  return Array.from(result);
}


/**
 * Check if a subject is valid for a given student
 * @param {string} subject - Subject name
 * @param {Object} student - Student object
 * @returns {boolean}
 */
function isSubjectValidForStudent(subject, student) {
  if (!subject || !student) return false;
  const valid = getValidSubjectsForStudent(student);
  const target = String(subject).trim().toLowerCase();
  return valid.some(v => String(v).trim().toLowerCase() === target);
}


/**
 * Seed default classes (6-12)
 * Sections: 6-10 → A1-A11; 11-12 → A1-A12
 */
function seedDefaultClasses() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Classes");
  
  const classes = [];
  const academicYear = "2024-2025";
  const sections6_10 = SECTIONS_6_10.join(",");
  const sections11_12 = SECTIONS_11_12.join(",");
  
  // Class 6-10: no stream (single common entry per class)
  for (let cls = 6; cls <= 10; cls++) {
    classes.push([
      `CLS${cls}`,
      `Class ${cls}`,
      sections6_10,
      "",  // no stream for 6-10
      academicYear,
      true
    ]);
  }
  
  // Class 11-12 with stream-based entries
  const streams = ["Science", "Computer Science", "Commerce"];
  for (let cls = 11; cls <= 12; cls++) {
    streams.forEach((stream, idx) => {
      classes.push([
        `CLS${cls}${idx + 1}`,
        `Class ${cls}`,
        sections11_12,
        stream,
        academicYear,
        true
      ]);
    });
  }

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, classes.length, 6).setValues(classes);
  }
}


/**
 * Reset school data (keeps structure, clears data)
 */
function resetSchool() {
  const ss = SpreadsheetApp.getActive();
  
  const dataSheets = [
    "Students",
    "Marks_Master",
    "Aggregates",
    "Alerts",
    "Logs"
  ];

  dataSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  });

  logAction("Reset School", "Student and marks data cleared");
  
  return { success: true, message: "School data reset successfully!" };
}


/**
 * Archive current year data and reset
 */
function archiveAndReset() {
  const ss = SpreadsheetApp.getActive();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy_MM_dd_HHmm");
  
  const sheetsToArchive = ["Students", "Marks_Master", "Aggregates"];
  
  sheetsToArchive.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      const copy = sheet.copyTo(ss);
      copy.setName(`${name}_ARCHIVE_${timestamp}`);
    }
  });

  resetSchool();
  logAction("Archive & Reset", `Data archived with timestamp: ${timestamp}`);
  
  return { success: true, message: `Data archived with timestamp: ${timestamp}` };
}


/**
 * Log action to Logs sheet
 */
function logAction(action, details) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Logs");
  const user = Session.getActiveUser().getEmail() || "System";
  const logId = `LOG${Date.now()}`;
  
  sheet.appendRow([
    logId,
    action,
    user,
    details || "",
    new Date()
  ]);
}


/**
 * Create custom menu on spreadsheet open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('MVM Report Tracker')
    .addItem('Open Dashboard', 'openDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Master Data')
      .addItem('Upload Students', 'showStudentUpload')
      .addItem('Upload Teachers', 'showTeacherUpload')
      .addItem('Manage Subjects', 'showSubjects')
      .addItem('Manage Classes', 'showClasses'))
    .addSubMenu(ui.createMenu('Exams')
      .addItem('Create Exam', 'showCreateExam')
      .addItem('Enter Marks', 'showMarksEntry')
      .addItem('View Marks', 'showViewMarks')
      .addItem('Lock/Unlock Exam', 'showExamLock'))
    .addSubMenu(ui.createMenu('Analytics')
      .addItem('Rebuild Analytics', 'rebuildAggregates')
      .addItem('View Weak Students', 'showWeakStudents')
      .addItem('View Toppers', 'showToppers'))
    .addSubMenu(ui.createMenu('Reports')
      .addItem('Subject Report', 'showSubjectReport')
      .addItem('Class Report', 'showClassReport')
      .addItem('Student Report', 'showStudentReport'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Settings')
      .addItem('Grade Ranges', 'showGradeRanges')
      .addItem('School Info', 'showSchoolInfo'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Admin')
      .addItem('Initialize App', 'initializeApp')
      .addItem('Reset Subjects to Default', 'resetSubjectsToDefault')
      .addItem('Reset School Data', 'resetSchool')
      .addSeparator()
      .addItem('Archive Academic Year', 'showArchiveYearPrompt')
      .addItem('Reset for New Academic Year', 'showResetNewYearPrompt')
      .addItem('Switch Academic Year', 'showSwitchYearPrompt')
      .addItem('Archive & Reset Year (Legacy)', 'archiveAndReset'))
    .addToUi();
}


/**
 * Open the main dashboard
 */
function openDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1400)
    .setHeight(900)
    .setTitle('MVM Report Tracker - Dashboard');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'MVM Report Tracker');
}


/**
 * doGet - Entry point for Web App deployment
 * This function is called when the web app URL is accessed
 * @param {Object} e - Event parameter
 * @returns {HtmlOutput} The dashboard HTML
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('MVM Report Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/**
 * Get the spreadsheet URL to share
 * @returns {string} Spreadsheet URL
 */
function getSpreadsheetUrl() {
  return SpreadsheetApp.getActive().getUrl();
}


/**
 * Create class-wise student sheets for Class 11 & 12
 * Creates: Students_11_A1, Students_11_A2, ... Students_12_A12
 */
function createClassWiseSheets() {
  const ss = SpreadsheetApp.getActive();
  
  const headers = [
    "SNo", "Name", "RollNo", "Stream", "ElectiveSubject", 
    "ParentName", "ParentPhone", "ParentEmail", "Address", "Status"
  ];
  
  const sections = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12"];
  const classes = [11, 12];
  
  let sheetsCreated = 0;
  
  classes.forEach(cls => {
    sections.forEach(section => {
      const sheetName = `Students_${cls}_${section}`;
      
      // Check if sheet already exists
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheetsCreated++;
      }
      
      // Set headers if first row is empty
      if (sheet.getRange(1, 1).getValue() === "") {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        // Style header row
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setFontWeight("bold");
        headerRange.setBackground("#059669");
        headerRange.setFontColor("white");
        headerRange.setHorizontalAlignment("center");
        
        // Add data validation for Stream
        const streamRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["Science", "Computer Science", "Commerce"], true)
          .build();
        sheet.getRange(2, 4, 500, 1).setDataValidation(streamRule);
        
        // Add data validation for ElectiveSubject
        const electiveRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["Mathematics", "Applied Mathematics", "Hindi", "History", "Sanskrit"], true)
          .build();
        sheet.getRange(2, 5, 500, 1).setDataValidation(electiveRule);
        
        // Add data validation for Status
        const statusRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["Active", "Inactive", "TC", "Promoted"], true)
          .build();
        sheet.getRange(2, 10, 500, 1).setDataValidation(statusRule);
        
        // Auto-resize columns
        for (let i = 1; i <= headers.length; i++) {
          sheet.autoResizeColumn(i);
        }
        
        // Set column widths
        sheet.setColumnWidth(2, 200); // Name
        sheet.setColumnWidth(9, 250); // Address
        
        // Add sheet description
        sheet.getRange(1, headers.length + 2).setValue(`Class ${cls} - Section ${section}`);
        sheet.getRange(1, headers.length + 2).setFontWeight("bold").setFontColor("#059669");
      }
    });
  });
  
  // Create Teachers Master sheet
  createTeachersMasterSheet();
  
  // Create Class Teachers sheet
  createClassTeachersSheet();
  
  logAction("Create Sheets", `Created ${sheetsCreated} class-wise student sheets`);
  
  return { 
    success: true, 
    message: `Created ${sheetsCreated} class-wise sheets for Class 11 & 12 (A1-A12)` 
  };
}


/**
 * Create Teachers Master Data sheet
 */
function createTeachersMasterSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheetName = "Teachers_Master";
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  const headers = [
    "SNo", "Name", "Subject", "Qualification", "Experience", 
    "Classes", "Sections", "Email", "Phone", "Address", 
    "JoinDate", "IsClassTeacher", "ClassTeacherOf", "Status"
  ];
  
  if (sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#059669");
    headerRange.setFontColor("white");
    headerRange.setHorizontalAlignment("center");
    
    // Data validation for IsClassTeacher
    const yesNoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Yes", "No"], true)
      .build();
    sheet.getRange(2, 12, 200, 1).setDataValidation(yesNoRule);
    
    // Data validation for Status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Active", "Inactive", "On Leave", "Resigned"], true)
      .build();
    sheet.getRange(2, 14, 200, 1).setDataValidation(statusRule);
    
    // Auto-resize
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    sheet.setColumnWidth(2, 200); // Name
    sheet.setColumnWidth(10, 250); // Address
  }
  
  return sheet;
}


/**
 * Create Class Teachers Assignment sheet
 */
function createClassTeachersSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheetName = "Class_Teachers";
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  const headers = ["Class", "Section", "ClassTeacherName", "ClassTeacherEmail", "ClassTeacherPhone"];
  
  if (sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Style header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#059669");
    headerRange.setFontColor("white");
    
    // Pre-populate class/section combinations
    const sections = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12"];
    const classes = [11, 12];
    
    let rowNum = 2;
    classes.forEach(cls => {
      sections.forEach(section => {
        sheet.getRange(rowNum, 1).setValue(cls);
        sheet.getRange(rowNum, 2).setValue(section);
        rowNum++;
      });
    });
    
    // Auto-resize
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  
  return sheet;
}


/**
 * Sync students from class-wise sheets to main Students sheet
 * Writes 15 cols (with AcademicYear + 3 language slots)
 */
function syncStudentsFromClassSheets() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const ss = SpreadsheetApp.getActive();
    const mainSheet = ss.getSheetByName("Students");
    
    const sections = SECTIONS_11_12;
    const classes = [11, 12];
    const academicYear = getCurrentAcademicYear();
    
    let totalSynced = 0;
    let allStudents = [];
    
    classes.forEach(cls => {
      sections.forEach(section => {
        const sheetName = `Students_${cls}_${section}`;
        const sheet = ss.getSheetByName(sheetName);
        
        if (sheet && sheet.getLastRow() > 1) {
          const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
          
          data.forEach((row, idx) => {
            if (row[1]) { // If name exists
              const studentId = `STU${cls}${section}${String(idx + 1).padStart(3, '0')}`;
              allStudents.push([
                studentId,
                row[1],  // Name
                cls,     // Class
                section, // Section
                row[3] || "Science",  // Stream
                row[2] || idx + 1,    // RollNo
                row[7] || "",         // ParentEmail
                row[6] || "",         // Phone
                new Date(),           // JoinDate
                row[9] || "Active",   // Status
                row[4] || "",         // ElectiveSubject
                academicYear,         // AcademicYear
                "English",            // LanguageL1
                "",                   // LanguageL2
                ""                    // LanguageL3
              ]);
              totalSynced++;
            }
          });
        }
      });
    });
    
    if (allStudents.length > 0) {
      // Clear existing data (keep header)
      if (mainSheet.getLastRow() > 1) {
        mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 15).clearContent();
      }
      
      // Write all students
      mainSheet.getRange(2, 1, allStudents.length, 15).setValues(allStudents);
    }
    
    logAction("Sync Students", `Synced ${totalSynced} students from class-wise sheets`);
    
    return {
      success: true,
      message: `Synced ${totalSynced} students from class-wise sheets to main Students sheet`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get count of rows in Marks_Master (for performance safety warning)
 * @returns {Object} { count, threshold, warning, message }
 */
function getMarksRowCount() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  if (!sheet) return { count: 0, threshold: MARKS_ROW_WARNING_THRESHOLD, warning: false };
  const count = Math.max(0, sheet.getLastRow() - 1);
  const warning = count >= MARKS_ROW_WARNING_THRESHOLD;
  return {
    count: count,
    threshold: MARKS_ROW_WARNING_THRESHOLD,
    warning: warning,
    message: warning
      ? `Warning: Marks_Master has ${count} rows (threshold: ${MARKS_ROW_WARNING_THRESHOLD}). Consider running Year-End Archive to keep performance optimal.`
      : `Marks_Master: ${count} rows (well within ${MARKS_ROW_WARNING_THRESHOLD} threshold).`
  };
}


/**
 * Year-End: Archive academic year data to a separate spreadsheet
 * Exports Students, Marks_Master, Exams (filtered by year)
 * @param {string} year - Academic year to archive (e.g., "2024-2025")
 * @returns {Object} Result with archive URL
 */
function archiveAcademicYear(year) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  if (!year) {
    return { success: false, message: "Academic year is required." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActive();
    const safeYear = String(year).replace(/[^0-9A-Za-z\-]/g, '_');
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
    const archiveName = `MVM_Archive_${safeYear}_${stamp}`;
    const archiveSS = SpreadsheetApp.create(archiveName);

    // Copy Students for this year
    const studentsSheet = ss.getSheetByName("Students");
    if (studentsSheet && studentsSheet.getLastRow() > 1) {
      const data = studentsSheet.getRange(1, 1, studentsSheet.getLastRow(), 15).getValues();
      const filtered = [data[0]].concat(data.slice(1).filter(r => String(r[11]) === String(year)));
      const target = archiveSS.insertSheet("Students");
      if (filtered.length > 0) target.getRange(1, 1, filtered.length, 15).setValues(filtered);
    }

    // Copy Marks_Master for this year
    const marksSheet = ss.getSheetByName("Marks_Master");
    if (marksSheet && marksSheet.getLastRow() > 1) {
      const data = marksSheet.getRange(1, 1, marksSheet.getLastRow(), 18).getValues();
      const filtered = [data[0]].concat(data.slice(1).filter(r => String(r[17]) === String(year)));
      const target = archiveSS.insertSheet("Marks_Master");
      if (filtered.length > 0) target.getRange(1, 1, filtered.length, 18).setValues(filtered);
    }

    // Copy Exams for this year
    const examsSheet = ss.getSheetByName("Exams");
    if (examsSheet && examsSheet.getLastRow() > 1) {
      const data = examsSheet.getRange(1, 1, examsSheet.getLastRow(), 18).getValues();
      const filtered = [data[0]].concat(data.slice(1).filter(r => String(r[11]) === String(year)));
      const target = archiveSS.insertSheet("Exams");
      if (filtered.length > 0) target.getRange(1, 1, filtered.length, 18).setValues(filtered);
    }

    // Drop default sheet "Sheet1"
    const defaultSheet = archiveSS.getSheetByName("Sheet1");
    if (defaultSheet && archiveSS.getSheets().length > 1) archiveSS.deleteSheet(defaultSheet);

    logAction("Archive Year", `Archived ${year} -> ${archiveSS.getUrl()}`);

    return {
      success: true,
      url: archiveSS.getUrl(),
      fileName: archiveName,
      message: `Year ${year} archived successfully.`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Year-End: Reset Marks_Master and Exams for new academic year (keeps Students)
 * @param {string} newYear - New academic year (e.g., "2025-2026")
 * @returns {Object} Result
 */
function resetForNewYear(newYear) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  if (!newYear) {
    return { success: false, message: "New academic year is required." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActive();

    const marksSheet = ss.getSheetByName("Marks_Master");
    if (marksSheet && marksSheet.getLastRow() > 1) {
      marksSheet.getRange(2, 1, marksSheet.getLastRow() - 1, marksSheet.getLastColumn()).clearContent();
    }

    const examsSheet = ss.getSheetByName("Exams");
    if (examsSheet && examsSheet.getLastRow() > 1) {
      examsSheet.getRange(2, 1, examsSheet.getLastRow() - 1, examsSheet.getLastColumn()).clearContent();
    }

    const aggSheet = ss.getSheetByName("Aggregates");
    if (aggSheet && aggSheet.getLastRow() > 1) {
      aggSheet.getRange(2, 1, aggSheet.getLastRow() - 1, aggSheet.getLastColumn()).clearContent();
    }

    switchAcademicYear(newYear);

    logAction("Reset For New Year", `Cleared Marks/Exams; set academic year = ${newYear}`);

    return {
      success: true,
      message: `Reset complete. New academic year is ${newYear}. Students preserved (use Promote Students to advance class).`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Switch the active academic year in Settings_School
 * @param {string} newYear
 * @returns {Object} Result
 */
function switchAcademicYear(newYear) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  if (!newYear) {
    return { success: false, message: "New academic year is required." };
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
  if (!sheet) {
    return { success: false, message: "Settings_School sheet not found. Run Initialize App first." };
  }

  const data = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "AcademicYear") {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2).setValue(newYear);
    sheet.getRange(foundRow, 3).setValue(new Date());
  } else {
    sheet.appendRow(["AcademicYear", newYear, new Date()]);
  }

  logAction("Switch Academic Year", `Active academic year set to ${newYear}`);

  return { success: true, message: `Active academic year is now ${newYear}.` };
}


/**
 * Sync teachers from Teachers_Master to main Teachers sheet
 */
function syncTeachersFromMaster() {
  const ss = SpreadsheetApp.getActive();
  const masterSheet = ss.getSheetByName("Teachers_Master");
  const mainSheet = ss.getSheetByName("Teachers");
  
  if (!masterSheet || masterSheet.getLastRow() <= 1) {
    return { success: false, message: "No teachers found in Teachers_Master sheet" };
  }
  
  const data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 14).getValues();
  let allTeachers = [];
  
  data.forEach((row, idx) => {
    if (row[1] && row[7]) { // If name and email exist
      const teacherId = `TCH${String(idx + 1).padStart(4, '0')}`;
      allTeachers.push([
        teacherId,
        row[1],  // Name
        row[2],  // Subject
        row[5],  // Classes
        row[6],  // Sections
        row[7],  // Email
        row[8],  // Phone
        row[10] || new Date(), // JoinDate
        row[13] || "Active",   // Status
        row[11] || "No",       // IsClassTeacher
        row[12] || ""          // ClassTeacherOf
      ]);
    }
  });
  
  if (allTeachers.length > 0) {
    // Clear existing data (keep header)
    if (mainSheet.getLastRow() > 1) {
      mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 11).clearContent();
    }
    
    // Write all teachers
    mainSheet.getRange(2, 1, allTeachers.length, 11).setValues(allTeachers);
  }
  
  logAction("Sync Teachers", `Synced ${allTeachers.length} teachers from master sheet`);
  
  return {
    success: true,
    message: `Synced ${allTeachers.length} teachers from Teachers_Master sheet`
  };
}



/**
 * Menu helper: Prompt admin for academic year to archive
 */
function showArchiveYearPrompt() {
  const ui = SpreadsheetApp.getUi();
  const current = getCurrentAcademicYear();
  const resp = ui.prompt(
    'Archive Academic Year',
    `Enter the academic year to archive (default: ${current}). A new spreadsheet will be created with that year's Students, Marks, and Exams.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const year = resp.getResponseText().trim() || current;
  const result = archiveAcademicYear(year);
  ui.alert(result.success ? `Archived: ${result.url}` : `Error: ${result.message}`);
}


/**
 * Menu helper: Prompt admin for new year to reset to
 */
function showResetNewYearPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Reset for New Academic Year',
    'This will CLEAR Marks_Master, Exams, and Aggregates (Students preserved). Enter the NEW academic year (e.g., 2025-2026):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const newYear = resp.getResponseText().trim();
  if (!newYear) { ui.alert('Cancelled — no year entered.'); return; }
  const confirm = ui.alert('Confirm', `Clear Marks/Exams and switch to ${newYear}? This cannot be undone (run Archive first).`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  const result = resetForNewYear(newYear);
  ui.alert(result.message);
}


/**
 * Menu helper: Prompt admin to switch academic year (no data clearing)
 */
function showSwitchYearPrompt() {
  const ui = SpreadsheetApp.getUi();
  const current = getCurrentAcademicYear();
  const resp = ui.prompt(
    'Switch Academic Year',
    `Current: ${current}. Enter the academic year to switch to (e.g., 2025-2026):`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const newYear = resp.getResponseText().trim();
  if (!newYear) { ui.alert('Cancelled.'); return; }
  const result = switchAcademicYear(newYear);
  ui.alert(result.message);
}


/* ========================================================================
   AUDIT TRAIL, YEAR FREEZE, SETTINGS, AGGREGATES TIMESTAMP
   ======================================================================== */

/**
 * Append entry to Audit_Trail sheet
 * @param {string} action - e.g. "UPDATE_MARKS", "DELETE_STUDENT", "FREEZE_YEAR"
 * @param {string} entityType - "Marks" | "Student" | "Exam" | "Auth" | "System"
 * @param {string} entityId
 * @param {string} field - field changed (or "*" for full row)
 * @param {*} oldValue
 * @param {*} newValue
 * @param {Object} [context] - optional extra info (will be JSON-stringified)
 */
function writeAudit(action, entityType, entityId, field, oldValue, newValue, context) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Audit_Trail");
    if (!sheet) return; // graceful: not yet initialized
    const changedBy = (typeof getActualUserEmail === 'function' && getActualUserEmail()) || "System";
    sheet.appendRow([
      new Date(),
      String(action || ""),
      String(entityType || ""),
      String(entityId || ""),
      String(field || ""),
      oldValue === undefined || oldValue === null ? "" : String(oldValue),
      newValue === undefined || newValue === null ? "" : String(newValue),
      changedBy,
      context ? JSON.stringify(context) : ""
    ]);
  } catch (e) {
    // never break a write because audit failed
  }
}


/**
 * Get a school setting value
 * @param {string} key
 * @returns {string} value or "" if not found
 */
function getSchoolSetting(key) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
  if (!sheet || sheet.getLastRow() <= 1) return "";
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) return String(data[i][1] !== undefined ? data[i][1] : "");
  }
  return "";
}


/**
 * Set a school setting value
 * @param {string} key
 * @param {*} value
 */
function updateSchoolSetting(key, value) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      sheet.getRange(i + 2, 3).setValue(new Date());
      return;
    }
  }
  sheet.appendRow([key, value, new Date()]);
}


/**
 * Admin-only wrapper for setting updates from UI
 */
function adminUpdateSchoolSetting(key, value) {
  if (!isAdmin()) return { success: false, message: "Access denied." };
  updateSchoolSetting(key, value);
  logAction("Update Setting", `${key} = ${value}`);
  writeAudit("UPDATE_SETTING", "Settings", key, key, "(prev)", String(value), {});
  return { success: true, message: "Setting updated successfully!" };
}


/**
 * Check if current academic year is finalized (frozen)
 * @returns {boolean}
 */
function isYearFinalized() {
  const v = String(getSchoolSetting("IsYearFinalized") || "").toLowerCase();
  return v === "true" || v === "yes" || v === "1";
}


/**
 * Guard: throw error if year is finalized (called by all write paths)
 */
function ensureYearNotFinalized(actionDesc) {
  if (isYearFinalized()) {
    throw new Error(`Academic year is FINALIZED. ${actionDesc || 'This change'} is not allowed. Admin must unfreeze first.`);
  }
}


/**
 * Finalize current academic year (admin only)
 * @returns {Object}
 */
function finalizeAcademicYear() {
  if (!isAdmin()) return { success: false, message: "Access denied. Admin only." };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const wasFinalized = isYearFinalized();
    updateSchoolSetting("IsYearFinalized", "true");
    writeAudit("FREEZE_YEAR", "System", "AcademicYear", "IsYearFinalized", String(wasFinalized), "true", { year: getCurrentAcademicYear() });
    logAction("Finalize Year", `Year ${getCurrentAcademicYear()} finalized by ${getActualUserEmail()}`);
    return { success: true, message: `Academic year ${getCurrentAcademicYear()} is now FINALIZED. No edits allowed until unfrozen.` };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Unfreeze a finalized year (super-admin only, requires literal confirmation text)
 * @param {string} confirmText - must equal "UNFREEZE YEAR"
 * @returns {Object}
 */
function unfreezeAcademicYear(confirmText) {
  if (!isAdmin()) return { success: false, message: "Access denied. Admin only." };
  // Super-admin = first email in ADMIN_EMAIL_LIST (script owner)
  const me = (getActualUserEmail() || "").toLowerCase();
  const superAdmin = (ADMIN_EMAIL_LIST[0] || "").toLowerCase();
  if (me !== superAdmin) {
    return { success: false, message: `Only the super-admin (${superAdmin}) can unfreeze a finalized year.` };
  }
  if (confirmText !== "UNFREEZE YEAR") {
    return { success: false, message: 'Confirmation text mismatch. Type exactly: UNFREEZE YEAR' };
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    updateSchoolSetting("IsYearFinalized", "false");
    writeAudit("UNFREEZE_YEAR", "System", "AcademicYear", "IsYearFinalized", "true", "false", { year: getCurrentAcademicYear(), unfrozenBy: me });
    logAction("Unfreeze Year", `Year ${getCurrentAcademicYear()} UNFROZEN by ${me}`);
    return { success: true, message: `Academic year ${getCurrentAcademicYear()} is now editable.` };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get aggregates last-updated info for dashboard display
 * @returns {Object} { lastUpdated, isStale, message }
 */
function getAggregatesStatus() {
  const ts = getSchoolSetting("LastAggregatesUpdatedAt");
  if (!ts) return { lastUpdated: "", isStale: true, message: "Aggregates have never been built. Run 'Rebuild Analytics'." };
  const lastDate = new Date(ts);
  const ageMs = Date.now() - lastDate.getTime();
  const ageHours = ageMs / 3600000;
  return {
    lastUpdated: ts,
    ageHours: ageHours,
    isStale: ageHours > 24,
    message: `Analytics last updated: ${lastDate.toLocaleString()} (${ageHours < 1 ? Math.round(ageHours * 60) + ' min' : Math.round(ageHours) + ' hr'} ago)`
  };
}


/**
 * Mark aggregates as just rebuilt (called from rebuildAggregates and after auto-triggers)
 */
function markAggregatesUpdated() {
  updateSchoolSetting("LastAggregatesUpdatedAt", new Date().toISOString());
}


/* ========================================================================
   PASSWORD HASHING + AUTH SHEET HELPERS
   ======================================================================== */

/**
 * Hash a password with a salt using SHA-256
 * @param {string} password
 * @param {string} salt
 * @returns {string} hex digest
 */
function hashPassword(password, salt) {
  const text = String(salt || "") + String(password || "");
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return digest.map(function (b) { return ((b < 0 ? b + 256 : b)).toString(16).padStart(2, '0'); }).join('');
}


/**
 * Generate a cryptographically random salt (UUID-based)
 */
function generateSalt() {
  return Utilities.getUuid().replace(/-/g, '');
}


/**
 * Find a row in Auth sheet for given email
 * @returns {Object|null} { rowNum, row } or null
 */
function _findAuthRow(email) {
  if (!email) return null;
  const sheet = SpreadsheetApp.getActive().getSheetByName("Auth");
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const target = String(email).trim().toLowerCase();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim().toLowerCase() === target) {
      return { rowNum: i + 2, row: data[i] };
    }
  }
  return null;
}

