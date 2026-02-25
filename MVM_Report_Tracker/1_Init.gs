/************************************************
 MVM REPORT TRACKER - INITIALIZATION & SETUP
 File 1 of 7
************************************************/

// Admin email whitelist
const ADMIN_EMAILS = [
  "rishisans83@gmail.com",
  "mvmseniors26@gmail.com"
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

// Streams configuration
const STREAMS = {
  "9": ["Science", "Computer Science", "Commerce"],
  "10": ["Science", "Computer Science", "Commerce"],
  "11": ["Science", "Computer Science", "Commerce"],
  "12": ["Science", "Computer Science", "Commerce"]
};

// Class sections
const SECTIONS = ["A", "B", "C", "D"];

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
      "RollNo", "ParentEmail", "Phone", "JoinDate", "Status"
    ],
    Teachers: [
      "TeacherID", "Name", "Subject", "Classes", "Sections", 
      "Email", "Phone", "JoinDate", "Status"
    ],
    Subjects: [
      "SubjectID", "SubjectName", "SubjectCode", "Class", "Stream", 
      "MaxMarks", "PassingMarks", "IsActive"
    ],
    Classes: [
      "ClassID", "ClassName", "Sections", "Stream", "AcademicYear", "IsActive"
    ],
    Exams: [
      "ExamID", "ExamName", "ExamType", "Class", "MaxMarks", 
      "Weightage", "StartDate", "EndDate", "Locked", "CreatedBy", "CreatedAt"
    ],
    Marks_Master: [
      "EntryID", "StudentID", "StudentName", "Subject", "SubjectCode",
      "TeacherID", "TeacherName", "ExamID", "ExamName", "Class", "Section",
      "MaxMarks", "MarksObtained", "Percentage", "Grade", "UpdatedAt", "UpdatedBy"
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
  
  logAction("Initialize App", "Application initialized successfully");
  
  SpreadsheetApp.flush();
  
  return { success: true, message: "MVM Report Tracker initialized successfully!" };
}


/**
 * Seed default grade ranges
 */
function seedDefaultRanges() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_Ranges");
  
  const ranges = [
    ["A+", "Excellent", 91, 100, "#22c55e"],
    ["A", "Very Good", 81, 90, "#16a34a"],
    ["B+", "Good", 71, 80, "#3b82f6"],
    ["B", "Above Average", 61, 70, "#0ea5e9"],
    ["C", "Average", 51, 60, "#f59e0b"],
    ["D", "Below Average", 41, 50, "#f97316"],
    ["F", "Fail", 0, 40, "#ef4444"]
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
    ["Email", "", new Date()]
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, settings.length, 3).setValues(settings);
  }
}


/**
 * Seed default subjects for all streams
 */
function seedDefaultSubjects() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Subjects");
  
  const subjects = [
    // Class 9 & 10 - Science
    ["SUB001", "Mathematics", "MATH", "9,10", "Science", 100, 40, true],
    ["SUB002", "Physics", "PHY", "9,10", "Science", 100, 40, true],
    ["SUB003", "Chemistry", "CHEM", "9,10", "Science", 100, 40, true],
    ["SUB004", "Biology", "BIO", "9,10", "Science", 100, 40, true],
    ["SUB005", "English", "ENG", "9,10", "Science", 100, 40, true],
    ["SUB006", "Hindi", "HIN", "9,10", "Science", 100, 40, true],
    
    // Class 9 & 10 - Computer Science
    ["SUB007", "Mathematics", "MATH", "9,10", "Computer Science", 100, 40, true],
    ["SUB008", "Computer Science", "CS", "9,10", "Computer Science", 100, 40, true],
    ["SUB009", "Physics", "PHY", "9,10", "Computer Science", 100, 40, true],
    ["SUB010", "English", "ENG", "9,10", "Computer Science", 100, 40, true],
    ["SUB011", "Hindi", "HIN", "9,10", "Computer Science", 100, 40, true],
    
    // Class 9 & 10 - Commerce
    ["SUB012", "Mathematics", "MATH", "9,10", "Commerce", 100, 40, true],
    ["SUB013", "Business Studies", "BS", "9,10", "Commerce", 100, 40, true],
    ["SUB014", "Economics", "ECO", "9,10", "Commerce", 100, 40, true],
    ["SUB015", "English", "ENG", "9,10", "Commerce", 100, 40, true],
    ["SUB016", "Hindi", "HIN", "9,10", "Commerce", 100, 40, true],
    
    // Class 11 & 12 - Science
    ["SUB017", "Mathematics", "MATH", "11,12", "Science", 100, 40, true],
    ["SUB018", "Physics", "PHY", "11,12", "Science", 100, 40, true],
    ["SUB019", "Chemistry", "CHEM", "11,12", "Science", 100, 40, true],
    ["SUB020", "Biology", "BIO", "11,12", "Science", 100, 40, true],
    ["SUB021", "English", "ENG", "11,12", "Science", 100, 40, true],
    
    // Class 11 & 12 - Computer Science
    ["SUB022", "Mathematics", "MATH", "11,12", "Computer Science", 100, 40, true],
    ["SUB023", "Computer Science", "CS", "11,12", "Computer Science", 100, 40, true],
    ["SUB024", "Physics", "PHY", "11,12", "Computer Science", 100, 40, true],
    ["SUB025", "English", "ENG", "11,12", "Computer Science", 100, 40, true],
    ["SUB026", "Informatics Practices", "IP", "11,12", "Computer Science", 100, 40, true],
    
    // Class 11 & 12 - Commerce
    ["SUB027", "Accountancy", "ACC", "11,12", "Commerce", 100, 40, true],
    ["SUB028", "Business Studies", "BS", "11,12", "Commerce", 100, 40, true],
    ["SUB029", "Economics", "ECO", "11,12", "Commerce", 100, 40, true],
    ["SUB030", "English", "ENG", "11,12", "Commerce", 100, 40, true],
    ["SUB031", "Mathematics", "MATH", "11,12", "Commerce", 100, 40, true]
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, subjects.length, 8).setValues(subjects);
  }
}


/**
 * Seed default classes
 */
function seedDefaultClasses() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Classes");
  
  const classes = [];
  const streams = ["Science", "Computer Science", "Commerce"];
  
  for (let cls = 9; cls <= 12; cls++) {
    streams.forEach((stream, idx) => {
      classes.push([
        `CLS${cls}${idx + 1}`,
        `Class ${cls}`,
        "A,B,C,D",
        stream,
        "2024-2025",
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
      .addItem('Reset School Data', 'resetSchool')
      .addItem('Archive & Reset Year', 'archiveAndReset'))
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
