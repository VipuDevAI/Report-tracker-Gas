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

// Streams configuration
const STREAMS = {
  "9": ["Science", "Computer Science", "Commerce"],
  "10": ["Science", "Computer Science", "Commerce"],
  "11": ["Science", "Computer Science", "Commerce"],
  "12": ["Science", "Computer Science", "Commerce"]
};

// Class sections - Different for 9-10 vs 11-12
const SECTIONS_9_10 = ["A", "B", "C", "D"];
const SECTIONS_11_12 = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12"];

// Legacy constant for backward compatibility
const SECTIONS = ["A", "B", "C", "D"];

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
  return SECTIONS_9_10;
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
      "RollNo", "ParentEmail", "Phone", "JoinDate", "Status", "ElectiveSubject"
    ],
    Teachers: [
      "TeacherID", "Name", "Subject", "Classes", "Sections", 
      "Email", "Phone", "JoinDate", "Status", "IsClassTeacher", "ClassTeacherOf"
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
      "Weightage", "StartDate", "EndDate", "Locked", "CreatedBy", "CreatedAt", "AcademicYear",
      "HasInternals", "Internal1", "Internal2", "Internal3", "Internal4", "TotalMaxMarks"
    ],
    Marks_Master: [
      "EntryID", "StudentID", "StudentName", "Subject", "SubjectCode",
      "TeacherID", "TeacherName", "ExamID", "ExamName", "Class", "Section",
      "MaxMarks", "MarksObtained", "Percentage", "Grade", "UpdatedAt", "UpdatedBy", "AcademicYear"
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
    
    // Class 11 & 12 - Computer Science (Mandatory: Physics, Chemistry, CS, English)
    ["SUB022", "Computer Science", "CS", "11,12", "Computer Science", 100, 40, true],
    ["SUB023", "Physics", "PHY", "11,12", "Computer Science", 100, 40, true],
    ["SUB024", "Chemistry", "CHEM", "11,12", "Computer Science", 100, 40, true],
    ["SUB025", "English", "ENG", "11,12", "Computer Science", 100, 40, true],
    
    // Class 11 & 12 - Commerce (Mandatory: Accountancy, Business Studies, Economics, English)
    ["SUB026", "Accountancy", "ACC", "11,12", "Commerce", 100, 40, true],
    ["SUB027", "Business Studies", "BS", "11,12", "Commerce", 100, 40, true],
    ["SUB028", "Economics", "ECO", "11,12", "Commerce", 100, 40, true],
    ["SUB029", "English", "ENG", "11,12", "Commerce", 100, 40, true],
    
    // Class 11 & 12 - ELECTIVE SUBJECTS (Student chooses ONE)
    // Maths/Applied Maths OR Hindi OR History OR Sanskrit
    ["SUB030", "Mathematics", "MATH", "11,12", "Elective", 100, 40, true],
    ["SUB031", "Applied Mathematics", "AMATH", "11,12", "Elective", 100, 40, true],
    ["SUB032", "Hindi", "HIN", "11,12", "Elective", 100, 40, true],
    ["SUB033", "History", "HIST", "11,12", "Elective", 100, 40, true],
    ["SUB034", "Sanskrit", "SANS", "11,12", "Elective", 100, 40, true]
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
  
  // Class 9 & 10 with sections A, B, C, D
  for (let cls = 9; cls <= 10; cls++) {
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
  
  // Class 11 & 12 with sections A1-A12
  for (let cls = 11; cls <= 12; cls++) {
    streams.forEach((stream, idx) => {
      classes.push([
        `CLS${cls}${idx + 1}`,
        `Class ${cls}`,
        "A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12",
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
 */
function syncStudentsFromClassSheets() {
  const ss = SpreadsheetApp.getActive();
  const mainSheet = ss.getSheetByName("Students");
  
  const sections = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12"];
  const classes = [11, 12];
  
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
              row[4] || ""          // ElectiveSubject
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
      mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 11).clearContent();
    }
    
    // Write all students
    mainSheet.getRange(2, 1, allStudents.length, 11).setValues(allStudents);
  }
  
  logAction("Sync Students", `Synced ${totalSynced} students from class-wise sheets`);
  
  return {
    success: true,
    message: `Synced ${totalSynced} students from class-wise sheets to main Students sheet`
  };
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
