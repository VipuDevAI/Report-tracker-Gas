/************************************************
 MVM REPORT TRACKER - MARKS ENTRY & MANAGEMENT
 File 5 of 7
 With Role-Based Access & Academic Year Support
************************************************/

/**
 * Add marks for a student
 * - LockService for concurrency
 * - Duplicate protection (update existing instead of inserting)
 * - Validates: student exists, exam exists, subject valid for student's class/stream/elective, exam not locked
 * @param {Object} marksData - Marks entry data
 * @returns {Object} Result object with action: 'created' | 'updated'
 */
function addMarks(marksData) {
  // Validate teacher/admin access
  if (!isAdmin() && !isTeacher()) {
    return { success: false, message: "Access denied. Teacher or Admin privileges required." };
  }
  
  // Validate required fields
  if (!marksData.studentId || !marksData.subject || !marksData.examId) {
    return { success: false, message: "Student, Subject, and Exam are required." };
  }
  
  if (marksData.marks === undefined || marksData.marks === null || marksData.marks === "") {
    return { success: false, message: "Marks value is required." };
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    // Check if exam is locked (Marks Edit Protection)
    const exam = getExamById(marksData.examId);
    if (!exam) {
      return { success: false, message: "Exam not found." };
    }
    if (exam.locked === true) {
      return { success: false, message: "Exam is locked. No edits allowed." };
    }
    
    // Validate marks range
    const marksNum = parseFloat(marksData.marks);
    if (isNaN(marksNum) || marksNum < 0 || marksNum > exam.maxMarks) {
      return { success: false, message: `Marks must be between 0 and ${exam.maxMarks}.` };
    }
    
    // Get student details (this already applies teacher filter)
    const students = getStudents({ status: "Active" });
    const student = students.find(s => s.studentId === marksData.studentId);
    
    if (!student) {
      return { success: false, message: "Student not found or you don't have access to this student." };
    }
    
    // Fail-safe Validation: subject must be valid for this student
    if (!isSubjectValidForStudent(marksData.subject, student)) {
      return { 
        success: false, 
        message: `Subject "${marksData.subject}" is not valid for ${student.name} (Class ${student.class} ${student.stream}${student.electiveSubject ? ', Elective: ' + student.electiveSubject : ''}).` 
      };
    }
    
    // Teacher-specific permission validation
    if (!isAdmin()) {
      const assignment = getTeacherAssignment();
      if (!assignment) {
        return { success: false, message: "Teacher assignment not found." };
      }
      if (assignment.subject !== "All" && assignment.subject !== marksData.subject) {
        return { success: false, message: `You can only enter marks for ${assignment.subject}.` };
      }
      if (!assignment.hasAllClasses && !assignment.classes.includes(String(student.class))) {
        return { success: false, message: "You don't have permission for this class." };
      }
      if (!assignment.hasAllSections && !assignment.sections.includes(student.section)) {
        return { success: false, message: "You don't have permission for this section." };
      }
    }
    
    // Get teacher details
    const teacher = getTeacherByEmail(getActualUserEmail());
    const teacherId = teacher ? teacher.teacherId : "ADMIN";
    const teacherName = teacher ? teacher.name : "Administrator";
    
    const academicYear = getCurrentAcademicYear();
    
    // Single bulk read of Marks_Master (no getDataRange in loop)
    const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
    const lastRow = sheet.getLastRow();
    const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 18).getValues() : [];
    
    // Find existing entry by composite key (studentId|examId|subject)
    let existingRowNum = -1;
    let existingEntryId = null;
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === marksData.studentId &&
          data[i][7] === marksData.examId &&
          String(data[i][3]).trim().toLowerCase() === String(marksData.subject).trim().toLowerCase()) {
        existingRowNum = i + 2; // +2 because data starts at row 2
        existingEntryId = data[i][0];
        break;
      }
    }
    
    const percentage = (marksNum / exam.maxMarks) * 100;
    const grade = calculateGrade(percentage);
    
    const entryData = [
      existingEntryId || `MRK${Date.now()}`,
      marksData.studentId,
      student.name,
      marksData.subject,
      marksData.subjectCode || "",
      teacherId,
      teacherName,
      marksData.examId,
      exam.name,
      student.class,
      student.section,
      exam.maxMarks,
      marksNum,
      percentage.toFixed(2),
      grade,
      new Date(),
      getActualUserEmail() || "System",
      academicYear
    ];
    
    if (existingRowNum > 0) {
      sheet.getRange(existingRowNum, 1, 1, 18).setValues([entryData]);
      logAction("Update Marks", `Updated marks for ${student.name} in ${marksData.subject}`);
      return { 
        success: true, 
        action: "updated",
        message: "Marks already exist — updating existing record." 
      };
    } else {
      sheet.appendRow(entryData);
      logAction("Add Marks", `Added marks for ${student.name} in ${marksData.subject}`);
      
      if (percentage < 40) {
        createWeakStudentAlert(student, marksData.subject, percentage, exam.name);
      }
      
      return { 
        success: true, 
        action: "created",
        message: "Marks added successfully!" 
      };
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Bulk add marks - optimized: single LockService, cached reads, batch writes
 * @param {Array} marksArray - Array of marks data objects
 * @returns {Object} Result object with success/fail counts
 */
function bulkAddMarks(marksArray) {
  if (!isAdmin() && !isTeacher()) {
    return { success: false, message: "Access denied." };
  }
  
  if (!Array.isArray(marksArray) || marksArray.length === 0) {
    return { success: false, message: "No marks provided." };
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const ss = SpreadsheetApp.getActive();
    const academicYear = getCurrentAcademicYear();
    
    // Cache: Students (one read)
    const studentsList = getStudents({ status: "Active" });
    const studentsMap = {};
    studentsList.forEach(s => { studentsMap[s.studentId] = s; });
    
    // Cache: Exams (one read)
    const examsSheet = ss.getSheetByName("Exams");
    const examsLastRow = examsSheet.getLastRow();
    const examsData = examsLastRow > 1 ? examsSheet.getRange(2, 1, examsLastRow - 1, 18).getValues() : [];
    const examsMap = {};
    examsData.forEach(row => {
      if (row[0]) examsMap[row[0]] = { name: row[1], maxMarks: row[4], locked: row[8] === true };
    });
    
    // Teacher info (one fetch)
    const teacher = getTeacherByEmail(getActualUserEmail());
    const teacherId = teacher ? teacher.teacherId : "ADMIN";
    const teacherName = teacher ? teacher.name : "Administrator";
    
    // Teacher assignment for permission filter
    const isUserAdmin = isAdmin();
    const assignment = isUserAdmin ? null : getTeacherAssignment();
    
    // Cache: Marks_Master (one read) + build composite key map
    const marksSheet = ss.getSheetByName("Marks_Master");
    const marksLastRow = marksSheet.getLastRow();
    const marksData = marksLastRow > 1 ? marksSheet.getRange(2, 1, marksLastRow - 1, 18).getValues() : [];
    const marksMap = {}; // key -> { rowNum, entryId }
    marksData.forEach((row, idx) => {
      const key = `${row[1]}|${row[7]}|${String(row[3]).trim().toLowerCase()}`;
      marksMap[key] = { rowNum: idx + 2, entryId: row[0] };
    });
    
    let createdCount = 0;
    let updatedCount = 0;
    let failCount = 0;
    const errors = [];
    const newRows = [];
    const updateRows = []; // { rowNum, data }
    
    let idCounter = 0;
    const baseTime = Date.now();
    
    marksArray.forEach((m, idx) => {
      // Required fields
      if (!m.studentId || !m.subject || !m.examId || m.marks === undefined || m.marks === null || m.marks === "") {
        failCount++;
        errors.push({ row: idx + 1, error: "Missing required field (studentId, subject, examId, marks)" });
        return;
      }
      
      const exam = examsMap[m.examId];
      if (!exam) {
        failCount++;
        errors.push({ row: idx + 1, error: `Exam ${m.examId} not found` });
        return;
      }
      if (exam.locked) {
        failCount++;
        errors.push({ row: idx + 1, error: `Exam "${exam.name}" is locked — no edits allowed` });
        return;
      }
      
      const student = studentsMap[m.studentId];
      if (!student) {
        failCount++;
        errors.push({ row: idx + 1, error: `Student ${m.studentId} not found or not active` });
        return;
      }
      
      // Subject validity check
      if (!isSubjectValidForStudent(m.subject, student)) {
        failCount++;
        errors.push({ row: idx + 1, error: `Subject "${m.subject}" not valid for ${student.name} (Class ${student.class} ${student.stream})` });
        return;
      }
      
      // Teacher permission check
      if (!isUserAdmin) {
        if (!assignment) {
          failCount++;
          errors.push({ row: idx + 1, error: "Teacher assignment not found" });
          return;
        }
        if (assignment.subject !== "All" && assignment.subject !== m.subject) {
          failCount++;
          errors.push({ row: idx + 1, error: `You can only enter marks for ${assignment.subject}` });
          return;
        }
        if (!assignment.hasAllClasses && !assignment.classes.includes(String(student.class))) {
          failCount++;
          errors.push({ row: idx + 1, error: "No permission for this class" });
          return;
        }
        if (!assignment.hasAllSections && !assignment.sections.includes(student.section)) {
          failCount++;
          errors.push({ row: idx + 1, error: "No permission for this section" });
          return;
        }
      }
      
      const marksNum = parseFloat(m.marks);
      if (isNaN(marksNum) || marksNum < 0 || marksNum > exam.maxMarks) {
        failCount++;
        errors.push({ row: idx + 1, error: `Marks must be 0-${exam.maxMarks}` });
        return;
      }
      
      const percentage = (marksNum / exam.maxMarks) * 100;
      const grade = calculateGrade(percentage);
      const key = `${m.studentId}|${m.examId}|${String(m.subject).trim().toLowerCase()}`;
      const existing = marksMap[key];
      
      const entryId = existing ? existing.entryId : `MRK${baseTime}${(idCounter++).toString(36)}`;
      const row = [
        entryId,
        m.studentId,
        student.name,
        m.subject,
        m.subjectCode || "",
        teacherId,
        teacherName,
        m.examId,
        exam.name,
        student.class,
        student.section,
        exam.maxMarks,
        marksNum,
        percentage.toFixed(2),
        grade,
        new Date(),
        getActualUserEmail() || "System",
        academicYear
      ];
      
      if (existing) {
        updateRows.push({ rowNum: existing.rowNum, data: row });
        updatedCount++;
      } else {
        newRows.push(row);
        // Track newly added so subsequent rows in this batch with same key UPDATE instead of duplicate
        marksMap[key] = { rowNum: -1, entryId: entryId, pendingNewIndex: newRows.length - 1 };
        createdCount++;
      }
    });
    
    // Batch write new rows
    if (newRows.length > 0) {
      const writeStart = marksSheet.getLastRow() + 1;
      marksSheet.getRange(writeStart, 1, newRows.length, 18).setValues(newRows);
    }
    
    // Apply updates row-by-row (cannot batch non-contiguous rows reliably)
    updateRows.forEach(u => {
      marksSheet.getRange(u.rowNum, 1, 1, 18).setValues([u.data]);
    });
    
    logAction("Bulk Add Marks", `Created: ${createdCount}, Updated: ${updatedCount}, Failed: ${failCount}`);
    
    return {
      success: failCount === 0,
      createdCount: createdCount,
      updatedCount: updatedCount,
      successCount: createdCount + updatedCount,
      failCount: failCount,
      errors: errors,
      message: `${createdCount} created, ${updatedCount} updated, ${failCount} failed.`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get marks with filters
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered marks
 */
function getMarks(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];
  
  // Single bulk read (no getDataRange in loops)
  const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  
  // Get current academic year for default filtering
  const currentYear = getCurrentAcademicYear();
  
  let marks = data.map(row => ({
    entryId: row[0],
    studentId: row[1],
    studentName: row[2],
    subject: row[3],
    subjectCode: row[4],
    teacherId: row[5],
    teacherName: row[6],
    examId: row[7],
    examName: row[8],
    class: row[9],
    section: row[10],
    maxMarks: row[11],
    marksObtained: row[12],
    percentage: parseFloat(row[13]),
    grade: row[14],
    updatedAt: row[15],
    updatedBy: row[16],
    academicYear: row[17] || currentYear
  })).filter(m => m.entryId);
  
  // Apply standard filters
  if (filters) {
    if (filters.studentId) {
      marks = marks.filter(m => m.studentId === filters.studentId);
    }
    if (filters.examId) {
      marks = marks.filter(m => m.examId === filters.examId);
    }
    if (filters.subject) {
      marks = marks.filter(m => m.subject === filters.subject);
    }
    if (filters.class) {
      marks = marks.filter(m => m.class == filters.class);
    }
    if (filters.section) {
      marks = marks.filter(m => m.section === filters.section);
    }
    if (filters.teacherId) {
      marks = marks.filter(m => m.teacherId === filters.teacherId);
    }
    if (filters.academicYear) {
      marks = marks.filter(m => m.academicYear === filters.academicYear);
    } else {
      // Default: filter by current academic year
      marks = marks.filter(m => m.academicYear === currentYear);
    }
  } else {
    // Default: filter by current academic year
    marks = marks.filter(m => m.academicYear === currentYear);
  }
  
  // Apply teacher assignment filter (server-side)
  marks = applyTeacherFilter(marks, { filterBySubject: true, subjectField: "subject" });
  
  return marks;
}


/**
 * Get student's marks summary
 * @param {string} studentId - Student ID
 * @returns {Object} Student marks summary
 */
function getStudentMarksSummary(studentId) {
  const marks = getMarks({ studentId: studentId });
  
  if (marks.length === 0) {
    return { studentId: studentId, totalMarks: 0, exams: [] };
  }
  
  const summary = {
    studentId: studentId,
    studentName: marks[0].studentName,
    class: marks[0].class,
    section: marks[0].section,
    totalMarksObtained: 0,
    totalMaxMarks: 0,
    overallPercentage: 0,
    overallGrade: "",
    subjectWise: {},
    examWise: {}
  };
  
  marks.forEach(m => {
    summary.totalMarksObtained += m.marksObtained;
    summary.totalMaxMarks += m.maxMarks;
    
    // Subject-wise aggregation
    if (!summary.subjectWise[m.subject]) {
      summary.subjectWise[m.subject] = { total: 0, max: 0, exams: [] };
    }
    summary.subjectWise[m.subject].total += m.marksObtained;
    summary.subjectWise[m.subject].max += m.maxMarks;
    summary.subjectWise[m.subject].exams.push({
      examName: m.examName,
      marks: m.marksObtained,
      max: m.maxMarks,
      percentage: m.percentage
    });
    
    // Exam-wise aggregation
    if (!summary.examWise[m.examId]) {
      summary.examWise[m.examId] = { examName: m.examName, total: 0, max: 0, subjects: [] };
    }
    summary.examWise[m.examId].total += m.marksObtained;
    summary.examWise[m.examId].max += m.maxMarks;
    summary.examWise[m.examId].subjects.push({
      subject: m.subject,
      marks: m.marksObtained,
      max: m.maxMarks
    });
  });
  
  summary.overallPercentage = summary.totalMaxMarks > 0 
    ? ((summary.totalMarksObtained / summary.totalMaxMarks) * 100).toFixed(2)
    : 0;
  summary.overallGrade = calculateGrade(parseFloat(summary.overallPercentage));
  
  return summary;
}


/**
 * Calculate grade range based on percentage (numeric ranges only)
 * @param {number} percentage - Percentage value
 * @returns {string} Range (e.g., "91-100", "81-90")
 */
function calculateGrade(percentage) {
  if (percentage >= 91) return "91-100";
  if (percentage >= 81) return "81-90";
  if (percentage >= 71) return "71-80";
  if (percentage >= 61) return "61-70";
  if (percentage >= 51) return "51-60";
  if (percentage >= 41) return "41-50";
  return "0-40";
}


/**
 * Get grade color based on range
 * @param {string} grade - Grade range string
 * @returns {string} Color hex code
 */
function getGradeColor(grade) {
  const colors = {
    "91-100": "#22c55e",
    "81-90": "#16a34a",
    "71-80": "#3b82f6",
    "61-70": "#0ea5e9",
    "51-60": "#f59e0b",
    "41-50": "#f97316",
    "0-40": "#ef4444"
  };
  return colors[grade] || "#6b7280";
}


/**
 * Delete marks entry (Admin only)
 * @param {string} entryId - Entry ID to delete
 * @returns {Object} Result object
 */
function deleteMarks(entryId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Entry not found." };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    
    let foundIdx = -1;
    let foundExamId = null;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === entryId) {
        foundIdx = i;
        foundExamId = data[i][7];
        break;
      }
    }
    
    if (foundIdx === -1) {
      return { success: false, message: "Entry not found." };
    }
    
    // Check if exam is locked (Marks Edit Protection)
    if (isExamLocked(foundExamId)) {
      return { success: false, message: "Exam is locked. No edits allowed." };
    }
    
    sheet.deleteRow(foundIdx + 2);
    
    logAction("Delete Marks", `Deleted marks entry: ${entryId}`);
    
    return { success: true, message: "Marks entry deleted successfully!" };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Create weak student alert
 * @param {Object} student - Student object
 * @param {string} subject - Subject name
 * @param {number} percentage - Percentage scored
 * @param {string} examName - Exam name
 */
function createWeakStudentAlert(student, subject, percentage, examName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Alerts");
  const alertId = `ALT${Date.now()}`;
  
  sheet.appendRow([
    alertId,
    "WEAK_STUDENT",
    student.studentId,
    student.name,
    student.class,
    subject,
    `${student.name} scored ${percentage.toFixed(1)}% in ${subject} (${examName})`,
    percentage < 25 ? "High" : "Medium",
    false,
    new Date()
  ]);
}


/**
 * Get recent marks entries
 * @param {number} limit - Number of entries to return
 * @returns {Array} Recent marks entries
 */
function getRecentMarks(limit) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const lastRow = sheet.getLastRow();
  const currentYear = getCurrentAcademicYear();
  
  if (lastRow <= 1) return [];
  
  const numRows = Math.min(limit || 20, lastRow - 1);
  const startRow = Math.max(2, lastRow - numRows + 1);
  
  const data = sheet.getRange(startRow, 1, numRows, 18).getValues();
  
  let results = data.map(row => ({
    entryId: row[0],
    studentId: row[1],
    studentName: row[2],
    subject: row[3],
    examName: row[8],
    class: row[9],
    section: row[10],
    maxMarks: row[11],
    marksObtained: row[12],
    percentage: parseFloat(row[13]),
    grade: row[14],
    updatedAt: row[15],
    teacherName: row[6],
    academicYear: row[17] || currentYear
  })).filter(m => m.academicYear === currentYear);
  
  // Apply teacher filter
  results = applyTeacherFilter(results, { filterBySubject: true, subjectField: "subject" });
  
  return results.reverse();
}


/**
 * Admin bulk upload marks from CSV
 * @param {Array} data - 2D array of marks data
 * @param {Object} columnMapping - Column index mapping { studentId, subject, examId, marks }
 * @param {Object} options - { preview: boolean }
 * @returns {Object} Result object
 */
function adminBulkUploadMarks(data, columnMapping, options) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!data || !Array.isArray(data) || data.length === 0) {
    return { success: false, message: "No data provided." };
  }
  
  const opts = options || {};
  const previewOnly = opts.preview || false;
  
  const lock = LockService.getScriptLock();
  try {
    if (!previewOnly) lock.waitLock(30000);
    
    // Default column mapping
    const mapping = columnMapping || {
      studentId: 0,
      subject: 1,
      examId: 2,
      marks: 3
    };
    
    const ss = SpreadsheetApp.getActive();
    
    // Cache: Students (one read)
    const studentsSheet = ss.getSheetByName("Students");
    const studentsLastRow = studentsSheet.getLastRow();
    const studentsData = studentsLastRow > 1 ? studentsSheet.getRange(2, 1, studentsLastRow - 1, 12).getValues() : [];
    const studentIndex = {};
    studentsData.forEach(row => {
      if (row[0]) {
        studentIndex[row[0]] = {
          name: row[1], class: row[2], section: row[3],
          stream: row[4], electiveSubject: row[10] || ''
        };
      }
    });
    
    // Cache: Exams (one read)
    const examsSheet = ss.getSheetByName("Exams");
    const examsLastRow = examsSheet.getLastRow();
    const examsData = examsLastRow > 1 ? examsSheet.getRange(2, 1, examsLastRow - 1, 18).getValues() : [];
    const examIndex = {};
    examsData.forEach(row => {
      if (row[0]) examIndex[row[0]] = { name: row[1], maxMarks: row[4], locked: row[8] === true };
    });
    
    // Cache: Marks_Master existing entries (for duplicate protection)
    const marksSheet = ss.getSheetByName("Marks_Master");
    const marksLastRow = marksSheet.getLastRow();
    const existingMarks = marksLastRow > 1 ? marksSheet.getRange(2, 1, marksLastRow - 1, 18).getValues() : [];
    const marksMap = {};
    existingMarks.forEach((row, idx) => {
      const key = `${row[1]}|${row[7]}|${String(row[3]).trim().toLowerCase()}`;
      marksMap[key] = { rowNum: idx + 2, entryId: row[0] };
    });
    
    const results = {
      preview: [],
      created: 0,
      updated: 0,
      failed: 0,
      errors: [],
      lockedExams: []
    };
    
    const newRows = [];
    const updateRows = []; // { rowNum, data }
    const academicYear = getCurrentAcademicYear();
    let idCounter = 0;
    const baseTime = Date.now();
    
    data.forEach((row, rowIdx) => {
      // Skip header row
      if (rowIdx === 0) {
        const firstCell = String(row[0] || "").toLowerCase();
        if (firstCell.includes("student") || firstCell.includes("id") || firstCell === "name") {
          return;
        }
      }
      
      const studentId = String(row[mapping.studentId] || "").trim();
      const subject = String(row[mapping.subject] || "").trim();
      const examId = String(row[mapping.examId] || "").trim();
      const marks = parseFloat(row[mapping.marks]);
      
      if (!studentId) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: "Student ID is required" });
        return;
      }
      if (!subject) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: "Subject is required" });
        return;
      }
      if (!examId) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: "Exam ID is required" });
        return;
      }
      if (isNaN(marks)) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: "Invalid marks value" });
        return;
      }
      
      const student = studentIndex[studentId];
      if (!student) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: `Student ${studentId} not found` });
        return;
      }
      
      const exam = examIndex[examId];
      if (!exam) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: `Exam ${examId} not found` });
        return;
      }
      
      if (exam.locked) {
        results.failed++;
        results.lockedExams.push({ row: rowIdx + 1, examId: examId, examName: exam.name });
        results.errors.push({ row: rowIdx + 1, error: `Exam "${exam.name}" is locked — no edits allowed` });
        return;
      }
      
      // Subject validity check
      if (!isSubjectValidForStudent(subject, student)) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: `Subject "${subject}" not valid for this student (Class ${student.class} ${student.stream})` });
        return;
      }
      
      if (marks < 0 || marks > exam.maxMarks) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: `Marks must be 0-${exam.maxMarks}` });
        return;
      }
      
      const percentage = (marks / exam.maxMarks) * 100;
      const grade = calculateGrade(percentage);
      const key = `${studentId}|${examId}|${subject.toLowerCase()}`;
      const existing = marksMap[key];
      const entryId = existing ? existing.entryId : `MRK${baseTime}${(idCounter++).toString(36)}`;
      
      const rowData = [
        entryId,
        studentId,
        student.name,
        subject,
        "",
        "ADMIN",
        "Administrator",
        examId,
        exam.name,
        student.class,
        student.section,
        exam.maxMarks,
        marks,
        percentage.toFixed(2),
        grade,
        new Date(),
        getActualUserEmail() || "ADMIN",
        academicYear
      ];
      
      if (existing) {
        if (!previewOnly) updateRows.push({ rowNum: existing.rowNum, data: rowData });
        results.updated++;
        results.preview.push({
          studentId: studentId, studentName: student.name, subject: subject,
          examName: exam.name, marks: marks, maxMarks: exam.maxMarks,
          percentage: percentage.toFixed(1) + "%", grade: grade, status: "UPDATE"
        });
      } else {
        if (!previewOnly) newRows.push(rowData);
        marksMap[key] = { rowNum: -1, entryId: entryId };
        results.created++;
        results.preview.push({
          studentId: studentId, studentName: student.name, subject: subject,
          examName: exam.name, marks: marks, maxMarks: exam.maxMarks,
          percentage: percentage.toFixed(1) + "%", grade: grade, status: "NEW"
        });
      }
    });
    
    // Preview: return without writing
    if (previewOnly) {
      return {
        success: true,
        preview: true,
        results: results,
        message: `Preview: ${results.created} new, ${results.updated} updates, ${results.failed} failed`
      };
    }
    
    // Batch write new rows
    if (newRows.length > 0) {
      const writeStart = marksSheet.getLastRow() + 1;
      marksSheet.getRange(writeStart, 1, newRows.length, 18).setValues(newRows);
    }
    
    // Apply updates
    updateRows.forEach(u => {
      marksSheet.getRange(u.rowNum, 1, 1, 18).setValues([u.data]);
    });
    
    logAction("Admin Bulk Upload Marks", `Created: ${results.created}, Updated: ${results.updated}, Failed: ${results.failed}`);
    
    return {
      success: true,
      preview: false,
      results: results,
      message: `Import complete: ${results.created} created, ${results.updated} updated, ${results.failed} failed`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get column headers from first row of data for mapping UI
 * @param {Array} firstRow - First row of CSV data
 * @returns {Array} Column options for mapping
 */
function getColumnMappingOptions(firstRow) {
  return firstRow.map((header, idx) => ({
    index: idx,
    header: header,
    suggested: suggestMapping(header)
  }));
}


/**
 * Suggest column mapping based on header name
 * @param {string} header - Column header
 * @returns {string} Suggested field name
 */
function suggestMapping(header) {
  const h = String(header).toLowerCase();
  if (h.includes("student") && h.includes("id")) return "studentId";
  if (h.includes("subject")) return "subject";
  if (h.includes("exam") && h.includes("id")) return "examId";
  if (h.includes("mark") || h.includes("score")) return "marks";
  if (h.includes("name")) return "studentName";
  return "";
}
