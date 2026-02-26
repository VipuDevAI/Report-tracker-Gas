/************************************************
 MVM REPORT TRACKER - MARKS ENTRY & MANAGEMENT
 File 5 of 7
 With Role-Based Access & Academic Year Support
************************************************/

/**
 * Add marks for a student
 * @param {Object} marksData - Marks entry data
 * @returns {Object} Result object
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
  
  if (marksData.marks === undefined || marksData.marks === null) {
    return { success: false, message: "Marks value is required." };
  }
  
  // Check if exam is locked
  if (isExamLocked(marksData.examId)) {
    return { success: false, message: "Exam is locked. Cannot add/modify marks." };
  }
  
  // Get exam details
  const exam = getExamById(marksData.examId);
  if (!exam) {
    return { success: false, message: "Exam not found." };
  }
  
  // Validate marks range
  if (marksData.marks < 0 || marksData.marks > exam.maxMarks) {
    return { success: false, message: `Marks must be between 0 and ${exam.maxMarks}.` };
  }
  
  // Get student details (this already applies teacher filter)
  const students = getStudents({ status: "Active" });
  const student = students.find(s => s.studentId === marksData.studentId);
  
  if (!student) {
    return { success: false, message: "Student not found or you don't have access to this student." };
  }
  
  // Teacher-specific validation: Check if teacher can enter marks for this subject
  if (!isAdmin()) {
    const assignment = getTeacherAssignment();
    if (!assignment) {
      return { success: false, message: "Teacher assignment not found." };
    }
    
    // Check subject permission
    if (assignment.subject !== "All" && assignment.subject !== marksData.subject) {
      return { success: false, message: `You can only enter marks for ${assignment.subject}.` };
    }
    
    // Check class permission
    if (!assignment.hasAllClasses && !assignment.classes.includes(String(student.class))) {
      return { success: false, message: "You don't have permission for this class." };
    }
    
    // Check section permission
    if (!assignment.hasAllSections && !assignment.sections.includes(student.section)) {
      return { success: false, message: "You don't have permission for this section." };
    }
  }
  
  // Get teacher details
  const teacher = getTeacherByEmail(getCurrentUser());
  const teacherId = teacher ? teacher.teacherId : "ADMIN";
  const teacherName = teacher ? teacher.name : "Administrator";
  
  // Get current academic year
  const academicYear = getCurrentAcademicYear();
  
  // Check for existing entry
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const data = sheet.getDataRange().getValues();
  
  const existingIndex = data.findIndex(r => 
    r[1] === marksData.studentId && 
    r[3] === marksData.subject && 
    r[7] === marksData.examId
  );
  
  // Calculate percentage and grade
  const percentage = (marksData.marks / exam.maxMarks) * 100;
  const grade = calculateGrade(percentage);
  
  const entryData = [
    existingIndex > 0 ? data[existingIndex][0] : `MRK${Date.now()}`,
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
    marksData.marks,
    percentage.toFixed(2),
    grade,
    new Date(),
    getCurrentUser(),
    academicYear  // Academic year field
  ];
  
  if (existingIndex > 0) {
    // Update existing entry
    sheet.getRange(existingIndex + 1, 1, 1, 18).setValues([entryData]);
    logAction("Update Marks", `Updated marks for ${student.name} in ${marksData.subject}`);
    return { success: true, message: "Marks updated successfully!" };
  } else {
    // Add new entry
    sheet.appendRow(entryData);
    logAction("Add Marks", `Added marks for ${student.name} in ${marksData.subject}`);
    
    // Check if weak student alert needed
    if (percentage < 40) {
      createWeakStudentAlert(student, marksData.subject, percentage, exam.name);
    }
    
    return { success: true, message: "Marks added successfully!" };
  }
}


/**
 * Bulk add marks for multiple students
 * @param {Array} marksArray - Array of marks data objects
 * @returns {Object} Result object with success/fail counts
 */
function bulkAddMarks(marksArray) {
  if (!isAdmin() && !isTeacher()) {
    return { success: false, message: "Access denied." };
  }
  
  let successCount = 0;
  let failCount = 0;
  const errors = [];
  
  marksArray.forEach((marks, index) => {
    const result = addMarks(marks);
    if (result.success) {
      successCount++;
    } else {
      failCount++;
      errors.push({ row: index + 1, error: result.message });
    }
  });
  
  return {
    success: failCount === 0,
    message: `${successCount} entries added, ${failCount} failed.`,
    successCount: successCount,
    failCount: failCount,
    errors: errors
  };
}


/**
 * Get marks with filters
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered marks
 */
function getMarks(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  // Get current academic year for default filtering
  const currentYear = getCurrentAcademicYear();
  
  let marks = data.slice(1).map(row => ({
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
  }));
  
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
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const data = sheet.getDataRange().getValues();
  
  const rowIndex = data.findIndex(r => r[0] === entryId);
  
  if (rowIndex === -1) {
    return { success: false, message: "Entry not found." };
  }
  
  // Check if exam is locked
  const examId = data[rowIndex][7];
  if (isExamLocked(examId)) {
    return { success: false, message: "Exam is locked. Cannot delete marks." };
  }
  
  sheet.deleteRow(rowIndex + 1);
  
  logAction("Delete Marks", `Deleted marks entry: ${entryId}`);
  
  return { success: true, message: "Marks entry deleted successfully!" };
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
  
  // Default column mapping
  const mapping = columnMapping || {
    studentId: 0,
    subject: 1,
    examId: 2,
    marks: 3
  };
  
  // Get existing data for validation
  const studentsSheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const studentsData = studentsSheet.getDataRange().getValues();
  const studentIndex = {};
  studentsData.slice(1).forEach(row => {
    studentIndex[row[0]] = { name: row[1], class: row[2], section: row[3] };
  });
  
  // Get exams for validation
  const examsSheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const examsData = examsSheet.getDataRange().getValues();
  const examIndex = {};
  examsData.slice(1).forEach(row => {
    examIndex[row[0]] = { name: row[1], maxMarks: row[4], locked: row[8] };
  });
  
  const results = {
    preview: [],
    created: 0,
    updated: 0,
    failed: 0,
    errors: [],
    lockedExams: []
  };
  
  const validMarks = [];
  
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
    
    // Validation
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
    
    // Check student exists
    const student = studentIndex[studentId];
    if (!student) {
      results.failed++;
      results.errors.push({ row: rowIdx + 1, error: `Student ${studentId} not found` });
      return;
    }
    
    // Check exam exists
    const exam = examIndex[examId];
    if (!exam) {
      results.failed++;
      results.errors.push({ row: rowIdx + 1, error: `Exam ${examId} not found` });
      return;
    }
    
    // Check exam is not locked
    if (exam.locked) {
      results.failed++;
      results.lockedExams.push({ row: rowIdx + 1, examId: examId, examName: exam.name });
      results.errors.push({ row: rowIdx + 1, error: `Exam ${exam.name} is locked` });
      return;
    }
    
    // Validate marks range
    if (marks < 0 || marks > exam.maxMarks) {
      results.failed++;
      results.errors.push({ row: rowIdx + 1, error: `Marks must be 0-${exam.maxMarks}` });
      return;
    }
    
    const percentage = (marks / exam.maxMarks) * 100;
    const grade = calculateGrade(percentage);
    
    validMarks.push({
      studentId: studentId,
      studentName: student.name,
      subject: subject,
      examId: examId,
      examName: exam.name,
      class: student.class,
      section: student.section,
      maxMarks: exam.maxMarks,
      marks: marks,
      percentage: percentage,
      grade: grade
    });
    
    results.created++;
    
    results.preview.push({
      studentId: studentId,
      studentName: student.name,
      subject: subject,
      examName: exam.name,
      marks: marks,
      maxMarks: exam.maxMarks,
      percentage: percentage.toFixed(1) + "%",
      grade: grade,
      status: "VALID"
    });
  });
  
  // If preview only, return without writing
  if (previewOnly) {
    return {
      success: true,
      preview: true,
      results: results,
      message: `Preview: ${results.created} valid, ${results.failed} failed`
    };
  }
  
  // Write marks
  const marksSheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const academicYear = getCurrentAcademicYear();
  
  validMarks.forEach(m => {
    marksSheet.appendRow([
      `MRK${Date.now()}${Math.random().toString(36).substr(2, 5)}`,
      m.studentId,
      m.studentName,
      m.subject,
      "",  // subjectCode
      "ADMIN",
      "Administrator",
      m.examId,
      m.examName,
      m.class,
      m.section,
      m.maxMarks,
      m.marks,
      m.percentage.toFixed(2),
      m.grade,
      new Date(),
      getCurrentUser(),
      academicYear
    ]);
  });
  
  logAction("Admin Bulk Upload Marks", `Created: ${results.created}, Failed: ${results.failed}`);
  
  return {
    success: true,
    preview: false,
    results: results,
    message: `Import complete: ${results.created} entries added, ${results.failed} failed`
  };
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
