/************************************************
 MVM REPORT TRACKER - EXAM MANAGEMENT
 File 4 of 7
 With Academic Year Support
************************************************/

/**
 * Create a new exam (Admin only)
 * @param {Object} examData - Exam details
 * @returns {Object} Result object
 */
function createExam(examData) {
  // Validate admin access
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to create exams." };
  }
  
  // Validate input
  if (!examData || !examData.name || !examData.maxMarks) {
    return { success: false, message: "Exam name and max marks are required." };
  }
  
  if (examData.maxMarks <= 0) {
    return { success: false, message: "Max marks must be greater than 0." };
  }
  
  if (examData.weightage && (examData.weightage < 0 || examData.weightage > 100)) {
    return { success: false, message: "Weightage must be between 0 and 100." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const examId = `EXM${Date.now()}`;
  const academicYear = getCurrentAcademicYear();
  
  // Calculate total max marks including internals
  const internal1 = examData.internal1 || 0;
  const internal2 = examData.internal2 || 0;
  const internal3 = examData.internal3 || 0;
  const internal4 = examData.internal4 || 0;
  const totalMaxMarks = examData.maxMarks + internal1 + internal2 + internal3 + internal4;
  
  sheet.appendRow([
    examId,
    examData.name,
    examData.examType || "Regular",
    examData.class || "All",
    examData.maxMarks,
    examData.weightage || 100,
    examData.startDate || new Date(),
    examData.endDate || new Date(),
    false,  // Not locked
    getLoggedInUser() || getCurrentUser(),
    new Date(),
    academicYear,
    examData.hasInternals || false,
    internal1,
    internal2,
    internal3,
    internal4,
    totalMaxMarks
  ]);
  
  const internalInfo = examData.hasInternals ? ` (Theory: ${examData.maxMarks}, Internals: ${internal1+internal2+internal3+internal4})` : '';
  logAction("Create Exam", `Created exam: ${examData.name} (${examId}) for ${academicYear}${internalInfo}`);
  
  return { 
    success: true, 
    message: `Exam "${examData.name}" created successfully! Total marks: ${totalMaxMarks}`,
    examId: examId
  };
}


/**
 * Get all exams with optional filters
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered exams
 */
function getExams(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const data = sheet.getDataRange().getValues();
  const currentYear = getCurrentAcademicYear();
  
  if (data.length <= 1) return [];
  
  let exams = data.slice(1).map(row => ({
    examId: row[0],
    name: row[1],
    examType: row[2],
    class: row[3],
    maxMarks: row[4],
    weightage: row[5],
    startDate: row[6],
    endDate: row[7],
    locked: row[8],
    createdBy: row[9],
    createdAt: row[10],
    academicYear: row[11] || currentYear,
    hasInternals: row[12] || false,
    internal1: row[13] || 0,
    internal2: row[14] || 0,
    internal3: row[15] || 0,
    internal4: row[16] || 0,
    totalMaxMarks: row[17] || row[4]
  }));
  
  if (filters) {
    if (filters.examType) {
      exams = exams.filter(e => e.examType === filters.examType);
    }
    if (filters.class) {
      exams = exams.filter(e => e.class === filters.class || e.class === "All");
    }
    if (filters.locked !== undefined) {
      exams = exams.filter(e => e.locked === filters.locked);
    }
    if (filters.academicYear) {
      exams = exams.filter(e => e.academicYear === filters.academicYear);
    } else {
      // Default: filter by current academic year
      exams = exams.filter(e => e.academicYear === currentYear);
    }
  } else {
    // Default: filter by current academic year
    exams = exams.filter(e => e.academicYear === currentYear);
  }
  
  return exams;
}


/**
 * Get single exam by ID
 * @param {string} examId - Exam ID
 * @returns {Object|null} Exam object or null
 */
function getExamById(examId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const data = sheet.getDataRange().getValues();
  const currentYear = getCurrentAcademicYear();
  
  if (data.length <= 1) return null;
  
  const row = data.find(r => r[0] === examId);
  if (!row) return null;
  
  return {
    examId: row[0],
    name: row[1],
    examType: row[2],
    class: row[3],
    maxMarks: row[4],
    weightage: row[5],
    startDate: row[6],
    endDate: row[7],
    locked: row[8],
    createdBy: row[9],
    createdAt: row[10],
    academicYear: row[11] || currentYear
  };
}


/**
 * Update exam details (Admin only)
 * @param {string} examId - Exam ID to update
 * @param {Object} updates - Fields to update
 * @returns {Object} Result object
 */
function updateExam(examId, updates) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const data = sheet.getDataRange().getValues();
  
  const rowIndex = data.findIndex(r => r[0] === examId);
  
  if (rowIndex === -1) {
    return { success: false, message: "Exam not found." };
  }
  
  const row = data[rowIndex];
  
  // Check if exam is locked
  if (row[8] === true && !updates.forceUpdate) {
    return { success: false, message: "Exam is locked. Cannot modify." };
  }
  
  const updatedRow = [
    examId,
    updates.name || row[1],
    updates.examType || row[2],
    updates.class || row[3],
    updates.maxMarks || row[4],
    updates.weightage || row[5],
    updates.startDate || row[6],
    updates.endDate || row[7],
    row[8],  // Keep lock status
    row[9],
    row[10],
    row[11]  // Keep academic year
  ];
  
  sheet.getRange(rowIndex + 1, 1, 1, 12).setValues([updatedRow]);
  
  logAction("Update Exam", `Updated exam: ${examId}`);
  
  return { success: true, message: "Exam updated successfully!" };
}


/**
 * Lock an exam (Admin only - prevents marks modification)
 * @param {string} examId - Exam ID to lock
 * @returns {Object} Result object
 */
function lockExam(examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to lock exams." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const data = sheet.getDataRange().getValues();
  
  const rowIndex = data.findIndex(r => r[0] === examId);
  
  if (rowIndex === -1) {
    return { success: false, message: "Exam not found." };
  }
  
  sheet.getRange(rowIndex + 1, 9).setValue(true);
  
  logAction("Lock Exam", `Locked exam: ${examId}`);
  
  return { success: true, message: "Exam locked successfully!" };
}


/**
 * Unlock an exam (Admin only)
 * @param {string} examId - Exam ID to unlock
 * @returns {Object} Result object
 */
function unlockExam(examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to unlock exams." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const data = sheet.getDataRange().getValues();
  
  const rowIndex = data.findIndex(r => r[0] === examId);
  
  if (rowIndex === -1) {
    return { success: false, message: "Exam not found." };
  }
  
  sheet.getRange(rowIndex + 1, 9).setValue(false);
  
  logAction("Unlock Exam", `Unlocked exam: ${examId}`);
  
  return { success: true, message: "Exam unlocked successfully!" };
}


/**
 * Delete an exam and associated marks (Admin only)
 * @param {string} examId - Exam ID to delete
 * @returns {Object} Result object
 */
function deleteExam(examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to delete exams." };
  }
  
  const ss = SpreadsheetApp.getActive();
  
  // Delete exam
  const examSheet = ss.getSheetByName("Exams");
  const examData = examSheet.getDataRange().getValues();
  const examRowIndex = examData.findIndex(r => r[0] === examId);
  
  if (examRowIndex === -1) {
    return { success: false, message: "Exam not found." };
  }
  
  examSheet.deleteRow(examRowIndex + 1);
  
  // Delete associated marks
  const marksSheet = ss.getSheetByName("Marks_Master");
  const marksData = marksSheet.getDataRange().getValues();
  
  // Delete from bottom to top to maintain row indices
  for (let i = marksData.length - 1; i >= 1; i--) {
    if (marksData[i][7] === examId) {
      marksSheet.deleteRow(i + 1);
    }
  }
  
  logAction("Delete Exam", `Deleted exam: ${examId} and associated marks`);
  
  return { success: true, message: "Exam and associated marks deleted successfully!" };
}


/**
 * Check if exam is locked
 * @param {string} examId - Exam ID
 * @returns {boolean} True if locked
 */
function isExamLocked(examId) {
  const exam = getExamById(examId);
  return exam ? exam.locked : false;
}


/**
 * Get exam types
 * @returns {Array} Available exam types
 */
function getExamTypes() {
  return [
    "Unit Test 1",
    "Unit Test 2",
    "Unit Test 3",
    "Unit Test 4",
    "Midterm",
    "Final",
    "Pre-Board",
    "Board Practice",
    "Assignment",
    "Project",
    "Practical",
    "Other"
  ];
}


/**
 * Get max marks options
 * @returns {Array} Common max marks values
 */
function getMaxMarksOptions() {
  return [10, 20, 25, 30, 40, 50, 70, 80, 100];
}


/**
 * Get exams for current academic year (for dropdown)
 * @returns {Array} Current year exams
 */
function getCurrentYearExams() {
  return getExams({ academicYear: getCurrentAcademicYear() });
}
