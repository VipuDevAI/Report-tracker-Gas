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
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    try { ensureYearNotFinalized("Create exam"); } catch (e) { return { success: false, message: e.message }; }
    
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
      totalMaxMarks,
      false  // IsDeleted
    ]);
    
    const internalInfo = examData.hasInternals ? ` (Theory: ${examData.maxMarks}, Internals: ${internal1+internal2+internal3+internal4})` : '';
    try { writeAudit("CREATE_EXAM", "Exam", examId, "*", "", examData.name, { class: examData.class, maxMarks: examData.maxMarks }); } catch (e) {}
    logAction("Create Exam", `Created exam: ${examData.name} (${examId}) for ${academicYear}${internalInfo}`);
    
    return { 
      success: true, 
      message: `Exam "${examData.name}" created successfully! Total marks: ${totalMaxMarks}`,
      examId: examId
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get all exams with optional filters
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered exams
 */
function getExams(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const lastRow = sheet.getLastRow();
  const currentYear = getCurrentAcademicYear();
  
  if (lastRow <= 1) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  const includeDeleted = filters && filters.includeDeleted === true;
  
  let exams = data.map(row => ({
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
    totalMaxMarks: row[17] || row[4],
    isDeleted: row[18] === true
  })).filter(e => e.examId && (includeDeleted || !e.isDeleted));
  
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
  const lastRow = sheet.getLastRow();
  const currentYear = getCurrentAcademicYear();
  if (lastRow <= 1) return null;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === examId && data[i][18] !== true) {
      const row = data[i];
      return {
        examId: row[0], name: row[1], examType: row[2], class: row[3],
        maxMarks: row[4], weightage: row[5], startDate: row[6], endDate: row[7],
        locked: row[8], createdBy: row[9], createdAt: row[10],
        academicYear: row[11] || currentYear,
        hasInternals: row[12] || false, internal1: row[13] || 0,
        internal2: row[14] || 0, internal3: row[15] || 0, internal4: row[16] || 0,
        totalMaxMarks: row[17] || row[4]
      };
    }
  }
  return null;
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
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { ensureYearNotFinalized("Update exam"); } catch (e) { return { success: false, message: e.message }; }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Exam not found." };
    const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
    
    let foundIdx = -1;
    for (let i = 0; i < data.length; i++) { if (data[i][0] === examId && data[i][18] !== true) { foundIdx = i; break; } }
    if (foundIdx === -1) return { success: false, message: "Exam not found." };
    
    const row = data[foundIdx];
    if (row[8] === true && !updates.forceUpdate) {
      return { success: false, message: "Exam is locked. Cannot modify." };
    }
    
    const i1 = updates.internal1 !== undefined ? updates.internal1 : row[13];
    const i2 = updates.internal2 !== undefined ? updates.internal2 : row[14];
    const i3 = updates.internal3 !== undefined ? updates.internal3 : row[15];
    const i4 = updates.internal4 !== undefined ? updates.internal4 : row[16];
    const max = updates.maxMarks || row[4];
    
    const updatedRow = [
      examId,
      updates.name || row[1],
      updates.examType || row[2],
      updates.class || row[3],
      max,
      updates.weightage || row[5],
      updates.startDate || row[6],
      updates.endDate || row[7],
      row[8], row[9], row[10], row[11],
      updates.hasInternals !== undefined ? updates.hasInternals : row[12],
      i1, i2, i3, i4,
      max + i1 + i2 + i3 + i4,
      false
    ];
    
    sheet.getRange(foundIdx + 2, 1, 1, 19).setValues([updatedRow]);
    try { writeAudit("UPDATE_EXAM", "Exam", examId, "*", row[1], updatedRow[1], { changes: updates }); } catch (e) {}
    logAction("Update Exam", `Updated exam: ${examId}`);
    return { success: true, message: "Exam updated successfully!" };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
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
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { ensureYearNotFinalized("Lock exam"); } catch (e) { return { success: false, message: e.message }; }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Exam not found." };
    const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
    let foundIdx = -1;
    for (let i = 0; i < data.length; i++) { if (data[i][0] === examId && data[i][18] !== true) { foundIdx = i; break; } }
    if (foundIdx === -1) return { success: false, message: "Exam not found." };
    
    sheet.getRange(foundIdx + 2, 9).setValue(true);
    try { writeAudit("LOCK_EXAM", "Exam", examId, "Locked", "false", "true", { examName: data[foundIdx][1] }); } catch (e) {}
    logAction("Lock Exam", `Locked exam: ${examId}`);
    
    // Auto-trigger aggregates rebuild after lock
    try { if (typeof rebuildAggregates === 'function') rebuildAggregates(); } catch (e) {}
    
    return { success: true, message: "Exam locked successfully!" };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Unlock an exam (Admin only)
 */
function unlockExam(examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to unlock exams." };
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { ensureYearNotFinalized("Unlock exam"); } catch (e) { return { success: false, message: e.message }; }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Exam not found." };
    const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
    let foundIdx = -1;
    for (let i = 0; i < data.length; i++) { if (data[i][0] === examId && data[i][18] !== true) { foundIdx = i; break; } }
    if (foundIdx === -1) return { success: false, message: "Exam not found." };
    
    sheet.getRange(foundIdx + 2, 9).setValue(false);
    try { writeAudit("UNLOCK_EXAM", "Exam", examId, "Locked", "true", "false", { examName: data[foundIdx][1] }); } catch (e) {}
    logAction("Unlock Exam", `Unlocked exam: ${examId}`);
    return { success: true, message: "Exam unlocked successfully!" };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Soft-delete an exam (sets IsDeleted=true). Associated marks are also soft-deleted.
 */
function deleteExam(examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to delete exams." };
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { ensureYearNotFinalized("Delete exam"); } catch (e) { return { success: false, message: e.message }; }
    
    const ss = SpreadsheetApp.getActive();
    const examSheet = ss.getSheetByName("Exams");
    const lastRow = examSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Exam not found." };
    const data = examSheet.getRange(2, 1, lastRow - 1, 19).getValues();
    let foundIdx = -1;
    for (let i = 0; i < data.length; i++) { if (data[i][0] === examId && data[i][18] !== true) { foundIdx = i; break; } }
    if (foundIdx === -1) return { success: false, message: "Exam not found." };
    
    // Soft-delete exam
    examSheet.getRange(foundIdx + 2, 19).setValue(true);
    
    // Soft-delete associated marks (batch update)
    const marksSheet = ss.getSheetByName("Marks_Master");
    const marksLastRow = marksSheet.getLastRow();
    let marksDeleted = 0;
    if (marksLastRow > 1) {
      const marksData = marksSheet.getRange(2, 1, marksLastRow - 1, 20).getValues();
      for (let i = 0; i < marksData.length; i++) {
        if (marksData[i][7] === examId && marksData[i][19] !== true) {
          marksSheet.getRange(i + 2, 20).setValue(true);
          marksDeleted++;
        }
      }
    }
    
    try { writeAudit("DELETE_EXAM", "Exam", examId, "IsDeleted", "false", "true", { examName: data[foundIdx][1], cascadedMarks: marksDeleted }); } catch (e) {}
    logAction("Delete Exam", `Soft-deleted exam: ${examId} (cascaded ${marksDeleted} marks)`);
    return { success: true, message: `Exam moved to trash. ${marksDeleted} associated marks also moved to trash.` };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Restore soft-deleted exam (and optionally cascaded marks)
 */
function restoreExam(examId, restoreMarks) {
  if (!isAdmin()) return { success: false, message: "Access denied. Admin only." };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActive();
    const examSheet = ss.getSheetByName("Exams");
    const lastRow = examSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Exam not found." };
    const data = examSheet.getRange(2, 1, lastRow - 1, 19).getValues();
    let foundIdx = -1;
    for (let i = 0; i < data.length; i++) { if (data[i][0] === examId && data[i][18] === true) { foundIdx = i; break; } }
    if (foundIdx === -1) return { success: false, message: "Exam not found in trash." };
    examSheet.getRange(foundIdx + 2, 19).setValue(false);
    
    let marksRestored = 0;
    if (restoreMarks === true) {
      const marksSheet = ss.getSheetByName("Marks_Master");
      const marksLastRow = marksSheet.getLastRow();
      if (marksLastRow > 1) {
        const md = marksSheet.getRange(2, 1, marksLastRow - 1, 20).getValues();
        for (let i = 0; i < md.length; i++) {
          if (md[i][7] === examId && md[i][19] === true) {
            marksSheet.getRange(i + 2, 20).setValue(false);
            marksRestored++;
          }
        }
      }
    }
    
    try { writeAudit("RESTORE_EXAM", "Exam", examId, "IsDeleted", "true", "false", { restoredMarks: marksRestored }); } catch (e) {}
    return { success: true, message: `Exam restored.${marksRestored ? ' ' + marksRestored + ' marks restored.' : ''}` };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get soft-deleted exams (for Trash UI)
 */
function getDeletedExams() {
  if (!isAdmin()) return [];
  const sheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  return data.filter(r => r[0] && r[18] === true).map(r => ({
    examId: r[0], name: r[1], examType: r[2], class: r[3], maxMarks: r[4], academicYear: r[11]
  }));
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
