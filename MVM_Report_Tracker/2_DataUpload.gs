/************************************************
 MVM REPORT TRACKER - DATA UPLOAD & MANAGEMENT
 File 2 of 7
 With Bulk Upload Features
************************************************/

/**
 * Internal: assert the current user can write to a student in the given class.
 * Admin → always allowed.
 * Wing Admin → allowed if class is within their wing scope.
 * Teacher → not allowed (use marks endpoints).
 * Principal → never allowed.
 *
 * @param {string|number} cls - Class number of the student being touched
 * @returns {Object|null} null if allowed, else { success:false, message }
 */
function _denyIfNoStudentWrite(cls) {
  const role = getRole();
  if (role === "admin") return null;
  if (role === "wing_admin") {
    const wa = getWingAdminAssignment();
    const allowed = wa && wa.classes.map(String).includes(String(cls || ""));
    if (!allowed) return { success: false, message: "Access denied. This class is outside your wing scope." };
    return null;
  }
  return { success: false, message: "Access denied. Insufficient privileges." };
}

/**
 * Bulk upload students from CSV/array data
 * @param {Array} data - 2D array of student data
 * @param {Object} options - { updateExisting: boolean, preview: boolean }
 * @returns {Object} Result object with summary
 */
function bulkUploadStudents(data, options) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!data || !Array.isArray(data) || data.length === 0) {
    return { success: false, message: "No data provided." };
  }
  
  const opts = options || {};
  const updateExisting = opts.updateExisting || false;
  const previewOnly = opts.preview || false;
  
  const lock = LockService.getScriptLock();
  try {
    if (!previewOnly) {
      lock.waitLock(30000);
      try { ensureYearNotFinalized("Bulk upload students"); } catch (e) { return { success: false, message: e.message }; }
    }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 16);
    const existingData = lastRow > 0
      ? sheet.getRange(1, 1, lastRow, lastCol).getValues()
      : [[]];
    const academicYear = getCurrentAcademicYear();
    
    const existingIndex = {};
    const existingIdIndex = {};
    existingData.slice(1).forEach((row, idx) => {
      if (row[15] === true) return; // skip deleted
      const key = `${row[2]}-${row[3]}-${row[5]}`;
      existingIndex[key] = { row: row, index: idx + 2 };
      if (row[0]) existingIdIndex[row[0]] = { row: row, index: idx + 2 };
    });
    
    const results = {
      preview: [],
      created: 0,
      updated: 0,
      failed: 0,
      errors: [],
      duplicates: []
    };
    
    const toCreate = [];
    const toUpdate = [];
    
    data.forEach((row, rowIdx) => {
      // Skip header row if detected
      if (rowIdx === 0 && (row[0] === "StudentID" || row[0] === "StudentId" || row[1] === "Name")) {
        return;
      }
      
      // Input columns: StudentID, Name, Class, Section, Stream, RollNo, Status, ElectiveSubject, LanguageL1, LanguageL2, LanguageL3, ParentEmail, Phone
      const studentId = row[0] || `STU${Date.now()}${rowIdx}`;
      const name = row[1] || "";
      const cls = String(row[2] || "");
      const section = row[3] || "A1";
      const stream = row[4] || "";
      const rollNo = row[5] || rowIdx;
      const status = row[6] || "Active";
      const electiveSubject = row[7] || "";
      const langL1 = row[8] || "English";
      const langL2 = row[9] || "";
      const langL3 = row[10] || "";
      const parentEmail = row[11] || "";
      const phone = row[12] || "";
      
      if (!name) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: "Name is required" });
        return;
      }
      
      if (!cls) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: "Class is required" });
        return;
      }
      
      // Validate elective for class 11 & 12
      const validElectives = ['Mathematics', 'Applied Mathematics', 'Hindi', 'History', 'Sanskrit', 'Computer Science', 'Biology', ''];
      if ((cls == '11' || cls == '12') && electiveSubject && !validElectives.includes(electiveSubject)) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: `Invalid elective: ${electiveSubject}` });
        return;
      }
      
      const key = `${cls}-${section}-${rollNo}`;
      const existingByKey = existingIndex[key];
      const existingById = existingIdIndex[studentId];
      
      const studentData = [
        studentId,
        name,
        cls,
        section,
        stream,
        rollNo,
        parentEmail,
        phone,
        new Date(),
        status,
        electiveSubject,
        academicYear,
        langL1,
        langL2,
        langL3,
        false  // IsDeleted
      ];
      
      if (existingByKey || existingById) {
        if (updateExisting) {
          const existingRecord = existingByKey || existingById;
          toUpdate.push({
            rowIndex: existingRecord.index,
            data: studentData,
            original: existingRecord.row
          });
          results.updated++;
        } else {
          results.duplicates.push({
            row: rowIdx + 1, name: name, class: cls, section: section, rollNo: rollNo
          });
          results.failed++;
        }
      } else {
        toCreate.push(studentData);
        results.created++;
      }
      
      results.preview.push({
        studentId: studentId,
        name: name,
        class: cls,
        section: section,
        stream: stream,
        rollNo: rollNo,
        status: existingByKey || existingById ? (updateExisting ? "UPDATE" : "DUPLICATE") : "NEW"
      });
    });
    
    if (previewOnly) {
      return {
        success: true,
        preview: true,
        results: results,
        message: `Preview: ${results.created} new, ${results.updated} updates, ${results.failed} failed`
      };
    }
    
    if (toCreate.length > 0) {
      const writeRow = sheet.getLastRow() + 1;
      sheet.getRange(writeRow, 1, toCreate.length, 16).setValues(toCreate);
    }
    
    toUpdate.forEach(item => {
      sheet.getRange(item.rowIndex, 1, 1, 16).setValues([item.data]);
    });
    
    logAction("Bulk Upload Students", `Created: ${results.created}, Updated: ${results.updated}, Failed: ${results.failed}`);
    
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
 * Bulk upload teachers (Google email-based auth)
 * @param {Array} data - 2D array of teacher data
 * @param {Object} options - { updateExisting: boolean, preview: boolean }
 * @returns {Object} Result object
 */
function bulkUploadTeachers(data, options) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!data || !Array.isArray(data) || data.length === 0) {
    return { success: false, message: "No data provided." };
  }
  
  // Idempotent migration for the Role column
  try { ensureTeachersRoleColumn(); } catch (e) {}
  
  const opts = options || {};
  const updateExisting = opts.updateExisting || false;
  const previewOnly = opts.preview || false;
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const existingData = sheet.getDataRange().getValues();
  
  // Build index of existing teachers by email
  const existingIndex = {};
  existingData.slice(1).forEach((row, idx) => {
    if (row[5]) existingIndex[row[5].toLowerCase()] = { row: row, index: idx + 2 };
  });
  
  const results = {
    preview: [],
    created: 0,
    updated: 0,
    failed: 0,
    errors: []
  };
  
  const toCreate = [];
  const toUpdate = [];
  
  data.forEach((row, rowIdx) => {
    // Skip header row
    if (rowIdx === 0 && (row[0] === "TeacherID" || row[0] === "TeacherId" || row[1] === "Name")) {
      return;
    }
    
    const teacherId = row[0] || `TCH${Date.now()}${rowIdx}`;
    const name = row[1] || "";
    const subject = row[2] || "";
    // Normalize classes/sections: trim, standardize separators to comma
    const splitNorm = (s) => String(s || "").split(/[,;|]/).map(x => x.trim()).filter(x => x).join(",");
    const classes = splitNorm(row[3]);
    const sections = splitNorm(row[4]);
    const email = row[5] || "";
    const phone = row[6] || "";
    const status = row[7] || "Active";
    // Optional Role column (index 8) — supports new CSVs; defaults to TEACHER
    const validRoles = ["ADMIN", "PRINCIPAL", "WING_ADMIN", "TEACHER"];
    let role = String(row[8] || "TEACHER").trim().toUpperCase();
    if (role === "WINGADMIN") role = "WING_ADMIN";
    if (!validRoles.includes(role)) role = "TEACHER";
    
    // Validation
    if (!name) {
      results.failed++;
      results.errors.push({ row: rowIdx + 1, error: "Name is required" });
      return;
    }
    
    if (!email) {
      results.failed++;
      results.errors.push({ row: rowIdx + 1, error: "Email is required (used for Google login)" });
      return;
    }
    
    // Validate email format
    if (!email.includes("@")) {
      results.failed++;
      results.errors.push({ row: rowIdx + 1, error: "Invalid email format" });
      return;
    }
    
    const teacherData = [
      teacherId,
      name,
      subject,
      classes,
      sections,
      email,
      phone,
      new Date(),
      status,
      "",   // IsClassTeacher (index 9)
      "",   // ClassTeacherOf (index 10)
      false, // IsDeleted (index 11)
      role   // Role (index 12)
    ];
    
    const existingByEmail = existingIndex[email.toLowerCase()];
    
    if (existingByEmail) {
      if (updateExisting) {
        toUpdate.push({
          rowIndex: existingByEmail.index,
          data: teacherData,
          original: existingByEmail.row
        });
        results.updated++;
      } else {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: `Email ${email} already exists` });
        return;
      }
    } else {
      toCreate.push(teacherData);
      results.created++;
    }
    
    results.preview.push({
      teacherId: teacherId,
      name: name,
      email: email,
      subject: subject,
      classes: classes,
      sections: sections,
      status: existingByEmail ? (updateExisting ? "UPDATE" : "DUPLICATE") : "NEW"
    });
  });
  
  // If preview only, return without writing
  if (previewOnly) {
    return {
      success: true,
      preview: true,
      results: results,
      message: `Preview: ${results.created} new, ${results.updated} updates, ${results.failed} failed`
    };
  }
  
  // Write new teachers (13 cols including Role)
  if (toCreate.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, toCreate.length, 13).setValues(toCreate);
  }
  
  // Update existing teachers (13 cols including Role)
  toUpdate.forEach(item => {
    sheet.getRange(item.rowIndex, 1, 1, 13).setValues([item.data]);
  });
  
  logAction("Bulk Upload Teachers", `Created: ${results.created}, Updated: ${results.updated}, Failed: ${results.failed}`);
  
  return {
    success: true,
    preview: false,
    results: results,
    message: `Import complete: ${results.created} created, ${results.updated} updated, ${results.failed} failed`
  };
}


// Password functions removed - using Google email-based authentication

/**
 * Upload/Replace students data (Legacy - kept for compatibility)
 * @param {Array} data - 2D array of student data
 * @returns {Object} Result object
 */
function replaceStudents(data) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  return bulkUploadStudents(data, { updateExisting: true });
}


/**
 * Add single student
 * @param {Object} student - Student data object
 * @returns {Object} Result object
 */
function addStudent(student) {
  if (!student || !student.name || !student.class) {
    return { success: false, message: "Name and Class are required." };
  }
  
  const deny = _denyIfNoStudentWrite(student.class);
  if (deny) return deny;
  
  // Validate elective subject for class 11 & 12
  const validElectives = ['Mathematics', 'Applied Mathematics', 'Hindi', 'History', 'Sanskrit', 'Computer Science', 'Biology'];
  if ((student.class == 11 || student.class == 12) && student.electiveSubject) {
    if (!validElectives.includes(student.electiveSubject)) {
      return { success: false, message: "Invalid elective subject. Choose from: " + validElectives.join(", ") };
    }
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { ensureYearNotFinalized("Add student"); } catch (e) { return { success: false, message: e.message }; }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const studentId = `STU${Date.now()}`;
    const academicYear = getCurrentAcademicYear();
    
    sheet.appendRow([
      studentId,
      student.name,
      student.class,
      student.section || "A1",
      student.stream || "",
      student.rollNo || sheet.getLastRow(),
      student.parentEmail || "",
      student.phone || "",
      new Date(),
      "Active",
      student.electiveSubject || "",
      academicYear,
      student.languageL1 || "English",
      student.languageL2 || "",
      student.languageL3 || "",
      false  // IsDeleted
    ]);
    
    try { writeAudit("CREATE_STUDENT", "Student", studentId, "*", "", student.name, { class: student.class, section: student.section }); } catch (e) {}
    logAction("Add Student", `Added student: ${student.name} (Class ${student.class})`);
    
    return { success: true, message: "Student added successfully!", studentId: studentId };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Update student
 * @param {string} studentId - Student ID to update
 * @param {Object} updates - Fields to update
 * @returns {Object} Result object
 */
function updateStudent(studentId, updates) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { ensureYearNotFinalized("Update student"); } catch (e) { return { success: false, message: e.message }; }
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Student not found." };
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    
    let foundIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === studentId && data[i][15] !== true) { foundIdx = i; break; }
    }
    
    if (foundIdx === -1) {
      return { success: false, message: "Student not found." };
    }
    
    const row = data[foundIdx];
    
    // Scope-check: writer must own current class AND target class
    const currentClass = row[2];
    const targetClass = updates.class || currentClass;
    let deny = _denyIfNoStudentWrite(currentClass);
    if (deny) return deny;
    if (String(targetClass) !== String(currentClass)) {
      deny = _denyIfNoStudentWrite(targetClass);
      if (deny) return deny;
    }
    const updatedRow = [
      studentId,
      updates.name || row[1],
      updates.class || row[2],
      updates.section || row[3],
      updates.stream !== undefined ? updates.stream : row[4],
      updates.rollNo || row[5],
      updates.parentEmail !== undefined ? updates.parentEmail : row[6],
      updates.phone !== undefined ? updates.phone : row[7],
      row[8],
      updates.status || row[9],
      updates.electiveSubject !== undefined ? updates.electiveSubject : (row[10] || ""),
      updates.academicYear || row[11] || getCurrentAcademicYear(),
      updates.languageL1 !== undefined ? updates.languageL1 : (row[12] || ""),
      updates.languageL2 !== undefined ? updates.languageL2 : (row[13] || ""),
      updates.languageL3 !== undefined ? updates.languageL3 : (row[14] || ""),
      false
    ];
    
    sheet.getRange(foundIdx + 2, 1, 1, 16).setValues([updatedRow]);
    
    try { writeAudit("UPDATE_STUDENT", "Student", studentId, "*", JSON.stringify({ name: row[1], class: row[2], section: row[3] }), JSON.stringify({ name: updatedRow[1], class: updatedRow[2], section: updatedRow[3] }), {}); } catch (e) {}
    logAction("Update Student", `Updated student: ${studentId}`);
    
    return { success: true, message: "Student updated successfully!" };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Soft-delete a student (sets IsDeleted=true)
 */
function deleteStudent(studentId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { ensureYearNotFinalized("Delete student"); } catch (e) { return { success: false, message: e.message }; }
    const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Student not found." };
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === studentId && data[i][15] !== true) {
        const deny = _denyIfNoStudentWrite(data[i][2]);
        if (deny) return deny;
        sheet.getRange(i + 2, 16).setValue(true);
        try { writeAudit("DELETE_STUDENT", "Student", studentId, "IsDeleted", "false", "true", { name: data[i][1], class: data[i][2], section: data[i][3] }); } catch (e) {}
        logAction("Delete Student", `Soft-deleted student: ${studentId}`);
        return { success: true, message: "Student moved to trash." };
      }
    }
    return { success: false, message: "Student not found or already deleted." };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Restore a soft-deleted student
 */
function restoreStudent(studentId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "Student not found." };
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === studentId && data[i][15] === true) {
        const deny = _denyIfNoStudentWrite(data[i][2]);
        if (deny) return deny;
        sheet.getRange(i + 2, 16).setValue(false);
        try { writeAudit("RESTORE_STUDENT", "Student", studentId, "IsDeleted", "true", "false", {}); } catch (e) {}
        return { success: true, message: "Student restored." };
      }
    }
    return { success: false, message: "Student not found in trash." };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get soft-deleted students for Trash UI
 * Admin/Principal: full list. Wing Admin: only their wing's classes. Teacher: empty.
 */
function getDeletedStudents() {
  const role = getRole();
  if (!role || role === "teacher") return [];
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  let items = data.filter(r => r[0] && r[15] === true).map(r => ({
    studentId: r[0], name: r[1], class: r[2], section: r[3], stream: r[4],
    rollNo: r[5], academicYear: r[11]
  }));
  if (role === "wing_admin") {
    const wa = getWingAdminAssignment();
    const allowed = wa ? wa.classes.map(String) : [];
    items = items.filter(s => allowed.includes(String(s.class)));
  }
  return items;
}


// (Removed duplicate deleteStudent that shadowed the soft-delete implementation above.)


/**
 * Get all students with optional filters
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters (class, section, stream, status)
 * @returns {Array} Filtered students
 */
function getStudents(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];
  
  // Single bulk read (no getDataRange in loops)
  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const currentYear = getCurrentAcademicYear();
  const includeDeleted = filters && filters.includeDeleted === true;
  
  let students = data.map(row => ({
    studentId: row[0],
    name: row[1],
    class: row[2],
    section: row[3],
    stream: row[4],
    rollNo: row[5],
    parentEmail: row[6],
    phone: row[7],
    joinDate: row[8],
    status: row[9],
    electiveSubject: row[10] || '',
    academicYear: row[11] || currentYear,
    languageL1: row[12] || '',
    languageL2: row[13] || '',
    languageL3: row[14] || '',
    isDeleted: row[15] === true
  })).filter(s => s.studentId && (includeDeleted || !s.isDeleted));
  
  // Apply filters
  if (filters) {
    if (filters.class) {
      students = students.filter(s => s.class == filters.class);
    }
    if (filters.section) {
      students = students.filter(s => s.section === filters.section);
    }
    if (filters.stream) {
      students = students.filter(s => s.stream === filters.stream);
    }
    if (filters.status) {
      students = students.filter(s => s.status === filters.status);
    } else {
      students = students.filter(s => s.status === "Active");
    }
    if (filters.academicYear) {
      students = students.filter(s => s.academicYear === filters.academicYear);
    } else if (filters.academicYear !== null) {
      students = students.filter(s => s.academicYear === currentYear);
    }
    if (filters.search) {
      const q = String(filters.search).toLowerCase();
      students = students.filter(s =>
        String(s.name || '').toLowerCase().includes(q) ||
        String(s.studentId || '').toLowerCase().includes(q) ||
        String(s.rollNo || '').toLowerCase().includes(q)
      );
    }
  } else {
    students = students.filter(s => s.status === "Active" && s.academicYear === currentYear);
  }
  
  // Apply teacher assignment filter (server-side)
  students = applyTeacherFilter(students, { filterBySubject: false });
  
  return students;
}


/**
 * Server-side paginated getStudents
 * @param {Object} filters - Same filters as getStudents + { search }
 * @param {number} page - 1-indexed page number (default 1)
 * @param {number} limit - Rows per page (default 100)
 * @returns {Object} { data, total, page, limit, totalPages }
 */
function getStudentsPage(filters, page, limit) {
  const all = getStudents(filters);
  const total = all.length;
  const lim = Math.max(1, parseInt(limit) || 100);
  const totalPages = Math.max(1, Math.ceil(total / lim));
  const pg = Math.min(Math.max(1, parseInt(page) || 1), totalPages);
  const start = (pg - 1) * lim;
  return {
    data: all.slice(start, start + lim),
    total: total,
    page: pg,
    limit: lim,
    totalPages: totalPages
  };
}


/**
 * Upload/Replace teachers data (Legacy - kept for compatibility)
 * @param {Array} data - 2D array of teacher data
 * @returns {Object} Result object
 */
function replaceTeachers(data) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  return bulkUploadTeachers(data, { updateExisting: true });
}


/**
 * Add single teacher (Google email-based auth)
 * @param {Object} teacher - Teacher data object
 * @returns {Object} Result object
 */
function addTeacher(teacher) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!teacher || !teacher.name) {
    return { success: false, message: "Name is required." };
  }
  
  if (!teacher.email) {
    return { success: false, message: "Email is required (used for Google login)." };
  }
  
  // Idempotent migration for the Role column
  try { ensureTeachersRoleColumn(); } catch (e) {}
  
  // Check if email already exists
  const existingTeacher = getTeacherByEmail(teacher.email);
  if (existingTeacher) {
    return { success: false, message: "A teacher with this email already exists." };
  }
  
  // Normalize role
  const validRoles = ["ADMIN", "PRINCIPAL", "WING_ADMIN", "TEACHER"];
  let role = String(teacher.role || "TEACHER").trim().toUpperCase();
  if (role === "WINGADMIN") role = "WING_ADMIN";
  if (!validRoles.includes(role)) role = "TEACHER";
  
  // Normalize classes/sections: trim, standardize separators to comma
  const splitNorm = (s) => String(s || "").split(/[,;|]/).map(x => x.trim()).filter(x => x).join(",");
  const classes = splitNorm(teacher.classes);
  const sections = splitNorm(teacher.sections);
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const teacherId = `TCH${Date.now()}`;
  
  sheet.appendRow([
    teacherId,
    teacher.name,
    teacher.subject || "",
    classes,
    sections,
    teacher.email,
    teacher.phone || "",
    new Date(),
    "Active",
    "",      // IsClassTeacher
    "",      // ClassTeacherOf
    false,   // IsDeleted
    role     // Role
  ]);
  
  logAction("Add Teacher", `Added teacher: ${teacher.name} (${teacher.email}) [${role}]`);
  
  return { 
    success: true, 
    message: `${role.replace('_',' ').toLowerCase()} added successfully! They can now login with their email & password.`,
    teacherId: teacherId
  };
}


/**
 * Get all teachers with optional filters
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered teachers
 */
function getTeachers(filters) {
  const role = getRole();
  if (!role) return [];
  
  // Non-admin/principal: limited visibility
  if (role !== "admin" && role !== "principal") {
    if (role === "wing_admin") {
      // Wing admin sees teachers whose Classes intersect their wing
      const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return [];
      const wa = getWingAdminAssignment();
      const wingClasses = wa ? wa.classes.map(String) : [];
      const splitNorm = (s) => String(s || "").split(/[,;|]/).map(x => x.trim()).filter(x => x);
      
      let teachers = data.slice(1)
        .filter(r => r[0] && (r[11] !== true)) // skip deleted
        .map(r => ({
          teacherId: r[0], name: r[1], subject: r[2],
          classes: r[3], sections: r[4], email: r[5], phone: r[6],
          joinDate: r[7], status: r[8],
          role: (String(r[12] || "TEACHER").toUpperCase())
        }))
        .filter(t => {
          const tc = splitNorm(t.classes).map(String);
          return tc.some(c => wingClasses.includes(c));
        });
      return teachers;
    }
    
    // Plain teacher: only their own info
    const assignment = getTeacherAssignment();
    if (assignment) {
      return [{
        teacherId: assignment.teacherId,
        name: assignment.name,
        subject: assignment.subject,
        classes: assignment.classes.join(","),
        sections: assignment.sections.join(","),
        email: assignment.email,
        status: "Active",
        role: (assignment.role || "TEACHER")
      }];
    }
    return [];
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  let teachers = data.slice(1)
    .filter(row => row[0] && (row[11] !== true)) // skip deleted
    .map(row => ({
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
      role: (String(row[12] || "TEACHER").toUpperCase())
    }));
  
  if (filters) {
    if (filters.subject) {
      teachers = teachers.filter(t => t.subject === filters.subject);
    }
    if (filters.role) {
      const r = String(filters.role).toUpperCase();
      teachers = teachers.filter(t => t.role === r);
    }
    if (filters.status) {
      teachers = teachers.filter(t => t.status === filters.status);
    } else {
      teachers = teachers.filter(t => t.status === "Active");
    }
  }
  
  return teachers;
}


/**
 * Get subjects with optional filters
 * @param {Object} filters - Optional filters (class, stream)
 * @returns {Array} Filtered subjects
 */
function getSubjects(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Subjects");
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  let subjects = data.map(row => ({
    subjectId: row[0],
    subjectName: row[1],
    subjectCode: row[2],
    classes: row[3],
    stream: row[4],
    maxMarks: row[5],
    passingMarks: row[6],
    isActive: row[7],
    languageGroup: row[8] || '',
    isOptional: row[9] === true
  })).filter(s => s.subjectId);
  
  if (filters) {
    if (filters.class) {
      subjects = subjects.filter(s => String(s.classes).split(',').map(c => c.trim()).includes(String(filters.class)));
    }
    if (filters.stream) {
      subjects = subjects.filter(s => s.stream === filters.stream);
    }
    if (filters.isActive !== undefined) {
      subjects = subjects.filter(s => s.isActive === filters.isActive);
    } else {
      subjects = subjects.filter(s => s.isActive === true);
    }
  }
  
  return subjects;
}


/**
 * Get classes with optional filters
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered classes
 */
function getClasses(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Classes");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  let classes = data.slice(1).map(row => ({
    classId: row[0],
    className: row[1],
    sections: row[2].split(","),
    stream: row[3],
    academicYear: row[4],
    isActive: row[5]
  }));
  
  if (filters) {
    if (filters.stream) {
      classes = classes.filter(c => c.stream === filters.stream);
    }
    if (filters.isActive !== undefined) {
      classes = classes.filter(c => c.isActive === filters.isActive);
    }
  }
  
  return classes;
}


/**
 * Get unique class numbers (9, 10, 11, 12)
 * @returns {Array} Unique class numbers
 */
function getClassNumbers() {
  return ["6", "7", "8", "9", "10", "11", "12"];
}


/**
 * Get sections
 * @returns {Array} Available sections
 */
function getSections() {
  return ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11"];
}


/**
 * Get streams
 * @returns {Array} Available streams
 */
function getStreams() {
  return ["Science", "Computer Science", "Commerce"];
}


/**
 * Auto-promote students 6 → 7 → 8 → 9 → 10 → 11 → 12
 * Section is preserved unless changed by admin afterwards.
 * Class 12 students are marked Alumni (graduated).
 * @param {string} fromYear - Source academic year (informational, used in log)
 * @param {string} toYear - Target academic year
 * @param {Object} options - { resetRollNumbers: boolean }
 * @returns {Object} Result object
 */
function promoteStudents(fromYear, toYear, options) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  if (!toYear) {
    return { success: false, message: "Target academic year is required." };
  }
  
  const opts = options || {};
  const resetRoll = opts.resetRollNumbers === true;
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, message: "No students to promote.", promoted: 0, graduated: 0 };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    
    let promoted = 0;
    let graduated = 0;
    let skipped = 0;
    
    // Group by class+section for roll number reset
    const rollCounters = {};
    
    // Build new data array (sorted by class then section then current rollno for stable roll reset)
    const updatedData = data.map(row => row.slice());
    
    // First pass: promote/graduate
    updatedData.forEach(row => {
      const currentClass = parseInt(row[2]);
      const status = String(row[9] || "").trim();
      if (status !== "Active") { skipped++; return; }
      if (isNaN(currentClass)) { skipped++; return; }
      
      if (currentClass >= 12) {
        row[9] = "Alumni";
        graduated++;
      } else if (currentClass >= 6 && currentClass <= 11) {
        row[2] = currentClass + 1;
        row[11] = toYear; // AcademicYear updated
        promoted++;
      } else {
        skipped++;
      }
    });
    
    // Second pass: reset roll numbers if requested
    if (resetRoll) {
      // Sort active students by class+section+oldRoll for stable order
      const indexed = updatedData
        .map((row, idx) => ({ row, idx }))
        .filter(x => x.row[9] === "Active")
        .sort((a, b) => {
          if (a.row[2] !== b.row[2]) return parseInt(a.row[2]) - parseInt(b.row[2]);
          if (a.row[3] !== b.row[3]) return String(a.row[3]).localeCompare(String(b.row[3]));
          return parseInt(a.row[5]) - parseInt(b.row[5]);
        });
      indexed.forEach(item => {
        const key = `${item.row[2]}-${item.row[3]}`;
        rollCounters[key] = (rollCounters[key] || 0) + 1;
        item.row[5] = rollCounters[key];
      });
    }
    
    // Write back in one batch
    sheet.getRange(2, 1, updatedData.length, 15).setValues(updatedData);
    
    // Update setting
    updateSchoolSetting("AcademicYear", toYear);
    
    logAction("Promote Students", `From ${fromYear || '?'} to ${toYear}: Promoted ${promoted}, Graduated ${graduated}, Skipped ${skipped}${resetRoll ? ' (rolls reset)' : ''}`);
    
    return {
      success: true,
      message: `Promotion complete: ${promoted} promoted, ${graduated} graduated, ${skipped} skipped${resetRoll ? ' (roll numbers reset)' : ''}`,
      promoted: promoted,
      graduated: graduated,
      skipped: skipped
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Archive year data
 * @param {string} academicYear - Year to archive
 * @returns {Object} Result object
 */
function archiveYearData(academicYear) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  const ss = SpreadsheetApp.getActive();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy_MM_dd");
  
  const sheetsToArchive = ["Students", "Marks_Master", "Exams"];
  
  sheetsToArchive.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      const copy = sheet.copyTo(ss);
      copy.setName(`${name}_${academicYear}_${timestamp}`);
    }
  });
  
  logAction("Archive Year", `Archived data for ${academicYear}`);
  
  return {
    success: true,
    message: `Data archived for ${academicYear}`
  };
}


/**
 * Export students to a new Google Sheet (Excel-compatible)
 * @param {string} classFilter - Filter by class
 * @param {string} sectionFilter - Filter by section
 * @param {string} streamFilter - Filter by stream
 * @returns {Object} Result with sheet URL
 */
function exportStudentsToSheet(classFilter, sectionFilter, streamFilter) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Filter data
  let filteredData = data.slice(1);
  
  if (classFilter) {
    filteredData = filteredData.filter(row => row[2] == classFilter);
  }
  if (sectionFilter) {
    filteredData = filteredData.filter(row => row[3] === sectionFilter);
  }
  if (streamFilter) {
    filteredData = filteredData.filter(row => row[4] === streamFilter);
  }
  
  if (filteredData.length === 0) {
    return { success: false, message: "No students found with the selected filters." };
  }
  
  // Create new spreadsheet
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const filterDesc = [classFilter ? `Class${classFilter}` : '', sectionFilter || '', streamFilter || ''].filter(Boolean).join('_') || 'All';
  const newSS = SpreadsheetApp.create(`MVM_Students_${filterDesc}_${timestamp}`);
  const newSheet = newSS.getActiveSheet();
  
  // Write headers and data
  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  newSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1a6b3a").setFontColor("white");
  
  if (filteredData.length > 0) {
    newSheet.getRange(2, 1, filteredData.length, headers.length).setValues(filteredData);
  }
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    newSheet.autoResizeColumn(i);
  }
  
  logAction("Export Students", `Exported ${filteredData.length} students to Excel`);
  
  return {
    success: true,
    url: newSS.getUrl(),
    fileName: newSS.getName(),
    message: `Exported ${filteredData.length} students`
  };
}


/**
 * Export teachers to a new Google Sheet (Excel-compatible)
 * @returns {Object} Result with sheet URL
 */
function exportTeachersToSheet() {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const teacherData = data.slice(1).filter(row => row[0]); // Filter empty rows
  
  if (teacherData.length === 0) {
    return { success: false, message: "No teachers found." };
  }
  
  // Create new spreadsheet
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const newSS = SpreadsheetApp.create(`MVM_Teachers_${timestamp}`);
  const newSheet = newSS.getActiveSheet();
  
  // Write headers and data
  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  newSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1a6b3a").setFontColor("white");
  newSheet.getRange(2, 1, teacherData.length, headers.length).setValues(teacherData);
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    newSheet.autoResizeColumn(i);
  }
  
  logAction("Export Teachers", `Exported ${teacherData.length} teachers to Excel`);
  
  return {
    success: true,
    url: newSS.getUrl(),
    fileName: newSS.getName(),
    message: `Exported ${teacherData.length} teachers`
  };
}
