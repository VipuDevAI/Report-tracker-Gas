/************************************************
 MVM REPORT TRACKER - DATA UPLOAD & MANAGEMENT
 File 2 of 7
 With Bulk Upload Features
************************************************/

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
    if (!previewOnly) lock.waitLock(30000);
    
    const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 12);
    const existingData = lastRow > 0
      ? sheet.getRange(1, 1, lastRow, lastCol).getValues()
      : [[]];
    const academicYear = getCurrentAcademicYear();
    
    // Build index of existing students (class-section-rollno as key)
    const existingIndex = {};
    const existingIdIndex = {};
    existingData.slice(1).forEach((row, idx) => {
      const key = `${row[2]}-${row[3]}-${row[5]}`; // class-section-rollno
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
      
      const studentId = row[0] || `STU${Date.now()}${rowIdx}`;
      const name = row[1] || "";
      const cls = String(row[2] || "");
      const section = row[3] || "A";
      const stream = row[4] || "Science";
      const rollNo = row[5] || rowIdx;
      const status = row[6] || "Active";
      const electiveSubject = row[7] || "";
      
      // Validation
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
      const validElectives = ['Mathematics', 'Applied Mathematics', 'Hindi', 'History', 'Sanskrit', ''];
      if ((cls == '11' || cls == '12') && electiveSubject && !validElectives.includes(electiveSubject)) {
        results.failed++;
        results.errors.push({ row: rowIdx + 1, error: `Invalid elective: ${electiveSubject}. Use: Mathematics, Applied Mathematics, Hindi, History, or Sanskrit` });
        return;
      }
      
      // Check for duplicates (same class + section + roll no)
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
        "",  // parentEmail
        "",  // phone
        new Date(),
        status,
        electiveSubject,
        academicYear
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
            row: rowIdx + 1,
            name: name,
            class: cls,
            section: section,
            rollNo: rollNo
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
    
    // If preview only, return without writing
    if (previewOnly) {
      return {
        success: true,
        preview: true,
        results: results,
        message: `Preview: ${results.created} new, ${results.updated} updates, ${results.failed} failed`
      };
    }
    
    // Write new students (batch)
    if (toCreate.length > 0) {
      const writeRow = sheet.getLastRow() + 1;
      sheet.getRange(writeRow, 1, toCreate.length, 12).setValues(toCreate);
    }
    
    // Update existing students (batch by individual row)
    toUpdate.forEach(item => {
      sheet.getRange(item.rowIndex, 1, 1, 12).setValues([item.data]);
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
    const classes = row[3] || "";
    const sections = row[4] || "";
    const email = row[5] || "";
    const phone = row[6] || "";
    const status = row[7] || "Active";
    
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
      status
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
  
  // Write new teachers
  if (toCreate.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, toCreate.length, 9).setValues(toCreate);
  }
  
  // Update existing teachers
  toUpdate.forEach(item => {
    sheet.getRange(item.rowIndex, 1, 1, 9).setValues([item.data]);
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
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!student || !student.name || !student.class) {
    return { success: false, message: "Name and Class are required." };
  }
  
  // Validate elective subject for class 11 & 12
  const validElectives = ['Mathematics', 'Applied Mathematics', 'Hindi', 'History', 'Sanskrit'];
  if ((student.class == 11 || student.class == 12) && student.electiveSubject) {
    if (!validElectives.includes(student.electiveSubject)) {
      return { success: false, message: "Invalid elective subject. Choose from: " + validElectives.join(", ") };
    }
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const studentId = `STU${Date.now()}`;
  const academicYear = getCurrentAcademicYear();
  
  sheet.appendRow([
    studentId,
    student.name,
    student.class,
    student.section || "A",
    student.stream || "Science",
    student.rollNo || sheet.getLastRow(),
    student.parentEmail || "",
    student.phone || "",
    new Date(),
    "Active",
    student.electiveSubject || "",
    academicYear
  ]);
  
  logAction("Add Student", `Added student: ${student.name} (Elective: ${student.electiveSubject || 'N/A'})`);
  
  return { success: true, message: "Student added successfully!", studentId: studentId };
}


/**
 * Update student
 * @param {string} studentId - Student ID to update
 * @param {Object} updates - Fields to update
 * @returns {Object} Result object
 */
function updateStudent(studentId, updates) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  
  const rowIndex = data.findIndex(r => r[0] === studentId);
  
  if (rowIndex === -1) {
    return { success: false, message: "Student not found." };
  }
  
  const row = data[rowIndex];
  const updatedRow = [
    studentId,
    updates.name || row[1],
    updates.class || row[2],
    updates.section || row[3],
    updates.stream || row[4],
    updates.rollNo || row[5],
    updates.parentEmail || row[6],
    updates.phone || row[7],
    row[8],
    updates.status || row[9],
    updates.electiveSubject !== undefined ? updates.electiveSubject : (row[10] || ""),
    updates.academicYear || row[11] || getCurrentAcademicYear()
  ];
  
  sheet.getRange(rowIndex + 1, 1, 1, 12).setValues([updatedRow]);
  
  logAction("Update Student", `Updated student: ${studentId}`);
  
  return { success: true, message: "Student updated successfully!" };
}


/**
 * Delete student (soft delete - set status to Inactive)
 * @param {string} studentId - Student ID to delete
 * @returns {Object} Result object
 */
function deleteStudent(studentId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  return updateStudent(studentId, { status: "Inactive" });
}


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
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  const currentYear = getCurrentAcademicYear();
  
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
    academicYear: row[11] || currentYear
  })).filter(s => s.studentId);
  
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
      // Default: filter by current academic year (null disables filter)
      students = students.filter(s => s.academicYear === currentYear);
    }
  } else {
    students = students.filter(s => s.status === "Active" && s.academicYear === currentYear);
  }
  
  // Apply teacher assignment filter (server-side)
  students = applyTeacherFilter(students, { filterBySubject: false });
  
  return students;
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
  
  // Check if email already exists
  const existingTeacher = getTeacherByEmail(teacher.email);
  if (existingTeacher) {
    return { success: false, message: "A teacher with this email already exists." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const teacherId = `TCH${Date.now()}`;
  
  sheet.appendRow([
    teacherId,
    teacher.name,
    teacher.subject || "",
    teacher.classes || "",
    teacher.sections || "",
    teacher.email,
    teacher.phone || "",
    new Date(),
    "Active"
  ]);
  
  logAction("Add Teacher", `Added teacher: ${teacher.name} (${teacher.email})`);
  
  return { 
    success: true, 
    message: "Teacher added successfully! They can now login using their Google account.", 
    teacherId: teacherId
  };
}


/**
 * Get all teachers with optional filters
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered teachers
 */
function getTeachers(filters) {
  if (!isAdmin()) {
    // Teachers can only see their own info
    const assignment = getTeacherAssignment();
    if (assignment) {
      return [{
        teacherId: assignment.teacherId,
        name: assignment.name,
        subject: assignment.subject,
        classes: assignment.classes.join(","),
        sections: assignment.sections.join(","),
        email: assignment.email,
        status: "Active"
      }];
    }
    return [];
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  let teachers = data.slice(1).map(row => ({
    teacherId: row[0],
    name: row[1],
    subject: row[2],
    classes: row[3],
    sections: row[4],
    email: row[5],
    phone: row[6],
    joinDate: row[7],
    status: row[8]
  }));
  
  if (filters) {
    if (filters.subject) {
      teachers = teachers.filter(t => t.subject === filters.subject);
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
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  let subjects = data.slice(1).map(row => ({
    subjectId: row[0],
    subjectName: row[1],
    subjectCode: row[2],
    classes: row[3],
    stream: row[4],
    maxMarks: row[5],
    passingMarks: row[6],
    isActive: row[7]
  }));
  
  if (filters) {
    if (filters.class) {
      subjects = subjects.filter(s => s.classes.includes(filters.class));
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
  return ["9", "10", "11", "12"];
}


/**
 * Get sections
 * @returns {Array} Available sections
 */
function getSections() {
  return ["A", "B", "C", "D"];
}


/**
 * Get streams
 * @returns {Array} Available streams
 */
function getStreams() {
  return ["Science", "Computer Science", "Commerce"];
}


/**
 * Auto promote students to next class
 * @param {string} fromYear - Source academic year
 * @param {string} toYear - Target academic year
 * @returns {Object} Result object
 */
function promoteStudents(fromYear, toYear) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  
  let promoted = 0;
  let graduated = 0;
  
  data.forEach((row, idx) => {
    if (idx === 0) return; // Skip header
    
    const currentClass = parseInt(row[2]);
    if (isNaN(currentClass)) return;
    
    if (currentClass >= 12) {
      // Graduate (mark as Alumni)
      sheet.getRange(idx + 1, 10).setValue("Alumni");
      graduated++;
    } else {
      // Promote to next class
      sheet.getRange(idx + 1, 3).setValue(currentClass + 1);
      promoted++;
    }
  });
  
  // Update academic year in settings
  updateSchoolSetting("AcademicYear", toYear);
  
  logAction("Promote Students", `Promoted: ${promoted}, Graduated: ${graduated}`);
  
  return {
    success: true,
    message: `Promotion complete: ${promoted} promoted, ${graduated} graduated`,
    promoted: promoted,
    graduated: graduated
  };
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
