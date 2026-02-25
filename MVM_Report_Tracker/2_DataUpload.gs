/************************************************
 MVM REPORT TRACKER - DATA UPLOAD & MANAGEMENT
 File 2 of 7
************************************************/

/**
 * Upload/Replace students data
 * @param {Array} data - 2D array of student data
 * @returns {Object} Result object
 */
function replaceStudents(data) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!data || !Array.isArray(data) || data.length === 0) {
    return { success: false, message: "Invalid data provided." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const headers = ["StudentID", "Name", "Class", "Section", "Stream", "RollNo", "ParentEmail", "Phone", "JoinDate", "Status"];
  
  // Clear existing data (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  // Process and validate data
  const processedData = data.map((row, index) => {
    const studentId = row[0] || `STU${Date.now()}${index}`;
    return [
      studentId,
      row[1] || "",           // Name
      row[2] || "",           // Class
      row[3] || "A",          // Section
      row[4] || "Science",    // Stream
      row[5] || index + 1,    // RollNo
      row[6] || "",           // ParentEmail
      row[7] || "",           // Phone
      row[8] || new Date(),   // JoinDate
      row[9] || "Active"      // Status
    ];
  });
  
  sheet.getRange(2, 1, processedData.length, 10).setValues(processedData);
  
  logAction("Students Upload", `${processedData.length} students uploaded`);
  
  return { 
    success: true, 
    message: `${processedData.length} students uploaded successfully!`,
    count: processedData.length
  };
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
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const studentId = `STU${Date.now()}`;
  
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
    "Active"
  ]);
  
  logAction("Add Student", `Added student: ${student.name}`);
  
  return { success: true, message: "Student added successfully!", studentId: studentId };
}


/**
 * Update student
 * @param {string} studentId - Student ID to update
 * @param {Object} updates - Fields to update
 * @returns {Object} Result object
 */
function updateStudent(studentId, updates) {
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
    updates.status || row[9]
  ];
  
  sheet.getRange(rowIndex + 1, 1, 1, 10).setValues([updatedRow]);
  
  logAction("Update Student", `Updated student: ${studentId}`);
  
  return { success: true, message: "Student updated successfully!" };
}


/**
 * Delete student (soft delete - set status to Inactive)
 * @param {string} studentId - Student ID to delete
 * @returns {Object} Result object
 */
function deleteStudent(studentId) {
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
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  let students = data.slice(1).map(row => ({
    studentId: row[0],
    name: row[1],
    class: row[2],
    section: row[3],
    stream: row[4],
    rollNo: row[5],
    parentEmail: row[6],
    phone: row[7],
    joinDate: row[8],
    status: row[9]
  }));
  
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
  }
  
  // Apply teacher assignment filter (server-side)
  students = applyTeacherFilter(students, { filterBySubject: false });
  
  return students;
}


/**
 * Upload/Replace teachers data
 * @param {Array} data - 2D array of teacher data
 * @returns {Object} Result object
 */
function replaceTeachers(data) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!data || !Array.isArray(data) || data.length === 0) {
    return { success: false, message: "Invalid data provided." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Teachers");
  
  // Clear existing data (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  // Process and validate data
  const processedData = data.map((row, index) => {
    const teacherId = row[0] || `TCH${Date.now()}${index}`;
    return [
      teacherId,
      row[1] || "",           // Name
      row[2] || "",           // Subject
      row[3] || "",           // Classes
      row[4] || "",           // Sections
      row[5] || "",           // Email
      row[6] || "",           // Phone
      row[7] || new Date(),   // JoinDate
      row[8] || "Active"      // Status
    ];
  });
  
  sheet.getRange(2, 1, processedData.length, 9).setValues(processedData);
  
  logAction("Teachers Upload", `${processedData.length} teachers uploaded`);
  
  return { 
    success: true, 
    message: `${processedData.length} teachers uploaded successfully!`,
    count: processedData.length
  };
}


/**
 * Add single teacher
 * @param {Object} teacher - Teacher data object
 * @returns {Object} Result object
 */
function addTeacher(teacher) {
  if (!teacher || !teacher.name || !teacher.email) {
    return { success: false, message: "Name and Email are required." };
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
  
  logAction("Add Teacher", `Added teacher: ${teacher.name}`);
  
  return { success: true, message: "Teacher added successfully!", teacherId: teacherId };
}


/**
 * Get all teachers with optional filters
 * @param {Object} filters - Optional filters
 * @returns {Array} Filtered teachers
 */
function getTeachers(filters) {
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
