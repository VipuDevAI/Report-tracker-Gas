/************************************************
 MVM REPORT TRACKER - REPORTS & EXPORTS
 File 7 of 7
************************************************/

/**
 * Generate subject report
 * @param {string} subject - Subject name
 * @param {string} examId - Optional exam filter
 * @returns {Object} Subject report data
 */
function generateSubjectReport(subject, examId) {
  const filters = { subject: subject };
  if (examId) filters.examId = examId;
  
  const marks = getMarks(filters);
  const students = getStudents();
  
  // Group by class
  const classWise = {};
  marks.forEach(m => {
    const key = `${m.class}-${m.section}`;
    if (!classWise[key]) {
      classWise[key] = {
        class: m.class,
        section: m.section,
        students: [],
        totalPercentage: 0,
        count: 0
      };
    }
    classWise[key].students.push({
      studentId: m.studentId,
      studentName: m.studentName,
      marksObtained: m.marksObtained,
      maxMarks: m.maxMarks,
      percentage: m.percentage,
      grade: m.grade,
      examName: m.examName
    });
    classWise[key].totalPercentage += m.percentage;
    classWise[key].count++;
  });
  
  // Calculate class averages
  Object.keys(classWise).forEach(key => {
    classWise[key].avgPercentage = (classWise[key].totalPercentage / classWise[key].count).toFixed(2);
    classWise[key].students.sort((a, b) => b.percentage - a.percentage);
  });
  
  return {
    subject: subject,
    examId: examId,
    generatedAt: new Date(),
    totalStudents: marks.length,
    overallAverage: marks.length > 0 
      ? (marks.reduce((sum, m) => sum + m.percentage, 0) / marks.length).toFixed(2) 
      : 0,
    classWise: classWise
  };
}


/**
 * Generate class report
 * @param {string} classNum - Class number
 * @param {string} section - Section
 * @param {string} examId - Optional exam filter
 * @returns {Object} Class report data
 */
function generateClassReport(classNum, section, examId) {
  const filters = { class: classNum };
  if (section) filters.section = section;
  if (examId) filters.examId = examId;
  
  const marks = getMarks(filters);
  const students = getStudents(filters);
  
  // Group by student
  const studentWise = {};
  marks.forEach(m => {
    if (!studentWise[m.studentId]) {
      studentWise[m.studentId] = {
        studentId: m.studentId,
        studentName: m.studentName,
        section: m.section,
        subjects: [],
        totalMarks: 0,
        totalMax: 0
      };
    }
    studentWise[m.studentId].subjects.push({
      subject: m.subject,
      marksObtained: m.marksObtained,
      maxMarks: m.maxMarks,
      percentage: m.percentage,
      grade: m.grade,
      examName: m.examName
    });
    studentWise[m.studentId].totalMarks += m.marksObtained;
    studentWise[m.studentId].totalMax += m.maxMarks;
  });
  
  // Calculate percentages and rank
  const rankedStudents = Object.values(studentWise).map(s => ({
    ...s,
    percentage: s.totalMax > 0 ? ((s.totalMarks / s.totalMax) * 100).toFixed(2) : 0,
    grade: calculateGrade(s.totalMax > 0 ? (s.totalMarks / s.totalMax) * 100 : 0)
  })).sort((a, b) => parseFloat(b.percentage) - parseFloat(a.percentage));
  
  // Assign ranks
  rankedStudents.forEach((s, index) => {
    s.rank = index + 1;
  });
  
  // Subject-wise analysis
  const subjectAnalysis = {};
  marks.forEach(m => {
    if (!subjectAnalysis[m.subject]) {
      subjectAnalysis[m.subject] = {
        subject: m.subject,
        totalPercentage: 0,
        count: 0,
        topScore: 0,
        lowestScore: 100
      };
    }
    subjectAnalysis[m.subject].totalPercentage += m.percentage;
    subjectAnalysis[m.subject].count++;
    subjectAnalysis[m.subject].topScore = Math.max(subjectAnalysis[m.subject].topScore, m.percentage);
    subjectAnalysis[m.subject].lowestScore = Math.min(subjectAnalysis[m.subject].lowestScore, m.percentage);
  });
  
  Object.keys(subjectAnalysis).forEach(key => {
    subjectAnalysis[key].avgPercentage = (subjectAnalysis[key].totalPercentage / subjectAnalysis[key].count).toFixed(2);
  });
  
  return {
    class: classNum,
    section: section || "All",
    examId: examId,
    generatedAt: new Date(),
    totalStudents: rankedStudents.length,
    classAverage: rankedStudents.length > 0 
      ? (rankedStudents.reduce((sum, s) => sum + parseFloat(s.percentage), 0) / rankedStudents.length).toFixed(2) 
      : 0,
    students: rankedStudents,
    subjectAnalysis: Object.values(subjectAnalysis)
  };
}


/**
 * Generate student report card
 * @param {string} studentId - Student ID
 * @param {string} examId - Optional exam filter
 * @returns {Object} Student report card data
 */
function generateStudentReport(studentId, examId) {
  const students = getStudents();
  const student = students.find(s => s.studentId === studentId);
  
  if (!student) {
    return { success: false, message: "Student not found." };
  }
  
  const filters = { studentId: studentId };
  if (examId) filters.examId = examId;
  
  const marks = getMarks(filters);
  
  // Group by exam
  const examWise = {};
  marks.forEach(m => {
    if (!examWise[m.examId]) {
      examWise[m.examId] = {
        examId: m.examId,
        examName: m.examName,
        subjects: [],
        totalMarks: 0,
        totalMax: 0
      };
    }
    examWise[m.examId].subjects.push({
      subject: m.subject,
      marksObtained: m.marksObtained,
      maxMarks: m.maxMarks,
      percentage: m.percentage,
      grade: m.grade
    });
    examWise[m.examId].totalMarks += m.marksObtained;
    examWise[m.examId].totalMax += m.maxMarks;
  });
  
  // Calculate exam totals
  Object.keys(examWise).forEach(key => {
    const exam = examWise[key];
    exam.percentage = exam.totalMax > 0 ? ((exam.totalMarks / exam.totalMax) * 100).toFixed(2) : 0;
    exam.grade = calculateGrade(parseFloat(exam.percentage));
  });
  
  // Overall calculations
  const overallTotal = marks.reduce((sum, m) => sum + m.marksObtained, 0);
  const overallMax = marks.reduce((sum, m) => sum + m.maxMarks, 0);
  const overallPercentage = overallMax > 0 ? ((overallTotal / overallMax) * 100).toFixed(2) : 0;
  
  return {
    success: true,
    student: {
      studentId: student.studentId,
      name: student.name,
      class: student.class,
      section: student.section,
      stream: student.stream,
      rollNo: student.rollNo
    },
    generatedAt: new Date(),
    examWise: Object.values(examWise),
    overall: {
      totalMarks: overallTotal,
      maxMarks: overallMax,
      percentage: overallPercentage,
      grade: calculateGrade(parseFloat(overallPercentage))
    }
  };
}


/**
 * Export marks data to CSV format
 * @param {Object} filters - Optional filters
 * @returns {string} CSV string
 */
function exportMarksToCSV(filters) {
  const marks = getMarks(filters);
  
  const headers = [
    "Student ID", "Student Name", "Class", "Section", "Subject", 
    "Exam Name", "Max Marks", "Marks Obtained", "Percentage", "Grade", 
    "Teacher", "Updated At"
  ];
  
  const rows = marks.map(m => [
    m.studentId,
    m.studentName,
    m.class,
    m.section,
    m.subject,
    m.examName,
    m.maxMarks,
    m.marksObtained,
    m.percentage,
    m.grade,
    m.teacherName,
    m.updatedAt
  ]);
  
  const csvContent = [headers, ...rows]
    .map(row => row.map(cell => `"${cell}"`).join(","))
    .join("\n");
  
  return csvContent;
}


/**
 * Download marks as spreadsheet
 * @param {Object} filters - Optional filters
 * @returns {Object} Result with download URL
 */
function downloadMarksReport(filters) {
  const marks = getMarks(filters);
  
  // Create new spreadsheet
  const ss = SpreadsheetApp.create("MVM_Marks_Report_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
  const sheet = ss.getActiveSheet();
  
  const headers = [
    "Student ID", "Student Name", "Class", "Section", "Subject", 
    "Exam Name", "Max Marks", "Marks Obtained", "Percentage", "Grade", 
    "Teacher", "Updated At"
  ];
  
  const data = marks.map(m => [
    m.studentId,
    m.studentName,
    m.class,
    m.section,
    m.subject,
    m.examName,
    m.maxMarks,
    m.marksObtained,
    m.percentage,
    m.grade,
    m.teacherName,
    m.updatedAt
  ]);
  
  sheet.appendRow(headers);
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  
  // Format header
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1a6b3a").setFontColor("#ffffff");
  
  logAction("Download Report", `Exported ${marks.length} marks entries`);
  
  return {
    success: true,
    url: ss.getUrl(),
    message: `Report created with ${marks.length} entries.`
  };
}


/**
 * Get alerts
 * @param {Object} filters - Optional filters (type, isRead, priority)
 * @returns {Array} Alerts list
 */
function getAlerts(filters) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Alerts");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  let alerts = data.slice(1).map(row => ({
    alertId: row[0],
    alertType: row[1],
    studentId: row[2],
    studentName: row[3],
    class: row[4],
    subject: row[5],
    message: row[6],
    priority: row[7],
    isRead: row[8],
    createdAt: row[9]
  }));
  
  if (filters) {
    if (filters.alertType) {
      alerts = alerts.filter(a => a.alertType === filters.alertType);
    }
    if (filters.isRead !== undefined) {
      alerts = alerts.filter(a => a.isRead === filters.isRead);
    }
    if (filters.priority) {
      alerts = alerts.filter(a => a.priority === filters.priority);
    }
  }
  
  return alerts.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}


/**
 * Mark alert as read
 * @param {string} alertId - Alert ID
 * @returns {Object} Result object
 */
function markAlertAsRead(alertId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Alerts");
  const data = sheet.getDataRange().getValues();
  
  const rowIndex = data.findIndex(r => r[0] === alertId);
  
  if (rowIndex === -1) {
    return { success: false, message: "Alert not found." };
  }
  
  sheet.getRange(rowIndex + 1, 9).setValue(true);
  
  return { success: true, message: "Alert marked as read." };
}


/**
 * Get activity logs
 * @param {number} limit - Number of logs to return
 * @returns {Array} Activity logs
 */
function getActivityLogs(limit) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Logs");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  const logs = data.slice(1).map(row => ({
    logId: row[0],
    action: row[1],
    user: row[2],
    details: row[3],
    timestamp: row[4]
  })).reverse();
  
  return logs.slice(0, limit || 50);
}


/**
 * Get school settings
 * @returns {Object} School settings
 */
function getSchoolSettings() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return {};
  
  const settings = {};
  data.slice(1).forEach(row => {
    settings[row[0]] = row[1];
  });
  
  return settings;
}


/**
 * Update school setting
 * @param {string} key - Setting key
 * @param {string} value - Setting value
 * @returns {Object} Result object
 */
function updateSchoolSetting(key, value) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied." };
  }
  
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_School");
  const data = sheet.getDataRange().getValues();
  
  const rowIndex = data.findIndex(r => r[0] === key);
  
  if (rowIndex === -1) {
    sheet.appendRow([key, value, new Date()]);
  } else {
    sheet.getRange(rowIndex + 1, 2, 1, 2).setValues([[value, new Date()]]);
  }
  
  logAction("Update Setting", `${key} = ${value}`);
  
  return { success: true, message: "Setting updated successfully!" };
}


/**
 * Get grade ranges
 * @returns {Array} Grade ranges
 */
function getGradeRanges() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Settings_Ranges");
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  return data.slice(1).map(row => ({
    rangeName: row[0],
    gradeLabel: row[1],
    minMarks: row[2],
    maxMarks: row[3],
    color: row[4]
  }));
}
