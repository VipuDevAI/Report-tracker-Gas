/************************************************
 MVM REPORT TRACKER - REPORTS & EXPORTS
 File 7 of 7
 With Report Card Generation (Admin Only)
************************************************/

/**
 * Generate student report card (Admin only)
 * @param {string} studentId - Student ID
 * @param {string} examId - Optional exam filter
 * @returns {Object} Report card data
 */
function generateStudentReport(studentId, examId) {
  const students = getStudents({ status: "Active" });
  const student = students.find(s => s.studentId === studentId);
  
  if (!student) {
    return { success: false, message: "Student not found." };
  }
  
  const filters = { studentId: studentId };
  if (examId) filters.examId = examId;
  
  // For report card, get all marks (bypass teacher filter for admin)
  const marksSheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const marksData = marksSheet.getDataRange().getValues();
  const currentYear = getCurrentAcademicYear();
  
  let marks = marksData.slice(1)
    .filter(row => row[1] === studentId && (row[17] || currentYear) === currentYear)
    .map(row => ({
      subject: row[3],
      examId: row[7],
      examName: row[8],
      maxMarks: row[11],
      marksObtained: row[12],
      percentage: parseFloat(row[13]),
      grade: row[14],
      teacherName: row[6]
    }));
  
  if (examId) {
    marks = marks.filter(m => m.examId === examId);
  }
  
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
      grade: m.grade,
      teacherName: m.teacherName
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
  
  // Get school info
  const schoolSettings = getSchoolSettings();
  
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
    school: {
      name: schoolSettings.SchoolName || "MVM School",
      academicYear: currentYear
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
 * Generate report cards for entire class (Admin only)
 * @param {string} classNum - Class number
 * @param {string} section - Section (optional)
 * @param {string} examId - Exam ID (optional)
 * @returns {Object} Report cards data
 */
function generateClassReportCards(classNum, section, examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to generate report cards." };
  }
  
  // Get students for the class
  const studentsSheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const studentsData = studentsSheet.getDataRange().getValues();
  
  let students = studentsData.slice(1)
    .filter(row => row[2] == classNum && row[9] === "Active")
    .map(row => ({
      studentId: row[0],
      name: row[1],
      class: row[2],
      section: row[3],
      stream: row[4],
      rollNo: row[5]
    }));
  
  if (section) {
    students = students.filter(s => s.section === section);
  }
  
  // Generate report for each student
  const reports = [];
  const errors = [];
  
  students.forEach(student => {
    const report = generateStudentReport(student.studentId, examId);
    if (report.success) {
      reports.push(report);
    } else {
      errors.push({ studentId: student.studentId, error: report.message });
    }
  });
  
  // Calculate ranks
  const sortedReports = [...reports].sort((a, b) => 
    parseFloat(b.overall.percentage) - parseFloat(a.overall.percentage)
  );
  
  sortedReports.forEach((report, idx) => {
    report.rank = idx + 1;
    report.totalStudents = sortedReports.length;
  });
  
  return {
    success: true,
    class: classNum,
    section: section || "All",
    examId: examId,
    totalStudents: students.length,
    reportsGenerated: reports.length,
    reports: sortedReports,
    errors: errors
  };
}


/**
 * Generate PDF report card HTML template
 * @param {Object} reportData - Report card data
 * @returns {string} HTML content
 */
function generateReportCardHTML(reportData) {
  const student = reportData.student;
  const school = reportData.school;
  const overall = reportData.overall;
  
  let subjectsHTML = "";
  
  reportData.examWise.forEach(exam => {
    subjectsHTML += `
      <tr class="exam-header">
        <td colspan="5" style="background: #f0f0f0; font-weight: bold;">${exam.examName}</td>
      </tr>
    `;
    exam.subjects.forEach(sub => {
      subjectsHTML += `
        <tr>
          <td>${sub.subject}</td>
          <td style="text-align: center;">${sub.maxMarks}</td>
          <td style="text-align: center;">${sub.marksObtained}</td>
          <td style="text-align: center;">${sub.percentage.toFixed(1)}%</td>
          <td style="text-align: center;">${sub.grade}</td>
        </tr>
      `;
    });
    subjectsHTML += `
      <tr class="exam-total">
        <td style="font-weight: bold;">Exam Total</td>
        <td style="text-align: center; font-weight: bold;">${exam.totalMax}</td>
        <td style="text-align: center; font-weight: bold;">${exam.totalMarks}</td>
        <td style="text-align: center; font-weight: bold;">${exam.percentage}%</td>
        <td style="text-align: center; font-weight: bold;">${exam.grade}</td>
      </tr>
    `;
  });
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Times New Roman', serif; padding: 20px; max-width: 800px; margin: 0 auto; }
        .header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #1a6b3a; padding-bottom: 15px; }
        .logo { width: 80px; height: 80px; margin-bottom: 10px; }
        .school-name { font-size: 24px; font-weight: bold; color: #1a6b3a; }
        .motto { font-style: italic; color: #666; font-size: 12px; }
        .report-title { font-size: 18px; font-weight: bold; margin: 15px 0; text-transform: uppercase; background: #1a6b3a; color: white; padding: 8px; }
        .student-info { display: flex; justify-content: space-between; margin-bottom: 20px; }
        .info-group { }
        .info-label { font-weight: bold; color: #333; }
        .info-value { margin-left: 5px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #333; padding: 8px; text-align: left; }
        th { background: #1a6b3a; color: white; }
        .overall { background: #f5f5f5; }
        .overall td { font-weight: bold; font-size: 14px; }
        .grade-box { text-align: center; margin: 20px 0; }
        .grade-value { font-size: 48px; font-weight: bold; color: #1a6b3a; }
        .rank-info { text-align: center; margin: 10px 0; font-size: 14px; }
        .footer { margin-top: 30px; display: flex; justify-content: space-between; }
        .signature { text-align: center; }
        .signature-line { border-top: 1px solid #333; width: 150px; margin-top: 40px; }
        .print-date { text-align: right; font-size: 10px; color: #666; margin-top: 20px; }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="school-name">${school.name}</div>
        <div class="motto">"Knowledge is Structured in Consciousness"</div>
      </div>
      
      <div class="report-title">Progress Report Card - ${school.academicYear}</div>
      
      <div class="student-info">
        <div class="info-group">
          <span class="info-label">Name:</span>
          <span class="info-value">${student.name}</span>
        </div>
        <div class="info-group">
          <span class="info-label">Class:</span>
          <span class="info-value">${student.class} - ${student.section}</span>
        </div>
        <div class="info-group">
          <span class="info-label">Roll No:</span>
          <span class="info-value">${student.rollNo}</span>
        </div>
        <div class="info-group">
          <span class="info-label">Stream:</span>
          <span class="info-value">${student.stream}</span>
        </div>
      </div>
      
      <table>
        <thead>
          <tr>
            <th>Subject</th>
            <th style="text-align: center; width: 80px;">Max Marks</th>
            <th style="text-align: center; width: 80px;">Obtained</th>
            <th style="text-align: center; width: 80px;">Percentage</th>
            <th style="text-align: center; width: 60px;">Grade</th>
          </tr>
        </thead>
        <tbody>
          ${subjectsHTML}
          <tr class="overall">
            <td>GRAND TOTAL</td>
            <td style="text-align: center;">${overall.maxMarks}</td>
            <td style="text-align: center;">${overall.totalMarks}</td>
            <td style="text-align: center;">${overall.percentage}%</td>
            <td style="text-align: center;">${overall.grade}</td>
          </tr>
        </tbody>
      </table>
      
      <div class="grade-box">
        <div style="font-size: 14px;">Overall Grade</div>
        <div class="grade-value">${overall.grade}</div>
        <div style="font-size: 12px;">${overall.percentage}%</div>
      </div>
      
      ${reportData.rank ? `
        <div class="rank-info">
          <strong>Class Rank:</strong> ${reportData.rank} out of ${reportData.totalStudents} students
        </div>
      ` : ''}
      
      <div class="footer">
        <div class="signature">
          <div class="signature-line"></div>
          <div>Class Teacher</div>
        </div>
        <div class="signature">
          <div class="signature-line"></div>
          <div>Principal</div>
        </div>
        <div class="signature">
          <div class="signature-line"></div>
          <div>Parent's Signature</div>
        </div>
      </div>
      
      <div class="print-date">Generated on: ${new Date().toLocaleDateString()}</div>
    </body>
    </html>
  `;
}


/**
 * Generate PDF report card for a student (Admin only)
 * @param {string} studentId - Student ID
 * @param {string} examId - Optional exam filter
 * @returns {Object} Result with PDF blob
 */
function generateReportCardPDF(studentId, examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to generate report cards." };
  }
  
  const reportData = generateStudentReport(studentId, examId);
  if (!reportData.success) {
    return reportData;
  }
  
  const html = generateReportCardHTML(reportData);
  
  // Create PDF
  const blob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .setName(`ReportCard_${reportData.student.name}_${reportData.student.class}${reportData.student.section}.pdf`)
    .getAs('application/pdf');
  
  // Save to Drive
  const folder = getOrCreateReportFolder();
  const file = folder.createFile(blob);
  
  return {
    success: true,
    message: "Report card generated successfully!",
    fileUrl: file.getUrl(),
    fileName: file.getName()
  };
}


/**
 * Generate PDF report cards for entire class (Admin only)
 * @param {string} classNum - Class number
 * @param {string} section - Section (optional)
 * @param {string} examId - Exam ID (optional)
 * @returns {Object} Result with ZIP download URL
 */
function generateClassReportCardsPDF(classNum, section, examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required to generate report cards." };
  }
  
  const classData = generateClassReportCards(classNum, section, examId);
  if (!classData.success) {
    return classData;
  }
  
  const folder = getOrCreateReportFolder();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const subFolder = folder.createFolder(`ReportCards_Class${classNum}${section || ''}_${timestamp}`);
  
  const generatedFiles = [];
  
  classData.reports.forEach(report => {
    try {
      const html = generateReportCardHTML(report);
      const blob = HtmlService.createHtmlOutput(html)
        .getBlob()
        .setName(`${report.rank}_${report.student.name}_${report.student.rollNo}.pdf`)
        .getAs('application/pdf');
      
      const file = subFolder.createFile(blob);
      generatedFiles.push({
        studentId: report.student.studentId,
        name: report.student.name,
        fileName: file.getName(),
        url: file.getUrl()
      });
    } catch (e) {
      classData.errors.push({ studentId: report.student.studentId, error: e.message });
    }
  });
  
  logAction("Generate Class Report Cards", `Class ${classNum}${section || ''}: ${generatedFiles.length} reports generated`);
  
  return {
    success: true,
    message: `Generated ${generatedFiles.length} report cards`,
    folderUrl: subFolder.getUrl(),
    files: generatedFiles,
    errors: classData.errors
  };
}


/**
 * Get or create reports folder in Drive
 * @returns {Folder} Reports folder
 */
function getOrCreateReportFolder() {
  const folderName = "MVM_Report_Cards";
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  }
  
  return DriveApp.createFolder(folderName);
}


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


/**
 * Check if user can generate report cards
 * @returns {boolean} True if admin
 */
function canGenerateReportCards() {
  return isAdmin();
}
