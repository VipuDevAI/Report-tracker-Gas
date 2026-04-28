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
 * Download class-wise marks report for an exam (Admin only)
 * Format: Student | Roll | Subject1 | Subject2 | Subject3... | Total | %
 * @param {string} classNum - Class number
 * @param {string} section - Section
 * @param {string} examId - Exam ID
 * @returns {Object} Result with download URL
 */
function downloadClassExamMarks(classNum, section, examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!classNum || !examId) {
    return { success: false, message: "Class and Exam are required." };
  }
  
  // Get exam details
  const exam = getExamById(examId);
  if (!exam) {
    return { success: false, message: "Exam not found." };
  }
  
  // Get students for the class
  const studentsSheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const studentsData = studentsSheet.getDataRange().getValues();
  
  let students = studentsData.slice(1)
    .filter(row => row[2] == classNum && row[9] === "Active")
    .map(row => ({
      studentId: row[0],
      name: row[1],
      rollNo: row[5],
      section: row[3]
    }));
  
  if (section) {
    students = students.filter(s => s.section === section);
  }
  
  // Sort by roll number
  students.sort((a, b) => a.rollNo - b.rollNo);
  
  // Get marks for this exam
  const marksSheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const marksData = marksSheet.getDataRange().getValues();
  
  const examMarks = marksData.slice(1)
    .filter(row => row[7] === examId && row[9] == classNum)
    .filter(row => !section || row[10] === section);
  
  // Get unique subjects
  const subjects = [...new Set(examMarks.map(m => m[3]))].sort();
  
  if (subjects.length === 0) {
    return { success: false, message: "No marks found for this exam and class." };
  }
  
  // Build marks index: studentId -> { subject: marks }
  const marksIndex = {};
  examMarks.forEach(row => {
    const studentId = row[1];
    const subject = row[3];
    const marks = row[12];
    const maxMarks = row[11];
    
    if (!marksIndex[studentId]) {
      marksIndex[studentId] = {};
    }
    marksIndex[studentId][subject] = { marks: marks, max: maxMarks };
  });
  
  // Create spreadsheet
  const ss = SpreadsheetApp.create(`${exam.name}_Class${classNum}${section || ''}_Marks_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd")}`);
  const sheet = ss.getActiveSheet();
  sheet.setName("Marks Report");
  
  // Headers
  const headers = ["Roll No", "Student Name", ...subjects, "Total", "Max", "Percentage", "Range"];
  sheet.appendRow(headers);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1a6b3a");
  headerRange.setFontColor("#ffffff");
  
  // Data rows
  const dataRows = [];
  students.forEach(student => {
    const studentMarks = marksIndex[student.studentId] || {};
    let total = 0;
    let maxTotal = 0;
    
    const row = [student.rollNo, student.name];
    
    subjects.forEach(subject => {
      if (studentMarks[subject]) {
        row.push(studentMarks[subject].marks);
        total += studentMarks[subject].marks;
        maxTotal += studentMarks[subject].max;
      } else {
        row.push("-");
      }
    });
    
    const percentage = maxTotal > 0 ? ((total / maxTotal) * 100).toFixed(2) : 0;
    const range = calculateGrade(parseFloat(percentage));
    
    row.push(total > 0 ? total : "-");
    row.push(maxTotal > 0 ? maxTotal : "-");
    row.push(total > 0 ? percentage + "%" : "-");
    row.push(total > 0 ? range : "-");
    
    dataRows.push(row);
  });
  
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }
  
  // Add summary row
  const summaryRow = sheet.getLastRow() + 2;
  sheet.getRange(summaryRow, 1).setValue("Class Summary");
  sheet.getRange(summaryRow, 1).setFontWeight("bold");
  
  // Calculate class averages per subject
  const avgRow = ["", "Class Average"];
  subjects.forEach((subject, idx) => {
    const subjectMarks = dataRows
      .map(r => r[2 + idx])
      .filter(m => m !== "-");
    const avg = subjectMarks.length > 0 
      ? (subjectMarks.reduce((a, b) => a + b, 0) / subjectMarks.length).toFixed(1)
      : "-";
    avgRow.push(avg);
  });
  avgRow.push("", "", "", "");
  sheet.getRange(summaryRow + 1, 1, 1, avgRow.length).setValues([avgRow]);
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Add exam info
  const infoRow = summaryRow + 3;
  sheet.getRange(infoRow, 1).setValue("Exam: " + exam.name);
  sheet.getRange(infoRow + 1, 1).setValue("Class: " + classNum + (section ? "-" + section : " (All Sections)"));
  sheet.getRange(infoRow + 2, 1).setValue("Max Marks per Subject: " + exam.maxMarks);
  sheet.getRange(infoRow + 3, 1).setValue("Generated: " + new Date().toLocaleString());
  
  logAction("Download Class Marks", `${exam.name} - Class ${classNum}${section || ''}`);
  
  return {
    success: true,
    url: ss.getUrl(),
    fileName: ss.getName(),
    message: `Report generated for ${students.length} students, ${subjects.length} subjects`
  };
}


/**
 * Download class-wise marks as CSV (Admin only)
 * Returns CSV string for client-side download
 * @param {string} classNum - Class number
 * @param {string} section - Section
 * @param {string} examId - Exam ID
 * @returns {Object} Result with CSV data
 */
function downloadClassExamMarksCSV(classNum, section, examId) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin privileges required." };
  }
  
  if (!classNum || !examId) {
    return { success: false, message: "Class and Exam are required." };
  }
  
  // Get exam details
  const exam = getExamById(examId);
  if (!exam) {
    return { success: false, message: "Exam not found." };
  }
  
  // Get students for the class
  const studentsSheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const studentsData = studentsSheet.getDataRange().getValues();
  
  let students = studentsData.slice(1)
    .filter(row => row[2] == classNum && row[9] === "Active")
    .map(row => ({
      studentId: row[0],
      name: row[1],
      rollNo: row[5],
      section: row[3]
    }));
  
  if (section) {
    students = students.filter(s => s.section === section);
  }
  
  // Sort by roll number
  students.sort((a, b) => a.rollNo - b.rollNo);
  
  // Get marks for this exam
  const marksSheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  const marksData = marksSheet.getDataRange().getValues();
  
  const examMarks = marksData.slice(1)
    .filter(row => row[7] === examId && row[9] == classNum)
    .filter(row => !section || row[10] === section);
  
  // Get unique subjects
  const subjects = [...new Set(examMarks.map(m => m[3]))].sort();
  
  if (subjects.length === 0) {
    return { success: false, message: "No marks found for this exam and class." };
  }
  
  // Build marks index: studentId -> { subject: marks }
  const marksIndex = {};
  examMarks.forEach(row => {
    const studentId = row[1];
    const subject = row[3];
    const marks = row[12];
    const maxMarks = row[11];
    
    if (!marksIndex[studentId]) {
      marksIndex[studentId] = {};
    }
    marksIndex[studentId][subject] = { marks: marks, max: maxMarks };
  });
  
  // Build CSV
  const headers = ["Roll No", "Student Name", ...subjects, "Total", "Max", "Percentage", "Range"];
  const csvRows = [headers.map(h => `"${h}"`).join(",")];
  
  students.forEach(student => {
    const studentMarks = marksIndex[student.studentId] || {};
    let total = 0;
    let maxTotal = 0;
    
    const row = [`"${student.rollNo}"`, `"${student.name}"`];
    
    subjects.forEach(subject => {
      if (studentMarks[subject]) {
        row.push(studentMarks[subject].marks);
        total += studentMarks[subject].marks;
        maxTotal += studentMarks[subject].max;
      } else {
        row.push('"-"');
      }
    });
    
    const percentage = maxTotal > 0 ? ((total / maxTotal) * 100).toFixed(2) : 0;
    const range = calculateGrade(parseFloat(percentage));
    
    row.push(total > 0 ? total : '"-"');
    row.push(maxTotal > 0 ? maxTotal : '"-"');
    row.push(total > 0 ? `"${percentage}%"` : '"-"');
    row.push(total > 0 ? `"${range}"` : '"-"');
    
    csvRows.push(row.join(","));
  });
  
  // Add summary row
  csvRows.push("");
  csvRows.push('"Class Summary"');
  
  // Calculate class averages per subject
  const avgRow = ['""', '"Class Average"'];
  subjects.forEach(subject => {
    const subjectMarks = students
      .map(s => marksIndex[s.studentId]?.[subject]?.marks)
      .filter(m => m !== undefined);
    const avg = subjectMarks.length > 0 
      ? (subjectMarks.reduce((a, b) => a + b, 0) / subjectMarks.length).toFixed(1)
      : "-";
    avgRow.push(avg === "-" ? '"-"' : avg);
  });
  avgRow.push('""', '""', '""', '""');
  csvRows.push(avgRow.join(","));
  
  // Add exam info
  csvRows.push("");
  csvRows.push(`"Exam: ${exam.name}"`);
  csvRows.push(`"Class: ${classNum}${section ? '-' + section : ' (All Sections)'}"`);
  csvRows.push(`"Max Marks per Subject: ${exam.maxMarks}"`);
  csvRows.push(`"Generated: ${new Date().toLocaleString()}"`);
  
  const csvData = csvRows.join("\n");
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const fileName = `${exam.name}_Class${classNum}${section || ''}_Marks_${timestamp}.csv`;
  
  logAction("Download CSV Marks", `${exam.name} - Class ${classNum}${section || ''}`);
  
  return {
    success: true,
    csvData: csvData,
    fileName: fileName,
    message: `CSV generated for ${students.length} students, ${subjects.length} subjects`
  };
}


/**
 * Get available exams for download dropdown
 * @returns {Array} Exams list
 */
function getExamsForDownload() {
  return getExams();
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
function updateSchoolSettingFromUI(key, value) {
  return adminUpdateSchoolSetting(key, value);
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


/* ========================================================================
   BATCHED PDF REPORT CARD GENERATION
   - Per class/section, chunked processing with resume capability
   - Job state in DocumentProperties (survives across executions)
   - Folder layout: MVM_Report_Cards / {AcademicYear} / Class_{X}{Section}
   ======================================================================== */

const PDF_JOB_PREFIX = "pdf_job_";
const PDF_DEFAULT_CHUNK_SIZE = 25;
const PDF_MAX_CHUNK_RUNTIME_MS = 4 * 60 * 1000; // stop chunk early at 4 min to be safe vs 6-min limit


/**
 * Get the year-scoped report cards root folder, creating if missing.
 * Layout: MVM_Report_Cards / {AcademicYear}
 */
function getYearScopedReportFolder(academicYear) {
  const root = getOrCreateReportFolder(); // existing function returns "MVM_Report_Cards"
  const yearName = String(academicYear || getCurrentAcademicYear());
  const it = root.getFoldersByName(yearName);
  if (it.hasNext()) return it.next();
  return root.createFolder(yearName);
}


/**
 * Get the class/section folder under the year folder.
 * Layout: MVM_Report_Cards / {AcademicYear} / Class_{X}{Section}
 */
function getClassReportFolder(academicYear, classNum, section) {
  const yearFolder = getYearScopedReportFolder(academicYear);
  const folderName = `Class_${classNum}${section ? section : ''}`;
  const it = yearFolder.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return yearFolder.createFolder(folderName);
}


function _pdfJobKey(jobId) {
  return PDF_JOB_PREFIX + String(jobId);
}


function _pdfJobScope() {
  // DocumentProperties is shared across users for this script — appropriate for admin batch jobs
  return PropertiesService.getDocumentProperties();
}


function _pdfLoadJob(jobId) {
  const raw = _pdfJobScope().getProperty(_pdfJobKey(jobId));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}


function _pdfSaveJob(job) {
  _pdfJobScope().setProperty(_pdfJobKey(job.jobId), JSON.stringify(job));
}


function _pdfDeleteJob(jobId) {
  _pdfJobScope().deleteProperty(_pdfJobKey(jobId));
}


/**
 * List active PDF jobs (admin only) — for resume UI
 */
function listPdfJobs() {
  if (!isAdmin()) return [];
  const props = _pdfJobScope().getProperties();
  const jobs = [];
  Object.keys(props).forEach(k => {
    if (k.indexOf(PDF_JOB_PREFIX) === 0) {
      try {
        const j = JSON.parse(props[k]);
        jobs.push({
          jobId: j.jobId,
          status: j.status,
          classNum: j.classNum,
          section: j.section,
          examId: j.examId,
          examName: j.examName,
          academicYear: j.academicYear,
          totalStudents: j.totalStudents,
          completed: j.completed,
          failed: (j.errors || []).length,
          startedAt: j.startedAt,
          updatedAt: j.updatedAt,
          folderUrl: j.folderUrl
        });
      } catch (e) {}
    }
  });
  jobs.sort((a, b) => new Date(b.updatedAt || b.startedAt).getTime() - new Date(a.updatedAt || a.startedAt).getTime());
  return jobs;
}


/**
 * Start (or restart) a batched report-card generation job.
 * @param {string|number} classNum
 * @param {string} section - section name; empty/null for whole class
 * @param {string} examId
 * @param {number} [chunkSize] - default 25
 * @returns {Object} { success, jobId, totalStudents, message, ... }
 */
function startBatchedReportCardGeneration(classNum, section, examId, chunkSize) {
  if (!isAdmin()) {
    return { success: false, message: "Access denied. Admin only." };
  }
  if (!classNum) return { success: false, message: "Class is required." };
  if (!examId) return { success: false, message: "Exam is required." };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    // Validate exam
    const exam = getExamById(examId);
    if (!exam) return { success: false, message: "Exam not found." };
    
    // Build student list (one read, filtered)
    const studentsSheet = SpreadsheetApp.getActive().getSheetByName("Students");
    const lastRow = studentsSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "No students found." };
    const data = studentsSheet.getRange(2, 1, lastRow - 1, 16).getValues();
    
    let students = data
      .filter(row => row[0] && row[15] !== true && row[9] === "Active" && String(row[2]) == String(classNum))
      .map(row => ({ studentId: row[0], name: row[1], rollNo: row[5], section: row[3] }));
    
    if (section) students = students.filter(s => s.section === section);
    if (students.length === 0) {
      return { success: false, message: `No active students found for Class ${classNum}${section ? ' ' + section : ''}.` };
    }
    
    students.sort((a, b) => parseInt(a.rollNo) - parseInt(b.rollNo));
    
    const academicYear = getCurrentAcademicYear();
    const folder = getClassReportFolder(academicYear, classNum, section);
    
    const jobId = `J${Date.now()}${Math.floor(Math.random() * 1000)}`;
    const job = {
      jobId,
      status: "RUNNING",
      classNum: String(classNum),
      section: section || "",
      examId,
      examName: exam.name,
      academicYear,
      chunkSize: Math.max(5, parseInt(chunkSize) || PDF_DEFAULT_CHUNK_SIZE),
      students,                  // ordered list of { studentId, name, rollNo, section }
      totalStudents: students.length,
      completed: 0,
      cursor: 0,                  // next index to process
      generatedFiles: [],         // { studentId, name, fileName, url }
      errors: [],                 // { studentId, error }
      folderId: folder.getId(),
      folderUrl: folder.getUrl(),
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    
    _pdfSaveJob(job);
    try { writeAudit("PDF_JOB_START", "Reports", jobId, "*", "", `${students.length} students`, { classNum, section, examId }); } catch (e) {}
    logAction("PDF Job Start", `${jobId}: Class ${classNum}${section ? ' ' + section : ''}, exam ${exam.name}, ${students.length} students`);
    
    return {
      success: true,
      jobId,
      totalStudents: students.length,
      folderUrl: job.folderUrl,
      chunkSize: job.chunkSize,
      message: `Job started: ${students.length} students. Call processBatchedReportCardChunk('${jobId}') to begin.`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Process ONE chunk of an existing job. Designed to be called repeatedly from the UI.
 * Stops early if elapsed time exceeds PDF_MAX_CHUNK_RUNTIME_MS, so the next call resumes safely.
 * @param {string} jobId
 * @returns {Object} { success, status: 'RUNNING'|'COMPLETE'|'PAUSED', completed, totalStudents, percentage, errors, folderUrl, message }
 */
function processBatchedReportCardChunk(jobId) {
  if (!isAdmin()) return { success: false, message: "Access denied. Admin only." };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const job = _pdfLoadJob(jobId);
    if (!job) return { success: false, message: "Job not found. Start a new job." };
    
    if (job.status === "COMPLETE") {
      return {
        success: true, status: "COMPLETE", jobId, completed: job.completed, totalStudents: job.totalStudents,
        percentage: 100, folderUrl: job.folderUrl, errors: job.errors,
        message: "Job already complete."
      };
    }
    
    job.status = "RUNNING";
    
    // Re-fetch folder
    let folder;
    try { folder = DriveApp.getFolderById(job.folderId); }
    catch (e) {
      folder = getClassReportFolder(job.academicYear, job.classNum, job.section);
      job.folderId = folder.getId();
      job.folderUrl = folder.getUrl();
    }
    
    const chunkStart = Date.now();
    let processedThisChunk = 0;
    
    while (job.cursor < job.totalStudents && processedThisChunk < job.chunkSize) {
      // Time-budget guard: leave 90 sec headroom for finalization
      if (Date.now() - chunkStart > PDF_MAX_CHUNK_RUNTIME_MS) {
        break;
      }
      
      const student = job.students[job.cursor];
      try {
        const report = generateStudentReport(student.studentId, job.examId);
        if (!report.success) {
          job.errors.push({ studentId: student.studentId, name: student.name, error: report.message || "report failed" });
        } else {
          // Rank within this class job is computed on completion (or skipped to keep chunks simple)
          const html = generateReportCardHTML(report);
          const fileName = `${String(student.rollNo || job.cursor + 1).padStart(3, '0')}_${(student.name || 'student').replace(/[^A-Za-z0-9 _-]/g, '')}.pdf`;
          const blob = HtmlService.createHtmlOutput(html)
            .getBlob()
            .setName(fileName)
            .getAs('application/pdf');
          const file = folder.createFile(blob);
          job.generatedFiles.push({
            studentId: student.studentId,
            name: student.name,
            fileName: file.getName(),
            url: file.getUrl()
          });
        }
      } catch (e) {
        job.errors.push({ studentId: student.studentId, name: student.name, error: String(e && e.message || e) });
      }
      
      job.cursor++;
      job.completed++;
      processedThisChunk++;
      
      // Persist mid-chunk every 10 to make resume robust to hard timeouts
      if (processedThisChunk % 10 === 0) {
        job.updatedAt = new Date().toISOString();
        _pdfSaveJob(job);
      }
    }
    
    // Finalize state
    if (job.cursor >= job.totalStudents) {
      job.status = "COMPLETE";
      try { writeAudit("PDF_JOB_COMPLETE", "Reports", jobId, "*", "", `${job.completed} done, ${job.errors.length} failed`, {}); } catch (e) {}
      logAction("PDF Job Complete", `${jobId}: ${job.completed}/${job.totalStudents} (${job.errors.length} failed)`);
    } else {
      job.status = "PAUSED"; // pauses until UI calls again
    }
    job.updatedAt = new Date().toISOString();
    _pdfSaveJob(job);
    
    const pct = job.totalStudents > 0 ? Math.round((job.completed / job.totalStudents) * 100) : 100;
    return {
      success: true,
      jobId,
      status: job.status,
      completed: job.completed,
      totalStudents: job.totalStudents,
      remaining: job.totalStudents - job.completed,
      percentage: pct,
      folderUrl: job.folderUrl,
      errors: job.errors,
      processedThisChunk,
      message: job.status === "COMPLETE"
        ? `Job complete. ${job.completed} report cards generated${job.errors.length ? ', ' + job.errors.length + ' failed' : ''}.`
        : `Processed ${processedThisChunk} this chunk. ${job.completed}/${job.totalStudents} (${pct}%) done. Click Continue to resume.`
    };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


/**
 * Get current status of a job (read-only, for polling/UI)
 */
function getPdfJobStatus(jobId) {
  if (!isAdmin()) return { success: false, message: "Access denied." };
  const job = _pdfLoadJob(jobId);
  if (!job) return { success: false, message: "Job not found." };
  const pct = job.totalStudents > 0 ? Math.round((job.completed / job.totalStudents) * 100) : 0;
  return {
    success: true,
    jobId: job.jobId,
    status: job.status,
    classNum: job.classNum,
    section: job.section,
    examName: job.examName,
    completed: job.completed,
    totalStudents: job.totalStudents,
    remaining: job.totalStudents - job.completed,
    percentage: pct,
    folderUrl: job.folderUrl,
    errors: job.errors,
    startedAt: job.startedAt,
    updatedAt: job.updatedAt
  };
}


/**
 * Cancel and remove a job (does NOT delete already-generated PDFs in Drive)
 */
function cancelPdfJob(jobId) {
  if (!isAdmin()) return { success: false, message: "Access denied." };
  const job = _pdfLoadJob(jobId);
  if (!job) return { success: false, message: "Job not found." };
  _pdfDeleteJob(jobId);
  try { writeAudit("PDF_JOB_CANCEL", "Reports", jobId, "status", job.status, "CANCELLED", { completed: job.completed, total: job.totalStudents }); } catch (e) {}
  return { success: true, message: `Job cancelled. ${job.completed}/${job.totalStudents} files remain in Drive folder.`, folderUrl: job.folderUrl };
}


/**
 * Cleanup helper: remove all completed jobs older than N days
 */
function cleanupOldPdfJobs(olderThanDays) {
  if (!isAdmin()) return { success: false, message: "Access denied." };
  const days = parseInt(olderThanDays) || 7;
  const cutoff = Date.now() - days * 86400000;
  const props = _pdfJobScope().getProperties();
  let removed = 0;
  Object.keys(props).forEach(k => {
    if (k.indexOf(PDF_JOB_PREFIX) !== 0) return;
    try {
      const j = JSON.parse(props[k]);
      if (j.status === "COMPLETE" && j.updatedAt && new Date(j.updatedAt).getTime() < cutoff) {
        _pdfJobScope().deleteProperty(k);
        removed++;
      }
    } catch (e) {}
  });
  return { success: true, message: `Removed ${removed} completed jobs older than ${days} day(s).` };
}

