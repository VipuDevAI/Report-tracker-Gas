/************************************************
 MVM REPORT TRACKER - ANALYTICS & AGGREGATES
 File 6 of 7
 With Academic Year Filtering & Role-Based Access
************************************************/

/**
 * Rebuild all aggregates/analytics
 * @returns {Object} Result object
 */
function rebuildAggregates() {
  const ss = SpreadsheetApp.getActive();
  const marksSheet = ss.getSheetByName("Marks_Master");
  const aggSheet = ss.getSheetByName("Aggregates");
  const currentYear = getCurrentAcademicYear();
  
  // Clear existing aggregates
  if (aggSheet.getLastRow() > 1) {
    aggSheet.getRange(2, 1, aggSheet.getLastRow() - 1, aggSheet.getLastColumn()).clearContent();
  }
  
  const marksData = marksSheet.getDataRange().getValues();
  if (marksData.length <= 1) {
    return { success: true, message: "No marks data to analyze." };
  }
  
  // Filter by current academic year
  const marks = marksData.slice(1).filter(row => {
    const rowYear = row[17] || currentYear;
    return rowYear === currentYear;
  });
  
  if (marks.length === 0) {
    return { success: true, message: "No marks data for current academic year." };
  }
  
  const aggregates = [];
  const now = new Date();
  
  // Subject averages
  const subjectStats = {};
  marks.forEach(row => {
    const subject = row[3];
    const percentage = parseFloat(row[13]);
    
    if (!subjectStats[subject]) {
      subjectStats[subject] = { total: 0, count: 0, min: 100, max: 0 };
    }
    subjectStats[subject].total += percentage;
    subjectStats[subject].count++;
    subjectStats[subject].min = Math.min(subjectStats[subject].min, percentage);
    subjectStats[subject].max = Math.max(subjectStats[subject].max, percentage);
  });
  
  Object.keys(subjectStats).forEach(subject => {
    const stats = subjectStats[subject];
    aggregates.push(["SUBJECT_AVG", subject, "", (stats.total / stats.count).toFixed(2), stats.count, now]);
    aggregates.push(["SUBJECT_MIN", subject, "", stats.min.toFixed(2), stats.count, now]);
    aggregates.push(["SUBJECT_MAX", subject, "", stats.max.toFixed(2), stats.count, now]);
  });
  
  // Class averages
  const classStats = {};
  marks.forEach(row => {
    const cls = row[9];
    const section = row[10];
    const key = `${cls}-${section}`;
    const percentage = parseFloat(row[13]);
    
    if (!classStats[key]) {
      classStats[key] = { total: 0, count: 0, passed: 0 };
    }
    classStats[key].total += percentage;
    classStats[key].count++;
    if (percentage >= 40) classStats[key].passed++;
  });
  
  Object.keys(classStats).forEach(key => {
    const stats = classStats[key];
    aggregates.push(["CLASS_AVG", key, "", (stats.total / stats.count).toFixed(2), stats.count, now]);
    aggregates.push(["CLASS_PASS_RATE", key, "", ((stats.passed / stats.count) * 100).toFixed(2), stats.count, now]);
  });
  
  // Teacher averages
  const teacherStats = {};
  marks.forEach(row => {
    const teacherId = row[5];
    const teacherName = row[6];
    const percentage = parseFloat(row[13]);
    
    if (!teacherStats[teacherId]) {
      teacherStats[teacherId] = { name: teacherName, total: 0, count: 0 };
    }
    teacherStats[teacherId].total += percentage;
    teacherStats[teacherId].count++;
  });
  
  Object.keys(teacherStats).forEach(teacherId => {
    const stats = teacherStats[teacherId];
    aggregates.push(["TEACHER_AVG", teacherId, stats.name, (stats.total / stats.count).toFixed(2), stats.count, now]);
  });
  
  // Grade distribution
  const gradeDistribution = { "A+": 0, "A": 0, "B+": 0, "B": 0, "C": 0, "D": 0, "F": 0 };
  marks.forEach(row => {
    const grade = row[14];
    if (gradeDistribution[grade] !== undefined) {
      gradeDistribution[grade]++;
    }
  });
  
  Object.keys(gradeDistribution).forEach(grade => {
    aggregates.push(["GRADE_DIST", grade, "", gradeDistribution[grade], marks.length, now]);
  });
  
  // Range distribution
  const rangeDistribution = {
    "91-100": 0,
    "81-90": 0,
    "71-80": 0,
    "61-70": 0,
    "51-60": 0,
    "41-50": 0,
    "0-40": 0
  };
  
  marks.forEach(row => {
    const percentage = parseFloat(row[13]);
    if (percentage >= 91) rangeDistribution["91-100"]++;
    else if (percentage >= 81) rangeDistribution["81-90"]++;
    else if (percentage >= 71) rangeDistribution["71-80"]++;
    else if (percentage >= 61) rangeDistribution["61-70"]++;
    else if (percentage >= 51) rangeDistribution["51-60"]++;
    else if (percentage >= 41) rangeDistribution["41-50"]++;
    else rangeDistribution["0-40"]++;
  });
  
  Object.keys(rangeDistribution).forEach(range => {
    aggregates.push(["RANGE_DIST", range, "", rangeDistribution[range], marks.length, now]);
  });
  
  // Write aggregates
  if (aggregates.length > 0) {
    aggSheet.getRange(2, 1, aggregates.length, 6).setValues(aggregates);
  }
  
  logAction("Rebuild Analytics", `Generated ${aggregates.length} aggregate entries for ${currentYear}`);
  
  return { success: true, message: `Analytics rebuilt. ${aggregates.length} entries generated for ${currentYear}.` };
}


/**
 * Get weak students (below threshold)
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters (class, section, subject)
 * @returns {Array} Weak students list
 */
function getWeakStudents(filters) {
  const marks = getMarks(filters); // Already filtered by teacher assignment
  const threshold = 40;
  
  const weakStudents = marks.filter(m => m.percentage < threshold);
  
  // Group by student
  const studentMap = {};
  weakStudents.forEach(m => {
    if (!studentMap[m.studentId]) {
      studentMap[m.studentId] = {
        studentId: m.studentId,
        studentName: m.studentName,
        class: m.class,
        section: m.section,
        weakSubjects: [],
        lowestPercentage: 100
      };
    }
    studentMap[m.studentId].weakSubjects.push({
      subject: m.subject,
      examName: m.examName,
      percentage: m.percentage
    });
    studentMap[m.studentId].lowestPercentage = Math.min(
      studentMap[m.studentId].lowestPercentage, 
      m.percentage
    );
  });
  
  // Sort by lowest percentage (most critical first)
  return Object.values(studentMap).sort((a, b) => a.lowestPercentage - b.lowestPercentage);
}


/**
 * Get toppers
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters
 * @param {number} limit - Number of toppers to return
 * @returns {Array} Toppers list
 */
function getToppers(filters, limit) {
  const marks = getMarks(filters); // Already filtered by teacher assignment
  
  // Group by student and calculate overall percentage
  const studentMap = {};
  marks.forEach(m => {
    if (!studentMap[m.studentId]) {
      studentMap[m.studentId] = {
        studentId: m.studentId,
        studentName: m.studentName,
        class: m.class,
        section: m.section,
        totalMarks: 0,
        maxMarks: 0,
        examCount: 0
      };
    }
    studentMap[m.studentId].totalMarks += m.marksObtained;
    studentMap[m.studentId].maxMarks += m.maxMarks;
    studentMap[m.studentId].examCount++;
  });
  
  // Calculate percentages and sort
  const toppers = Object.values(studentMap).map(s => ({
    ...s,
    percentage: s.maxMarks > 0 ? ((s.totalMarks / s.maxMarks) * 100).toFixed(2) : 0,
    grade: calculateGrade(s.maxMarks > 0 ? (s.totalMarks / s.maxMarks) * 100 : 0)
  })).sort((a, b) => parseFloat(b.percentage) - parseFloat(a.percentage));
  
  return toppers.slice(0, limit || 10);
}


/**
 * Get subject performance analytics
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters
 * @returns {Array} Subject performance data
 */
function getSubjectPerformance(filters) {
  const marks = getMarks(filters); // Already filtered by teacher assignment
  
  const subjectMap = {};
  marks.forEach(m => {
    if (!subjectMap[m.subject]) {
      subjectMap[m.subject] = {
        subject: m.subject,
        totalPercentage: 0,
        count: 0,
        passed: 0,
        failed: 0,
        grades: { "A+": 0, "A": 0, "B+": 0, "B": 0, "C": 0, "D": 0, "F": 0 }
      };
    }
    subjectMap[m.subject].totalPercentage += m.percentage;
    subjectMap[m.subject].count++;
    if (m.percentage >= 40) {
      subjectMap[m.subject].passed++;
    } else {
      subjectMap[m.subject].failed++;
    }
    if (subjectMap[m.subject].grades[m.grade] !== undefined) {
      subjectMap[m.subject].grades[m.grade]++;
    }
  });
  
  return Object.values(subjectMap).map(s => ({
    ...s,
    avgPercentage: (s.totalPercentage / s.count).toFixed(2),
    passRate: ((s.passed / s.count) * 100).toFixed(2)
  })).sort((a, b) => parseFloat(b.avgPercentage) - parseFloat(a.avgPercentage));
}


/**
 * Get class performance analytics
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters
 * @returns {Array} Class performance data
 */
function getClassPerformance(filters) {
  const marks = getMarks(filters); // Already filtered by teacher assignment
  
  const classMap = {};
  marks.forEach(m => {
    const key = `${m.class}-${m.section}`;
    if (!classMap[key]) {
      classMap[key] = {
        class: m.class,
        section: m.section,
        totalPercentage: 0,
        count: 0,
        passed: 0,
        topScore: 0,
        students: new Set()
      };
    }
    classMap[key].totalPercentage += m.percentage;
    classMap[key].count++;
    if (m.percentage >= 40) classMap[key].passed++;
    classMap[key].topScore = Math.max(classMap[key].topScore, m.percentage);
    classMap[key].students.add(m.studentId);
  });
  
  return Object.values(classMap).map(c => ({
    class: c.class,
    section: c.section,
    avgPercentage: (c.totalPercentage / c.count).toFixed(2),
    passRate: ((c.passed / c.count) * 100).toFixed(2),
    topScore: c.topScore.toFixed(2),
    studentCount: c.students.size,
    examEntries: c.count
  })).sort((a, b) => parseFloat(b.avgPercentage) - parseFloat(a.avgPercentage));
}


/**
 * Get teacher performance analytics
 * Admin sees all, Teacher sees only their own
 * @returns {Array} Teacher performance data
 */
function getTeacherPerformance() {
  const marks = getMarks(); // Already filtered by teacher assignment
  
  const teacherMap = {};
  marks.forEach(m => {
    if (!teacherMap[m.teacherId]) {
      teacherMap[m.teacherId] = {
        teacherId: m.teacherId,
        teacherName: m.teacherName,
        totalPercentage: 0,
        count: 0,
        passed: 0,
        subjects: new Set(),
        classes: new Set()
      };
    }
    teacherMap[m.teacherId].totalPercentage += m.percentage;
    teacherMap[m.teacherId].count++;
    if (m.percentage >= 40) teacherMap[m.teacherId].passed++;
    teacherMap[m.teacherId].subjects.add(m.subject);
    teacherMap[m.teacherId].classes.add(`${m.class}-${m.section}`);
  });
  
  return Object.values(teacherMap).map(t => ({
    teacherId: t.teacherId,
    teacherName: t.teacherName,
    avgPercentage: (t.totalPercentage / t.count).toFixed(2),
    passRate: ((t.passed / t.count) * 100).toFixed(2),
    subjects: Array.from(t.subjects),
    classes: Array.from(t.classes),
    examEntries: t.count
  })).sort((a, b) => parseFloat(b.avgPercentage) - parseFloat(a.avgPercentage));
}


/**
 * Get range distribution
 * Applies teacher filtering for non-admin users
 * @param {Object} filters - Optional filters
 * @returns {Object} Range distribution data
 */
function getRangeDistribution(filters) {
  const marks = getMarks(filters); // Already filtered by teacher assignment
  
  const distribution = {
    "91-100": { count: 0, label: "A+ (91-100%)", color: "#22c55e" },
    "81-90": { count: 0, label: "A (81-90%)", color: "#16a34a" },
    "71-80": { count: 0, label: "B+ (71-80%)", color: "#3b82f6" },
    "61-70": { count: 0, label: "B (61-70%)", color: "#0ea5e9" },
    "51-60": { count: 0, label: "C (51-60%)", color: "#f59e0b" },
    "41-50": { count: 0, label: "D (41-50%)", color: "#f97316" },
    "0-40": { count: 0, label: "F (0-40%)", color: "#ef4444" }
  };
  
  marks.forEach(m => {
    const pct = m.percentage;
    if (pct >= 91) distribution["91-100"].count++;
    else if (pct >= 81) distribution["81-90"].count++;
    else if (pct >= 71) distribution["71-80"].count++;
    else if (pct >= 61) distribution["61-70"].count++;
    else if (pct >= 51) distribution["51-60"].count++;
    else if (pct >= 41) distribution["41-50"].count++;
    else distribution["0-40"].count++;
  });
  
  const total = marks.length;
  Object.keys(distribution).forEach(key => {
    distribution[key].percentage = total > 0 
      ? ((distribution[key].count / total) * 100).toFixed(1) 
      : 0;
  });
  
  return {
    distribution: distribution,
    total: total
  };
}


/**
 * Get dashboard KPIs
 * Filtered by user role and academic year
 * @returns {Object} Dashboard KPI data
 */
function getDashboardKPIs() {
  const subjectPerf = getSubjectPerformance();
  const classPerf = getClassPerformance();
  const teacherPerf = getTeacherPerformance();
  const rangeData = getRangeDistribution();
  const currentYear = getCurrentAcademicYear();
  const userRole = getCurrentUserRole();
  
  // Find lowest performing subject
  const lowestSubject = subjectPerf.length > 0 
    ? subjectPerf[subjectPerf.length - 1] 
    : null;
  
  // Find top class
  const topClass = classPerf.length > 0 ? classPerf[0] : null;
  
  // Find top teacher
  const topTeacher = teacherPerf.length > 0 ? teacherPerf[0] : null;
  
  // Calculate above 80%
  const above80 = rangeData.total > 0 
    ? (((rangeData.distribution["91-100"].count + rangeData.distribution["81-90"].count) / rangeData.total) * 100).toFixed(1)
    : 0;
  
  return {
    lowestSubject: lowestSubject ? {
      name: lowestSubject.subject,
      avgScore: lowestSubject.avgPercentage
    } : null,
    topClass: topClass ? {
      name: `Class ${topClass.class}${topClass.section}`,
      avgScore: topClass.avgPercentage
    } : null,
    topTeacher: topTeacher ? {
      name: topTeacher.teacherName,
      avgScore: topTeacher.avgPercentage
    } : null,
    rangeAbove80: above80,
    totalEntries: rangeData.total,
    academicYear: currentYear,
    userRole: userRole,
    dataScope: userRole === "admin" ? "All Data" : "Your Assigned Data"
  };
}


/**
 * Get analytics summary for teacher's own classes
 * @returns {Object} Teacher-specific analytics
 */
function getMyAnalytics() {
  if (isAdmin()) {
    return {
      message: "Admin view - showing all data",
      subjectPerformance: getSubjectPerformance(),
      classPerformance: getClassPerformance(),
      weakStudents: getWeakStudents(),
      toppers: getToppers(null, 10)
    };
  }
  
  const assignment = getTeacherAssignment();
  if (!assignment) {
    return { message: "No assignment found", data: null };
  }
  
  return {
    message: `Showing data for ${assignment.subject} - Classes: ${assignment.classes.join(", ")}`,
    teacher: assignment,
    subjectPerformance: getSubjectPerformance(),
    classPerformance: getClassPerformance(),
    weakStudents: getWeakStudents(),
    toppers: getToppers(null, 10)
  };
}
