/************************************************
 MVM REPORT TRACKER - ONBOARDING / NEXT STEPS
 File 8 of 8

 Lightweight role-aware "what should I do next?" helper.
 No new sheets — pure aggregation over existing data.

 Returns:
   {
     role: 'admin'|'principal'|'wing_admin'|'teacher'|'unauthorized',
     headline: string,                  // 1-line title for the hero card
     subline:  string,                  // 1-line description
     steps:    [{ key, label, count, done, ctaPage, ctaLabel, context? }],
     primary:  { ctaPage, ctaLabel, context? } | null,
     allDone:  boolean
   }

 `context` is an opaque object the UI can pass back to pre-select dropdowns
 on the destination page (e.g., { class:"9", section:"A1", examId:"EXM..." }).
************************************************/

/**
 * Public entry point — called from Dashboard on load.
 * @returns {Object}
 */
function getNextSteps() {
  const role = getRole();
  if (!role) {
    return {
      role: "unauthorized",
      headline: "Access denied",
      subline: "Your email is not registered. Contact admin.",
      steps: [], primary: null, allDone: false
    };
  }
  
  if (role === "admin")     return _nextStepsAdmin();
  if (role === "principal") return _nextStepsPrincipal();
  if (role === "wing_admin") return _nextStepsWingAdmin();
  if (role === "teacher")   return _nextStepsTeacher();
  
  return { role, headline: "Welcome", subline: "", steps: [], primary: null, allDone: false };
}


// ─── ADMIN ────────────────────────────────────────────────────────────────

function _nextStepsAdmin() {
  // Cheap counts (each is 1 sheet read)
  const teachersTotal = _countNonDeleted("Teachers", 11);
  const studentsTotal = _countNonDeleted("Students", 15);
  const examsTotal    = _countNonDeleted("Exams", 18);
  const marksTotal    = _countNonDeleted("Marks_Master", 19);
  
  const steps = [
    {
      key: "teachers",
      label: "Add Teachers, Wing Admins & Principal",
      count: teachersTotal,
      done: teachersTotal > 0,
      ctaPage: "teachers",
      ctaLabel: teachersTotal > 0 ? "Manage Users" : "Add First User"
    },
    {
      key: "students",
      label: "Add Students",
      count: studentsTotal,
      done: studentsTotal > 0,
      ctaPage: "students",
      ctaLabel: studentsTotal > 0 ? "Manage Students" : "Add Students"
    },
    {
      key: "exams",
      label: "Create Exam",
      count: examsTotal,
      done: examsTotal > 0,
      ctaPage: "create-exam",
      ctaLabel: examsTotal > 0 ? "Manage Exams" : "Create Exam"
    },
    {
      key: "marks",
      label: "Enter Marks",
      count: marksTotal,
      done: marksTotal > 0,
      ctaPage: "marks",
      ctaLabel: marksTotal > 0 ? "Continue Marks Entry" : "Start Marks Entry"
    },
    {
      key: "reports",
      label: "Generate Report Cards",
      count: 0,
      done: false, // never marked "done" — always available once marks exist
      ctaPage: "report-cards",
      ctaLabel: "Generate Reports"
    }
  ];
  
  // First incomplete step is primary; if all data steps done, primary = reports
  let primary = null;
  for (let i = 0; i < steps.length; i++) {
    if (!steps[i].done) { primary = { ctaPage: steps[i].ctaPage, ctaLabel: steps[i].ctaLabel }; break; }
  }
  if (!primary) primary = { ctaPage: "report-cards", ctaLabel: "Generate Reports" };
  
  // Empty-state messaging
  const allDone = steps.slice(0, 4).every(s => s.done); // setup steps 1-4
  let headline, subline;
  if (teachersTotal === 0) {
    headline = "Welcome! Let's set up your school.";
    subline  = "Start by adding teachers and admins.";
  } else if (studentsTotal === 0) {
    headline = "Add students next.";
    subline  = "No students added yet — start by adding students.";
  } else if (examsTotal === 0) {
    headline = "Create your first exam.";
    subline  = "Define an exam (e.g., Unit Test 1) before teachers can enter marks.";
  } else if (marksTotal === 0) {
    headline = "Marks entry is ready to begin.";
    subline  = "Teachers can now enter marks for the created exam(s).";
  } else {
    headline = "Setup complete";
    subline  = "Track progress and generate report cards anytime.";
  }
  
  return { role: "admin", headline, subline, steps, primary, allDone };
}


// ─── PRINCIPAL ────────────────────────────────────────────────────────────

function _nextStepsPrincipal() {
  return {
    role: "principal",
    headline: "Read-only overview",
    subline: "View analytics, audit trail and download report cards. No edits.",
    steps: [
      { key: "analytics",  label: "Open analytics dashboards", done: false, ctaPage: "subject-perf", ctaLabel: "View Analytics" },
      { key: "reports",    label: "Generate report cards",     done: false, ctaPage: "report-cards", ctaLabel: "Generate Reports" },
      { key: "audit",      label: "Review audit trail",        done: false, ctaPage: "audit-log",    ctaLabel: "Open Audit Log" }
    ],
    primary: { ctaPage: "subject-perf", ctaLabel: "View Analytics" },
    allDone: false
  };
}


// ─── WING ADMIN ───────────────────────────────────────────────────────────

function _nextStepsWingAdmin() {
  const wa = getWingAdminAssignment();
  const wingClasses = (wa && wa.classes) ? wa.classes.map(String) : [];
  
  if (!wingClasses.length) {
    return {
      role: "wing_admin",
      headline: "Wing assignment missing",
      subline: "Ask an admin to set the Classes column on your Teachers row.",
      steps: [], primary: null, allDone: false
    };
  }
  
  // Collect unlocked, non-deleted exams in this wing
  const examsSheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const examsLastRow = examsSheet ? examsSheet.getLastRow() : 1;
  let pendingExams = [];
  if (examsLastRow > 1) {
    const ed = examsSheet.getRange(2, 1, examsLastRow - 1, 19).getValues();
    ed.forEach(r => {
      const examId   = r[0];
      const examName = r[1];
      const examCls  = String(r[3] || "");
      const locked   = r[8] === true;
      const deleted  = r[18] === true;
      if (!examId || locked || deleted) return;
      if (wingClasses.indexOf(examCls) === -1) return;
      pendingExams.push({ examId, examName, class: examCls });
    });
  }
  
  // Pending students per (class, exam) = total active students in class - distinct
  // students with non-deleted marks in that exam.
  const studentsByClass = _activeStudentCountByClass(wingClasses);
  const marksCoverage   = _distinctStudentsWithMarksByExamClass();
  
  pendingExams.forEach(e => {
    const total = studentsByClass[e.class] || 0;
    const key   = `${e.examId}|${e.class}`;
    const done  = marksCoverage[key] || 0;
    e.totalStudents = total;
    e.studentsWithMarks = done;
    e.pending = Math.max(0, total - done);
  });
  
  // Sort: most pending first
  pendingExams.sort((a, b) => b.pending - a.pending);
  pendingExams = pendingExams.slice(0, 5);
  
  const totalStudents = wingClasses.reduce((s, c) => s + (studentsByClass[c] || 0), 0);
  const totalPending  = pendingExams.reduce((s, e) => s + e.pending, 0);
  
  const steps = pendingExams.map(e => ({
    key: `${e.examId}-${e.class}`,
    label: `${e.examName} — Class ${e.class}`,
    count: e.pending,
    done: e.pending === 0,
    ctaPage: "marks",
    ctaLabel: e.pending === 0 ? "View Entered" : `Enter (${e.pending} left)`,
    context: { class: e.class, examId: e.examId }
  }));
  
  let primary = null;
  if (pendingExams.length && pendingExams[0].pending > 0) {
    primary = {
      ctaPage: "marks",
      ctaLabel: "Start Marks Entry",
      context: { class: pendingExams[0].class, examId: pendingExams[0].examId }
    };
  } else if (pendingExams.length) {
    primary = { ctaPage: "report-cards", ctaLabel: "Generate Reports" };
  } else if (totalStudents === 0) {
    primary = { ctaPage: "students", ctaLabel: "Add Students" };
  } else {
    primary = { ctaPage: "create-exam", ctaLabel: "Create Exam" };
  }
  
  let headline, subline;
  if (totalStudents === 0) {
    headline = "No students in your wing yet.";
    subline  = "No students added yet — start by adding students for classes " + wingClasses.join(", ") + ".";
  } else if (pendingExams.length === 0) {
    headline = "No active exams in your wing.";
    subline  = "Create an exam to begin marks entry.";
  } else if (totalPending === 0) {
    headline = "All caught up";
    subline  = "Every active exam has marks for all students. Generate reports anytime.";
  } else {
    headline = `${totalPending} student${totalPending === 1 ? "" : "s"} pending marks entry`;
    subline  = `Across ${pendingExams.length} active exam${pendingExams.length === 1 ? "" : "s"} in your wing (Classes ${wingClasses.join(", ")}).`;
  }
  
  return {
    role: "wing_admin",
    wing: wa.wing || null,
    headline, subline, steps, primary,
    allDone: totalPending === 0 && pendingExams.length > 0
  };
}


// ─── TEACHER ──────────────────────────────────────────────────────────────

function _nextStepsTeacher() {
  const a = getTeacherAssignment();
  if (!a) {
    return {
      role: "teacher",
      headline: "No assignment found",
      subline: "Contact admin to set your subject/classes/sections.",
      steps: [], primary: null, allDone: false
    };
  }
  
  const myClasses  = a.hasAllClasses  ? null : a.classes.map(String);    // null = all
  const mySubject  = a.subject;
  
  // Active unlocked exams matching my classes
  const examsSheet = SpreadsheetApp.getActive().getSheetByName("Exams");
  const examsLastRow = examsSheet ? examsSheet.getLastRow() : 1;
  let exams = [];
  if (examsLastRow > 1) {
    const ed = examsSheet.getRange(2, 1, examsLastRow - 1, 19).getValues();
    ed.forEach(r => {
      const examId = r[0];
      const examName = r[1];
      const cls = String(r[3] || "");
      const locked = r[8] === true;
      const deleted = r[18] === true;
      if (!examId || locked || deleted) return;
      // "All"-class exams are visible to every teacher
      if (cls.toLowerCase() === "all" || !myClasses || myClasses.indexOf(cls) >= 0) {
        exams.push({ examId, examName, class: cls });
      }
    });
  }
  
  // For each (exam, my class), compute students - studentsWithMyMarksForSubject
  const targetClasses = myClasses || _allActiveClasses();
  const studentsByClass = _activeStudentCountByClass(targetClasses);
  const myCoverage = _myMarksCoverage(mySubject, a.email);
  
  let steps = [];
  exams.forEach(e => {
    const classesForRow = (e.class.toLowerCase() === "all") ? targetClasses : [e.class];
    classesForRow.forEach(cls => {
      const total = studentsByClass[cls] || 0;
      if (total === 0) return;
      const key = `${e.examId}|${cls}`;
      const done = myCoverage[key] || 0;
      const pending = Math.max(0, total - done);
      // Pick a representative section: first one in assignment
      const section = (a.sections && a.sections.length && !a.hasAllSections) ? a.sections[0] : "";
      steps.push({
        key: `${e.examId}-${cls}`,
        label: `${e.examName} — Class ${cls}${mySubject && mySubject !== "All" ? " · " + mySubject : ""}`,
        count: pending,
        done: pending === 0,
        ctaPage: "marks",
        ctaLabel: pending === 0 ? "Edit / Review" : `Enter (${pending} left)`,
        context: { class: cls, section: section, examId: e.examId, subject: mySubject }
      });
    });
  });
  
  // Sort: most pending first; cap at 5
  steps.sort((s1, s2) => s2.count - s1.count);
  steps = steps.slice(0, 5);
  
  const totalPending = steps.reduce((s, x) => s + (x.count || 0), 0);
  
  let primary = null;
  if (steps.length && steps[0].count > 0) {
    primary = { ctaPage: "marks", ctaLabel: "Enter Marks", context: steps[0].context };
  } else if (steps.length) {
    primary = { ctaPage: "view-marks", ctaLabel: "View My Marks" };
  } else {
    primary = { ctaPage: "view-marks", ctaLabel: "View My Marks" };
  }
  
  let headline, subline;
  const subjPart = (mySubject && mySubject !== "All") ? mySubject : "your subjects";
  if (steps.length === 0) {
    headline = "No active exams need your input.";
    subline  = `Assigned: ${subjPart} · Classes ${a.classes.join(", ") || "(all)"}.`;
  } else if (totalPending === 0) {
    headline = "All caught up";
    subline  = `You've entered ${subjPart} marks for every student in active exams.`;
  } else {
    headline = `${totalPending} student${totalPending === 1 ? "" : "s"} pending`;
    subline  = `${subjPart} · Class ${a.classes.join(", ") || "all"}${a.sections.length ? " · Section " + a.sections.join(", ") : ""}`;
  }
  
  return { role: "teacher", assignment: a, headline, subline, steps, primary, allDone: totalPending === 0 && steps.length > 0 };
}


// ─── HELPERS ──────────────────────────────────────────────────────────────

/** Count rows where IsDeleted (zero-based col index) is not strictly true. */
function _countNonDeleted(sheetName, isDeletedColIdx) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return 0;
  const last = sheet.getLastRow();
  if (last <= 1) return 0;
  const lastCol = sheet.getLastColumn();
  if (lastCol <= isDeletedColIdx) {
    // IsDeleted column doesn't exist yet (legacy schema) — count all rows with non-empty col 0
    const data = sheet.getRange(2, 1, last - 1, 1).getValues();
    return data.filter(r => r[0]).length;
  }
  const data = sheet.getRange(2, 1, last - 1, isDeletedColIdx + 1).getValues();
  return data.filter(r => r[0] && r[isDeletedColIdx] !== true).length;
}


/** Returns { "9": 240, "10": 198, ... } for the given class numbers. */
function _activeStudentCountByClass(classes) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  const out = {};
  classes.forEach(c => { out[String(c)] = 0; });
  if (!sheet || sheet.getLastRow() <= 1) return out;
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  const wantSet = new Set(classes.map(String));
  data.forEach(r => {
    if (!r[0] || r[15] === true) return;     // not deleted
    if ((r[9] || "").toString() !== "Active") return;
    const cls = String(r[2] || "");
    if (wantSet.has(cls)) out[cls] = (out[cls] || 0) + 1;
  });
  return out;
}


/** Returns { "EXM123|9": 12, ... } — distinct studentIds with non-deleted marks per (exam, class). */
function _distinctStudentsWithMarksByExamClass() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  if (!sheet || sheet.getLastRow() <= 1) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues();
  const seen = {};   // key -> Set
  data.forEach(r => {
    if (!r[0] || r[19] === true) return;
    const status = String(r[18] || "PRESENT").toUpperCase();
    if (status === "ABSENT" || status === "EXEMPT") {
      // Still counts as "entered" — the teacher acknowledged the student
    }
    const k = `${r[7]}|${String(r[9] || "")}`;
    if (!seen[k]) seen[k] = new Set();
    seen[k].add(r[1]);
  });
  const out = {};
  Object.keys(seen).forEach(k => { out[k] = seen[k].size; });
  return out;
}


/** Per-teacher coverage filtered by subject + updatedBy email (case-insensitive). */
function _myMarksCoverage(subject, email) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Marks_Master");
  if (!sheet || sheet.getLastRow() <= 1) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues();
  const subjLower  = String(subject || "").toLowerCase();
  const emailLower = String(email   || "").toLowerCase();
  const matchAllSubjects = subjLower === "all" || subjLower === "";
  const seen = {};
  data.forEach(r => {
    if (!r[0] || r[19] === true) return;
    const rSubj = String(r[3] || "").toLowerCase();
    if (!matchAllSubjects && rSubj !== subjLower) return;
    // Only count rows entered by *this* user when subject is shared
    if (!matchAllSubjects) {
      const updatedBy = String(r[16] || "").toLowerCase();
      if (emailLower && updatedBy && updatedBy !== emailLower) return;
    }
    const k = `${r[7]}|${String(r[9] || "")}`;
    if (!seen[k]) seen[k] = new Set();
    seen[k].add(r[1]);
  });
  const out = {};
  Object.keys(seen).forEach(k => { out[k] = seen[k].size; });
  return out;
}


/** All distinct active class numbers (used when teacher.hasAllClasses). */
function _allActiveClasses() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Students");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  const set = new Set();
  data.forEach(r => {
    if (!r[0] || r[15] === true) return;
    if ((r[9] || "").toString() !== "Active") return;
    const cls = String(r[2] || "");
    if (cls) set.add(cls);
  });
  return Array.from(set);
}
