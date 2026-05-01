# MVM Report Tracker - Product Requirements Document

## Overview
MVM Report Tracker is a Google Apps Script-based school report tracking system that runs entirely within Google Sheets. Designed to scale to **4,000 students + 200 teachers** for a full academic year without data corruption, duplicates, or performance issues.

## Original Problem Statement
Build a complete school report tracking system based on provided Google Apps Script files, featuring:
- Role-based access control (Admin / Principal / Wing Admin / Teacher)
- Student, teacher, exam, and marks management for Classes 6-12
- Analytics, reporting, and report card generation
- Bulk data upload
- Stable, optimized, config-driven architecture

## Technical Architecture
- **Platform**: Google Apps Script (GAS)
- **Database**: Google Sheets (single unified `Marks_Master` — NOT split per class)
- **UI**: Single Dashboard.html via HtmlService
- **Authentication**: Custom HTML email-based login + PropertiesService

## Project Structure
```
/app/MVM_Report_Tracker/
├── 0_DeploymentGuide.gs
├── 1_Init.gs             # Schema, Subject cache, Year-end (Archive/Reset/Switch), Performance warning
├── 2_DataUpload.gs       # Bulk uploads, getStudentsPage (paginated), promoteStudents 6→12
├── 3_Auth.gs             # RBAC, custom email login, getCurrentAcademicYear
├── 4_Exams.gs            # Exam CRUD with LockService
├── 5_Marks.gs            # Marks CRUD (LockService, duplicate protection, validation), getMarksPage, getAuditLog
├── 6_Analytics.gs
├── 7_Reports.gs
└── Dashboard.html        # Full UI (pagination, audit log, year-end UI, perf warning)
```

## Database Schema

### Students (15 columns)
StudentID, Name, Class, Section, Stream, RollNo, ParentEmail, Phone, JoinDate, Status, ElectiveSubject, AcademicYear, **LanguageL1**, **LanguageL2**, **LanguageL3**

### Subjects (10 columns) — Config-Driven
SubjectID, SubjectName, SubjectCode, Class, Stream (nullable for 6-10), MaxMarks, PassingMarks, IsActive, **LanguageGroup** (L1/L2/L3 or comma list), **IsOptional**

### Marks_Master (18 columns)
EntryID, StudentID, StudentName, Subject, SubjectCode, TeacherID, TeacherName, ExamID, ExamName, Class, Section, MaxMarks, MarksObtained, Percentage, Grade, UpdatedAt, UpdatedBy, AcademicYear

### Teachers (11), Exams (18), Classes, Settings_*, Logs, Aggregates, Alerts

## Class & Section Configuration
- **Class 6-10**: 11 sections (A1-A11), no stream
- **Class 11-12**: 12 sections (A1-A12), with stream (Science / Computer Science / Commerce)
- **Subjects per class**: read from `Subjects` sheet (NOT hardcoded)

## Languages
- **6-8**: 3-language system → LanguageL1 + LanguageL2 + LanguageL3 (English fixed L1; L2/L3 chosen from Hindi, Sanskrit, Tamil)
- **9-10**: 2-language system → LanguageL1 + LanguageL2
- **11-12**: ElectiveSubject (Maths / Applied Maths / Hindi / History / Sanskrit)

## ✅ Implemented (Feb 2026 Stabilization Mandate)

### Concurrency, Integrity & Performance
- **Duplicate Protection**: All marks operations use composite key `studentId|examId|subject` → UPDATE existing instead of insert
- **LockService**: All write paths wrap with `LockService.getScriptLock().waitLock(30000)` + `try/finally` release. Files: addMarks, bulkAddMarks, adminBulkUploadMarks, deleteMarks, bulkUploadStudents, addStudent, updateStudent, createExam, syncStudentsFromClassSheets, archiveAcademicYear, resetForNewYear, promoteStudents, resetSubjectsToDefault
- **Optimized Reads**: Eliminated `getDataRange()` inside loops. Bulk operations read each sheet ONCE → in-memory Map → single `setValues` batch write
- **Server-side Pagination**: `getStudentsPage(filters, page, limit=100)`, `getMarksPage(...)`, `getAuditLog(...)` — default 100 rows
- **Performance Safety**: Dashboard shows warning when `Marks_Master` ≥ 200,000 rows via `getMarksRowCount()`

### Validation Layer (final, before any marks save)
- Student exists & is Active
- Exam exists & not locked (Marks Edit Protection — locked exams reject add/update/delete)
- Subject is valid for student's class/stream/elective via config-driven `isSubjectValidForStudent()`
- Marks ≥ 0 and ≤ MaxMarks
- No null/missing required fields
- All errors return clear, actionable messages

### Subject Config (config-driven)
- `_getSubjectsCache()` reads Subjects sheet once per execution
- `getValidSubjectsForStudent(student)` derives valid subjects from config (mandatory + chosen languages + elective)
- Admin "Reset Subjects to Default" button re-seeds defaults
- Default seed covers Class 6-12 with proper language groups and stream-based subjects

### Year-End System
- `archiveAcademicYear(year)` — exports Students + Marks + Exams to a NEW spreadsheet (returns URL)
- `resetForNewYear(newYear)` — clears Marks_Master + Exams + Aggregates, switches academic year, preserves Students
- `switchAcademicYear(newYear)` — non-destructive year switch
- Admin menu items + dedicated UI page (Year-End Operations) with 3 cards

### Student Promotion (6 → 7 → 8 → 9 → 10 → 11 → 12)
- `promoteStudents(fromYear, toYear, { resetRollNumbers })` — section preserved, optional roll-number reset, Class 12 → Alumni
- All 6 → 12 supported

### Audit View (Admin only)
- Dashboard tab "Audit Log" reads Marks_Master with filters: class, subject, updatedBy, fromDate, toDate, search
- Server-side paginated, sorted by UpdatedAt desc
- Columns: Updated At, Student, Class/Section, Subject, Exam, Marks, %, Updated By, Teacher

### UI Polish
- Class dropdowns: 6, 7, 8, 9, 10, 11, 12 everywhere
- Section dropdowns: dynamically populated by `getSectionsForClass()` (A1-A11 or A1-A12)
- Add Student modal: Language L1/L2/L3 fields + optional stream/elective
- Search box + debounced inputs on Students and Audit views
- Performance warning banner on Dashboard when row threshold crossed

### Architecture Constraints (User Requirements)
- ✅ **Marks sheet remains UNIFIED** — no per-class/per-section duplication
- ✅ **System is config-driven** — no hardcoded subjects, sections (except defaults at init)
- ✅ **Single source of truth** for everything

## API Functions (Server-side, Key)

### Student / Subject
- `getStudentsPage(filters, page, limit)` → `{ data, total, page, limit, totalPages }`
- `getStudents(filters)` → array (legacy/unpaginated, used internally)
- `addStudent`, `updateStudent`, `bulkUploadStudents` — all write 15 cols, LockService-guarded
- `getSubjects(filters)` — returns 10-col data including languageGroup, isOptional
- `getValidSubjectsForStudent(student)` / `isSubjectValidForStudent(subject, student)`

### Marks
- `addMarks(data)` → `{ success, action: 'created'|'updated', message }`
- `bulkAddMarks(array)` → `{ createdCount, updatedCount, failCount, errors }`
- `adminBulkUploadMarks(data, mapping, options)` — preview/import with full validation
- `getMarksPage(filters, page, limit)` — paginated
- `getAuditLog(filters, page, limit)` — admin-only audit history

### Year-End / Promotion
- `archiveAcademicYear(year)` → `{ success, url, fileName, message }`
- `resetForNewYear(newYear)` — clears Marks/Exams, preserves Students
- `switchAcademicYear(newYear)` — settings only
- `promoteStudents(fromYear, toYear, { resetRollNumbers })` — 6→12
- `getMarksRowCount()` → `{ count, threshold, warning, message }`

### Admin Tools
- `resetSubjectsToDefault()` — clears + reseeds Subjects
- `initializeApp()` — creates all sheets

## Test Credentials
See `/app/memory/test_credentials.md`. Admin emails: `rishisans83@gmail.com`, `mvmseniors@gmail.com`, `anithasivanesan4604@gmail.com`. Teachers: any active email in Teachers sheet.

## Backlog (Future)
- Auto-promote scheduled trigger
- Time-driven trigger for `rebuildAggregates()`
- Per-language separate Subjects view in admin
- Class teacher dashboard with class-only view
- Multi-year analytics across Archive sheets

## Changelog
### Feb 2026 — Wing Admin Teacher Management + Guided Dashboard
- **Backend (Wing Admin user mgmt):**
  - New gate `_denyIfNoTeacherWrite(targetRole, targetClasses)` in `2_DataUpload.gs` — admin: any; wing_admin: only `Role===TEACHER` AND every target class within wing.
  - `addTeacher` opened to wing_admin (TEACHER-only, in-wing classes).
  - New `updateTeacher(email, updates)` — wing_admin cannot change `email` or `role` (silently ignored); enforces wing scope on current AND target classes.
  - New `deleteTeacher(email)` — soft-delete via `IsDeleted`, marks history preserved, refuses to delete the last remaining ADMIN, invalidates active session.
  - New `restoreTeacher(email)` — admin-only.
- **Backend (Onboarding):**
  - New file `8_Onboarding.gs` with `getNextSteps()` returning role-aware checklist.
  - Admin: 5-step setup checklist (Teachers→Students→Exam→Marks→Reports) using existing sheet counts.
  - Wing Admin: top 5 unlocked exams in wing with pending-students count.
  - Teacher: top 5 active exams in their assignment with their personal pending-students count (filtered by `subject` and `updatedBy` email).
- **Frontend (Dashboard.html):**
  - "Next Step" hero card replaces dashboard top; KPI grid demoted below (kept for post-setup analytics).
  - CTA pre-selects context (class/section/exam/subject) on the marks page.
  - Auto-fill + lock subject dropdown on marks page for teachers (their assignment dictates).
  - Teachers nav re-enabled for Wing Admin (Principal still hidden — read-only).
  - Edit/Delete buttons in teachers list (wired to `updateTeacher`/`deleteTeacher`).
  - Same modal serves Add + Edit (email field locked in edit mode).
  - Role dropdown auto-hidden for Wing Admin in the modal (forced TEACHER).

### Feb 2026 — 4-Role System (Option B)
- Added `Role` column to Teachers sheet: `ADMIN | PRINCIPAL | WING_ADMIN | TEACHER` (single source of truth for non-super-admin roles).
- New helpers in `3_Auth.gs`: `getRole()`, `isPrincipal()`, `isWingAdmin()`, `canWrite(action)`, `canRead()`, `getWingForClass()`, `getClassesForWing()`, `getWingAdminAssignment()`.
- Refactored `applyTeacherFilter` → `applyScopeFilter` (4-way: admin/principal=full, wing_admin=wing classes, teacher=existing). `applyTeacherFilter` retained as backward-compat alias.
- Wing config in `Settings_School` (`Wing_Primary=6,7,8`, `Wing_Secondary=9,10`, `Wing_Senior=11,12`) — auto-seeded; idempotent migration via `seedDefaultSchoolSettings()`.
- Idempotent migration `ensureTeachersRoleColumn()` (auto-called from `initializeApp`, `addTeacher`, `bulkUploadTeachers`).
- Write-gate refactor: replaced `isAdmin()` with `canWrite()` + scope helpers (`_denyIfNoStudentWrite`, `_denyIfNoExamWrite`, `_denyIfNoReportAccess`) on student CRUD, marks CRUD, exam CRUD, report generation. Lock-exam, year freeze, archive, settings remain admin-only.
- Teacher assignment normalization: trim + standardize separators (`,`, `;`, `|` → comma) on save (`addTeacher`, `bulkUploadTeachers`, `getTeacherAssignment`).
- `Dashboard.html`: 4-role aware UI (`applyRoleRestrictions`), Role dropdown in Add Teacher modal, Role badge column in teachers list, Principal read-only mode (CSS-disables write actions), wing/scope label in header.
- Removed shadowing duplicate `deleteStudent` (the second declaration was hiding the proper soft-delete with audit + LockService).
- `getTeacherByEmail` now respects `IsDeleted` and returns `role`. `getAuditLog` now allows Principal too.

## Deployment
Refer to `0_DeploymentGuide.gs`. Web app deploy: Execute as "Me", Access "Anyone with Google account". Share only the web app URL.
