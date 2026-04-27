# MVM Report Tracker - Product Requirements Document

## Overview
MVM Report Tracker is a Google Apps Script-based school report tracking system that runs entirely within Google Sheets. It provides a comprehensive dashboard for managing students, teachers, exams, marks, and generating report cards. Currently being stabilized to scale to **4,000 students + 200 teachers**.

## Original Problem Statement
Build a complete school report tracking system based on provided Google Apps Script files, featuring:
- Role-based access control (Admin/Teacher)
- Student and teacher management
- Exam creation and marks entry
- Analytics and reporting
- Report card generation
- Bulk data upload capabilities

## Technical Architecture
- **Platform**: Google Apps Script (GAS)
- **Database**: Google Sheets (acts as database)
- **UI**: Single Dashboard.html file served via HtmlService
- **Authentication**: Custom HTML email-based login + PropertiesService (bypasses Google account identity sharing limitations)

## Project Structure
```
/app/MVM_Report_Tracker/
├── 0_DeploymentGuide.gs  # Deployment instructions
├── 1_Init.gs             # Sheet initialization, Subject cache helpers, Reset Subjects
├── 2_DataUpload.gs       # Bulk upload functions (Students with AcademicYear)
├── 3_Auth.gs             # Authentication and RBAC
├── 4_Exams.gs            # Exam management (LockService on createExam)
├── 5_Marks.gs            # Marks entry (LockService + Duplicate protection + cached reads)
├── 6_Analytics.gs        # Performance analytics
├── 7_Reports.gs          # Report generation and exports
└── Dashboard.html        # Complete UI
```

## Key Features

### Implemented
1. **Role-Based Access Control** - Admin & Teacher with server-side filtering
2. **Student Management** - Individual + Bulk upload, Class/Section/Stream + Elective Subject + AcademicYear
3. **Teacher Management** - Email-based login, Subject/Class/Section assignments
4. **Exam Management** - Create exams with internals, Lock/Unlock, Academic year filtering
5. **Marks Entry** - Individual + Bulk upload + Admin Bulk CSV import
6. **Analytics** - Subject/Class/Teacher performance, Weak students, Toppers
7. **Report Generation** - Student PDFs, Class-wise PDFs, Download Marksheet (Excel + CSV)
8. **Admin Tools** - Archive, Promote, Reset Subjects to Default

### Stabilization Mandate (Feb 2026) — Backend Done ✅
- ✅ **Issue 1: Duplicate Protection on Marks** — `addMarks`, `bulkAddMarks`, `adminBulkUploadMarks` now use composite key `${studentId}|${examId}|${subject}` to UPDATE existing rows instead of inserting duplicates
- ✅ **Issue 2: LockService on writes** — All write operations (`addMarks`, `bulkAddMarks`, `adminBulkUploadMarks`, `bulkUploadStudents`, `createExam`, `deleteMarks`, `resetSubjectsToDefault`) wrapped in `LockService.getScriptLock().waitLock(30000)` with `try/finally` release
- ✅ **Issue 3: Optimized Sheet Reads** — Replaced `getDataRange()` inside loops with single ranged reads + in-memory Map lookups across `5_Marks.gs` and `2_DataUpload.gs`. `bulkAddMarks` now reads each sheet ONCE and uses cached maps; batch-writes new rows in one `setValues` call
- ✅ **Issue 5: AcademicYear in Students sheet** — Added 12th column; `addStudent`, `updateStudent`, `bulkUploadStudents`, `syncStudentsFromClassSheets`, `getStudents` all populate/read it; default filter by current year
- ✅ **Issue 6: Subject Config System** — Subjects now read from `Subjects` sheet via `_getSubjectsCache()` + `getValidSubjectsForStudent()` helpers in `1_Init.gs`. Admin menu has **"Reset Subjects to Default"** button (`resetSubjectsToDefault()`). System gracefully auto-populates defaults
- ✅ **Issue 7: Fail-safe Validation on Marks** — `addMarks`, `bulkAddMarks`, `adminBulkUploadMarks` validate (a) student exists & active, (b) exam exists, (c) exam not locked, (d) subject is valid for student's class/stream/elective via `isSubjectValidForStudent()`
- ✅ **Marks Edit Protection** — Locked exam blocks all writes (add/update/delete/bulk import) with explicit "Exam is locked. No edits allowed." error message
- ✅ **Duplicate Safety UI** — `addMarks` returns `{ action: 'created' | 'updated' }` and friendly message "Marks already exist — updating existing record." for UI to display

### Pending — Pagination UI (Issue 4)
- Add `limit/offset` to `getStudents()` and marks listing
- Pagination controls in `Dashboard.html` (default 100 rows/page)

## Database Schema (Sheet Headers)
- **Students**: StudentID, Name, Class, Section, Stream, RollNo, ParentEmail, Phone, JoinDate, Status, ElectiveSubject, **AcademicYear** (NEW)
- **Teachers**: TeacherID, Name, Subject, Classes, Sections, Email, Phone, JoinDate, Status, IsClassTeacher, ClassTeacherOf
- **Subjects**: SubjectID, SubjectName, SubjectCode, Class, Stream, MaxMarks, PassingMarks, IsActive
- **Exams**: ExamID, ExamName, ExamType, Class, MaxMarks, Weightage, StartDate, EndDate, Locked, CreatedBy, CreatedAt, AcademicYear, HasInternals, Internal1-4, TotalMaxMarks
- **Marks_Master**: EntryID, StudentID, StudentName, Subject, SubjectCode, TeacherID, TeacherName, ExamID, ExamName, Class, Section, MaxMarks, MarksObtained, Percentage, Grade, UpdatedAt, UpdatedBy, AcademicYear

## API Functions (Server-side, Key)
- `addMarks(marksData)` → returns `{ success, action: 'created'|'updated', message }`
- `bulkAddMarks(marksArray)` → returns `{ success, createdCount, updatedCount, failCount, errors }`
- `adminBulkUploadMarks(data, columnMapping, options)` → preview/import with duplicate protection
- `getStudents(filters)` → optimized single read; default filters by current academic year + Active
- `resetSubjectsToDefault()` → admin-only, clears+reseeds Subjects sheet
- `getValidSubjectsForStudent(student)` / `isSubjectValidForStudent(subject, student)`
- `downloadClassExamMarks(classNum, section, examId)` / `downloadClassExamMarksCSV(...)`

## Backlog (P1)
- **Pagination UI** for students/marks tables (next up)
- Add Classes 6-8 and 9-10 (Sections A,B,C,D) — *user explicitly said wait until stabilization complete*
- Year-End Archive & Reset (clear marks, keep structure, move to Archive Spreadsheet)

## Backlog (P2)
- Auto-promote students (11 → 12)
- Password security for teachers (currently email-only)

## Deployment Notes
Refer to `0_DeploymentGuide.gs` for setting up as web app, managing user access, and configuring admin emails.
