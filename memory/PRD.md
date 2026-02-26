# MVM Report Tracker - Product Requirements Document

## Overview
MVM Report Tracker is a Google Apps Script-based school report tracking system that runs entirely within Google Sheets. It provides a comprehensive dashboard for managing students, teachers, exams, marks, and generating report cards.

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
- **Authentication**: Google Account email-based (Session.getActiveUser().getEmail())

## Project Structure
```
/app/MVM_Report_Tracker/
├── 0_DeploymentGuide.gs  # Deployment instructions
├── 1_Init.gs             # Sheet initialization and setup
├── 2_DataUpload.gs       # Bulk upload functions
├── 3_Auth.gs             # Authentication and RBAC
├── 4_Exams.gs            # Exam management
├── 5_Marks.gs            # Marks entry and calculations
├── 6_Analytics.gs        # Performance analytics
├── 7_Reports.gs          # Report generation and exports
└── Dashboard.html        # Complete UI
```

## Key Features

### Implemented
1. **Role-Based Access Control**
   - Admin: rishisans83@gmail.com, mvmseniors@gmail.com
   - Teacher: Any email in Teachers sheet
   - Server-side filtering for all data

2. **Student Management**
   - Add individual students
   - Bulk upload via CSV
   - Class/Section/Stream organization

3. **Teacher Management**
   - Google account-based login
   - Subject/Class/Section assignments
   - Bulk upload via CSV

4. **Exam Management**
   - Create exams with type, max marks, weightage
   - Lock/Unlock exams
   - Academic year filtering

5. **Marks Entry**
   - Individual marks entry
   - Bulk marks upload
   - Grade calculation (numeric ranges: 91-100, 81-90, etc.)

6. **Analytics**
   - Subject performance
   - Class performance
   - Teacher performance
   - Weak students identification
   - Toppers list

7. **Report Generation**
   - Individual student report cards (PDF)
   - Class-wise report cards (PDF batch)
   - **Download Class-wise Marks (Excel + CSV)** - NEW

8. **Admin Tools**
   - Archive year data
   - Promote students (placeholder)

### Download Marks Feature (Latest - Dec 2025)
- Admin can download complete class-wise marksheet for any exam
- Includes all subjects as columns with marks for each student
- Two download formats:
  - **Excel (Google Sheets)**: Opens in new tab, can export to .xlsx
  - **CSV**: Direct browser download
- Includes: Roll No, Student Name, Subject marks, Total, Max, Percentage, Range
- Class averages calculated per subject

## Database Schema (Sheet Headers)
- **Students**: StudentID, Name, Class, Section, Stream, RollNo, ParentEmail, Phone, JoinDate, Status
- **Teachers**: TeacherID, Name, Subject, Classes, Sections, Email, Phone, JoinDate, Status
- **Exams**: ExamID, ExamName, ExamType, Class, MaxMarks, Weightage, StartDate, EndDate, Locked, CreatedBy, CreatedAt, AcademicYear
- **Marks_Master**: EntryID, StudentID, StudentName, Subject, SubjectCode, TeacherID, TeacherName, ExamID, ExamName, Class, Section, MaxMarks, MarksObtained, Percentage, Grade, UpdatedAt, UpdatedBy, AcademicYear

## API Functions (Server-side)
- `downloadClassExamMarks(classNum, section, examId)` - Returns Google Sheets URL
- `downloadClassExamMarksCSV(classNum, section, examId)` - Returns CSV data string

## Future Tasks (Backlog)
- P1: Auto-promote students to next class at year end
- P1: Archive previous year data (more robust system)

## Deployment Notes
Refer to `0_DeploymentGuide.gs` for detailed deployment instructions including:
- Setting up as web app
- Managing user access permissions
- Configuring admin emails
