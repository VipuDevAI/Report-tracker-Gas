# MVM Report Tracker - Test Credentials

## Authentication (NEW: Email + Password)

The system now uses **email + password** authentication. On first login, every user must set their own password.

### Admin Emails (configured in `3_Auth.gs` → ADMIN_EMAIL_LIST)
- rishisans83@gmail.com
- mvmseniors@gmail.com
- anithasivanesan4604@gmail.com

**Super-admin** (only one who can UNFREEZE a finalized year): the FIRST email in the list → `rishisans83@gmail.com`

### Teacher Logins
Any email present in the `Teachers` sheet (status = Active, IsDeleted ≠ true) can log in. Add teachers via Admin → Teachers_Master sheet, then click "Sync Teachers" in the dashboard.

### Roles (NEW — Role column on Teachers sheet)
The Teachers sheet now has a `Role` column (column 13). Supported values:
- `ADMIN`     — full school-wide access
- `PRINCIPAL` — full READ-ONLY access (analytics, reports, audit, trash). Cannot create/update/delete.
- `WING_ADMIN` — admin powers but restricted to a class range (their wing). Set the wing via the `Classes` column, either explicit (e.g. `9,10`) or by wing name (`PRIMARY`/`SECONDARY`/`SENIOR`).
- `TEACHER`   — existing assignment-based access (subject + classes + sections). DEFAULT if column is empty.

Wing class ranges are stored in `Settings_School`:
- `Wing_Primary`   → `6,7,8`
- `Wing_Secondary` → `9,10`
- `Wing_Senior`    → `11,12`

Existing data is auto-migrated: rows with empty Role are treated as `TEACHER`. Admin emails in `ADMIN_EMAIL_LIST` continue to override the sheet (super-admin escape hatch).

### Manual test plan for the new role system
1. Add 3 users via the "Add Teacher" modal:
   - Principal (Role=PRINCIPAL, blank classes)
   - Wing Admin (Role=WING_ADMIN, Classes="9,10")
   - Teacher (Role=TEACHER, Subject=Math, Classes="9", Sections="A1,A2")
2. Each user sets their password on first login.
3. Verify:
   - Principal: dashboard loads with all data; all action buttons disabled (read-only mode); Trash & Audit visible.
   - Wing Admin: can add/edit/delete students for class 9 & 10 only; cannot touch class 11/12.
   - Wing Admin: can create exams for class 9/10 only; cannot lock exams; reports limited to wing.
   - Wing Admin: can save/delete marks for any subject in their wing's classes.
   - Teacher: only sees their assigned classes/sections; can only enter marks for their subject.

## First-Time Login Flow
1. Enter email → click Continue
2. System prompts: "First-time login: please set a password"
3. User enters new password (min 8 chars) + confirms
4. Logged in immediately, session valid for 8 hours

## Existing Login Flow
1. Enter email → click Continue
2. System prompts for password
3. Enter password → logged in

## Password Reset (Admin)
- System Admin page → Password Management → enter target email → "Reset User Password"
- Target user is forced to set a new password on next login

## Session
- Token + expiry stored in `Auth` sheet AND `localStorage`
- Default expiry: 8 hours (configurable in Settings_School → SessionDurationHours)
- Logout button in dashboard header invalidates token

## Lockout
- 5 failed password attempts → 15-minute soft lockout
