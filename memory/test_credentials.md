# MVM Report Tracker - Test Credentials

## Authentication (NEW: Email + Password)

The system now uses **email + password** authentication. On first login, every user must set their own password.

### Admin Emails (configured in `3_Auth.gs` → ADMIN_EMAIL_LIST)
- rishisans83@gmail.com
- mvmseniors@gmail.com
- anithasivanesan4604@gmail.com

**Super-admin** (only one who can UNFREEZE a finalized year): the FIRST email in the list → `rishisans83@gmail.com`

### Teacher Logins
Any email present in the `Teachers` sheet (status = Active) can log in. Add teachers via Admin → Teachers_Master sheet, then click "Sync Teachers" in the dashboard.

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
