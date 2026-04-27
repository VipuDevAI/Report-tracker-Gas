# MVM Report Tracker - Test Credentials

## Admin Emails (login via Dashboard custom HTML login form)
- rishisans83@gmail.com
- mvmseniors@gmail.com
- anithasivanesan4604@gmail.com

## Teacher Login
Any email present in the `Teachers` sheet (status = Active).
The teacher must first be added via Admin → Teachers_Master sheet, then synced via "Sync Teachers" button in Dashboard.

## Notes
- No password is required (custom email-based login only).
- Custom HTML login form stores email in `PropertiesService.getUserProperties().setProperty('loggedInEmail', email)`.
- Admin emails are hardcoded in `3_Auth.gs` → `ADMIN_EMAIL_LIST`.
