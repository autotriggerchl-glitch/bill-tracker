# Bill Tracker

A household bill tracker: monthly, weekly and custom bills, paycheck planning, debts, history, reminders and separate logins.

- **The app:** https://autotriggerchl-glitch.github.io/bill-tracker/ — open on your phone and choose **Add to Home Screen**.
- **The data** lives in a private Google Sheet and is reached through a Google Apps Script web app (`apps-script/`). Every request needs a login, so this public page holds no bill data.

## Files
| File | What it is |
|---|---|
| `index.html` | The full-screen app (same as `apps-script/Index.html`, plus the server link and home-screen icon) |
| `manifest.webmanifest`, `*.png` | Home-screen name and icons |
| `apps-script/Code.gs`, `apps-script/Index.html` | Copy of the Apps Script project (backup / reference) |
