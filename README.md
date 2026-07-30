# Captains — National Day Crew Form

نموذج تسجيل بيانات طاقم مشروع اليوم الوطني.

## Files

| File | Purpose |
|---|---|
| `index.html` | The form itself. Served by GitHub Pages. |
| `captains-logo-white.png` | Logo used in the header. |
| `captains_google_sheet_apps_script.js` | Apps Script code that lives in the Google Sheet, not on this site. |

## How it works

The form POSTs JSON to a Google Apps Script Web App, which appends a row to the
Captains spreadsheet. No submitted data is ever stored in this repository.

## Updating the Apps Script

The `.js` file here is a copy for reference. To change the live behaviour:

1. Open the spreadsheet → **Extensions → Apps Script**
2. Paste the updated code
3. **Deploy → Manage deployments → Edit → New version → Deploy**
4. Keep access set to **Anyone** so the public form can submit

If you create a *new* deployment instead of a new version, the URL changes and
`SHEET_WEB_APP_URL` in `index.html` must be updated to match.
