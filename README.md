# GCal Reminders — Excel Online Add-in

Push marketing follow-up dates from Excel Online directly to Google Calendar.

---

## Files in this project

| File | Purpose |
|---|---|
| `manifest.xml` | Tells Excel about the add-in (upload this to M365 admin) |
| `taskpane.html` | Main UI shown in the Excel side panel |
| `taskpane.css` | Styles |
| `taskpane.js` | All logic — reads Excel, maps columns, pushes to Google Calendar |
| `auth-dialog.html` | Handles Google OAuth sign-in popup |
| `commands.html` | Required stub for the manifest |

---

## Setup — 3 stages

### Stage 1 — Host on GitHub Pages (free, ~5 min)

1. Create a free account at https://github.com if you don't have one
2. Create a new repository — name it `excel-gcal-addin`
3. Upload all files from this folder into the repository
4. Go to **Settings → Pages → Branch: main → Save**
5. Your add-in is now live at:
   `https://YOUR-GITHUB-USERNAME.github.io/excel-gcal-addin/`

6. Open `manifest.xml` and replace every instance of `YOUR-GITHUB-USERNAME` with your actual GitHub username.
   Then re-upload the updated `manifest.xml` to GitHub.

---

### Stage 2 — Get a free Google Calendar API Client ID (~10 min)

1. Go to https://console.cloud.google.com
2. Create a new project (top-left dropdown → New Project)
3. Go to **APIs & Services → Library**
4. Search for **Google Calendar API** → Enable it
5. Go to **APIs & Services → OAuth consent screen**
   - Choose **External**
   - Fill in App name (e.g. "GCal Reminders"), support email
   - Save and continue (skip scopes for now)
   - Add your own email as a test user
6. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**
   - Application type: **Web application**
   - Name: GCal Reminders
   - Authorised JavaScript origins — add:
     ```
     https://YOUR-GITHUB-USERNAME.github.io
     ```
   - Click Create
7. Copy the **Client ID** (looks like: `123456789-abc.apps.googleusercontent.com`)

> **Cost:** The Google Calendar API free tier allows 1,000,000 calls per day.
> Set a $0 spending cap at Billing → Budgets & Alerts to guarantee zero charges.

---

### Stage 3 — Deploy to your M365 team (~2 min, done by admin)

1. Go to https://admin.microsoft.com
2. Navigate to **Settings → Integrated Apps → Upload custom apps**
3. Choose **Upload manifest file** and upload `manifest.xml`
4. Assign to **Everyone** (or specific users/groups)
5. The add-in will appear in the **Home** tab ribbon in Excel Online for all assigned users within ~24 hours (usually much faster)

---

## How team members use it

1. Open any Excel Online workbook with your marketing data
2. Click **GCal Reminders** in the Home ribbon → the side panel opens
3. **Step 1:** Paste the Google Client ID → click Connect → sign in with Google
4. **Step 2:** Click "Read active sheet"
5. **Step 3:** Confirm column mapping (auto-detected) and set your title template
6. **Step 4:** Review the events, uncheck any you don't want → Push to Google Calendar

Each team member connects their own Google account — events go into their own calendar.

---

## Title template syntax

Use `{Column Name}` to insert column values into the event title.

Examples:
- `Follow up: {Contact} @ {Company}` → "Follow up: Rahul Sharma @ Infosys"
- `Call {Contact} — {Company}` → "Call Priya Patel — TCS"
- `{Company} follow-up` → "Reliance follow-up"

Column names are case-sensitive and must match your sheet headers exactly.

---

## Troubleshooting

| Issue | Fix |
|---|---|
| Add-in not showing in ribbon | Wait up to 1 hour after admin deploys; try refreshing Excel |
| Sign-in popup blocked | Allow popups for excel.office.com in your browser |
| "Not a valid date" on a row | Make sure the date column is formatted as a date in Excel, not text |
| Duplicate events | The add-in checks for existing events on the same day — duplicates are skipped |
| Client ID error | Make sure your GitHub Pages URL is in the Authorised Origins in Google Cloud Console |
