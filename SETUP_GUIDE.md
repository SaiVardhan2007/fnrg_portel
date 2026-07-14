# FNRG Portal — Setup Guide

Static PWA frontend + Google Sheets as database via Apps Script Web App.

## Step 1: Create the Google Sheet & paste the script

1. Create a **new blank Google Sheet** (name it anything, e.g. `FNRG Portal DB`).
2. **Extensions → Apps Script** → delete the placeholder code.
3. Copy everything from [google-apps-script.js](google-apps-script.js) and paste it in. Save.

## Step 2: Deploy as a Web App

1. Click **Deploy → New deployment**.
2. Gear icon → select **Web app**.
3. Set:
   - **Execute as:** `Me`
   - **Who has access:** `Anyone`  ← required, otherwise the frontend gets blocked
4. Click **Deploy**, authorize when asked, and copy the **Web app URL** (`…/exec`).

> After any future script change: **Deploy → Manage deployments → edit pencil → Version: New version → Deploy** (same URL stays).

## Step 3: Connect the frontend

Open [app.js](app.js) and replace line 7:

```js
const API_URL = "PASTE_YOUR_WEB_APP_URL_HERE";
```

## Step 4: First run (auto-setup)

Just open the Web app URL once in your browser (or load the site). The script **auto-creates and seeds** two tabs:

| Tab | Purpose | Columns |
|---|---|---|
| `Users` | Who can log in | Name, Phone Number |
| `Thursday Calling` | Page data | Name, Phone Number, W/S, Sessions Attended |

Sample login phone numbers seeded: `9876543210` (Sai Vardhan), `9876543211` (Dushmanth), `9876543212` (Srinivas). Edit the `Users` tab to manage real users — no code changes needed.

## Step 5: Run the frontend locally

The service worker needs http(s), so serve the folder instead of double-clicking index.html:

```bash
cd ~/fnrg_portel && python3 -m http.server 8787
```

Open http://localhost:8787 — log in with any phone from the `Users` tab.

> Note: port 8080 is often occupied by other software (e.g. Docker), which shows its own page instead — that's why we use 8787.

## Step 6: Host it (free) for the PWA install

Push the folder to GitHub and enable **GitHub Pages** (or drop it on Netlify). Open the site on a phone → browser menu → **Add to Home Screen** → it installs as an app.

## How the sheet is managed day-to-day

- **Add/remove a login user:** edit rows in the `Users` tab.
- **Add/remove a contact:** edit rows in the `Thursday Calling` tab.
- **Edits from the app** (W/S dropdown, Sessions) save straight into `Thursday Calling` — protected by a script lock so simultaneous users can't corrupt rows, and reads are cached for 60s so 20–30 users don't slow the sheet down.
