 Kapture CRM → Google Sheets Automation
 
Runs daily at **9:00 AM IST** via GitHub Actions.  
Fetches the last 30 days of report data from Kapture CRM and writes it to Google Sheets.
 
---
 
## Setup (one-time, ~15 minutes)
 
### 1. Create a GitHub repository
 
- Go to https://github.com/new
- Name it e.g. `kapture-sync`
- Keep it **Private**
- Click **Create repository**
 
Upload both files into the repo:
```
fetch_report.py
.github/workflows/fetch_report.yml
```
 
---
 
### 2. Get your Kapture Cookie (refresh monthly)
 
1. Open Chrome → go to `fitelo.kapturecrm.com`
2. Press **F12** → Network tab
3. Trigger any report → click the `generate` request
4. Under **Headers** → find the `Cookie:` line
5. Copy the **entire cookie string** (it's long)
 
---
 
### 3. Set up Google Service Account
 
1. Go to https://console.cloud.google.com
2. Create a new project (or use existing)
3. Enable **Google Sheets API**:
   - APIs & Services → Enable APIs → search "Google Sheets API" → Enable
4. Create a Service Account:
   - APIs & Services → Credentials → Create Credentials → Service Account
   - Name it `kapture-sync` → click Done
5. Click the service account → **Keys** tab → Add Key → JSON
6. Download the JSON file — keep it safe!
7. Copy the `client_email` from the JSON (looks like `kapture-sync@your-project.iam.gserviceaccount.com`)
 
---
 
### 4. Share your Google Sheet with the service account
 
1. Open your Google Sheet
2. Click **Share**
3. Paste the `client_email` from step 3
4. Give it **Editor** access
5. Click Send
 
---
 
### 5. Add GitHub Secrets
 
Go to your GitHub repo → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**
 
Add these 3 secrets:
 
| Secret Name | Value |
|---|---|
| `KAPTURE_COOKIE` | The full cookie string from Step 2 |
| `GOOGLE_CREDENTIALS_JSON` | The entire contents of the JSON file from Step 3 |
| `GOOGLE_SHEET_ID` | The ID from your Sheet URL (see below) |
 
**Finding your Sheet ID:**
```
https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit
                                        ^^^^^^^^^^^^^^^^^^^^^^^^
```
 
---
 
## What it does daily
 
1. Calculates date range: today minus 30 days → today
2. Calls the Kapture API to generate a CSV report
3. Downloads the CSV from S3
4. Writes data to Google Sheets:
   - A tab named **today's date** (e.g. `2026-04-06`) — historical archive
   - A tab named **Latest** — always has the most recent run
 
---
 
## Manual trigger
 
Go to your GitHub repo → **Actions** tab → **Kapture CRM → Google Sheets Daily Sync** → **Run workflow**
 
---
 
## Cookie expiry warning
 
Kapture session cookies expire periodically (usually every 30–60 days).  
When the job fails, update `KAPTURE_COOKIE` in GitHub Secrets with a fresh cookie.
 
You'll get an email from GitHub Actions when the job fails.
