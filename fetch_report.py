import os
import io
import csv
import time
import requests
import gspread
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import json
 
# ── Date range: last 30 days ──────────────────────────────────────────────────
today = datetime.now()
start = today - timedelta(days=30)
start_date = start.strftime("%m/%d/%Y")
end_date   = today.strftime("%m/%d/%Y")
 
print(f"Fetching report: {start_date} → {end_date}")
 
# ── Kapture API config ────────────────────────────────────────────────────────
KAPTURE_URL = "https://fitelo.kapturecrm.com/ms/kreport/generic-report/generate"
COOKIE      = os.environ["KAPTURE_COOKIE"]
 
HEADERS = {
    "Content-Type":    "application/x-www-form-urlencoded",
    "Accept":          "application/json, text/plain, */*",
    "Origin":          "https://fitelo.kapturecrm.com",
    "Referer":         "https://fitelo.kapturecrm.com/nui/report",
    "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Cookie":          COOKIE,
}
 
PAYLOAD = {
    "reportType":               "T",
    "cmId":                     "1000171",
    "offlinePath":              "1",
    "isS3":                     "1",
    "limit":                    "2",
    "substatus":                "",
    "query_type":               "C",
    "templateId":               "",
    "start_date":               start_date,
    "start_time":               "00:00:00",
    "end_date":                 end_date,
    "end_time":                 "23:55:55",
    "format":                   "CSVEXCEL",
    "folder_ids":               "",
    "employee_ids":             "",
    "queue_keys":               "",
    "filterType":               "",
    "activeType":               "",
    "note_type":                "",
    "all_pending_exclude_archive": "",
    "ticket_ids":               "",
    "ticketType":               "",
    "keyword":                  "",
}
 
# ── Step 1: Trigger report generation ────────────────────────────────────────
print("Triggering report generation...")
resp = requests.post(KAPTURE_URL, headers=HEADERS, data=PAYLOAD, timeout=60)
resp.raise_for_status()
result = resp.json()
print("API response:", result)
 
# ── Step 2: Poll for download URL (Kapture generates async S3 reports) ────────
# Kapture returns a file URL or a job ID depending on the response structure.
# Handles both patterns gracefully.
 
download_url = None
 
# Pattern A: direct URL in response
for key in ("fileUrl", "file_url", "url", "downloadUrl", "download_url", "s3Url"):
    if result.get(key):
        download_url = result[key]
        break
 
# Pattern B: job/task id — poll status endpoint
if not download_url:
    job_id = result.get("jobId") or result.get("job_id") or result.get("taskId")
    if job_id:
        status_url = f"https://fitelo.kapturecrm.com/ms/kreport/generic-report/status/{job_id}"
        print(f"Polling job {job_id}...")
        for attempt in range(12):   # up to ~2 minutes
            time.sleep(10)
            s = requests.get(status_url, headers=HEADERS, timeout=30)
            s.raise_for_status()
            sdata = s.json()
            print(f"  poll {attempt+1}: {sdata}")
            for key in ("fileUrl", "file_url", "url", "downloadUrl", "s3Url"):
                if sdata.get(key):
                    download_url = sdata[key]
                    break
            if download_url:
                break
        else:
            raise TimeoutError("Report not ready after 2 minutes.")
 
if not download_url:
    # Dump full response so you can inspect and adjust the key name above
    raise ValueError(f"Could not find download URL in response: {result}")
 
# ── Step 3: Download CSV ──────────────────────────────────────────────────────
print(f"Downloading CSV from: {download_url}")
csv_resp = requests.get(download_url, timeout=120)
csv_resp.raise_for_status()
 
content = csv_resp.content.decode("utf-8-sig")   # strips BOM if present
reader  = csv.reader(io.StringIO(content))
rows    = list(reader)
 
if not rows:
    raise ValueError("Downloaded CSV is empty.")
 
print(f"CSV rows (including header): {len(rows)}")
 
# ── Step 4: Push to Google Sheets ────────────────────────────────────────────
creds_info = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
scopes     = ["https://www.googleapis.com/auth/spreadsheets"]
creds      = Credentials.from_service_account_info(creds_info, scopes=scopes)
gc         = gspread.authorize(creds)
 
sheet_id   = os.environ["GOOGLE_SHEET_ID"]
sh         = gc.open_by_key(sheet_id)
 
# Write to a tab named after today's date, e.g. "2026-04-06"
tab_name   = today.strftime("%Y-%m-%d")
try:
    ws = sh.worksheet(tab_name)
    ws.clear()
    print(f"Cleared existing tab: {tab_name}")
except gspread.exceptions.WorksheetNotFound:
    ws = sh.add_worksheet(title=tab_name, rows=max(len(rows)+10, 100), cols=max(len(rows[0])+5, 26))
    print(f"Created new tab: {tab_name}")
 
ws.update(rows, value_input_option="RAW")
print(f"✅ Done! {len(rows)-1} data rows written to tab '{tab_name}'.")
 
# ── Step 5: Also keep a 'Latest' tab always up to date ───────────────────────
try:
    latest_ws = sh.worksheet("Latest")
    latest_ws.clear()
except gspread.exceptions.WorksheetNotFound:
    latest_ws = sh.add_worksheet(title="Latest", rows=max(len(rows)+10, 100), cols=max(len(rows[0])+5, 26))
 
latest_ws.update(rows, value_input_option="RAW")
print("✅ 'Latest' tab also updated.")
