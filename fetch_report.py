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
    "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Cookie":          COOKIE,
    "x-kaptrace-id":   "c05e33d4-a013-421e-a899-8b0bd1d85e0c##1000171###181179",
}
 
PAYLOAD = {
    "reportType":                "T",
    "cmId":                      "1000171",
    "offlinePath":               "1",
    "isS3":                      "1",
    "limit":                     "2",
    "substatus":                 "",
    "query_type":                "C",
    "templateId":                "",
    "start_date":                start_date,
    "start_time":                "00:00:00",
    "end_date":                  end_date,
    "end_time":                  "23:55:55",
    "format":                    "CSVEXCEL",
    "folder_ids":                "",
    "employee_ids":              "",
    "queue_keys":                "",
    "filterType":                "",
    "activeType":                "",
    "note_type":                 "",
    "all_pending_exclude_archive": "",
    "ticket_ids":                "",
    "ticketType":                "",
    "keyword":                   "",
}
 
# ── Step 1: Trigger report generation ────────────────────────────────────────
print("Triggering report generation...")
resp = requests.post(KAPTURE_URL, headers=HEADERS, data=PAYLOAD, timeout=60)
 
# Always print full response details for debugging
print(f"Status code : {resp.status_code}")
print(f"Content-Type: {resp.headers.get('Content-Type', 'not set')}")
print(f"Response body (first 1000 chars):\n{resp.text[:1000]}")
 
# Check for auth failure / redirect
if resp.status_code in (401, 403):
    raise Exception("Auth failed — cookie is invalid or expired. Update KAPTURE_COOKIE secret.")
 
if resp.status_code == 302 or "login" in resp.url.lower():
    raise Exception(f"Redirected to login page ({resp.url}) — cookie expired.")
 
if not resp.text.strip():
    raise Exception(
        "Empty response from Kapture API.\n"
        "Most likely cause: cookie is expired or was not copied fully.\n"
        "Fix: go to Chrome → DevTools → copy the full Cookie header again → update KAPTURE_COOKIE secret."
    )
 
# ── Step 2: Parse response ────────────────────────────────────────────────────
try:
    result = resp.json()
except Exception:
    # Not JSON — maybe a direct CSV download?
    content_type = resp.headers.get("Content-Type", "")
    if "text/csv" in content_type or "application/octet-stream" in content_type:
        print("Response is a direct CSV download — processing immediately.")
        rows = list(csv.reader(io.StringIO(resp.content.decode("utf-8-sig"))))
        result = None   # skip polling, jump straight to Sheets
    else:
        raise Exception(
            f"Could not parse API response as JSON.\n"
            f"Status: {resp.status_code}\n"
            f"Content-Type: {content_type}\n"
            f"Body: {resp.text[:500]}"
        )
else:
    rows = None   # will be populated after polling/download
    print(f"Parsed JSON response: {result}")
 
# ── Step 3: Get download URL (if not direct CSV) ──────────────────────────────
download_url = None
 
if rows is None:
    # Pattern A: URL directly in response
    for key in ("fileUrl", "file_url", "url", "downloadUrl", "download_url", "s3Url", "s3_url", "data"):
        val = result.get(key)
        if val and isinstance(val, str) and val.startswith("http"):
            download_url = val
            print(f"Found download URL under key '{key}': {download_url}")
            break
 
    # Pattern B: nested under result.data or result.response
    if not download_url:
        for wrapper in ("data", "response", "result"):
            nested = result.get(wrapper)
            if isinstance(nested, dict):
                for key in ("fileUrl", "file_url", "url", "downloadUrl", "s3Url"):
                    val = nested.get(key)
                    if val and isinstance(val, str) and val.startswith("http"):
                        download_url = val
                        print(f"Found download URL at {wrapper}.{key}: {download_url}")
                        break
 
    # Pattern C: job/task ID — poll status endpoint
    if not download_url:
        job_id = (result.get("jobId") or result.get("job_id")
                  or result.get("taskId") or result.get("task_id"))
        if job_id:
            status_url = f"https://fitelo.kapturecrm.com/ms/kreport/generic-report/status/{job_id}"
            print(f"Polling job {job_id} at {status_url}")
            for attempt in range(18):   # up to 3 minutes
                time.sleep(10)
                s = requests.get(status_url, headers=HEADERS, timeout=30)
                print(f"  poll {attempt+1}: status={s.status_code} body={s.text[:200]}")
                try:
                    sdata = s.json()
                except Exception:
                    continue
                for key in ("fileUrl", "file_url", "url", "downloadUrl", "s3Url"):
                    val = sdata.get(key)
                    if val and isinstance(val, str) and val.startswith("http"):
                        download_url = val
                        break
                if download_url:
                    break
            else:
                raise TimeoutError("Report not ready after 3 minutes of polling.")
 
    if not download_url:
        raise ValueError(
            f"Could not find a download URL in the API response.\n"
            f"Full response: {json.dumps(result, indent=2)}\n\n"
            "Please share this output so the script can be updated with the correct key name."
        )
 
# ── Step 4: Download CSV ──────────────────────────────────────────────────────
if rows is None:
    print(f"Downloading CSV from: {download_url}")
    csv_resp = requests.get(download_url, timeout=120)
    csv_resp.raise_for_status()
    content = csv_resp.content.decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(content)))
 
if not rows:
    raise ValueError("Downloaded CSV is empty.")
 
print(f"CSV rows (including header): {len(rows)}")
 
# ── Step 5: Push to Google Sheets ────────────────────────────────────────────
creds_info = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
scopes     = ["https://www.googleapis.com/auth/spreadsheets"]
creds      = Credentials.from_service_account_info(creds_info, scopes=scopes)
gc         = gspread.authorize(creds)
 
sheet_id   = os.environ["GOOGLE_SHEET_ID"]
sh         = gc.open_by_key(sheet_id)
 
# Write to a tab named after today's date e.g. "2026-04-06"
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
 
# Also keep a 'Latest' tab always up to date
try:
    latest_ws = sh.worksheet("Latest")
    latest_ws.clear()
except gspread.exceptions.WorksheetNotFound:
    latest_ws = sh.add_worksheet(title="Latest", rows=max(len(rows)+10, 100), cols=max(len(rows[0])+5, 26))
 
latest_ws.update(rows, value_input_option="RAW")
print("✅ 'Latest' tab also updated.")
