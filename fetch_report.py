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
KAPTURE_URL     = "https://fitelo.kapturecrm.com/ms/kreport/generic-report/generate"
REPORT_LIST_URL = "https://fitelo.kapturecrm.com/ms/kreport/generic-report/list"
COOKIE          = os.environ["KAPTURE_COOKIE"]
 
HEADERS = {
    "Content-Type":  "application/x-www-form-urlencoded",
    "Accept":        "application/json, text/plain, */*",
    "Origin":        "https://fitelo.kapturecrm.com",
    "Referer":       "https://fitelo.kapturecrm.com/nui/report",
    "User-Agent":    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Cookie":        COOKIE,
    "x-kaptrace-id": "c05e33d4-a013-421e-a899-8b0bd1d85e0c##1000171###181179",
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
 
# ── Helper: search report list for a completed download URL ──────────────────
def find_download_url_from_list():
    print("Checking report list for a completed job...")
    try:
        r = requests.get(REPORT_LIST_URL, headers={**HEADERS, "Content-Type": "application/json"}, timeout=30)
        print(f"  List status: {r.status_code} | body[:500]: {r.text[:500]}")
        data = r.json()
        items = data if isinstance(data, list) else (
            data.get("data") or data.get("reports") or data.get("list") or []
        )
        for item in items:
            url = (item.get("fileUrl") or item.get("file_url") or
                   item.get("url") or item.get("downloadUrl") or
                   item.get("s3Url") or item.get("s3_url") or "")
            if url and url.startswith("http"):
                print(f"  Found URL in list: {url}")
                return url
    except Exception as e:
        print(f"  Could not fetch report list: {e}")
    return None
 
# ── Step 1: Trigger report generation ────────────────────────────────────────
print("Triggering report generation...")
resp = requests.post(KAPTURE_URL, headers=HEADERS, data=PAYLOAD, timeout=60)
 
print(f"Status code : {resp.status_code}")
print(f"Content-Type: {resp.headers.get('Content-Type', 'not set')}")
print(f"Response body: {resp.text[:500]}")
 
duplicate_request = "duplicate" in resp.text.lower() or "already processing" in resp.text.lower()
 
# ── Step 2: Parse response or handle duplicate ────────────────────────────────
download_url = None
rows = None
 
if duplicate_request:
    print("Duplicate request detected — polling report list for existing job result...")
    for attempt in range(18):
        time.sleep(10)
        print(f"  Attempt {attempt + 1}/18...")
        download_url = find_download_url_from_list()
        if download_url:
            break
    if not download_url:
        raise TimeoutError("Existing report job did not complete within 3 minutes. Try again shortly.")
else:
    try:
        result = resp.json()
        print(f"Parsed JSON: {result}")
    except Exception:
        content_type = resp.headers.get("Content-Type", "")
        if "text/csv" in content_type or "application/octet-stream" in content_type:
            print("Direct CSV response — processing immediately.")
            rows = list(csv.reader(io.StringIO(resp.content.decode("utf-8-sig"))))
            result = None
        else:
            raise Exception(f"Unexpected response:\nStatus: {resp.status_code}\nBody: {resp.text[:500]}")
 
    if rows is None and result is not None:
        for key in ("fileUrl", "file_url", "url", "downloadUrl", "download_url", "s3Url", "s3_url"):
            val = result.get(key)
            if val and isinstance(val, str) and val.startswith("http"):
                download_url = val
                print(f"Found URL under key '{key}': {download_url}")
                break
 
        if not download_url:
            for wrapper in ("data", "response", "result"):
                nested = result.get(wrapper)
                if isinstance(nested, dict):
                    for key in ("fileUrl", "file_url", "url", "downloadUrl", "s3Url"):
                        val = nested.get(key)
                        if val and isinstance(val, str) and val.startswith("http"):
                            download_url = val
                            print(f"Found URL at {wrapper}.{key}: {download_url}")
                            break
 
        if not download_url:
            job_id = (result.get("jobId") or result.get("job_id") or
                      result.get("taskId") or result.get("task_id"))
            if job_id:
                status_url = f"https://fitelo.kapturecrm.com/ms/kreport/generic-report/status/{job_id}"
                print(f"Polling job {job_id}...")
                for attempt in range(18):
                    time.sleep(10)
                    s = requests.get(status_url, headers=HEADERS, timeout=30)
                    print(f"  poll {attempt+1}: {s.status_code} | {s.text[:200]}")
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
                    raise TimeoutError("Report job did not complete within 3 minutes.")
 
        if not download_url:
            print("No URL found in response — checking report list as fallback...")
            for attempt in range(12):
                time.sleep(10)
                download_url = find_download_url_from_list()
                if download_url:
                    break
 
        if not download_url:
            raise ValueError(
                f"Could not find a download URL anywhere.\n"
                f"Full API response: {json.dumps(result, indent=2)}\n\n"
                "Paste this output so the script can be updated."
            )
 
# ── Step 3: Download CSV ──────────────────────────────────────────────────────
if rows is None:
    print(f"Downloading CSV from: {download_url}")
    csv_resp = requests.get(download_url, timeout=120)
    csv_resp.raise_for_status()
    content = csv_resp.content.decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(content)))
 
if not rows:
    raise ValueError("Downloaded CSV is empty.")
 
print(f"CSV rows (including header): {len(rows)}")
 
# ── Step 4: Push to Google Sheets ────────────────────────────────────────────
creds_info = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
scopes     = ["https://www.googleapis.com/auth/spreadsheets"]
creds      = Credentials.from_service_account_info(creds_info, scopes=scopes)
gc         = gspread.authorize(creds)
 
sheet_id = os.environ["GOOGLE_SHEET_ID"]
sh       = gc.open_by_key(sheet_id)
 
tab_name = today.strftime("%Y-%m-%d")
try:
    ws = sh.worksheet(tab_name)
    ws.clear()
    print(f"Cleared existing tab: {tab_name}")
except gspread.exceptions.WorksheetNotFound:
    ws = sh.add_worksheet(title=tab_name, rows=max(len(rows)+10, 100), cols=max(len(rows[0])+5, 26))
    print(f"Created new tab: {tab_name}")
 
ws.update(rows, value_input_option="RAW")
print(f"✅ Done! {len(rows)-1} data rows written to tab '{tab_name}'.")
 
try:
    latest_ws = sh.worksheet("Latest")
    latest_ws.clear()
except gspread.exceptions.WorksheetNotFound:
    latest_ws = sh.add_worksheet(title="Latest", rows=max(len(rows)+10, 100), cols=max(len(rows[0])+5, 26))
 
latest_ws.update(rows, value_input_option="RAW")
print("✅ 'Latest' tab also updated.")
