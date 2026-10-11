"""UP PCC PDF Downloader: downloads the UP CCTNS character-certificate PDF
for every row of an Excel file (ELEMENT ID + APPLICATION columns).

Uses the login session of your own browser: log in to the CCTNS portal in
Edge/Chrome, copy the AppID cookie (F12 > Application > Cookies) and paste it
when asked. Runs one row at a time. Running again with the same output folder
skips PDFs that were already downloaded."""
# imports
import argparse
import ssl
import requests
import re
import hashlib
import os
import time
import openpyxl
from requests.adapters import HTTPAdapter

# constants
BASE_URL = "https://cctnsup.gov.in/CitizenPortal/"
SEARCH_URL = BASE_URL + "SearchStatus/SearchStatus.aspx"
PDF_URL = BASE_URL + "PopFileUpload.aspx"
P = "ctl00$ContentPlaceHolder1$"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0"
DELAY = 0  # seconds between rows, so the firewall doesn't block us

#LegacyTLSAdapter
class LegacyTLSAdapter(HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.options |= ssl.OP_LEGACY_SERVER_CONNECT
        kwargs["ssl_context"] = ctx
        return super().init_poolmanager(*args, **kwargs)

# get_hidden
def get_hidden(html):
    pairs = re.findall(r'<input type="hidden" name="(.*?)" id=".*?" value="(.*?)"', html)
    return dict(pairs)

# search / print form (same fields, only the button changes)
def make_form(html, application_no, button_name, button_value):
    form = dict(get_hidden(html))
    form["__EVENTTARGET"] = ""
    form["__EVENTARGUMENT"] = ""
    form[P + "ddlServiceType"] = "3"
    form[P + "ddlServiceRequestType"] = "9"
    form[P + "ddlYearSR"] = ""
    form[P + "txtServiceNo"] = application_no
    form[P + button_name] = button_value
    return form

# command-line arguments (Argus passes: excel_path --out folder --workers n)
parser = argparse.ArgumentParser(description="Download UP CCTNS character-certificate PDFs for each row of an Excel file.")
parser.add_argument("excel_path", help="Excel file with ELEMENT ID and APPLICATION columns")
parser.add_argument("--out", default="up_pcc_pdf_downloader_output", help="Output folder")
parser.add_argument("--workers", type=int, default=1, help="Accepted for Argus compatibility; always runs 1 at a time (one login session)")
args = parser.parse_args()
excel_path = args.excel_path
OUT_DIR = args.out

# user input
print("[?25h", end="")  # show the cursor (Argus hides it while a tool runs)
app_id = input("AppID cookie value : ").strip()

# excel setup
excel = openpyxl.load_workbook(excel_path)
sheet = excel.worksheets[0]

header = [str(cell.value).strip().upper() for cell in sheet[1]]
if "ELEMENT ID" not in header or "APPLICATION" not in header:
    print("Error: required columns ELEMENT ID and APPLICATION not found. Columns in the file:", header)
    exit()
id_col = header.index("ELEMENT ID")
app_col = header.index("APPLICATION")

total_start = time.perf_counter()

# session setup
session = requests.Session()
session.mount("https://cctnsup.gov.in", LegacyTLSAdapter())
session.headers["Referer"] = SEARCH_URL
session.headers["User-Agent"] = USER_AGENT
session.cookies.set("AppID", app_id, domain="cctnsup.gov.in")

# GET, opening the search page
print("Opening search page", end=" ... ", flush=True)
r = session.get(SEARCH_URL, timeout=30)
if r.status_code == 200 and "ddlServiceType" in r.text:
    print("OK")
    # every page already has some alert(...) lines in its template JavaScript
    # (e.g. a commented-out alert("Internal")); remember them so they aren't
    # mistaken for a real popup later
    template_alerts = set(re.findall(r"alert\(['\"](.*?)['\"]\)", r.text))
elif r.status_code == 403:
    print(f"Failed (HTTP 403): blocked by the portal firewall. Wait a few minutes, or increase DELAY (currently {DELAY} s).")
    exit()
else:
    print(f"Failed (HTTP {r.status_code}): session expired or not logged in. Log in again and use a new AppID. Landed on: {r.url}")
    exit()

# POST 2 selecting the service type
print("Selecting service type", end=" ... ", flush=True)
form = dict(get_hidden(r.text))
form["__EVENTTARGET"] = P + "ddlServiceType"
form["__EVENTARGUMENT"] = ""
form[P + "ddlServiceType"] = "3"
form[P + "txtServiceNo"] = ""
r = session.post(SEARCH_URL, data=form, timeout=30)
if "चरित्र" in r.text:
    print("OK")
else:
    print(f"Failed (HTTP {r.status_code}): the request-type list did not load. Landed on: {r.url}")
    exit()

os.makedirs(OUT_DIR, exist_ok=True)
saved = 0
last_pdf_hash = None
skipped = 0
failed = []

# loop: one row of the Excel at a time
for row in sheet.iter_rows(min_row=2, values_only=True):
    if row[app_col] is None:
        continue
    element_id = str(row[id_col]).strip()
    application_no = str(row[app_col]).strip()
    file_path = os.path.join(OUT_DIR, f"{element_id}_{application_no}.pdf")
    row_start = time.perf_counter()
    print(element_id, application_no, end=" ... ", flush=True)

    # already downloaded in an earlier run -> skip
    if os.path.exists(file_path):
        print("Skipped: already downloaded")
        skipped += 1
        continue

    # POST 3 खोजें
    r = session.post(SEARCH_URL, data=make_form(r.text, application_no, "btnViewStatus", "खोजें"), timeout=30)
    if r.status_code == 403:
        print(f"Stopped (HTTP 403): blocked by the portal firewall. Wait a few minutes, or increase DELAY (currently {DELAY} s), then run again. Downloaded files will be skipped.")
        exit()
    if "ddlServiceType" not in r.text:
        print(f"Stopped (HTTP {r.status_code}): session expired. Log in again and run with a new AppID. Downloaded files will be skipped.")
        exit()
    if "btnPrint" not in r.text:
        print("Failed: no Print button (application not found or not approved)")
        failed.append((element_id, application_no, "No Print button (not found or not approved)"))
        time.sleep(DELAY)
        continue

    # POST 4 प्रिंट
    r = session.post(SEARCH_URL, data=make_form(r.text, application_no, "btnPrint", "प्रिंट"), timeout=30)
    popup = [a for a in re.findall(r"alert\(['\"](.*?)['\"]\)", r.text) if a not in template_alerts]
    if popup:
        print("Failed: portal message:", popup[0])
        failed.append((element_id, application_no, "Portal message: " + popup[0]))
        time.sleep(DELAY)
        continue
    # when print really worked, the server's reply opens the PDF window:
    # window.open('../PopFileUpload.aspx', ...). No such line = no new PDF for this row.
    if "PopFileUpload.aspx" not in r.text:
        print("Failed: Print did not open the PDF window (no certificate generated)")
        failed.append((element_id, application_no, "Print did not open the PDF window"))
        time.sleep(DELAY)
        continue

    # PDF GET
    pdf = session.get(PDF_URL, timeout=60)
    if not pdf.content.startswith(b"%PDF"):
        print("Failed: the portal returned an empty page instead of the PDF (certificate file not available)")
        failed.append((element_id, application_no, "Certificate file not available on the portal"))
        time.sleep(DELAY)
        continue
    # the server's file name doesn't contain the application number, so the
    # wrong-PDF check is: it must not be the exact same file as the last row's
    pdf_hash = hashlib.md5(pdf.content).hexdigest()
    if pdf_hash == last_pdf_hash:
        print("Failed: the portal returned the previous row's PDF again (not saved)")
        failed.append((element_id, application_no, "Portal returned the previous row's PDF"))
        time.sleep(DELAY)
        continue
    last_pdf_hash = pdf_hash

    # saving the pdf
    with open(file_path, "wb") as f:
        f.write(pdf.content)
    saved += 1
    served_name = pdf.headers.get("Content-Disposition", "").split("filename=")[-1]
    print(f"Saved ({len(pdf.content) // 1024} KB, {time.perf_counter() - row_start:.2f} s) | portal file name: {served_name}")

    time.sleep(DELAY)

# summary
print()
print("Summary")
print("Saved   :", saved)
print("Skipped :", skipped, "(already downloaded)")
print("Failed  :", len(failed))
for element_id, application_no, reason in failed:
    print("   ", element_id, application_no, "->", reason)
print("Folder  :", os.path.abspath(OUT_DIR))
print(f"Time    : {time.perf_counter() - total_start:.2f} s")
