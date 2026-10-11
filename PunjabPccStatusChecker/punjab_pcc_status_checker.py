import argparse
import getpass
import io
import os
import re
import sys
import time

import openpyxl
import pytesseract
from PIL import Image
import pymupdf as fitz
import requests
from playwright.sync_api import sync_playwright

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from colors import BOLD, RESET, dim, fail, heading, ok, warn

# When stdout isn't attached to a real console, Python falls back to the
# OS default encoding (cp1252 on Windows) instead of UTF-8 -- printing
# anything outside that narrow charset then crashes with
# UnicodeEncodeError. errors="replace" makes that merely show a
# substitute character instead of taking the whole tool down.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

BASE_URL = "https://saanjh.punjabpolice.gov.in"
NEW_STATUS_URL = f"{BASE_URL}/quick/character-verification/status"
OLD_LIST_URL = f"{BASE_URL}/PPSANJH/Forms/char_entry_data.jsp"

ELEMENT_ID_COL = "ELEMENT ID"
APP_NO_COL = "PCC APPLICATION NUMBER"
LOGIN_COL = "LOGIN ID"
PASSWORD_COL = "PASSWORD"

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')

TESSERACT_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]
for _candidate in TESSERACT_CANDIDATES:
    if os.path.exists(_candidate):
        pytesseract.pytesseract.tesseract_cmd = _candidate
        break


def parse_args():
    parser = argparse.ArgumentParser(description="Check Punjab (SAANJH) PCC/Character Verification status.")
    parser.add_argument("excel_file", help="Excel file with ELEMENT ID and PCC APPLICATION NUMBER columns")
    parser.add_argument("--out", default="punjab_pcc_status_checker_output", help="Output folder")
    parser.add_argument("--mobile", default=None, help="Default mobile number for login (prompted if not given)")
    parser.add_argument("--password", default=None, help="Default password for login (prompted if not given)")
    return parser.parse_args()


def sanitize_filename(name) -> str:
    name = str(name).strip()
    return INVALID_FILENAME_CHARS.sub("_", name)


def find_column_index(header_row, target: str):
    for idx, cell in enumerate(header_row):
        if cell is not None and str(cell).strip().lower() == target.lower():
            return idx
    return None


def is_blank(value) -> bool:
    if value is None:
        return True
    return str(value).strip() in ("", "-")


def parse_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        header = next(rows_iter)
    except StopIteration:
        sys.exit("Error: the Excel file is empty.")

    eid_idx = find_column_index(header, ELEMENT_ID_COL)
    app_idx = find_column_index(header, APP_NO_COL)
    login_idx = find_column_index(header, LOGIN_COL)
    pass_idx = find_column_index(header, PASSWORD_COL)
    missing = []
    if eid_idx is None:
        missing.append(ELEMENT_ID_COL)
    if app_idx is None:
        missing.append(APP_NO_COL)
    if missing:
        sys.exit(f"Error: missing required column(s): {', '.join(missing)}")

    rows = []
    for raw_row in rows_iter:
        if raw_row is None or all(v is None for v in raw_row):
            continue
        rows.append({
            "element_id": raw_row[eid_idx],
            "app_no": raw_row[app_idx],
            "login": raw_row[login_idx] if login_idx is not None and login_idx < len(raw_row) else None,
            "password": raw_row[pass_idx] if pass_idx is not None and pass_idx < len(raw_row) else None,
        })
    return rows


def is_new_system_id(app_no: str) -> bool:
    return str(app_no).strip().upper().startswith("CV-")


# ---------------------------------------------------------------------------
# OCR: reading the downloaded certificate PDFs. These PDFs use a custom/
# subset font that direct text extraction decodes into garbage -- OCR on a
# rendered image of the page is what actually works reliably here.
# ---------------------------------------------------------------------------

DATE_RE = re.compile(r"\b(\d{1,2}[-/][A-Za-z0-9]{2,9}[-/]\d{2,4})\b")


def ocr_pdf_bytes(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text_parts = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text_parts.append(pytesseract.image_to_string(img))
    doc.close()
    return "\n".join(text_parts)


def parse_certificate_text(text: str):
    """Returns (decision, reason, date) from OCR'd certificate text.
    decision is 'APPROVED' or 'REJECTED'; reason is None unless rejected."""
    upper = text.upper()
    rejected = "NOT RECOMMEND" in upper or "NOT-RECOMMEND" in upper
    decision = "REJECTED" if rejected else "APPROVED"

    reason = None
    if rejected:
        # Old-system wording: "Reason for denial is : XXXX ."
        m = re.search(r"reason for denial is\s*[:\-]?\s*(.+?)(?:\.|$)", text, re.IGNORECASE | re.DOTALL)
        if m:
            reason = " ".join(m.group(1).split())
        else:
            # New-system wording: the "observations" box, right after the
            # "Verification Result" line.
            m = re.search(r"observations are as follows\s*:?\s*(.+?)(?:This certificate|$)", text, re.IGNORECASE | re.DOTALL)
            if m:
                reason = " ".join(m.group(1).split())
        if reason:
            # OCR sometimes drops or garbles the reason line's very first
            # letter (a font-rendering quirk on this specific document,
            # not fixed by higher resolution or padding) -- strip stray
            # leading punctuation, then repair the single most common
            # resulting artifact: "NOT ..." losing its leading N.
            reason = re.sub(r"^[^A-Za-z0-9]+", "", reason).strip()
            if re.match(r"^OT\s", reason, re.IGNORECASE):
                reason = "N" + reason

    date = None
    m = re.search(r"on\s+(\d{1,2}-\d{1,2}-\d{2,4})\.", text)  # new-system "...on DD-MM-YYYY."
    if m:
        date = m.group(1)
    else:
        m = re.search(r"Date\s*[:\-]?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})", text, re.IGNORECASE)  # old-system footer
        if m:
            date = m.group(1)

    return decision, reason, date


ROW_COL = 6      # "999   "
ID_COL = 16      # element id
APP_COL_W = 18   # application/token no.
TIME_COL = 7     # "12.34s"


def format_row(row_num, element_id, app_no, time_text, status_text):
    row = str(row_num).ljust(ROW_COL)
    eid = str(element_id).ljust(ID_COL)[:ID_COL]
    app = str(app_no).ljust(APP_COL_W)[:APP_COL_W]
    time_ = time_text.rjust(TIME_COL)
    return f"{dim(row)}{eid}{app}{time_}  {status_text}"


def show_status(text):
    # This tool processes cases one at a time (login/CAPTCHA rules out
    # parallel workers), so a single line is enough -- no need for the
    # multi-slot machinery the parallel Argus tools use. The row's final
    # result (printed with a trailing newline) naturally overwrites this.
    print(f"\r\033[2K{text}", end="", flush=True)


def color_result(status):
    if status == "APPROVED":
        return ok(status)
    if status in ("REJECTED", "NOT_FOUND", "ERROR"):
        return fail(status)
    return warn(status)


def download_bytes(url: str) -> bytes:
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content


# ---------------------------------------------------------------------------
# Login -- mobile number and password are auto-filled by the script itself
# (typed into the terminal beforehand, same as before). The only thing that
# still needs a human is the CAPTCHA: it's saved as an image file, its path
# is printed, and the answer is typed back into the terminal -- the user
# never has to open or click anything in the browser.
# ---------------------------------------------------------------------------

CAPTCHA_ANSWER_SELECTOR = "input[placeholder='Enter captcha shown above']"


def _open_login_modal(page, out_folder=None, attempts=3):
    # Reaching the nav "Login" button after a fresh page load is
    # intermittently slow on this site (confirmed by repeated live testing:
    # most loads finish in 1-5s, but some stall well past that) -- so this
    # retries with a fresh navigation each time instead of failing on the
    # first stall, and only gives up (with real diagnostics saved to disk)
    # after several tries.
    last_error = None
    for i in range(attempts):
        try:
            page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60000)
            # force=True: a leftover toast/overlay from a prior failed login
            # attempt has been confirmed (via a live failure) to sit on top
            # of page elements and intercept normal clicks even after a
            # fresh navigation. The nav button itself is genuinely visible
            # and enabled at this point -- only a decorative overlay is in
            # the way -- so clicking through it is correct here, not a
            # workaround for a real gating condition.
            page.get_by_text("Login", exact=True).first.click(timeout=20000, force=True)
            page.locator("input[type='tel']").wait_for(state="visible", timeout=20000)
            return
        except Exception as e:
            last_error = e
            _save_debug_snapshot(out_folder, f"open_modal_fail_{i}", page, e)
    raise RuntimeError(f"Could not open the login modal after {attempts} attempts: {last_error}")


def _save_debug_snapshot(out_folder, label, page, error):
    # Always write at least the error text (a plain file write can't fail
    # the way a screenshot/content() call can when the page is mid-
    # navigation or otherwise unstable at the exact moment of failure --
    # which is exactly when a screenshot silently failed once before).
    if not out_folder:
        return
    debug_dir = os.path.join(out_folder, "debug")
    os.makedirs(debug_dir, exist_ok=True)
    with open(os.path.join(debug_dir, f"{label}.txt"), "w", encoding="utf-8") as f:
        f.write(f"{type(error).__name__}: {error}\n")
    try:
        page.screenshot(path=os.path.join(debug_dir, f"{label}.png"), full_page=True)
    except Exception as e2:
        with open(os.path.join(debug_dir, f"{label}.txt"), "a", encoding="utf-8") as f:
            f.write(f"\n(screenshot also failed: {e2})\n")
    try:
        with open(os.path.join(debug_dir, f"{label}.html"), "w", encoding="utf-8") as f:
            f.write(page.content())
    except Exception as e3:
        with open(os.path.join(debug_dir, f"{label}.txt"), "a", encoding="utf-8") as f:
            f.write(f"(content() also failed: {e3})\n")


def _capture_captcha(page, out_folder, login_id):
    # The captcha visual isn't a plain <img> (it renders async, via canvas or
    # a background image), so instead of hunting for its exact element, this
    # anchors on the captcha ANSWER input -- confirmed reliable -- and crops
    # the region just above it, where the captcha visual sits.
    answer_box = page.locator(CAPTCHA_ANSWER_SELECTOR)
    answer_box.wait_for(state="visible", timeout=20000)
    page.wait_for_timeout(2000)  # let the async captcha image finish rendering
    box = answer_box.bounding_box()
    os.makedirs(out_folder, exist_ok=True)
    path = os.path.join(out_folder, f"captcha_{sanitize_filename(login_id)}.png")
    clip = {
        "x": max(box["x"] - 10, 0),
        "y": max(box["y"] - 90, 0),
        "width": box["width"] + 80,
        "height": 80,
    }
    page.screenshot(path=path, clip=clip)
    return path


def _click_submit_login(page):
    # The nav button that OPENS the modal and the modal's own SUBMIT button
    # both show the text "Login", but they have different CSS classes
    # ("navbar-login-btn" vs "login-btn") -- confirmed via a live dry run --
    # so this targets the submit button specifically instead of guessing by
    # position or DOM order. force=True for the same overlay-interception
    # reason as the other clicks in this flow.
    page.locator("button.login-btn").click(force=True)


CAPTCHA_REFRESH_SELECTOR = "button.captcha-refresh-btn"


def _refresh_captcha(page):
    # Confirmed selector (title="Refresh Captcha"). It's briefly disabled
    # right after the captcha image loads, so wait for it to become
    # clickable rather than firing immediately. A generous timeout here
    # since a post-submit auto-refresh may take longer than a manual click's
    # cooldown (measured at well under 1s in isolation).
    btn = page.locator(CAPTCHA_REFRESH_SELECTOR)
    btn.wait_for(state="visible", timeout=15000)
    page.wait_for_function(
        "sel => { const b = document.querySelector(sel); return b && !b.disabled; }",
        arg=CAPTCHA_REFRESH_SELECTOR,
        timeout=15000,
    )
    # force=True: confirmed via a live failure that a leftover toast/overlay
    # can sit on top of this button and intercept a normal click even
    # though the button itself is visible and enabled.
    btn.click(force=True)


def _read_login_error(page):
    # Best-effort: grab whatever text the modal is showing so a genuine
    # wrong-CAPTCHA can be told apart from e.g. a wrong password -- no
    # confirmed selector for this site's error/toast element yet, so this
    # scans broadly rather than assuming one.
    for sel in [".error", ".text-danger", ".toast", "[role='alert']", ".ant-message", ".Toastify__toast"]:
        loc = page.locator(sel)
        if loc.count() > 0:
            try:
                text = " | ".join(t.strip() for t in loc.all_inner_texts() if t.strip())
                if text:
                    return text
            except Exception:
                pass
    return ""


def login(page, mobile, password, out_folder):
    captcha_dir = os.path.join(out_folder, "captcha")
    attempt = 0
    _open_login_modal(page, out_folder)
    page.locator("input[type='tel']").fill(mobile)
    page.locator("input[type='password']").fill(password)

    while True:
        attempt += 1
        captcha_path = _capture_captcha(page, captcha_dir, mobile)
        print(warn(f"\nCAPTCHA saved to: {captcha_path}"))
        answer = input(dim("Open that image and type the CAPTCHA text here: ")).strip()
        page.locator(CAPTCHA_ANSWER_SELECTOR).fill(answer)

        _click_submit_login(page)
        page.wait_for_timeout(2500)

        # The modal's own fields (not just the "tel" one -- some SPAs leave
        # a hidden copy mounted in the DOM) are checked for actual on-screen
        # visibility, not mere DOM presence, since a login modal that's
        # merely hidden (rather than removed) would otherwise always look
        # like a failure even on success.
        if not page.locator(CAPTCHA_ANSWER_SELECTOR).is_visible():
            return  # modal closed -- logged in

        error_text = _read_login_error(page)
        if error_text:
            print(warn(f"Site says: {error_text}"))

        if attempt >= 5:
            _save_debug_snapshot(out_folder, "login_gave_up", page, error_text or "no error text found")
            raise RuntimeError(
                f"Login failed after 5 attempts (bad CAPTCHA or credentials)."
                f"{' Site said: ' + error_text if error_text else ''}"
            )

        print(warn("Login didn't go through -- refreshing CAPTCHA and retrying...\n"))
        try:
            _refresh_captcha(page)
        except Exception:
            # Fallback if the refresh button isn't behaving as expected --
            # reopen the modal from scratch (with its own retry+diagnostics)
            # instead of getting stuck on a single attempt.
            _open_login_modal(page, out_folder)
            page.locator("input[type='tel']").fill(mobile)
            page.locator("input[type='password']").fill(password)


# ---------------------------------------------------------------------------
# New system: search-by-application-number.
# ---------------------------------------------------------------------------

def check_new_system(page, app_no, out_folder):
    t0 = time.time()
    show_status(f"{app_no}: searching...")
    page.goto(NEW_STATUS_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(1000)
    page.locator("input[placeholder*='CV-2026']").fill(app_no)
    page.get_by_role("button", name="Search", exact=True).click()
    page.wait_for_timeout(1500)

    if page.get_by_text("No Matching Applications").count() > 0:
        elapsed = time.time() - t0
        return {"status": "NOT_FOUND", "remarks": "No matching application (wrong login/account?)",
                "downloaded": "", "report_date": ""}, elapsed

    row = page.locator("table tbody tr").first
    status_text = " ".join(row.inner_text().split("\n")).strip()

    if "Certificate Ready" in status_text or row.get_by_text("Certificate", exact=False).count() > 0:
        show_status(f"{app_no}: downloading certificate...")
        try:
            with page.expect_download(timeout=30000) as dl_info:
                row.get_by_role("button", name="Certificate").click()
            download = dl_info.value
            path = download.path()
            with open(path, "rb") as f:
                pdf_bytes = f.read()
        except Exception:
            # Fallback: some browsers open the PDF in a new tab instead of downloading it.
            with page.context.expect_page(timeout=15000) as pop_info:
                row.get_by_role("button", name="Certificate").click()
            popup = pop_info.value
            popup.wait_for_load_state("domcontentloaded", timeout=15000)
            pdf_bytes = download_bytes(popup.url)
            popup.close()

        show_status(f"{app_no}: reading certificate (OCR)...")
        text = ocr_pdf_bytes(pdf_bytes)
        decision, reason, date = parse_certificate_text(text)
        elapsed = time.time() - t0
        return {"status": decision, "remarks": reason or "", "downloaded": "Y", "report_date": date or ""}, elapsed

    elapsed = time.time() - t0
    return {"status": "PENDING", "remarks": status_text, "downloaded": "", "report_date": ""}, elapsed


# ---------------------------------------------------------------------------
# Old system: the full listing has to be scanned for the matching token ID
# (there's no search box). Then either download+OCR the certificate, or
# fall back to the "Check Status" modal.
# ---------------------------------------------------------------------------

def _dump_debug(out_folder, token_id, label, **fields):
    debug_dir = os.path.join(out_folder, "debug")
    os.makedirs(debug_dir, exist_ok=True)
    path = os.path.join(debug_dir, "old_system_log.txt")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"=== {sanitize_filename(token_id)} [{label}] ===\n")
        for key, value in fields.items():
            f.write(f"{key}: {value!r}\n")
        f.write("\n")


def check_old_system(page, token_id, out_folder):
    t0 = time.time()

    # Navigating straight to the JSP URL breaks it ("YOUR APPLICATION HAS
    # GENERATED AN ERROR... PLEASE LOGIN AGAIN") -- this legacy page only
    # works when actually reached by clicking the button, which opens it
    # in a new tab. Direct page.goto() skips whatever session state that
    # click sets up server-side.
    show_status(f"{token_id}: opening old-system list...")
    page.goto(NEW_STATUS_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(1000)
    with page.context.expect_page(timeout=15000) as pop_info:
        page.get_by_text("Check Old Application Status", exact=False).click()
    old_page = pop_info.value
    old_page.wait_for_load_state("domcontentloaded", timeout=15000)
    old_page.wait_for_timeout(1500)

    try:
        show_status(f"{token_id}: scanning list for this token...")
        row = old_page.locator(f"tr:has(td:text-is('{token_id}'))").first
        if row.count() == 0:
            # Diagnostic: dump what the page actually showed, so a
            # NOT_FOUND can be told apart from "wrong page entirely" (not
            # logged in, empty table, different structure) without guessing.
            debug_dir = os.path.join(out_folder, "debug")
            os.makedirs(debug_dir, exist_ok=True)
            safe_id = sanitize_filename(token_id)
            old_page.screenshot(path=os.path.join(debug_dir, f"{safe_id}_not_found.png"), full_page=True)
            with open(os.path.join(debug_dir, f"{safe_id}_not_found.html"), "w", encoding="utf-8") as f:
                f.write(old_page.content())
            all_ids = old_page.locator("table tbody tr td:nth-child(2)").all_inner_texts()
            with open(os.path.join(debug_dir, f"{safe_id}_visible_ids.txt"), "w", encoding="utf-8") as f:
                f.write(f"URL: {old_page.url}\n")
                f.write(f"Looking for: {token_id}\n")
                f.write(f"Token IDs visible on page ({len(all_ids)}):\n")
                f.write("\n".join(all_ids))

            elapsed = time.time() - t0
            return {"status": "NOT_FOUND", "remarks": "Token ID not found in this account's list",
                    "downloaded": "", "report_date": ""}, elapsed

        # The Status column shows a "Check Status" link only while a case
        # is still pending -- once it's decided, that link is gone (this
        # is what caused the earlier timeout on an already-decided row),
        # and the raw status cell text is used as a fallback signal instead.
        status_cell = row.locator("td").nth(7)
        status_cell_text = status_cell.inner_text().strip()
        has_check_status = "Check Status" in status_cell_text

        ps_status = cpsc_status = due_date = reg_date = ""
        if has_check_status:
            show_status(f"{token_id}: opening Check Status...")
            row.get_by_text("Check Status").click()
            old_page.wait_for_timeout(800)
            modal = old_page.locator("text=PPSAANJH || Status").locator("..").locator("..")
            modal_text = modal.inner_text()
            old_page.keyboard.press("Escape")

            ps_status = _extract_modal_field(modal_text, "Police Station Status")
            cpsc_status = _extract_modal_field(modal_text, "CPSC Status")
            due_date = _extract_modal_field(modal_text, "Due Date")
            reg_date = _extract_modal_field(modal_text, "Registration Date")
            _dump_debug(out_folder, token_id, "modal", modal_text=modal_text, ps_status=ps_status,
                        cpsc_status=cpsc_status, due_date=due_date, reg_date=reg_date)
        modal_date = due_date or reg_date

        download_cell_text = row.inner_text()
        if "Download Character Certificate" in download_cell_text or (
            "Download" in download_cell_text and "Not Available" not in download_cell_text.split("Download")[-1][:40]
        ):
            link = row.get_by_role("link", name="Download Character Certificate")
            if link.count() > 0:
                show_status(f"{token_id}: downloading certificate...")
                href = link.get_attribute("href")
                pdf_bytes = download_bytes(href)
                show_status(f"{token_id}: reading certificate (OCR)...")
                text = ocr_pdf_bytes(pdf_bytes)
                decision, reason, ocr_date = parse_certificate_text(text)
                elapsed = time.time() - t0
                return {"status": decision, "remarks": reason or "", "downloaded": "Y",
                        "report_date": modal_date or ocr_date or ""}, elapsed

        if not has_check_status:
            # Decided (no "Check Status" link) but no certificate either --
            # fall back to whatever the status cell itself says, logged for
            # review since this is an unexpected combination.
            _dump_debug(out_folder, token_id, "no_check_status_no_cert", status_cell_text=status_cell_text)
            elapsed = time.time() - t0
            return {"status": "UNKNOWN", "remarks": status_cell_text, "downloaded": "",
                    "report_date": ""}, elapsed

        # Still pending, per "Check Status" -- decide from its fields.
        # Require an EXPLICIT positive/negative statement (not just a loose
        # substring like "RECOMMEND", which also appears in unrelated
        # template text and produced a false APPROVED earlier).
        combined = f"{ps_status} {cpsc_status}".upper()
        elapsed = time.time() - t0

        if "NOT RECOMMEND" in combined or "REJECT" in combined:
            shot_name = f"{sanitize_filename(token_id)}_rejected.png"
            shot_path = os.path.join(out_folder, "screenshots", shot_name)
            os.makedirs(os.path.dirname(shot_path), exist_ok=True)
            row.get_by_text("Check Status").click()
            old_page.wait_for_timeout(800)
            old_page.locator("text=PPSAANJH || Status").locator("..").screenshot(path=shot_path)
            old_page.keyboard.press("Escape")
            return {"status": "REJECTED", "remarks": ps_status or cpsc_status,
                    "downloaded": "", "report_date": modal_date}, elapsed

        if re.search(r"\bRECOMMENDED\b", combined) or "DISPATCH" in combined:
            return {"status": "APPROVED", "remarks": "", "downloaded": "",
                    "report_date": modal_date}, elapsed

        return {"status": "PENDING", "remarks": "", "downloaded": "", "report_date": ""}, elapsed
    finally:
        old_page.close()


def _extract_modal_field(modal_text, label):
    m = re.search(re.escape(label) + r"\s*(.+)", modal_text)
    if not m:
        return ""
    value = m.group(1).split("\n")[0].strip()
    return value


def format_duration(seconds):
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes, secs = divmod(seconds, 60)
    return f"{int(minutes)}m {secs:.0f}s"


def main():
    args = parse_args()
    if not os.path.exists(args.excel_file):
        sys.exit(f"Error: file not found: {args.excel_file}")

    rows = parse_excel(args.excel_file)
    if not rows:
        sys.exit("Error: no data rows found in the Excel file.")

    default_mobile = args.mobile or input(f"{BOLD}Default mobile number for login:{RESET} ").strip()
    default_password = args.password or getpass.getpass(f"{BOLD}Default password for login (hidden):{RESET} ")

    total = len(rows)
    print(heading(f"Found {total} rows.") + " Checking one at a time"
          " (each login saves a CAPTCHA image and asks you to type the answer here in the terminal).\n")

    header = (
        dim("Row".ljust(ROW_COL)) + "Element ID".ljust(ID_COL)
        + "App/Token No.".ljust(APP_COL_W) + "Time".rjust(TIME_COL) + "  Status"
    )
    print(header)
    print(dim("-" * (ROW_COL + ID_COL + APP_COL_W + TIME_COL + 9)))

    os.makedirs(args.out, exist_ok=True)

    # Group rows by (login, password) so we only log in once per account,
    # not once per case -- solving a CAPTCHA for every single row isn't
    # practical.
    groups = {}
    for i, row in enumerate(rows):
        login_id = row["login"] if not is_blank(row["login"]) else default_mobile
        password = row["password"] if not is_blank(row["password"]) else default_password
        groups.setdefault((login_id, password), []).append(i)

    results = {}
    succeeded = 0
    failures = []
    start_time = time.time()

    with sync_playwright() as p:
        for (login_id, password), indices in groups.items():
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = context.new_page()
            try:
                login(page, str(login_id).strip(), str(password).strip(), args.out)
                for i in indices:
                    row = rows[i]
                    app_no = row["app_no"]
                    try:
                        if is_new_system_id(app_no):
                            result, elapsed = check_new_system(page, str(app_no).strip(), args.out)
                        else:
                            result, elapsed = check_old_system(page, str(app_no).strip(), args.out)
                    except Exception as e:
                        result = {"status": "ERROR", "remarks": str(e), "downloaded": "", "report_date": ""}
                        elapsed = 0.0

                    results[i] = result
                    line = format_row(i + 1, row["element_id"], app_no, f"{elapsed:.2f}s", color_result(result["status"]))
                    print(line)
                    if result["status"] in ("APPROVED",):
                        succeeded += 1
                    elif result["status"] not in ("PENDING",):
                        failures.append((row["element_id"], app_no, result["status"], result["remarks"]))
            finally:
                context.close()
                browser.close()

    elapsed_total = time.time() - start_time
    summary = (
        f"\nDone: checked {len(results)}/{total}. "
        f"Total time: {format_duration(elapsed_total)}."
    )
    print(summary)

    report_wb = openpyxl.Workbook()
    ws = report_wb.active
    ws.title = "Punjab PCC Status"
    ws.append(["ELEMENT ID", "PCC APPLICATION NUMBER", "STATUS", "REMARKS", "DOWNLOADED", "REPORT DATE"])
    for i, row in enumerate(rows):
        r = results.get(i, {"status": "SKIPPED", "remarks": "", "downloaded": "", "report_date": ""})
        ws.append([row["element_id"], row["app_no"], r["status"], r["remarks"], r["downloaded"], r["report_date"]])
    report_wb.save(os.path.join(args.out, "punjab_pcc_status_report.xlsx"))

    print(ok(f"Saved to folder: {args.out}"))


if __name__ == "__main__":
    main()
