import argparse
import io
import os
import queue
import re
import sys
import threading
import time

import openpyxl
import requests
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from colors import BOLD, RESET, dim, fail, heading, ok, warn

# When stdout isn't attached to a real console, Python falls back to the
# OS default encoding (cp1252 on Windows) instead of UTF-8 -- printing
# anything outside that narrow charset then crashes with
# UnicodeEncodeError. errors="replace" makes that merely show a
# substitute character instead of taking the whole tool down.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

TRACK_ID_COL = "TRACK ID"
FRONT_COL = "FRONT URL"
BACK_COL = "BACK URL"

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
MERGE_GAP_PX = 12


def parse_args():
    parser = argparse.ArgumentParser(description="Merge Aadhaar front/back image links into one file per Track ID.")
    parser.add_argument("excel_file", help="Path to the Excel file (must have TRACK ID, FRONT URL, BACK URL columns)")
    parser.add_argument("--workers", type=int, default=5, help="Parallel downloads (default: 5)")
    parser.add_argument("--out", default="aadhaar_merger_output", help="Output folder (default: aadhaar_merger_output)")
    return parser.parse_args()


def sanitize_filename(name: str) -> str:
    name = str(name).strip()
    return INVALID_FILENAME_CHARS.sub("_", name)


def describe_request_error(e: requests.RequestException) -> str:
    if isinstance(e, requests.exceptions.Timeout):
        return "request timed out"
    if isinstance(e, requests.exceptions.HTTPError) and e.response is not None:
        return f"{e.response.status_code} {e.response.reason}"
    if isinstance(e, requests.exceptions.ConnectionError):
        return "could not connect to the link"
    return str(e)


def find_column_index(header_row, target: str):
    for idx, cell in enumerate(header_row):
        if cell is not None and str(cell).strip().lower() == target.lower():
            return idx
    return None


def parse_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        header = next(rows_iter)
    except StopIteration:
        sys.exit("Error: the Excel file is empty.")

    track_idx = find_column_index(header, TRACK_ID_COL)
    front_idx = find_column_index(header, FRONT_COL)
    back_idx = find_column_index(header, BACK_COL)
    missing = []
    if track_idx is None:
        missing.append(TRACK_ID_COL)
    if front_idx is None:
        missing.append(FRONT_COL)
    if back_idx is None:
        missing.append(BACK_COL)
    if missing:
        sys.exit(f"Error: missing required column(s): {', '.join(missing)}")

    rows = []
    for raw_row in rows_iter:
        if raw_row is None or all(v is None for v in raw_row):
            continue
        rows.append({
            "track_id": raw_row[track_idx] if track_idx < len(raw_row) else None,
            "front_url": raw_row[front_idx] if front_idx < len(raw_row) else None,
            "back_url": raw_row[back_idx] if back_idx < len(raw_row) else None,
        })
    return rows


def is_blank(value) -> bool:
    if value is None:
        return True
    return str(value).strip() in ("", "-")


def merge_images(front_bytes, back_bytes):
    front_img = Image.open(io.BytesIO(front_bytes)).convert("RGB")
    back_img = Image.open(io.BytesIO(back_bytes)).convert("RGB")
    width = max(front_img.width, back_img.width)

    def fit_width(img):
        if img.width == width:
            return img
        ratio = width / img.width
        return img.resize((width, max(1, round(img.height * ratio))))

    front_img = fit_width(front_img)
    back_img = fit_width(back_img)
    merged = Image.new("RGB", (width, front_img.height + MERGE_GAP_PX + back_img.height), "white")
    merged.paste(front_img, (0, 0))
    merged.paste(back_img, (0, front_img.height + MERGE_GAP_PX))
    out = io.BytesIO()
    merged.save(out, format="JPEG", quality=92)
    return out.getvalue()


ROW_COL = 6    # "999   "
ID_COL = 16    # track id
SIZE_COL = 10  # "1234.5 KB"
TYPE_COL = 14  # "image/jpeg"
TIME_COL = 7   # "12.34s"

# A cursor can only ever be moved up into rows that are STILL on screen --
# once a row scrolls past the top of the terminal (inevitable for any file
# bigger than one screenful), no escape code can address it again. So
# instead of one line per Excel row, there's a small FIXED block of
# "worker" lines (one per parallel worker) that always sits at the very
# bottom of the screen and is the only thing ever rewritten in place. Each
# row's final result is printed ONCE, as a permanent new line inserted
# just above that block -- never touched again. This scales to any row
# count, because the live block is always tiny.
_print_lock = threading.Lock()
_num_slots = 0
_slot_text = []

# Guards the on-disk filename each row picks, so two rows that happen to
# share a TRACK ID never race on the same "is this name already used" check.
_save_lock = threading.Lock()
_used_names = {}


def format_row(row_num, track_id, size_text, type_text, time_text, status_text):
    row = str(row_num).ljust(ROW_COL)
    tid = str(track_id).ljust(ID_COL)[:ID_COL]
    size = size_text.rjust(SIZE_COL)
    type_ = type_text.ljust(TYPE_COL)
    time_ = time_text.rjust(TIME_COL)
    return f"{dim(row)}{tid}{size}  {type_}{time_}  {status_text}"


IDLE_SLOT = format_row("", "", "", "", "", dim("..."))


def init_slots(n):
    global _num_slots, _slot_text
    _num_slots = n
    _slot_text = [IDLE_SLOT] * n
    for line in _slot_text:
        print(line)


def update_slot(slot_idx, text):
    with _print_lock:
        _slot_text[slot_idx] = text
        lines_up = _num_slots - slot_idx
        print(f"\033[{lines_up}A\r\033[2K{text}\033[{lines_up}B\r", end="", flush=True)


def log_completed(line):
    with _print_lock:
        print(f"\033[{_num_slots}A\033[J", end="")
        print(line)
        for slot_line in _slot_text:
            print(slot_line)


def clear_slots():
    with _print_lock:
        print(f"\033[{_num_slots}A\033[J", end="", flush=True)


def process_row(slot_idx, row_num, track_id_raw, front_url, back_url, session, out_folder):
    t0 = time.time()

    if is_blank(track_id_raw):
        return "failed", "TRACK ID missing", format_row(row_num, "(missing)", "", "", "0.00s", fail("TRACK ID missing"))
    if is_blank(front_url):
        return "failed", "FRONT URL missing", format_row(row_num, track_id_raw, "", "", "0.00s", fail("FRONT URL missing"))
    if is_blank(back_url):
        return "failed", "BACK URL missing", format_row(row_num, track_id_raw, "", "", "0.00s", fail("BACK URL missing"))

    track_id = sanitize_filename(track_id_raw)

    update_slot(slot_idx, format_row(row_num, track_id, "", "", "", warn("downloading front...")))
    try:
        r = session.get(str(front_url).strip(), timeout=15)
        r.raise_for_status()
        front_bytes = r.content
    except requests.RequestException as e:
        reason = f"Download failed (front): {describe_request_error(e)}"
        elapsed = time.time() - t0
        return "failed", reason, format_row(row_num, track_id, "", "", f"{elapsed:.2f}s", fail(reason))

    update_slot(slot_idx, format_row(row_num, track_id, "", "", "", warn("downloading back...")))
    try:
        r = session.get(str(back_url).strip(), timeout=15)
        r.raise_for_status()
        back_bytes = r.content
    except requests.RequestException as e:
        reason = f"Download failed (back): {describe_request_error(e)}"
        elapsed = time.time() - t0
        return "failed", reason, format_row(row_num, track_id, "", "", f"{elapsed:.2f}s", fail(reason))

    update_slot(slot_idx, format_row(row_num, track_id, "", "image/jpeg", "", warn("merging...")))
    try:
        merged = merge_images(front_bytes, back_bytes)
    except Exception as e:
        reason = f"Could not merge images: {e}"
        elapsed = time.time() - t0
        return "failed", reason, format_row(row_num, track_id, "", "", f"{elapsed:.2f}s", fail(reason))

    size_kb = len(merged) / 1024
    update_slot(slot_idx, format_row(row_num, track_id, f"{size_kb:.1f} KB", "image/jpeg", "", warn("saving...")))

    with _save_lock:
        filename = f"{track_id}_AADHAAR.jpg"
        if filename in _used_names:
            _used_names[filename] += 1
            filename = f"{track_id}_AADHAAR_{_used_names[filename]}.jpg"
        else:
            _used_names[filename] = 0
        save_path = os.path.join(out_folder, filename)

    with open(save_path, "wb") as f:
        f.write(merged)

    elapsed = time.time() - t0
    line = format_row(row_num, track_id, f"{size_kb:.1f} KB", "image/jpeg", f"{elapsed:.2f}s", ok("OK"))
    return "success", "", line


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

    total = len(rows)
    print(heading(f"Found {total} rows.") + f" Processing with {BOLD}{args.workers}{RESET} parallel workers...\n")

    header = (
        dim("Row".ljust(ROW_COL)) + "Track ID".ljust(ID_COL)
        + "Size".rjust(SIZE_COL) + "  " + "Type".ljust(TYPE_COL)
        + "Time".rjust(TIME_COL) + "  Status"
    )
    print(header)
    print(dim("-" * (ROW_COL + ID_COL + SIZE_COL + 2 + TYPE_COL + TIME_COL + 9)))

    os.makedirs(args.out, exist_ok=True)

    tasks = queue.Queue()
    for i, row in enumerate(rows):
        tasks.put((i, row))
    init_slots(args.workers)

    succeeded = 0
    failures = []
    results_lock = threading.Lock()

    # Rows finish in whatever order workers happen to complete them, but
    # the printed log should read in the same order as the Excel file.
    pending = {}
    next_to_log = 0
    order_lock = threading.Lock()

    def submit_result(i, line):
        nonlocal next_to_log
        with order_lock:
            pending[i] = line
            while next_to_log in pending:
                entry = pending.pop(next_to_log)
                try:
                    log_completed(entry)
                except Exception:
                    pass  # never let one bad line jam every row after it
                next_to_log += 1

    def worker_loop(slot_idx, session):
        nonlocal succeeded
        while True:
            try:
                i, row = tasks.get_nowait()
            except queue.Empty:
                return
            status, reason, line = process_row(
                slot_idx, i + 1, row["track_id"], row["front_url"], row["back_url"], session, args.out
            )
            with results_lock:
                if status == "success":
                    succeeded += 1
                else:
                    failures.append((row["track_id"], row["front_url"], row["back_url"], reason))
            submit_result(i, line)

    start_time = time.time()
    with requests.Session() as session:
        adapter = requests.adapters.HTTPAdapter(pool_connections=args.workers, pool_maxsize=args.workers)
        session.mount("https://", adapter)
        session.mount("http://", adapter)

        threads = [threading.Thread(target=worker_loop, args=(i, session)) for i in range(args.workers)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
    elapsed_total = time.time() - start_time
    clear_slots()

    summary = (
        f"\nDone: {ok(str(succeeded) + ' succeeded')}, "
        f"{fail(str(len(failures)) + ' failed') if failures else '0 failed'}. "
        f"Total time: {format_duration(elapsed_total)}."
    )
    print(summary)

    if failures:
        report_lines = ["TRACK ID,FRONT URL,BACK URL,Reason"]
        for track_id_raw, front_url, back_url, reason in failures:
            report_lines.append(f'"{track_id_raw}","{front_url}","{back_url}","{reason}"')
        with open(os.path.join(args.out, "failures.csv"), "w", encoding="utf-8") as f:
            f.write("\n".join(report_lines))

    print(ok(f"Saved to folder: {args.out}"))


if __name__ == "__main__":
    main()
