import argparse
import os
import queue
import re
import sys
import threading
import time
from urllib.parse import urlparse

import openpyxl
import requests

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

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')

CONTENT_TYPE_EXT = {
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
    "image/webp": ".webp",
    "image/bmp": ".bmp",
    "application/pdf": ".pdf",
}


def parse_args():
    parser = argparse.ArgumentParser(description="Download images from an Excel sheet of links.")
    parser.add_argument("excel_file", help="Path to the Excel file (must have TRACK ID and a photo/image URL column)")
    parser.add_argument("--photo-column", default="PHOTO URL", help="Column heading for the image URL (default: PHOTO URL)")
    parser.add_argument("--workers", type=int, default=5, help="Parallel downloads (default: 5)")
    parser.add_argument("--suffix", default="PHOTO", help="Filename suffix, e.g. TRACKID_PHOTO.jpg (default: PHOTO)")
    parser.add_argument("--out", default="photo_downloader_output", help="Output folder (default: photo_downloader_output)")
    return parser.parse_args()


def sanitize_filename(name: str) -> str:
    name = str(name).strip()
    return INVALID_FILENAME_CHARS.sub("_", name)


def guess_extension(url: str, content_type: str | None) -> str:
    if content_type:
        content_type = content_type.split(";")[0].strip().lower()
        if content_type in CONTENT_TYPE_EXT:
            return CONTENT_TYPE_EXT[content_type]
    path = urlparse(url).path
    tail = path.rsplit("/", 1)[-1]
    if "." in tail:
        ext = tail[tail.rfind("."):]
        if len(ext) <= 5:
            return ext.lower()
    return ".jpg"


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


def parse_excel(path, photo_col):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        header = next(rows_iter)
    except StopIteration:
        sys.exit("Error: the Excel file is empty.")

    photo_idx = find_column_index(header, photo_col)
    track_idx = find_column_index(header, TRACK_ID_COL)
    missing = []
    if photo_idx is None:
        missing.append(photo_col)
    if track_idx is None:
        missing.append(TRACK_ID_COL)
    if missing:
        sys.exit(f"Error: missing required column(s): {', '.join(missing)}")

    rows = []
    for raw_row in rows_iter:
        if raw_row is None or all(v is None for v in raw_row):
            continue
        track_id_raw = raw_row[track_idx] if track_idx < len(raw_row) else None
        url = raw_row[photo_idx] if photo_idx < len(raw_row) else None
        rows.append({"track_id": track_id_raw, "url": url})
    return rows


ROW_COL = 6    # "999   "
ID_COL = 16    # track id
SIZE_COL = 10  # "1234.5 KB"
TYPE_COL = 14  # "image/jpeg"
TIME_COL = 7   # "12.34s"

# A cursor can only ever be moved up into rows that are STILL on screen --
# once a row scrolls past the top of the terminal (inevitable for any file
# bigger than one screenful, e.g. 1000+ rows), no escape code can address
# it again. So instead of one line per Excel row, there's a small FIXED
# block of "worker" lines (one per parallel worker, e.g. 5) that always
# sits at the very bottom of the screen and is the only thing ever
# rewritten in place. Each row's final result is printed ONCE, as a
# permanent new line inserted just above that block -- never touched
# again. This scales to any row count, because the live block is always
# tiny and always the most recently printed thing on screen.
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
    # Inserted as a permanent line just above the live slots block, using
    # whatever the slots currently show -- callers decide WHEN to call
    # this (see the ordering buffer in main()), so a slot's own text may
    # have already moved on to a newer row by the time this runs. That's
    # fine: this only ever draws the current live state, it doesn't own
    # any particular slot.
    with _print_lock:
        print(f"\033[{_num_slots}A\033[J", end="")
        print(line)
        for slot_line in _slot_text:
            print(slot_line)


def clear_slots():
    with _print_lock:
        print(f"\033[{_num_slots}A\033[J", end="", flush=True)


def download_row(slot_idx, row_num, track_id_raw, url, session, out_folder, suffix):
    t0 = time.time()

    if track_id_raw is None or str(track_id_raw).strip() == "":
        line = format_row(row_num, "(missing)", "", "", "0.00s", fail("TRACK ID missing"))
        return "failed", "TRACK ID missing", None, line

    if url is None or str(url).strip() == "":
        line = format_row(row_num, track_id_raw, "", "", "0.00s", fail("URL missing"))
        return "failed", "URL missing", None, line

    url_str = str(url).strip()
    track_id = sanitize_filename(track_id_raw)

    update_slot(slot_idx, format_row(row_num, track_id, "", "", "", warn("connecting...")))
    try:
        response = session.get(url_str, timeout=15, stream=True)
        response.raise_for_status()
    except requests.RequestException as e:
        reason = describe_request_error(e)
        elapsed = time.time() - t0
        line = format_row(row_num, track_id, "", "", f"{elapsed:.2f}s", fail(reason))
        return "failed", f"Download failed: {reason}", None, line

    content_type = response.headers.get("Content-Type", "unknown")
    ext = guess_extension(url_str, content_type)

    update_slot(slot_idx, format_row(row_num, track_id, "", content_type, "", warn("downloading...")))
    try:
        content = response.content
    except requests.RequestException as e:
        reason = describe_request_error(e)
        elapsed = time.time() - t0
        line = format_row(row_num, track_id, "", "", f"{elapsed:.2f}s", fail(reason))
        return "failed", f"Download failed: {reason}", None, line

    size_kb = len(content) / 1024
    update_slot(slot_idx, format_row(row_num, track_id, f"{size_kb:.1f} KB", content_type, "", warn("saving...")))

    with _save_lock:
        filename = f"{track_id}_{suffix}{ext}"
        if filename in _used_names:
            _used_names[filename] += 1
            filename = f"{track_id}_{suffix}_{_used_names[filename]}{ext}"
        else:
            _used_names[filename] = 0
        save_path = os.path.join(out_folder, filename)

    with open(save_path, "wb") as f:
        f.write(content)

    elapsed = time.time() - t0
    line = format_row(row_num, track_id, f"{size_kb:.1f} KB", content_type, f"{elapsed:.2f}s", ok("OK"))
    return "success", "", track_id, line


def format_duration(seconds):
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes, secs = divmod(seconds, 60)
    return f"{int(minutes)}m {secs:.0f}s"


def main():
    args = parse_args()
    if not os.path.exists(args.excel_file):
        sys.exit(f"Error: file not found: {args.excel_file}")

    rows = parse_excel(args.excel_file, args.photo_column)
    if not rows:
        sys.exit("Error: no data rows found in the Excel file.")

    total = len(rows)
    print(heading(f"Found {total} rows.") + f" Downloading with {BOLD}{args.workers}{RESET} parallel workers...\n")

    header = (
        dim("Row".ljust(ROW_COL)) + "Track ID".ljust(ID_COL)
        + "Size".rjust(SIZE_COL) + "  " + "Type".ljust(TYPE_COL)
        + "Time".rjust(TIME_COL) + "  Status"
    )
    print(header)
    print(dim("-" * (ROW_COL + ID_COL + SIZE_COL + 2 + TYPE_COL + TIME_COL + 9)))

    os.makedirs(args.out, exist_ok=True)

    # A fixed-size queue of pending work, and a fixed pool of worker
    # threads (one per --workers) that each keep the SAME slot index for
    # their whole life -- that stable index is what makes update_slot()
    # safe to call from any thread, any number of rows in.
    tasks = queue.Queue()
    for i, row in enumerate(rows):
        tasks.put((i, row))
    init_slots(args.workers)

    succeeded = 0
    failures = []
    results_lock = threading.Lock()

    # Rows finish in whatever order workers happen to complete them, but
    # the printed log should read in the same order as the Excel file.
    # Each result is held here until every row before it has already been
    # printed, then it (and anything else now unblocked) is flushed --
    # this never touches the slots mechanism above, so it stays correct
    # for any file size.
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
            status, reason, track_id, line = download_row(
                slot_idx, i + 1, row["track_id"], row["url"], session, args.out, args.suffix
            )
            with results_lock:
                if status == "success":
                    succeeded += 1
                else:
                    failures.append((row["track_id"], row["url"], reason))
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
        report_lines = [f"TRACK ID,{args.photo_column},Reason"]
        for track_id_raw, url, reason in failures:
            report_lines.append(f'"{track_id_raw}","{url}","{reason}"')
        with open(os.path.join(args.out, "failures.csv"), "w", encoding="utf-8") as f:
            f.write("\n".join(report_lines))

    print(ok(f"Saved to folder: {args.out}"))


if __name__ == "__main__":
    main()
