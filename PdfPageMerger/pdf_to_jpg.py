import argparse
import glob
import io
import os
import queue
import re
import sys
import threading
import time

import pymupdf as fitz
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

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')
RENDER_DPI = 150
MERGE_GAP_PX = 10


def parse_args():
    parser = argparse.ArgumentParser(description="Stack each PDF's pages into a single JPG, named by Element ID.")
    parser.add_argument("pdf_folder", help="Folder containing the PDF files to process")
    parser.add_argument("--workers", type=int, default=4, help="Parallel renders (default: 4)")
    parser.add_argument("--out", default="pdf_to_jpg_output", help="Output folder (default: pdf_to_jpg_output)")
    return parser.parse_args()


def sanitize_filename(name: str) -> str:
    name = str(name).strip()
    return INVALID_FILENAME_CHARS.sub("_", name)


def render_page_to_image(page, dpi=RENDER_DPI):
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=matrix)
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def stack_images_vertically(images, gap=MERGE_GAP_PX):
    width = max(img.width for img in images)

    def fit_width(img):
        if img.width == width:
            return img
        ratio = width / img.width
        return img.resize((width, max(1, round(img.height * ratio))))

    images = [fit_width(img) for img in images]
    total_height = sum(img.height for img in images) + gap * (len(images) - 1)
    merged = Image.new("RGB", (width, total_height), "white")
    y = 0
    for img in images:
        merged.paste(img, (0, y))
        y += img.height + gap
    out = io.BytesIO()
    merged.save(out, format="JPEG", quality=92)
    return out.getvalue()


ROW_COL = 6    # "999   "
ID_COL = 16    # element id
SIZE_COL = 10  # "1234.5 KB"
TYPE_COL = 14  # "image/jpeg"
TIME_COL = 7   # "12.34s"

# A cursor can only ever be moved up into rows that are STILL on screen --
# once a row scrolls past the top of the terminal (inevitable for any
# folder with more files than fit on one screen), no escape code can
# address it again. So instead of one line per file, there's a small
# FIXED block of "worker" lines that always sits at the bottom of the
# screen and is the only thing ever rewritten in place. Each file's final
# result is printed ONCE, as a permanent new line inserted just above
# that block -- never touched again.
_print_lock = threading.Lock()
_num_slots = 0
_slot_text = []

# Guards the on-disk filename each file picks, so two files that happen
# to share an ELEMENT ID never race on the same "is this name used" check.
_save_lock = threading.Lock()
_used_names = {}


def format_row(row_num, element_id, size_text, type_text, time_text, status_text):
    row = str(row_num).ljust(ROW_COL)
    eid = str(element_id).ljust(ID_COL)[:ID_COL]
    size = size_text.rjust(SIZE_COL)
    type_ = type_text.ljust(TYPE_COL)
    time_ = time_text.rjust(TIME_COL)
    return f"{dim(row)}{eid}{size}  {type_}{time_}  {status_text}"


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


def process_pdf(slot_idx, row_num, filename, pdf_path, out_folder):
    t0 = time.time()
    stem = os.path.splitext(filename)[0]
    element_id = sanitize_filename(stem)
    if not element_id:
        reason = "Could not determine ELEMENT ID from filename"
        line = format_row(row_num, filename, "", "", "0.00s", fail(reason))
        return "failed", reason, line

    update_slot(slot_idx, format_row(row_num, element_id, "", "", "", warn("opening...")))
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        reason = f"Could not open PDF: {e}"
        elapsed = time.time() - t0
        line = format_row(row_num, element_id, "", "", f"{elapsed:.2f}s", fail(reason))
        return "failed", reason, line

    try:
        page_count = doc.page_count
        if page_count == 0:
            reason = "PDF has no pages"
            elapsed = time.time() - t0
            line = format_row(row_num, element_id, "", "", f"{elapsed:.2f}s", fail(reason))
            return "failed", reason, line
        try:
            update_slot(slot_idx, format_row(row_num, element_id, "", "image/jpeg", "", warn(f"rendering {page_count} page(s)...")))
            images = [render_page_to_image(doc[i]) for i in range(page_count)]
            update_slot(slot_idx, format_row(row_num, element_id, "", "image/jpeg", "", warn("stacking...")))
            merged = stack_images_vertically(images)
        except Exception as e:
            reason = f"Could not render/stack pages: {e}"
            elapsed = time.time() - t0
            line = format_row(row_num, element_id, "", "", f"{elapsed:.2f}s", fail(reason))
            return "failed", reason, line
    finally:
        doc.close()

    size_kb = len(merged) / 1024
    update_slot(slot_idx, format_row(row_num, element_id, f"{size_kb:.1f} KB", "image/jpeg", "", warn("saving...")))

    with _save_lock:
        out_filename = f"{element_id}.jpg"
        if out_filename in _used_names:
            _used_names[out_filename] += 1
            out_filename = f"{element_id}_{_used_names[out_filename]}.jpg"
        else:
            _used_names[out_filename] = 0
        save_path = os.path.join(out_folder, out_filename)

    with open(save_path, "wb") as f:
        f.write(merged)

    elapsed = time.time() - t0
    line = format_row(row_num, element_id, f"{size_kb:.1f} KB", "image/jpeg", f"{elapsed:.2f}s", ok("OK"))
    return "success", "", line


def format_duration(seconds):
    if seconds < 60:
        return f"{seconds:.1f}s"
    minutes, secs = divmod(seconds, 60)
    return f"{int(minutes)}m {secs:.0f}s"


def main():
    args = parse_args()
    if not os.path.isdir(args.pdf_folder):
        sys.exit(f"Error: folder not found: {args.pdf_folder}")

    pdf_paths = sorted(glob.glob(os.path.join(args.pdf_folder, "*.pdf")))
    if not pdf_paths:
        sys.exit(f"Error: no PDF files found in {args.pdf_folder}")

    total = len(pdf_paths)
    print(heading(f"Found {total} PDF files.") + f" Processing with {BOLD}{args.workers}{RESET} parallel workers...\n")

    header = (
        dim("Row".ljust(ROW_COL)) + "Element ID".ljust(ID_COL)
        + "Size".rjust(SIZE_COL) + "  " + "Type".ljust(TYPE_COL)
        + "Time".rjust(TIME_COL) + "  Status"
    )
    print(header)
    print(dim("-" * (ROW_COL + ID_COL + SIZE_COL + 2 + TYPE_COL + TIME_COL + 9)))

    os.makedirs(args.out, exist_ok=True)

    tasks = queue.Queue()
    for i, p in enumerate(pdf_paths):
        tasks.put((i, p))
    init_slots(args.workers)

    succeeded = 0
    failures = []
    results_lock = threading.Lock()

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

    def worker_loop(slot_idx):
        nonlocal succeeded
        while True:
            try:
                i, pdf_path = tasks.get_nowait()
            except queue.Empty:
                return
            filename = os.path.basename(pdf_path)
            status, reason, line = process_pdf(slot_idx, i + 1, filename, pdf_path, args.out)
            with results_lock:
                if status == "success":
                    succeeded += 1
                else:
                    failures.append((filename, reason))
            submit_result(i, line)

    start_time = time.time()
    threads = [threading.Thread(target=worker_loop, args=(i,)) for i in range(args.workers)]
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
        report_lines = ["PDF FILE,Reason"]
        for pdf_filename, reason in failures:
            report_lines.append(f'"{pdf_filename}","{reason}"')
        with open(os.path.join(args.out, "failures.csv"), "w", encoding="utf-8") as f:
            f.write("\n".join(report_lines))

    print(ok(f"Saved to folder: {args.out}"))


if __name__ == "__main__":
    main()
