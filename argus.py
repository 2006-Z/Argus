import msvcrt
import os
import re
import shutil
import subprocess
import sys
import textwrap

import openpyxl

from colors import BOLD, CYAN, RESET, YELLOW, dim, fail, heading, ok, warn

# When stdout isn't attached to a real console (or even in some real
# consoles), Python falls back to the OS default encoding (cp1252 on
# Windows) instead of UTF-8 -- printing anything outside that narrow
# charset (Devanagari text from a scraped site, certain symbols, etc.)
# then crashes with UnicodeEncodeError. errors="replace" makes that
# merely show a substitute character instead of taking the whole tool down.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ANSI_RE = re.compile(r"\033\[[0-9;]*[a-zA-Z]")

def validate_column_name(value, path):
    if not value:
        return "Enter the column heading exactly as it appears in the file."
    return validate_excel(path, [value])


INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def validate_suffix(value, _path):
    if INVALID_FILENAME_CHARS.search(value):
        return r'Suffix can\'t contain \ / : * ? " < > |'
    return None


# name, folder, script, what kind of path to ask for, field label for the
# form, required Excel columns (None if the tool doesn't take an Excel
# file), default output folder name, default parallel-workers count, extra
# tool-specific fields (each: key/label/prompt/default/validate, validate
# takes (value, excel_path) and returns an error string or None)
TOOLS = [
    ("Image Downloader", "ExcelPhotoDownloader", "photo_downloader.py", "Excel file", "Excel File",
     ["TRACK ID"], "photo_downloader_output", 5,
     [
         {"key": "photo-column", "label": "Photo URL Column",
          "prompt": "Column heading for the image URL (Enter = PHOTO URL):",
          "default": "PHOTO URL", "validate": validate_column_name,
          "note": "IMAGE URL (any column name -- specify below)"},
         {"key": "suffix", "label": "Suffix",
          "prompt": "Filename suffix (Enter = PHOTO):",
          "default": "PHOTO", "validate": validate_suffix},
     ]),
    ("Aadhaar Merger", "AadhaarMerger", "aadhaar_merger.py", "Excel file", "Excel File",
     ["TRACK ID", "FRONT URL", "BACK URL"], "aadhaar_merger_output", 5, []),
    ("Pdf To Jpg", "PdfPageMerger", "pdf_to_jpg.py", "folder of PDF files", "PDF Folder",
     None, "pdf_to_jpg_output", 4, []),
    ("Up Challan Downloader", "ChallanDownloader", "up_challan_downloader.py", "Excel file", "Excel File",
     ["ELEMENT ID", "CHALLAN"], "up_challan_downloader_output", 2, []),
    ("Up Pcc Status Checker", "CctnsStatusChecker", "up_pcc_status_checker.py", "Excel file", "Excel File",
     ["ELEMENT ID", "APPLICATION"], "up_pcc_status_checker_output", 2, []),
    ("Delhi Pcc Status Checker", "PccStatusChecker", "delhi_pcc_status_checker.py", "Excel file", "Excel File",
     ["ELEMENT ID", "PCC APPLICATION NUMBER", "NAME", "FATHER NAME"], "delhi_pcc_status_checker_output", 5, []),
]


BANNER = rf"""{CYAN}{BOLD} █████╗ ██████╗  ██████╗ ██╗   ██╗███████╗
██╔══██╗██╔══██╗██╔════╝ ██║   ██║██╔════╝
███████║██████╔╝██║  ███╗██║   ██║███████╗
██╔══██║██╔══██╗██║   ██║██║   ██║╚════██║
██║  ██║██║  ██║╚██████╔╝╚██████╔╝███████║
╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝  ╚═════╝ ╚══════╝
{RESET}"""


class GoBack(Exception):
    """Raised when the user types 0 at a field prompt -- cancel this tool
    and return to the main menu."""


class QuitArgus(Exception):
    """Raised when the user confirms they want to quit entirely."""


def row_count(text):
    # How many physical terminal rows `text` actually occupies -- NOT just
    # its "\n" count. A single line longer than the terminal's width wraps
    # onto extra rows with no "\n" in sight, and erase_lines()/update_at()
    # move the cursor by physical rows, not logical lines. Undercounting
    # here is exactly what left stray leftovers behind on a long error
    # message: the wipe moved up too few rows to actually reach it.
    width = shutil.get_terminal_size(fallback=(80, 24)).columns
    total = 0
    for segment in text.split("\n"):
        visible_len = len(ANSI_RE.sub("", segment))
        total += max(1, -(-visible_len // width))  # ceil division, min 1 row
    return total


def wrap_indented(text, indent="  "):
    # Wraps `text` to the terminal's own width and re-applies `indent` to
    # every line AFTER the first -- the terminal's OWN automatic wrapping
    # has no idea we indent our messages, so a line we wrap ourselves like
    # this is the only way the continuation keeps the same left margin
    # instead of starting flush against the edge.
    width = shutil.get_terminal_size(fallback=(80, 24)).columns
    wrapped = textwrap.wrap(text, width=max(20, width - len(indent))) or [""]
    return ("\n" + indent).join(wrapped)


def erase_lines(count):
    # Move the cursor up `count` lines, then clear from there to the end
    # of the screen. This only erases what THIS script itself just drew --
    # anything printed before that (like an earlier tool's own results)
    # stays exactly where it was in the scrollback, untouched.
    if count > 0:
        print(f"\033[{count}A\033[J", end="", flush=True)


def clear_screen():
    # Same effect as PowerShell's own `cls` / Clear-Host: \033[2J clears the
    # visible screen, and \033[3J additionally clears the scrollback itself
    # -- so nothing printed before this point is reachable by scrolling up
    # afterward. This is still the normal screen buffer (not the alternate
    # one), so anything printed AFTER this still scrolls normally within
    # itself if it runs longer than one screen.
    print("\033[H\033[2J\033[3J", end="", flush=True)


def hide_cursor():
    print("\033[?25l", end="", flush=True)


def show_cursor():
    print("\033[?25h", end="", flush=True)


def discard_typed_input():
    # While a tool subprocess runs, nothing is reading stdin -- any keys
    # the user presses (Enter, out of habit or impatience) just sit
    # queued in the console's input buffer. Without draining it, the
    # very next input() call (the "Press Enter" prompt) consumes that
    # queued keystroke instantly instead of waiting for a fresh one,
    # which looks like it "auto-pressed" itself. kbhit()/getch() drain
    # whatever's queued, without blocking, before that prompt is shown.
    while msvcrt.kbhit():
        msvcrt.getch()


def show_banner():
    # The big ASCII banner -- shown exactly once, at the very start of the
    # session, same as how Claude Code prints its own welcome banner once
    # and never repeats it for every prompt after that.
    for line in [""] + BANNER.split("\n") + [""]:
        print(line)


def show_tool_list():
    lines = [f"  {YELLOW}{BOLD}{i}{RESET}  {name}" for i, (name, *_) in enumerate(TOOLS, 1)] + [""]
    for line in lines:
        print(line)
    return len(lines)


def confirm_quit():
    # Self-contained: prints its own prompt AND cleans up after itself, so
    # every caller can use it without separately having to account for (or
    # forgetting to account for) the extra lines it puts on screen -- the
    # exact kind of gap that left stray lines behind elsewhere.
    prompt = f"\n{YELLOW}Quit Argus? (y/n):{RESET} "
    answer = input(prompt)
    erase_lines(row_count(prompt + answer))
    return answer.strip().lower() == "y"


def remove_from_path():
    ps_script = (
        "$dir = '" + BASE_DIR + "'; "
        "$old = [Environment]::GetEnvironmentVariable('PATH','User'); "
        "$parts = ($old -split ';') | Where-Object { $_ -ne $dir -and $_ -ne '' }; "
        "[Environment]::SetEnvironmentVariable('PATH', ($parts -join ';'), 'User')"
    )
    subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps_script])


def handle_delete():
    # A secret self-test switch, not a listed menu option: simulates a
    # corrupted install (removes Argus from PATH, deletes the setup
    # marker) so the self-healing path in argus.bat can be tested on
    # demand -- next launch detects the missing marker and repairs
    # everything (re-syncs from git, reinstalls, re-adds PATH) on its own.
    print(warn("\nDeleting Argus's local setup (PATH entry + setup marker)...\n"))
    remove_from_path()
    marker = os.path.join(BASE_DIR, ".argus_setup_ok")
    if os.path.exists(marker):
        os.remove(marker)
    print(fail("Setup deleted.") + " Argus is no longer on your PATH.")
    print(dim("Go back into the Argus folder and run argus.bat again -- it will notice and repair itself."))


def handle_update():
    # Another secret switch: pulls the latest code from git and restarts
    # into it immediately, so "typing update" really does update Argus
    # on the spot rather than just printing instructions.
    print(heading("\nChecking for updates...\n"))
    origin_check = subprocess.run(
        ["git", "remote", "get-url", "origin"], cwd=BASE_DIR, capture_output=True, text=True
    )
    if origin_check.returncode != 0:
        print(warn("No git remote configured yet -- nothing to update from."))
        input(dim("\nPress Enter to continue..."))
        return

    subprocess.run(["git", "fetch", "origin"], cwd=BASE_DIR)
    branch = subprocess.run(
        ["git", "rev-parse", "--abbrev-ref", "HEAD"], cwd=BASE_DIR, capture_output=True, text=True
    ).stdout.strip()
    reset = subprocess.run(["git", "reset", "--hard", f"origin/{branch}"], cwd=BASE_DIR, capture_output=True, text=True)
    if reset.returncode != 0:
        print(fail("Update failed:") + f" {reset.stderr.strip()}")
        input(dim("\nPress Enter to continue..."))
        return

    print(dim("Reinstalling dependencies (in case they changed)..."))
    subprocess.run([sys.executable, "-m", "pip", "install", "-r", os.path.join(BASE_DIR, "requirements.txt"), "-q"])
    print(ok("Updated. Restarting Argus...\n"))
    os.execv(sys.executable, [sys.executable, os.path.join(BASE_DIR, "argus.py")])


def validate_excel(path, required_columns):
    """Returns None if the file looks usable, or an error string explaining
    exactly what's wrong (unreadable file, or missing column(s))."""
    if not os.path.exists(path):
        return f"Not found: {path}"
    if not required_columns:
        return None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header = next(ws.iter_rows(values_only=True))
    except Exception as e:
        return f"Could not read this as an Excel file: {e}"

    header_names = [str(c).strip() for c in header if c is not None]
    header_lower = [h.lower() for h in header_names]
    missing = [c for c in required_columns if c.lower() not in header_lower]
    if missing:
        found = ", ".join(header_names) if header_names else "(no columns found)"
        line1 = wrap_indented(f"Missing column(s): {', '.join(missing)}.")
        line2 = wrap_indented(f"This file has: {found}")
        return f"{line1}\n  {line2}"
    return None


def run_session():
    def redraw():
        # A real clear, not cursor-position math -- that's what makes this
        # correct no matter how tall the terminal is or how much scrolled
        # by. Everything since the last "Choose a tool:" (the filled form,
        # the run, its output) is gone; only the banner and tool list ever
        # persist across the whole session.
        clear_screen()
        show_banner()
        show_tool_list()

    redraw()
    printed = [0]  # lines printed since the last redraw() -- always accurate

    while True:
        choose_marker = printed[0]

        def say(text=""):
            # flush=True matters here: a tool's own subprocess writes to
            # this same terminal right after "Running X..." -- without an
            # explicit flush, that line could still be sitting in Python's
            # own output buffer when the child starts writing, making it
            # appear in the wrong order.
            print(text, flush=True)
            printed[0] += row_count(text)
            return printed[0]  # position marker: line number this text ended on

        def ask(prompt):
            # Counted AFTER typing, using prompt+typed-text together --
            # what actually determines the row(s) consumed is the FULL
            # echoed line (the terminal doesn't know the typed part is
            # "different" from the prompt), so a long typed value can
            # push this onto extra wrapped rows just as easily as a long
            # prompt can.
            value = input(prompt)
            printed[0] += row_count(prompt + value)
            return value

        def ask_field(prompt):
            """Like ask(), but q means 'quit Argus' from anywhere."""
            value = ask(prompt).strip()
            if value.lower() == "q":
                if confirm_quit():
                    raise QuitArgus()
                return ask_field(prompt)
            return value

        def update_at(position, new_text):
            # Rewrites the line at `position` in place, then returns the
            # cursor to exactly where it was so normal printing continues
            # correctly below.
            lines_up = printed[0] - position + 1
            print(f"\033[{lines_up}A\r\033[2K{new_text}\033[{lines_up}B\r", end="", flush=True)

        def ask_until_valid(prompt, validate, strip_quotes=False, back_ok=True):
            """Asks the same question until `validate(value)` returns None
            (valid). A bad answer shows the reason in bold, waits for the
            user to acknowledge, then wipes that failed attempt clean and
            asks again from the exact same spot. A GOOD answer also wipes
            its own question+answer line once accepted -- nothing about
            the asking process stays on screen, only the result (written
            elsewhere, e.g. into the summary block above)."""
            while True:
                start = printed[0]
                value = ask_field(prompt)
                if back_ok and value == "0":
                    erase_lines(printed[0] - start)
                    printed[0] = start
                    raise GoBack()
                if strip_quotes:
                    value = value.strip('"')
                error = validate(value)
                if error is None:
                    erase_lines(printed[0] - start)
                    printed[0] = start
                    return value
                ack = f"  {fail(error)}\n\n  {dim('(Press Enter to try again)')}"
                printed[0] += row_count(ack)
                input(ack)
                erase_lines(printed[0] - start)
                printed[0] = start

        def validate_choice(v):
            if v.lower() in ("delete", "update"):
                return None
            try:
                i = int(v) - 1
            except ValueError:
                return "Enter a number from the list (or 0 to exit)."
            if v == "0" or 0 <= i < len(TOOLS):
                return None
            return "Enter a number from the list (or 0 to exit)."

        try:
            choice = ask_until_valid(
                f"{BOLD}Choose a tool:{RESET} ",
                validate_choice,
                back_ok=False,
            )
        except QuitArgus:
            return

        if choice == "0":
            return

        if choice.lower() == "delete":
            handle_delete()
            return

        if choice.lower() == "update":
            handle_update()  # only returns here if there was nothing to update / it failed
            continue

        index = int(choice) - 1
        name, folder, script, input_kind, field_label, columns, out_folder_name, default_workers, extra_fields = TOOLS[index]

        # Show everything this tool needs up front, as a small form, before
        # asking for any of it -- then fill each line in as it's answered.
        column_notes = [ef["note"] for ef in extra_fields if ef.get("note")]
        all_columns = (columns or []) + column_notes
        input_line = (
            f"  Excel Columns: {YELLOW}{', '.join(all_columns)}{RESET}" if all_columns
            else f"  Input: {YELLOW}Any Number Of PDF Files In The Folder{RESET}"
        )
        say(f"\n{heading(name)}")
        say(input_line)
        pos_path = say(f"  {field_label}: ")
        extra_positions = [say(f"  {ef['label']}: ") for ef in extra_fields]
        pos_out = say("  Save Folder: ")
        pos_workers = say("  No Of Workers: ")
        say("")

        try:
            path = ask_until_valid(
                f"{BOLD}Path to {input_kind}:{RESET} ",
                lambda v: validate_excel(v.strip('"'), columns),
                strip_quotes=True,
            )
            update_at(pos_path, f"  {field_label}: {ok(path)}")

            extra_values = {}
            for ef, pos in zip(extra_fields, extra_positions):
                def validate_extra(v, ef=ef):
                    v = v.strip() or ef["default"]
                    return ef["validate"](v, path) if ef["validate"] else None

                raw = ask_until_valid(f"{BOLD}{ef['prompt']}{RESET} ", validate_extra)
                value = raw.strip() or ef["default"]
                extra_values[ef["key"]] = value
                update_at(pos, f"  {ef['label']}: {ok(value)}")

            def validate_out_folder(v):
                v = v.strip('"') or os.getcwd()
                if not os.path.isdir(v):
                    return f"Folder not found: {v}"
                return None

            out_folder_raw = ask_until_valid(
                f"{BOLD}Save output in folder (Enter = this folder):{RESET} ",
                validate_out_folder,
            )
            out_folder = out_folder_raw.strip('"') or os.getcwd()
            update_at(pos_out, f"  Save Folder: {ok(out_folder)}")

            def validate_workers(v):
                if not v:
                    return None
                try:
                    if int(v) < 1:
                        raise ValueError
                except ValueError:
                    return "Enter a whole number of 1 or more (or leave blank)."
                return None

            workers_raw = ask_until_valid(
                f"{BOLD}No of workers (Enter = {default_workers}):{RESET} ",
                validate_workers,
            )
            workers = int(workers_raw) if workers_raw else default_workers
            update_at(pos_workers, f"  No Of Workers: {ok(str(workers))}")
        except GoBack:
            erase_lines(printed[0] - choose_marker)
            printed[0] = choose_marker
            continue
        except QuitArgus:
            return

        out_path = os.path.join(out_folder, out_folder_name)

        extra_args = []
        for key, value in extra_values.items():
            extra_args += [f"--{key}", value]

        script_path = os.path.join(BASE_DIR, folder, script)
        say(heading(f"Running {name}...\n"))
        proc = subprocess.Popen(
            [sys.executable, script_path, path, "--out", out_path, "--workers", str(workers)] + extra_args
        )
        hide_cursor()
        try:
            try:
                proc.wait()
            except KeyboardInterrupt:
                show_cursor()
                print(f"\n{YELLOW}Interrupted.{RESET}")
                if confirm_quit():
                    proc.terminate()
                    return
                print(dim("Continuing -- the tool is still running in the background."))
                hide_cursor()
                proc.wait()
        finally:
            show_cursor()
        discard_typed_input()

        ask(f"\n{dim('Done. Press Enter to return to the menu...')}")
        redraw()
        printed[0] = 0
        # Back to a clean "Choose a tool:" -- banner and tool list are the
        # only things that persist for the whole session; the form and run
        # output from this round are gone.


def main():
    run_session()


if __name__ == "__main__":
    main()
