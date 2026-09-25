import os

# Modern Windows terminals understand ANSI color codes, but classic cmd.exe
# needs this one-time nudge to turn that processing on -- this trick works
# with zero extra dependencies (no colorama needed).
os.system("")

RESET = "\033[0m"
BOLD = "\033[1m"
DIM = "\033[2m"

GREEN = "\033[32m"
RED = "\033[31m"
YELLOW = "\033[33m"
CYAN = "\033[36m"
MAGENTA = "\033[35m"
GRAY = "\033[90m"
WHITE = "\033[97m"

BG_CYAN = "\033[46m"
BLACK = "\033[30m"


def ok(text):
    return f"{GREEN}{text}{RESET}"


def fail(text):
    return f"{BOLD}{RED}{text}{RESET}"


def warn(text):
    return f"{YELLOW}{text}{RESET}"


def info(text):
    return f"{CYAN}{text}{RESET}"


def heading(text):
    return f"{BOLD}{CYAN}{text}{RESET}"


def dim(text):
    return f"{GRAY}{text}{RESET}"
