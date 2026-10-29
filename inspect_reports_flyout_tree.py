"""
Diagnostic stage 2: inspect_reports_menu.py's dump showed 8 top-level
windows after clicking "Reports", none titled "Reports" -- so the flyout is
NOT a separate top-level window. It must render as content inside the main
MEDENT window itself, so this dumps THAT window's full control tree while
the flyout is open and checks it for the menu entries navigation needs.

Read-only: opens the Reports flyout, dumps the tree, reports which expected
labels are present/missing. Does not click anything inside the flyout,
generate/export a report, or touch filter/export/logout logic.

Run this after login_only.py (requires already being at the main menu).
"""

from datetime import datetime
from pathlib import Path

from app import (
    LOGIN_WINDOW_TITLE_RE,
    MEDENT_CLASS_NAME,
    audit,
    click_when_ready,
    connect_to_app,
    is_logged_in,
)

DUMP_DIR = Path(r"C:\MedentAutomation\reports")

# Labels we need to find inside the Reports flyout to build real selectors
# for navigate_to_point_ar_report(). Matched case-insensitively as
# substrings against every title captured in the dump.
EXPECTED_LABELS = [
    "Open Item Reports",
    "Daysheet",
    "Insurance Reports",
    "Open Item: Charges For One or more Ins. Co",
]


def locate_main_window(app):
    """Re-resolve the live MEDENT main window by title pattern + class,
    never by a cached handle/pid. The window's title text changes with the
    current screen (it drops the "Main Menu" prefix once navigated away --
    see the title observed after Stage 1: "COMPASSIONATE WOMEN'S HEALTH
    PLLC UJ,Thr, 10/29/26", no "Main Menu" in it), so a selector resolved
    once at launch and reused later can point at a stale title.
    """
    return app.window(title_re=LOGIN_WINDOW_TITLE_RE, class_name=MEDENT_CLASS_NAME)


def main():
    app = connect_to_app()
    if not is_logged_in(app):
        raise RuntimeError("Not logged in / not at main menu -- run login_only.py first.")

    win = locate_main_window(app)

    click_when_ready(
        win, title="Reports", control_type="Text", dump_label="reports_tile_not_found"
    )
    audit("clicked_reports_tile")

    # Re-resolve rather than reuse `win` -- the tree rebuilds often enough
    # (see click_when_ready's docstring) that a pre-click reference risks
    # dumping stale content.
    win = locate_main_window(app)

    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dump_path = DUMP_DIR / f"main_window_tree_with_reports_open_{stamp}.txt"
    win.print_control_identifiers(filename=str(dump_path))
    audit("control_tree_dumped", path=str(dump_path), label="main_window_with_reports_open")

    dump_text = dump_path.read_text(encoding="utf-8", errors="replace")

    print(f"Dumped main window control tree (Reports flyout open) to {dump_path}\n")

    missing = []
    for label in EXPECTED_LABELS:
        if label.lower() in dump_text.lower():
            print(f"FOUND:   {label!r}")
        else:
            missing.append(label)
            print(f"MISSING: {label!r}")

    if missing:
        print(
            "\nSome expected labels are missing from the dump. This likely means "
            "either the flyout had already auto-dismissed by the time the dump "
            "ran, or it's rendered via an accessibility interface UIA doesn't "
            "surface -- some owner-drawn controls only implement the legacy "
            "MSAA/IAccessible interface (see inspect_msaa.py for that "
            "win32-backend approach). Inspect every title/control_type in "
            f"{dump_path} manually before guessing a selector."
        )
    else:
        print("\nAll expected labels found -- safe to build real selectors from this dump.")

    print("\nStopping here -- nothing inside the flyout was clicked, no report run.")


if __name__ == "__main__":
    main()
