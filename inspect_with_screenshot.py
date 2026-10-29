"""
Diagnostic: click "Reports", then IMMEDIATELY (same instant) both screenshot
the window and dump its full UIA control tree, so there's no ambiguity about
whether the flyout was actually on screen at the moment the dump was taken.

Does not click anything inside the flyout.
"""

from datetime import datetime
from pathlib import Path

from app import LOGIN_WINDOW_TITLE_RE, audit, click_when_ready, is_logged_in
from pywinauto import Application

DUMP_DIR = Path(r"C:\MedentAutomation\reports")


def main():
    app = Application(backend="uia").connect(title_re=LOGIN_WINDOW_TITLE_RE)
    if not is_logged_in(app):
        raise RuntimeError("Not logged in / not at main menu -- run login_only.py first.")

    win = app.top_window()

    click_when_ready(
        win, title="Reports", control_type="Text", dump_label="reports_tile_not_found"
    )
    audit("clicked_reports_tile")

    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    screenshot_path = DUMP_DIR / f"screenshot_{stamp}.png"
    win.wrapper_object().capture_as_image().save(str(screenshot_path))

    dump_path = DUMP_DIR / f"uia_dump_{stamp}.txt"
    win.print_control_identifiers(filename=str(dump_path))

    print(f"Window title at dump time: {win.window_text()!r}")
    print(f"Screenshot: {screenshot_path}")
    print(f"UIA dump:   {dump_path}")


if __name__ == "__main__":
    main()
