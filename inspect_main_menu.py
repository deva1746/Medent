"""
Diagnostic: log in, reach the main menu, and stop -- dumping the real
control tree instead of guessing at the "Reports" selector again.

Run this instead of app.py when navigate_to_point_ar_report's child_window()
lookups are failing. It reuses the already-confirmed login/main-menu-wait
code from app.py and does nothing risky beyond that.
"""

from datetime import datetime
from pathlib import Path

from app import (
    Application,
    LOGIN_WINDOW_TITLE_RE,
    audit,
    get_credentials,
    is_logged_in,
    login,
    wait_for_main_menu,
)

DUMP_DIR = Path(r"C:\MedentAutomation\reports")


def main():
    user_id, password = get_credentials()

    app = Application(backend="uia").connect(title_re=LOGIN_WINDOW_TITLE_RE)
    if is_logged_in(app):
        audit("already_logged_in")
    else:
        login(app, user_id, password)
    wait_for_main_menu(app)

    win = app.top_window()
    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    dump_path = DUMP_DIR / f"main_menu_tree_{datetime.now():%Y%m%d_%H%M%S}.txt"
    win.print_control_identifiers(filename=str(dump_path))
    print(f"Dumped control tree to {dump_path}")


if __name__ == "__main__":
    main()
