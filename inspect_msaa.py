"""
Diagnostic: click "Reports" (via the already-confirmed uia-backend click),
then dump the SAME window's tree via the legacy MSAA backend ("win32")
instead of UI Automation. The flyout has shown up in zero UIA-based dumps
and created zero new top-level windows (native Win32 EnumWindows check), so
it's likely custom-painted content with no UIA element at all. Some
owner-drawn Delphi/VCL controls still implement the older IAccessible
(MSAA) interface manually even when they don't bridge to UIA -- this checks
for that before falling back to coordinate/image-based clicking.

Does not click anything inside the flyout.
"""

from datetime import datetime
from pathlib import Path

from pywinauto import Application as UiaApplication
from pywinauto.application import Application as Win32Application

from app import LOGIN_WINDOW_TITLE_RE, audit, click_when_ready, is_logged_in

DUMP_DIR = Path(r"C:\MedentAutomation\reports")


def main():
    app = UiaApplication(backend="uia").connect(title_re=LOGIN_WINDOW_TITLE_RE)
    if not is_logged_in(app):
        raise RuntimeError("Not logged in / not at main menu -- run login_only.py first.")

    win = app.top_window()
    handle = win.handle

    click_when_ready(
        win, title="Reports", control_type="Text", dump_label="reports_tile_not_found"
    )
    audit("clicked_reports_tile")

    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dump_path = DUMP_DIR / f"msaa_dump_{stamp}.txt"

    msaa_app = Win32Application(backend="win32").connect(handle=handle)
    msaa_win = msaa_app.window(handle=handle)
    msaa_win.print_control_identifiers(filename=str(dump_path))

    print(f"Dumped MSAA (win32 backend) tree to {dump_path}")


if __name__ == "__main__":
    main()
