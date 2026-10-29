"""
Diagnostic: log in, reach the main menu, click "Reports", and stop --
dumping every top-level window instead of guessing at the flyout's
selector again. Does not click anything inside the flyout.

Run this when navigate_to_point_ar_report's "Open Item Reports" lookup is
failing. The Reports flyout turns out to render as a separate top-level
window rather than new content inside the main menu's own window, so a
plain print_control_identifiers() on the main window (as inspect_main_menu.py
does) never captures it.
"""

from datetime import datetime
from pathlib import Path

from pywinauto import Desktop
from pywinauto.timings import wait_until_passes

from app import (
    Application,
    LOGIN_WINDOW_TITLE_RE,
    audit,
    click_when_ready,
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
    baseline_handles = {w.handle for w in Desktop(backend="uia").windows()}

    click_when_ready(
        win, title="Reports", control_type="Text", dump_label="reports_tile_not_found"
    )
    audit("clicked_reports_tile")

    # len(windows()) >= 1 would pass trivially even before the click (the
    # main window alone satisfies it) and prove nothing. Wait for the actual
    # top-level window SET to change from the pre-click baseline instead --
    # that's the real signal the flyout has appeared.
    wait_until_passes(
        5, 0.5,
        lambda: {w.handle for w in Desktop(backend="uia").windows()} != baseline_handles,
    )

    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    windows = Desktop(backend="uia").windows()
    summary_path = DUMP_DIR / f"reports_flyout_windows_{stamp}.txt"
    with open(summary_path, "w", encoding="utf-8") as f:
        for i, w in enumerate(windows):
            line = (
                f"[{i}] title={w.window_text()!r} "
                f"class={w.friendly_class_name()!r} handle={w.handle} "
                f"is_new={w.handle not in baseline_handles}"
            )
            print(line)
            f.write(line + "\n")

    for i, w in enumerate(windows):
        dump_path = DUMP_DIR / f"reports_flyout_window_{i}_{stamp}.txt"
        try:
            # Desktop().windows() returns bare UIAWrapper objects, which
            # don't have print_control_identifiers -- only the
            # WindowSpecification returned by .window(...) does. Re-wrap by
            # handle to get one.
            Desktop(backend="uia").window(handle=w.handle).print_control_identifiers(
                filename=str(dump_path)
            )
        except Exception as exc:
            print(f"[{i}] could not dump: {exc}")

    print(f"Dumped {len(windows)} window(s) to {DUMP_DIR}")
    print("Stopping here -- nothing inside the flyout was clicked.")


if __name__ == "__main__":
    main()
