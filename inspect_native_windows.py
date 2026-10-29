"""
Diagnostic: click "Reports" and immediately enumerate ALL native Win32
windows via EnumWindows (not pywinauto's UIA-based Desktop().windows()),
including invisible/tool/no-title windows and native popup-menu windows
(class "#32768"). The Reports flyout has not shown up in several UIA-based
dumps despite being visibly open on screen, so this checks whether it's a
native popup menu or something else UI Automation isn't surfacing, at the
Win32 level instead.

Does not click anything inside the flyout.
"""

from datetime import datetime
from pathlib import Path

import win32gui

from app import Application, LOGIN_WINDOW_TITLE_RE, audit, click_when_ready, is_logged_in

DUMP_DIR = Path(r"C:\MedentAutomation\reports")


def enum_all_windows():
    results = []

    def callback(hwnd, _):
        results.append(
            {
                "hwnd": hwnd,
                "class": win32gui.GetClassName(hwnd),
                "title": win32gui.GetWindowText(hwnd),
                "visible": bool(win32gui.IsWindowVisible(hwnd)),
            }
        )
        return True

    win32gui.EnumWindows(callback, None)
    return results


def main():
    app = Application(backend="uia").connect(title_re=LOGIN_WINDOW_TITLE_RE)
    if not is_logged_in(app):
        raise RuntimeError("Not logged in / not at main menu -- run login_only.py first.")

    win = app.top_window()

    before = enum_all_windows()
    before_hwnds = {w["hwnd"] for w in before}

    click_when_ready(
        win, title="Reports", control_type="Text", dump_label="reports_tile_not_found"
    )
    audit("clicked_reports_tile")

    after = enum_all_windows()

    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = DUMP_DIR / f"native_windows_{stamp}.txt"

    new_windows = [w for w in after if w["hwnd"] not in before_hwnds]

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"Total windows before click: {len(before)}\n")
        f.write(f"Total windows after click:  {len(after)}\n")
        f.write(f"New windows (by hwnd) after click: {len(new_windows)}\n\n")
        f.write("=== NEW WINDOWS ===\n")
        for w in new_windows:
            line = f"hwnd={w['hwnd']} class={w['class']!r} title={w['title']!r} visible={w['visible']}"
            print(line)
            f.write(line + "\n")
        f.write("\n=== ALL WINDOWS AFTER CLICK ===\n")
        for w in after:
            f.write(
                f"hwnd={w['hwnd']} class={w['class']!r} title={w['title']!r} "
                f"visible={w['visible']} new={w['hwnd'] not in before_hwnds}\n"
            )

    print(f"\n{len(new_windows)} new window(s) appeared after the click. Full detail: {out_path}")


if __name__ == "__main__":
    main()
