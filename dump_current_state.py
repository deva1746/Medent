"""
Diagnostic: dump whatever Medent is showing right now, with zero clicks.
Use this when a menu/flyout is already open on screen and you just need its
real control tree, instead of trying to reproduce the click that opened it.
"""

from datetime import datetime
from pathlib import Path

from pywinauto import Desktop

from app import audit

DUMP_DIR = Path(r"C:\MedentAutomation\reports")
TITLE_NEEDLE = "COMPASSIONATE WOMEN'S HEALTH"


def main():
    # Application(backend="uia").connect(title_re=...) walks every open
    # top-level window's UIA tree to match the title, and with several other
    # apps open (VS Code, Brave, Teams, Notepad) that walk is flaky and keeps
    # raising ElementNotFoundError even though the window is right there.
    # Desktop().windows() has worked reliably both times so far -- match
    # against that instead of going through connect()/find_element.
    all_windows = Desktop(backend="uia").windows()
    candidates = [w for w in all_windows if TITLE_NEEDLE in w.window_text()]

    if not candidates:
        print(f"No open window with {TITLE_NEEDLE!r} in its title found.")
        print(f"Currently enumerable top-level windows ({len(all_windows)}):")
        for w in all_windows:
            print(f"  title={w.window_text()!r} class={w.friendly_class_name()!r} handle={w.handle}")
        return
    if len(candidates) > 1:
        titles = [w.window_text() for w in candidates]
        raise RuntimeError(f"Multiple matching windows found, expected one: {titles}")

    win = Desktop(backend="uia").window(handle=candidates[0].handle)

    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dump_path = DUMP_DIR / f"current_state_{stamp}.txt"
    win.print_control_identifiers(filename=str(dump_path))
    audit("control_tree_dumped", path=str(dump_path), label="current_state")
    print(f"Window title: {win.window_text()!r}")
    print(f"Dumped to {dump_path}")


if __name__ == "__main__":
    main()
