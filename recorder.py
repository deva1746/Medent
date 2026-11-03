"""
Click / keystroke / new-window recorder for manually walking through Medent
screens, so ambiguous steps (which control is really under a click, what
order screens actually appear in) can be confirmed from real data instead of
guessed or narrated from memory.

Runs in the foreground of this PowerShell window while you drive Medent by
hand elsewhere. It does not control anything, and -- unlike the first two
versions of this script -- it does NOT install any system-wide input hook.
It only POLLS a few times a second:

  - Left mouse button state (GetAsyncKeyState) + cursor position
    (GetCursorPos). On each press, it records which top-level window you
    clicked in (Medent, Chrome, or something else), the text nearby read via
    Windows' own OCR engine (same engine app.py already uses), and whatever
    MSAA can tell about the exact thing under the cursor (name/role/value) --
    the same low-level accessibility API app.py had to fall back to for
    Medent's custom-drawn controls, since plain UI Automation often sees
    nothing there.
  - A fixed set of key states (GetAsyncKeyState again). Named keys (Enter,
    Tab, Escape, Backspace, Delete, arrows, F-keys) are logged by name.
    Ordinary typing is logged only as an approximate "typing detected" tick
    count, never which characters -- this is a coarser signal than true
    keystroke logging on purpose, since polling can't reliably distinguish
    exact characters across shift/layout state anyway.
  - New top-level windows/popups appearing, with their child controls'
    titles and control types read via UI Automation.

WHY POLLING, NOT HOOKING: the first two versions used WH_KEYBOARD_LL /
WH_MOUSE_LL, which intercept every keystroke/click for the WHOLE PC,
synchronously, before Windows delivers them anywhere else. A stall anywhere
in that path -- a blocked print(), a slow COM/accessibility call -- froze
mouse and keyboard input system-wide, and that happened twice in testing.
GetAsyncKeyState/GetCursorPos are passive reads with no place in any OS
input-delivery chain -- they cannot block anything, so this version cannot
reproduce that failure mode no matter what bug this script has. The tradeoff
is it can miss a click/keypress faster than the poll interval (default 50ms)
-- irrelevant for a deliberate, human-paced manual walkthrough.

Everything is written to a local JSONL file under RECORDER_DIR (see below) --
never synced to OneDrive. Nothing is sent anywhere.

CAUTION -- the app screen itself may show real patient/claims data (PHI).
The OCR crop around each click is small (kept deliberately tight), but
review events.jsonl before sharing it, and avoid clicking on fields that
show real patient data during the walkthrough; stick to navigation chrome
(buttons, menus, checkboxes) where possible.

Stop with Ctrl+C in this window (now safe to use -- nothing here can block
input processing) or just close the window.
"""

import json
import queue
import threading
import time
from datetime import datetime
from pathlib import Path

import win32api
import win32con
import win32gui
from PIL import ImageGrab
from pywinauto import Desktop

# Reuses app.py's already-proven OCR pipeline and MSAA plumbing rather than
# re-implementing them -- importing app.py is safe (it only runs main() under
# `if __name__ == "__main__"`).
import ctypes
from ctypes import byref, wintypes
from app import Accessibility, VARIANT, VT_DISPATCH, _run_ocr

oleacc = ctypes.windll.oleacc

RECORDER_DIR = Path(r"C:\MedentAutomation\recorder") / datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_PATH = RECORDER_DIR / "events.jsonl"

POLL_INTERVAL = 0.05  # seconds; fine enough for a human-paced walkthrough

SPECIAL_KEYS = {
    0x0D: "Enter", 0x09: "Tab", 0x1B: "Escape", 0x08: "Backspace",
    0x2E: "Delete", 0x25: "Left", 0x26: "Up", 0x27: "Right", 0x28: "Down",
    **{0x70 + i: f"F{i + 1}" for i in range(12)},
}

# Broad-but-not-exhaustive set used only to flag "typing happened" -- never
# used to reconstruct what was typed.
PRINTABLE_VKS = list(range(0x30, 0x3A)) + list(range(0x41, 0x5B)) + [0x20]

_event_queue = queue.Queue()
_key_run_ticks = 0
_known_windows = set()
_log_lock = threading.Lock()


def log_event(kind, **fields):
    RECORDER_DIR.mkdir(parents=True, exist_ok=True)
    record = {"timestamp": datetime.now().isoformat(), "kind": kind, **fields}
    with _log_lock:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, default=str) + "\n")
    shown = {k: v for k, v in fields.items() if k != "children"}
    print(f"[{record['timestamp']}] {kind}: {shown}")


def _describe_window_at_point(x, y):
    hwnd = win32gui.WindowFromPoint((x, y))
    top_hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT) if hwnd else 0
    title = win32gui.GetWindowText(top_hwnd) if top_hwnd else ""
    cls = win32gui.GetClassName(top_hwnd) if top_hwnd else ""
    if cls == "MedentClient":
        app_name = "Medent"
    elif "Chrome" in cls:
        app_name = "Chrome"
    else:
        app_name = cls or "Unknown"
    return {"app": app_name, "window_title": title, "window_class": cls}


def _describe_point_msaa(x, y):
    try:
        pt = wintypes.POINT(x, y)
        ppacc = ctypes.POINTER(Accessibility.IAccessible)()
        child_var = VARIANT()
        res = oleacc.AccessibleObjectFromPoint(pt, byref(ppacc), byref(child_var))
        if res != 0 or not ppacc:
            return {"error": f"AccessibleObjectFromPoint failed, hresult={res}"}
        acc = ppacc.value
        child_id = 0 if child_var.vt == VT_DISPATCH else (child_var.value or 0)
        info = {}
        try:
            info["name"] = acc.accName(child_id) or ""
        except Exception as exc:
            info["name_error"] = repr(exc)
        try:
            role = acc.accRole(child_id)
            buf = ctypes.create_unicode_buffer(128)
            oleacc.GetRoleTextW(role, buf, 128)
            info["role"] = buf.value
        except Exception as exc:
            info["role_error"] = repr(exc)
        try:
            info["value"] = acc.accValue(child_id) or ""
        except Exception:
            pass
        return info
    except Exception as exc:
        return {"error": f"msaa describe failed: {exc!r}"}


def _ocr_near_point(x, y, half_width=120, half_height=20):
    try:
        box = (x - half_width, y - half_height, x + half_width, y + half_height)
        img = ImageGrab.grab(bbox=box, all_screens=True)
        words = _run_ocr(img)
        return " ".join(w["text"] for w in words).strip()
    except Exception as exc:
        return f"<ocr failed: {exc!r}>"


def _on_click(x, y):
    win_info = _describe_window_at_point(x, y)
    element = _describe_point_msaa(x, y)
    ocr_text = _ocr_near_point(x, y)
    log_event("click", x=x, y=y, element=element, ocr_text=ocr_text, **win_info)


def _flush_key_run():
    global _key_run_ticks
    if _key_run_ticks:
        log_event("typed", approx_ticks=_key_run_ticks)
        _key_run_ticks = 0


def _worker_loop():
    global _key_run_ticks
    while True:
        item = _event_queue.get()
        try:
            if item[0] == "key":
                _flush_key_run()
                log_event("key", name=SPECIAL_KEYS[item[1]])
            elif item[0] == "typed_tick":
                _key_run_ticks += 1
            elif item[0] == "click":
                _, x, y = item
                _flush_key_run()
                _on_click(x, y)
        except Exception as exc:
            print(f"worker error on {item!r}: {exc!r}")


def _poll_mouse():
    prev_down = False
    while True:
        down = bool(win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000)
        if down and not prev_down:
            x, y = win32gui.GetCursorPos()
            _event_queue.put_nowait(("click", x, y))
        prev_down = down
        time.sleep(POLL_INTERVAL)


def _poll_keys():
    prev_state = {vk: False for vk in list(SPECIAL_KEYS) + PRINTABLE_VKS}
    while True:
        for vk in SPECIAL_KEYS:
            down = bool(win32api.GetAsyncKeyState(vk) & 0x8000)
            if down and not prev_state[vk]:
                _event_queue.put_nowait(("key", vk))
            prev_state[vk] = down

        any_typed = False
        for vk in PRINTABLE_VKS:
            down = bool(win32api.GetAsyncKeyState(vk) & 0x8000)
            if down and not prev_state[vk]:
                any_typed = True
            prev_state[vk] = down
        if any_typed:
            _event_queue.put_nowait(("typed_tick",))

        time.sleep(POLL_INTERVAL)


def _describe_new_window_children(hwnd, limit=40):
    children = []
    try:
        w = Desktop(backend="uia").window(handle=hwnd)
        for d in w.descendants():
            try:
                children.append({
                    "title": d.window_text(),
                    "control_type": d.element_info.control_type,
                    "auto_id": getattr(d.element_info, "automation_id", None),
                })
            except Exception:
                pass
            if len(children) >= limit:
                break
    except Exception as exc:
        children = [{"error": repr(exc)}]
    return children


def _poll_new_windows():
    global _known_windows
    while True:
        current = []

        def _collect(hwnd, out):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                out.append(hwnd)

        win32gui.EnumWindows(_collect, current)
        current_set = set(current)
        for hwnd in current_set - _known_windows:
            title = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd)
            children = _describe_new_window_children(hwnd)
            log_event("new_window", title=title, class_name=cls, children=children)
        _known_windows = current_set
        time.sleep(0.4)


def _disable_quick_edit_mode():
    """Console 'Quick Edit Mode' makes Ctrl+C copy the current selection
    instead of interrupting the process whenever any text happens to be
    selected -- easy to trigger by accident just by clicking/dragging on the
    scrolling log output. Turning it off means Ctrl+C always means
    'stop', never 'copy'. Best-effort only: has no effect in terminals that
    don't use the classic Windows console (e.g. some VS Code integrated
    terminal configurations) -- closing the window always works regardless.
    """
    try:
        STD_INPUT_HANDLE = -10
        ENABLE_QUICK_EDIT_MODE = 0x0040
        ENABLE_EXTENDED_FLAGS = 0x0080
        h = ctypes.windll.kernel32.GetStdHandle(STD_INPUT_HANDLE)
        mode = wintypes.DWORD()
        ctypes.windll.kernel32.GetConsoleMode(h, byref(mode))
        new_mode = (mode.value & ~ENABLE_QUICK_EDIT_MODE) | ENABLE_EXTENDED_FLAGS
        ctypes.windll.kernel32.SetConsoleMode(h, new_mode)
    except Exception:
        pass


def main():
    RECORDER_DIR.mkdir(parents=True, exist_ok=True)
    _disable_quick_edit_mode()
    print(f"Recording to {LOG_PATH}")
    print("Drive Medent normally in its own window. Press Ctrl+C here (or close the window) to stop.")
    print(
        "This version polls instead of hooking -- it cannot freeze system "
        "input, unlike the earlier versions."
    )

    threading.Thread(target=_worker_loop, daemon=True).start()
    threading.Thread(target=_poll_mouse, daemon=True).start()
    threading.Thread(target=_poll_keys, daemon=True).start()
    threading.Thread(target=_poll_new_windows, daemon=True).start()

    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    finally:
        _flush_key_run()
        print("Stopped.")


if __name__ == "__main__":
    main()
