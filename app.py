"""
Medent "Point A/R Table Report" auto-pull script.

STATUS: the full flow -- login through Reports navigation, filter
configuration, export, Excel data capture, and logout -- is implemented in
navigate_to_point_ar_report() and confirmed against TWO independent live,
full start-to-end runs (2026-11-03), captured hands-off with recorder.py (a
separate click/window observer built specifically so the remaining steps
could be confirmed from real data instead of guessed from screenshots/
narration -- see its own docstring). Every step still halts via
NavigationCheckpoint with full diagnostics if what it finds doesn't match
what was confirmed live, same discipline as every earlier pass below.

Confirmed and implemented, end to end:
  - Login (2 Edit fields + "Sign In" pane)
  - Reports -> "Open Item Reports" -> "Open Item: Charges For One or More
    Ins. Co", "Print A/R Tasks on Report?" -> Yes -> two Enter keypresses
  - Insurance scope ("All" / "Type of Insurance" / "Select Ins. Co" / "All")
    and "Report by Patient Name" selection
  - Date-range calendar overlay (year arrows / month / day) for the "from"
    date, confirmed via the "Done" checkmark
  - "Data Export File" -> XML -> "Select XML Option" -> OK (confirmed to be
    a real UIA Button, not a touch-UI Pane)
  - Export confirm: the SAME "Done" checkmark used for the date range is
    reused here (there is no separate "Export" toolbar click) -> "Yes" on
    Medent's own print-warning prompt -> native "Print" dialog -> OK ->
    native "Save Print Output As" dialog (writes a PDF) -> Excel opens
    automatically with the real report data as an "XML Spreadsheet 2003"
    file
  - The actual report DATA is captured by reading Excel's ActiveWorkbook
    path via COM (win32com.client.GetActiveObject("Excel.Application")) and
    copying that file to OUTPUT_DIR, rather than UI-automating Excel's own
    File -> Save As -> Browse -> change format -> Save (the human
    walkthrough's manual path, and the single least-tested, highest-risk
    remaining surface -- one of the two live runs even hit an unexpected
    compatibility prompt there). The workbook is then closed via COM
    without resaving, and the file this script generated is left untouched.
  - Exit / logout ("Exit" -> "Yes", both confirmed as embedded Panes)

Known gap, NOT touched by the above (pre-existing): the --date-from/--date-to
CLI arguments are parsed and validated, but the date-range calendar clicks
inside navigate_to_point_ar_report() are still the original hardcoded
navigation -- CLI dates do not yet actually change what gets clicked. A "to"
date other than today has also never been observed live at all. Fixing
either requires capturing that specific interaction live first, same as
everything else in this file.

The Reports flyout is confirmed to be a native child window of class
"WP Menu Class" (not a top-level window, not UIA-visible content) whose
items are read via raw MSAA (IAccessible), not pywinauto.
navigate_to_point_ar_report() enumerates ALL same-class sibling hwnds (not
just the first), multi-signal polls them, and -- once a candidate is
found -- reads and identifies the "Open Item Reports" item via MSAA. As of
this revision (Pass 3), expanding that item's submenu is a dedicated,
verified sequence (see _expand_submenu_item): accDoDefaultAction, then
keyboard Right Arrow, then a coordinate click at the estimated position of
the ">" arrow drawn at the item's right edge (computed only from its own
captured rect) -- never Enter (which selects/commits rather than expands),
and never a hardcoded coordinate. Each attempt is checked against FIVE
independent signals (_find_submenu_evidence): a new, distinct menu window;
the parent window's own accessible content changing; its MSAA child-count
changing; its geometry changing; and a screenshot for human review --
because MEDENT may reuse the existing popup HWND instead of creating a new
one, so a new-window check alone isn't sufficient. No single signal is
treated as proof; the result is one of verified / inconclusive / not
expanded, and the function always halts via NavigationCheckpoint with that
evidence, since clicking into whatever the submenu actually contains has
never been observed live. A human needs to cross-check that evidence
against what's really on screen before a follow-up change implements
whatever comes next.

Pass 4 (this revision): the prior Pass 3 code never actually reached
_expand_submenu_item live -- audit logs from three real runs show
_capture_msaa_menu_items's "open item" search always failed first, because
(a) its accLocation call used the raw Win32 byref-pointer signature, which
comtypes' auto-generated IAccessible wrapper rejects (it returns the four
[out] params as a tuple instead), so `rect` silently failed to resolve for
EVERY item on EVERY run; and (b), confirmed from the same logs, the Reports
popup's hwnd-level MSAA root exposes only 1-2 unnamed pass-through
container children (role WINDOW), not the real rows -- "Open Item Reports"
and its siblings live one level deeper. _capture_msaa_menu_items now
recurses into such a container via _capture_msaa_children (capped at
MSAA_MAX_RECURSION_DEPTH), and each item carries an `accessible_path` so a
nested item can be correctly re-resolved (_resolve_msaa_accessible) for
accDoDefaultAction instead of being (silently wrongly) invoked against the
hwnd-level root. Neither fix has been dry-run verified yet -- the next live
run is what actually confirms whether 'Open Item Reports' is now found with
a real rect, and whether its submenu then expands.

Pass 5 (this revision): Pass 4's fix was verified live and DID work exactly
as intended -- MSAA now reads 14 items with zero errors, real rects, no
resolve/accLocation failures. But the 14 items are NOT real menu rows: their
names are "System", "", "Application", "Vertical", "Horizontal", "" with
roles TITLEBAR/MENUBAR/SCROLLBAR/GRIP -- the generic default accessible
object Windows provides for ANY ordinary top-level window when nothing more
specific registers its own accessibility support, duplicated twice because
the recursion walked into two separate pass-through containers that both
bottom out at this same generic object. There is no deeper "real" MSAA
content to recurse into after all; this control genuinely exposes none of
its owner-drawn row text via MSAA or UIA (UIA likewise only ever showed a
generic ScrollBar, which turned out to be this same default furniture, not
evidence of real list content). Locating "Open Item Reports" now goes
through OCR instead (_locate_submenu_arrow_via_ocr /
_expand_submenu_item_via_ocr), reading the popup's own rendered pixels via
PrintWindow (NOT a desktop screenshot -- a live test showed a desktop
screenshot can capture an unrelated overlapping window instead of the real
popup) and Windows' built-in OCR engine (winsdk). The existing MSAA-based
_expand_submenu_item/_click_msaa_menu_item functions are left in place as
confirmed-dead-end history (see above) but are no longer called from
navigate_to_point_ar_report.

Not yet confirmed at all (will raise NotImplementedError where relevant):
  - Selecting a SPECIFIC insurance company (only "ALL" path is stubbed)
  - A report "to"/"from" date other than whatever the hardcoded calendar
    navigation currently clicks (see the CLI-date gap noted above)
"""

import argparse
import asyncio
import ctypes
import getpass
import json
import logging
import os
import shutil
import time
import xml.etree.ElementTree as ET
from ctypes import byref
from datetime import datetime
from pathlib import Path

import comtypes
import numpy as np
import comtypes.client
import keyring
import pythoncom
import win32com.client
import win32con
import win32gui
import win32process
import win32ui
from PIL import Image, ImageGrab
from pywinauto import Application, Desktop
from pywinauto import mouse as pywinauto_mouse
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys
from pywinauto.timings import TimeoutError as PywinautoTimeoutError
from pywinauto.timings import wait_until_passes
from winsdk.windows.globalization import Language
from winsdk.windows.graphics.imaging import BitmapAlphaMode, BitmapPixelFormat, SoftwareBitmap
from winsdk.windows.media.ocr import OcrEngine
from winsdk.windows.storage.streams import DataWriter

# Generates comtypes.gen.Accessibility (IAccessible) from oleacc.dll's type
# library on first run, cached after. Needed because the Reports flyout's
# menu items are theorized to be owner-drawn with no HWNDs of their own (see
# navigate_to_point_ar_report) -- pywinauto's "win32" backend only walks real
# child HWNDs (EnumChildWindows) and cannot see them, so reading them
# requires calling the legacy MSAA (IAccessible) interface directly.
comtypes.client.GetModule("oleacc.dll")
from comtypes.automation import VARIANT, VT_DISPATCH  # noqa: E402
from comtypes.gen import Accessibility  # noqa: E402


class NavigationCheckpoint(Exception):
    """Raised to deliberately halt the script at a point where a selector
    needs human verification from real window data before proceeding."""

LOGIN_WINDOW_TITLE_RE = r".*COMPASSIONATE WOMEN'S HEALTH.*"
MAIN_MENU_TITLE_RE = r"Main Menu COMPASSIONATE WOMEN'S HEALTH.*"
REPORT_SCREEN_TITLE_RE = r"Reports COMPASSIONATE WOMEN'S HEALTH.*"

# win32con doesn't always define this classic winuser.h constant; it's a
# fixed value (0), not version-dependent, so fall back to it directly rather
# than failing to import.
OBJID_WINDOW = getattr(win32con, "OBJID_WINDOW", 0x00000000)

# Same fallback pattern: the oleacc STATE_SYSTEM_HASPOPUP bit (fixed value,
# not version-dependent) isn't always present on comtypes.gen.Accessibility.
STATE_SYSTEM_HASPOPUP = getattr(Accessibility, "STATE_SYSTEM_HASPOPUP", 0x40000000)

# Same fallback pattern, for the oleacc roles a live run showed sitting
# BETWEEN a menu's hwnd-level root and its real (named) item rows -- see
# _capture_msaa_children's container-recursion comment.
ROLE_SYSTEM_WINDOW = getattr(Accessibility, "ROLE_SYSTEM_WINDOW", 0x9)
ROLE_SYSTEM_CLIENT = getattr(Accessibility, "ROLE_SYSTEM_CLIENT", 0xA)
ROLE_SYSTEM_MENUPOPUP = getattr(Accessibility, "ROLE_SYSTEM_MENUPOPUP", 0xB)
MSAA_CONTAINER_ROLES = (ROLE_SYSTEM_WINDOW, ROLE_SYSTEM_CLIENT, ROLE_SYSTEM_MENUPOPUP)
MSAA_MAX_RECURSION_DEPTH = 4

# PW_RENDERFULLCONTENT (not in win32con): asks PrintWindow to render a
# window's real content regardless of on-screen occlusion/z-order -- plain
# occlusion (not fixed-version-dependent) rather than a Windows-version
# difference, same fallback-constant rationale as the MSAA roles above.
PW_RENDERFULLCONTENT = 0x00000002

# Windows' built-in OCR engine language -- confirmed available live on this
# machine (OcrEngine.try_create_from_language("en-US") succeeded).
OCR_ENGINE_LANGUAGE = "en-US"

# Normalized (lowercased, punctuation-stripped) word sequence OCR must match
# CONSECUTIVELY to identify the 'Open Item Reports' row -- see
# _find_ocr_phrase for why this replaces a numeric confidence check (Windows'
# OCR engine exposes none).
OCR_OPEN_ITEM_REPORTS_TOKENS = ("open", "item", "reports")

# Same normalization for the 'Open Item: Charges For One or More Ins. Co'
# row inside the (reused) popup's Open Item Reports content -- punctuation
# (the ':' after 'Item', the '.' after 'Ins') is stripped by _find_ocr_phrase's
# normalize(), same as every other token list here. UNVERIFIED until a live
# run's recognized-words list (surfaced in full on a no-match RuntimeError)
# confirms OCR actually splits this row into exactly these words.
OCR_CHARGES_FOR_INS_CO_TOKENS = ("open", "item", "charges", "for", "one", "or", "more", "ins", "co")

# Inset from the popup's real right-hand content edge (the scrollbar's left
# edge when present, else the window's own right edge -- see
# _locate_submenu_arrow_via_ocr) to the '>' arrow's estimated click point.
OCR_ARROW_MARGIN_PX = 10

# The app's single top-level window has an empty title at the pre-login
# screen (confirmed via a live UIA dump: the "CCloudLogin" pane sits under a
# Dialog whose window_text() is '') -- it only picks up the
# "...COMPASSIONATE WOMEN'S HEALTH" text once MAIN_MENU_TITLE_RE applies
# post-login. LOGIN_WINDOW_TITLE_RE can't match that pre-login state, so
# connect by class name instead, which is stable across both states.
MEDENT_CLASS_NAME = "MedentClient"

# Desktop shortcut used to launch Medent when no instance is already running
# (confirmed present at this path on this machine). os.startfile() opens it
# exactly like a user double-clicking it would -- no separate installed .exe
# path was assumed.
MEDENT_LAUNCH_SHORTCUT = Path(r"C:\Users\Public\Desktop\MEDENT.lnk")

# TODO: confirm with user -- should live outside OneDrive sync since output
# contains PHI (claims/insurance data).
OUTPUT_DIR = Path(r"C:\MedentAutomation\reports")
AUDIT_LOG_PATH = Path(r"C:\MedentAutomation\audit_log.jsonl")

# Per user request: PDF reports should land in the local Downloads folder
# (also the Save dialog's own default location, per every live screenshot).
PDF_OUTPUT_DIR = Path(r"C:\Users\javva\Downloads")

# Windows Credential Manager (via `keyring`) storage for Medent login.
# The user id itself isn't secret, but keyring only looks up a password by
# (service, username), so it's stashed under a fixed lookup key alongside the
# real password, keyed by that user id.
KEYRING_SERVICE = "MedentAutomation"
KEYRING_USER_ID_KEY = "medent_user_id"


def audit(event, **fields):
    AUDIT_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    record = {"timestamp": datetime.now().isoformat(), "event": event, **fields}
    with open(AUDIT_LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(record) + "\n")
    logging.info("%s %s", event, fields)


def get_credentials():
    user_id = keyring.get_password(KEYRING_SERVICE, KEYRING_USER_ID_KEY)
    password = keyring.get_password(KEYRING_SERVICE, user_id) if user_id else None

    if user_id and password:
        audit("credentials_loaded_from_keyring", user=user_id)
        return user_id, password

    user_id = input("Medent User ID: ")
    password = getpass.getpass("Medent Password: ")

    save = input(
        "Save these credentials to Windows Credential Manager for future runs? [y/N]: "
    )
    if save.strip().lower() == "y":
        keyring.set_password(KEYRING_SERVICE, KEYRING_USER_ID_KEY, user_id)
        keyring.set_password(KEYRING_SERVICE, user_id, password)
        audit("credentials_saved_to_keyring", user=user_id)

    return user_id, password


def connect_to_app():
    return Application(backend="uia").connect(class_name=MEDENT_CLASS_NAME)


def launch_medent_if_needed(timeout=60):
    """Added so the script can run start-to-end without Medent already being
    open. Tries connect_to_app() first (same connection logic as before,
    untouched); only if that finds nothing running does it start Medent from
    MEDENT_LAUNCH_SHORTCUT and wait for its window before connecting. login()
    and wait_for_main_menu() still run afterward exactly as before -- this
    only adds the "nothing running yet" case in front of connect_to_app().
    """
    try:
        app = connect_to_app()
        audit("medent_already_running")
        return app
    except ElementNotFoundError:
        pass

    audit("launching_medent", shortcut=str(MEDENT_LAUNCH_SHORTCUT))
    os.startfile(MEDENT_LAUNCH_SHORTCUT)

    # wait_until_passes only retries while func RAISES one of `exceptions` --
    # a func that just returns False (e.g. a bare .exists() check) looks
    # "done" on its very first call and doesn't actually wait. Retrying
    # connect_to_app() itself (which raises ElementNotFoundError until the
    # window is really there) is what makes this actually poll.
    app = wait_until_passes(timeout, 1, connect_to_app, exceptions=(ElementNotFoundError,))
    audit("medent_window_appeared")

    return app


def is_logged_in(app):
    return app.window(title_re=MAIN_MENU_TITLE_RE).exists()


def login(app, user_id, password):
    win = app.top_window()
    login_pane = win.child_window(title="CCloudLogin", control_type="Pane")
    edits = login_pane.children(control_type="Edit")
    if len(edits) < 2:
        raise RuntimeError(f"Expected 2 Edit fields on login screen, found {len(edits)}")
    user_field, pass_field = edits[0], edits[1]
    user_field.set_edit_text(user_id)
    pass_field.set_edit_text(password)
    login_pane.child_window(title="Sign In", control_type="Pane").click_input()
    audit("login_submitted", user=user_id)


def wait_for_main_menu(app, timeout=20):
    wait_until_passes(timeout, 0.5, lambda: app.window(title_re=MAIN_MENU_TITLE_RE).exists())
    audit("main_menu_reached")

    # The window title flips to "Main Menu..." before the WebView2-rendered
    # content underneath has finished building its accessibility tree, so
    # searching for elements like "Reports" immediately after this can fail
    # with ElementNotFoundError even though the element shows up a moment
    # later. Wait for the "Reports" text to actually be resolvable.
    #
    # Live run 20261103_061104 crashed the whole script here with
    # RuntimeError("No windows for that process could be found") from
    # app.top_window() -- and widening it to a 10s retry (run
    # 20261103_061508) still timed out the entire window, never finding
    # anything. That's not a brief flicker; app.top_window() genuinely
    # disagrees with app.window(title_re=MAIN_MENU_TITLE_RE).exists(),
    # which had just succeeded moments earlier. The two aren't equivalent
    # checks: top_window() applies no title filter at all AND (per
    # pywinauto's findwindows.find_elements signature) defaults to
    # visible_only=True, so during this app's slow post-login
    # content-loading transition it can easily miss the very window we
    # already confirmed exists by title. Fix: stop asking a weaker,
    # differently-scoped question -- reuse the exact selector that was
    # just proven to work instead.
    win = app.window(title_re=MAIN_MENU_TITLE_RE)
    wait_until_passes(
        15, 0.5,
        lambda: win.child_window(title="Reports", control_type="Text").exists()
    )
    audit("main_menu_content_ready")


def dump_control_tree(parent, label, stamp=None):
    """Write the current control tree to a file for correcting a bad selector.

    WebView2 screens rebuild their accessibility tree on re-render, so a
    selector that's merely stale looks identical to one that's simply wrong.
    Dumping the real tree on failure is the only way to tell them apart.

    If `parent` itself can't be resolved (e.g. a window-level selector never
    matched anything), there's no tree to dump -- that's a different failure
    than a bad child selector, so it's reported distinctly instead of
    raising out of this diagnostic helper and obscuring the original error.

    The UIA tree can also throw transient COM errors (e.g. "An event was
    unable to invoke any of the subscribers") while walking it during an
    active re-render -- a live run surfaced exactly this, and because it
    isn't an ElementNotFoundError/PywinautoTimeoutError, it used to escape
    this function and replace/mask whatever original error the caller was
    trying to diagnose. Caught broadly here for the same reason the two
    named exceptions are: this helper must never itself raise and obscure
    the real failure it was called to investigate.

    `stamp`, when given (e.g. by _capture_diagnostics), lets this dump share
    a timestamp with sibling dumps from the same diagnostic event; the
    "_controltree_" stem keeps it from colliding with dump_top_level_windows'
    own file even when both share a stamp and a label.
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = stamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    dump_path = OUTPUT_DIR / f"{label}_controltree_{stamp}.txt"
    try:
        parent.print_control_identifiers(filename=str(dump_path))
    except (ElementNotFoundError, PywinautoTimeoutError) as exc:
        audit("control_tree_dump_failed_parent_not_found", label=label, error=str(exc))
        print(
            f"[{label}] Could not dump control tree: the parent window/control "
            f"itself could not be resolved ({exc}). This means a window-level "
            "selector is wrong or the expected window never appeared -- not a "
            "bad child selector."
        )
        return None
    except Exception as exc:
        audit("control_tree_dump_failed_unexpected", label=label, error=repr(exc))
        print(
            f"[{label}] Could not dump control tree due to an unexpected "
            f"error (likely a transient COM/UIA error while the tree was "
            f"being walked): {exc!r}. Treating this as 'no tree available' "
            "rather than letting it replace the original failure."
        )
        return None
    audit("control_tree_dumped", path=str(dump_path), label=label)
    return dump_path


def dump_top_level_windows(label, stamp=None):
    """Enumerate every visible top-level window right now, for identifying
    an unknown window's real title/class/handle instead of guessing.

    `stamp`, when given, lets this dump share a timestamp with sibling
    dumps from the same diagnostic event; the "_toplevel_" stem keeps it
    from colliding with dump_control_tree's own file.
    """
    windows = Desktop(backend="uia").windows()
    lines = []
    for w in windows:
        try:
            lines.append(
                f"title={w.window_text()!r} class={w.class_name()!r} "
                f"handle={w.handle} pid={w.process_id()} visible={w.is_visible()}"
            )
        except Exception as exc:
            lines.append(f"<failed to read window info: {exc}>")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = stamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    dump_path = OUTPUT_DIR / f"{label}_toplevel_{stamp}.txt"
    dump_path.write_text("\n".join(lines), encoding="utf-8")
    audit("top_level_windows_dumped", path=str(dump_path), label=label, count=len(windows))

    print(f"[{label}] {len(windows)} top-level windows -- see {dump_path}")
    for line in lines:
        print(f"  {line}")
    return dump_path


def click_when_ready(parent, timeout=15, retry=0.5, dump_label=None, **criteria):
    """Wait for a child control to resolve, then click it.

    Mirrors the retry wait_for_main_menu already does for the same
    WebView2 content-not-ready race -- a bare child_window().click_input()
    only gets pywinauto's own short default timeout, which isn't long
    enough for these screens.

    Checking exists() first and then clicking separately isn't enough: each
    child_window() access re-walks the live UIA tree from scratch rather
    than reusing a cached reference, and this screen's tree rebuilds often
    enough that it can go from "exists" to stale again before the very next
    lookup for the click. So the lookup+click is retried as one atomic
    action instead of trusting an exists() check done a moment earlier.
    On final failure, dump the control tree so the selector (if it's
    actually wrong, not just slow) can be corrected.
    """
    def attempt():
        parent.child_window(**criteria).click_input()
        return True

    try:
        wait_until_passes(timeout, retry, attempt)
    except PywinautoTimeoutError:
        if dump_label:
            dump_control_tree(parent, dump_label)
        raise


def _enumerate_menu_candidate_hwnds(main_hwnd):
    """Enumerate candidate flyout windows from TWO sources, not just the
    original single theory:

    1. 'WP Menu Class' children of main_hwnd (the original theory -- a real
       child HWND, but every capture of it so far, across multiple live
       runs, shows a permanently degenerate (0, -9, 0, -9) rect).
    2. 'Touch Popup Menu' TOP-LEVEL windows owned by MEDENT's own process --
       a live run surfaced exactly such a window (title 'Touch Popup') that
       persisted through the whole poll and diagnostics-capture window, a
       much better fit for an actual touch-friendly flyout than a
       permanently-zero-size child. Found via EnumWindows (top-level only)
       plus a class-name match, not FindWindowEx(main_hwnd, ...), since
       it's NOT a child of main_hwnd. Plausibly missed by every prior UIA-
       based investigation because those diffed the top-level window
       HANDLE SET looking for something freshly created; if this window is
       pre-created and merely shown/repositioned (the same pattern already
       confirmed for 'WP Menu Class'), its handle never changes and a
       diff-based search would never flag it.

    Each candidate is described with rect, IsWindowVisible, the WS_VISIBLE
    style bit (these two can legitimately disagree), the layered/
    transparent/tool-window extended-style bits, and the owning PID --
    compared against MEDENT's own process ID (`pid_mismatch`), since a
    window sharing either class name from a different process would be a
    false match and should never be trusted as a real candidate.

    Returns a list of dicts tagged with "class_name"; a hwnd that fails to
    describe (e.g. destroyed mid-enumeration) is recorded as
    {"hwnd": h, "class_name": ..., "error": ...} rather than raising, so
    one bad hwnd doesn't lose the others.
    """
    _, main_pid = win32process.GetWindowThreadProcessId(main_hwnd)

    candidates = []  # (class_name, hwnd)

    found = 0
    while True:
        found = win32gui.FindWindowEx(main_hwnd, found, "WP Menu Class", None)
        if not found:
            break
        candidates.append(("WP Menu Class", found))

    def _check_top_level(hwnd, _extra):
        try:
            if win32gui.GetClassName(hwnd) == "Touch Popup Menu":
                candidates.append(("Touch Popup Menu", hwnd))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_check_top_level, None)

    described = []
    for class_name, hwnd in candidates:
        try:
            rect = win32gui.GetWindowRect(hwnd)
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            described.append({
                "hwnd": hwnd,
                "class_name": class_name,
                "pid": pid,
                "pid_mismatch": pid != main_pid,
                "rect": rect,
                "is_window_visible": bool(win32gui.IsWindowVisible(hwnd)),
                "style_visible_bit": bool(style & win32con.WS_VISIBLE),
                "ex_layered": bool(ex_style & win32con.WS_EX_LAYERED),
                "ex_transparent": bool(ex_style & win32con.WS_EX_TRANSPARENT),
                "ex_toolwindow": bool(ex_style & win32con.WS_EX_TOOLWINDOW),
            })
        except Exception as exc:
            described.append({"hwnd": hwnd, "class_name": class_name, "error": repr(exc)})
    return described


def _dump_candidate_window_tree(hwnd, label):
    """UIA-side dump of a specific candidate window. dump_control_tree(win,
    ...) elsewhere only walks descendants of the main app window and would
    never include a separate top-level window like a 'Touch Popup Menu'
    candidate -- this targets the candidate itself, in case UIA exposes
    content here that it never did for 'WP Menu Class'. Never raises.
    """
    try:
        candidate_win = Desktop(backend="uia").window(handle=hwnd)
        return dump_control_tree(candidate_win, f"{label}_candidate_uia")
    except Exception as exc:
        audit("candidate_uia_dump_failed", label=label, hwnd=hwnd, error=repr(exc))
        return None


def _probe_msaa_child_count(hwnd, objid):
    """Cheap presence probe for the poll loop: can an IAccessible be
    obtained for hwnd/objid at all, and does it report any children?
    Deliberately lighter than _capture_msaa_menu_items -- no per-child
    name/role/rect reads -- so it stays safe to call twice per tick
    (OBJID_CLIENT and OBJID_WINDOW) without slowing the ~20ms poll cadence.
    Returns (ok, child_count, error).
    """
    try:
        ppv = ctypes.POINTER(Accessibility.IAccessible)()
        hr = ctypes.windll.oleacc.AccessibleObjectFromWindow(
            hwnd, objid, byref(Accessibility.IAccessible._iid_), byref(ppv)
        )
        if hr != 0 or not ppv:
            return False, None, f"hresult={hr}"
        return True, ppv.value.accChildCount, None
    except Exception as exc:
        return False, None, repr(exc)


def _poll_menu_signals(main_hwnd, label, timeout=1.5, interval=0.02):
    """Multi-signal replacement for the old rect-only wait loop. Every
    ~20ms tick, re-enumerates ALL candidate windows fresh (WP Menu Class
    children + Touch Popup Menu top-level windows -- the set may change)
    and layers on a cheap MSAA accChildCount probe (OBJID_CLIENT and
    OBJID_WINDOW) per hwnd. Every tick's full snapshot is recorded -- not
    just a final pass/fail -- so a live run shows which signal, if any,
    actually correlates with the menu being open, instead of just
    re-confirming a prior hypothesis failed again.

    The loop body is strictly limited to these cheap native/MSAA-count
    calls -- no full item capture, no control-tree dump, no screenshot runs
    inside it, so the ~20ms cadence holds even if the menu is only open
    briefly. Those heavier dumps stay confined to one-time calls outside
    this loop. Never raises. Returns (trace, jsonl_path, txt_path).
    """
    trace = []
    start = time.monotonic()
    deadline = start + timeout
    tick = 0
    while time.monotonic() < deadline:
        tick += 1
        hwnd_infos = _enumerate_menu_candidate_hwnds(main_hwnd)
        for info in hwnd_infos:
            if "error" in info:
                continue
            hwnd = info["hwnd"]
            ok_c, cnt_c, err_c = _probe_msaa_child_count(hwnd, win32con.OBJID_CLIENT)
            ok_w, cnt_w, err_w = _probe_msaa_child_count(hwnd, OBJID_WINDOW)
            info.update({
                "msaa_client_ok": ok_c, "msaa_client_children": cnt_c, "msaa_client_err": err_c,
                "msaa_window_ok": ok_w, "msaa_window_children": cnt_w, "msaa_window_err": err_w,
            })
        trace.append({"tick": tick, "t": round(time.monotonic() - start, 4), "hwnds": hwnd_infos})
        time.sleep(interval)

    jsonl_path, txt_path = _write_menu_trace(label, trace)
    return trace, jsonl_path, txt_path


def _write_menu_trace(label, trace):
    """Writes the full per-tick multi-signal trace as both a .jsonl
    (structured, diffable) and a fixed-width .txt (human-readable, meant to
    be pasted back directly as evidence). Returns (jsonl_path, txt_path).
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    jsonl_path = OUTPUT_DIR / f"{label}_menutrace_{stamp}.jsonl"
    txt_path = OUTPUT_DIR / f"{label}_menutrace_{stamp}.txt"

    with open(jsonl_path, "w", encoding="utf-8") as f:
        for snap in trace:
            f.write(json.dumps(snap) + "\n")

    header = (
        "tick  t(s)    class              hwnd        rect                     pid_mismatch "
        "vis   style_vis layered transp  msaa_client(ok/n/err)         msaa_window(ok/n/err)"
    )
    lines = [header]
    for snap in trace:
        if not snap["hwnds"]:
            lines.append(f"{snap['tick']:>4}  {snap['t']:>6}   <no candidate hwnd found>")
            continue
        for h in snap["hwnds"]:
            if "rect" not in h:
                lines.append(
                    f"{snap['tick']:>4}  {snap['t']:>6}   {h.get('class_name', '?'):<18} "
                    f"hwnd={h['hwnd']} <describe failed: {h.get('error')}>"
                )
                continue
            lines.append(
                f"{snap['tick']:>4}  {snap['t']:>6}   {h['class_name']:<18} "
                f"{h['hwnd']:<10} {str(h['rect']):<24} "
                f"{str(h['pid_mismatch']):<12} "
                f"{str(h['is_window_visible']):<5} {str(h['style_visible_bit']):<9} "
                f"{str(h['ex_layered']):<7} {str(h['ex_transparent']):<7} "
                f"{h['msaa_client_ok']}/{h['msaa_client_children']}/{h['msaa_client_err']}   "
                f"{h['msaa_window_ok']}/{h['msaa_window_children']}/{h['msaa_window_err']}"
            )
    txt_path.write_text("\n".join(lines), encoding="utf-8")
    audit("menu_signal_trace_written", jsonl=str(jsonl_path), txt=str(txt_path), ticks=len(trace), label=label)
    return jsonl_path, txt_path


def _best_populated_candidate(trace):
    """Ranks observed hwnd+tick snapshots instead of trusting the first
    non-zero rect it sees: a non-zero rect alone doesn't prove a window is
    the active, MEDENT-owned Reports menu. Prefers a hwnd+tick where the
    rect is non-zero, the PID matches MEDENT's process (no pid_mismatch),
    and at least one other signal agrees (is_window_visible, or a non-zero
    MSAA child count on either OBJID). A rect-only match with no
    corroborating signal, or a PID mismatch, is still recorded and used
    only as a last-resort fallback -- flagged as such via "corroborated".

    Returns a dict describing the chosen candidate (hwnd, rect, tick,
    pid_mismatch, corroborated), or None if no non-zero rect was ever
    observed at all.
    """
    fallback = None
    for snap in trace:
        for h in snap["hwnds"]:
            if "rect" not in h:
                continue
            left, top, right, bottom = h["rect"]
            if not (right - left > 0 and bottom - top > 0):
                continue

            corroborated = (not h["pid_mismatch"]) and (
                h["is_window_visible"]
                or (h["msaa_client_ok"] and (h["msaa_client_children"] or 0) > 0)
                or (h["msaa_window_ok"] and (h["msaa_window_children"] or 0) > 0)
            )
            candidate = {
                "hwnd": h["hwnd"],
                "class_name": h["class_name"],
                "rect": h["rect"],
                "tick": snap["tick"],
                "pid_mismatch": h["pid_mismatch"],
                "corroborated": corroborated,
            }
            if corroborated:
                return candidate
            if fallback is None:
                fallback = candidate
    return fallback


def _find_new_menu_candidate(trace, known_hwnds):
    """Like _best_populated_candidate, but only considers hwnds NOT already
    in known_hwnds. A submenu reuses the same window classes as its parent
    flyout ('WP Menu Class' / 'Touch Popup Menu'), so _best_populated_candidate
    alone would just re-match the already-open parent (already corroborated
    from the first tick) instead of detecting that a second, distinct popup
    appeared beside it. Same ranking/fallback rules otherwise. Returns None
    if no non-zero-rect candidate outside known_hwnds was ever observed.
    """
    fallback = None
    for snap in trace:
        for h in snap["hwnds"]:
            if "rect" not in h or h["hwnd"] in known_hwnds:
                continue
            left, top, right, bottom = h["rect"]
            if not (right - left > 0 and bottom - top > 0):
                continue

            corroborated = (not h["pid_mismatch"]) and (
                h["is_window_visible"]
                or (h["msaa_client_ok"] and (h["msaa_client_children"] or 0) > 0)
                or (h["msaa_window_ok"] and (h["msaa_window_children"] or 0) > 0)
            )
            candidate = {
                "hwnd": h["hwnd"],
                "class_name": h["class_name"],
                "rect": h["rect"],
                "tick": snap["tick"],
                "pid_mismatch": h["pid_mismatch"],
                "corroborated": corroborated,
            }
            if corroborated:
                return candidate
            if fallback is None:
                fallback = candidate
    return fallback


def _capture_click_screenshot(label):
    """Full-screen screenshot taken immediately after a click, before any
    slower diagnostic work (hwnd enumeration, MSAA probes) runs -- the menu
    is most likely to actually be visible right at this instant. Kept as a
    piece of evidence independent of any hwnd/MSAA signal: a human can
    visually confirm whether this actually shows a menu at a later-chosen
    candidate's reported rect. Never raises.
    """
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    shot_path = OUTPUT_DIR / f"{label}_screenshot_{stamp}.png"
    try:
        ImageGrab.grab(all_screens=True).save(shot_path)
        audit("screenshot_captured", path=str(shot_path), label=label)
    except Exception as exc:
        audit("screenshot_capture_failed", label=label, error=str(exc))
    return shot_path


def _capture_msaa_menu_items(hwnd, objid=win32con.OBJID_CLIENT):
    """Read WP Menu Class's menu items via true MSAA (IAccessible), not
    pywinauto's win32-backend children() -- that only walks EnumChildWindows
    (real HWNDs), and owner-drawn menu items are expected to have none, like
    native Win32 popup menus. Returns (items, errors); never raises, so a
    capture failure can be diagnosed from the caller rather than crashing
    navigation outright. Caller may retry with objid=OBJID_WINDOW if
    OBJID_CLIENT (the default) yields nothing.

    Delegates the actual per-child walk to _capture_msaa_children, which
    recurses -- a live run showed hwnd/objid's own immediate children are
    sometimes just 1-2 unnamed pass-through containers (role WINDOW/CLIENT/
    MENUPOPUP), not the real rows, so stopping at depth 0 silently returned
    those containers (empty name, no usable rect) instead of 'Open Item
    Reports' and its siblings.
    """
    try:
        ppv = ctypes.POINTER(Accessibility.IAccessible)()
        res = ctypes.windll.oleacc.AccessibleObjectFromWindow(
            hwnd, objid, byref(Accessibility.IAccessible._iid_), byref(ppv)
        )
        if res != 0 or not ppv:
            return [], [f"AccessibleObjectFromWindow failed, hresult={res}"]
        return _capture_msaa_children(ppv.value)
    except Exception as exc:
        return [], [f"MSAA capture failed: {exc}"]


def _capture_msaa_children(acc, depth=0, path=()):
    """Read `acc`'s immediate MSAA children, recursing into any unnamed
    container child (see MSAA_CONTAINER_ROLES) instead of recording it as if
    it were a real menu row -- a live run found exactly this shape (the
    hwnd-level root's only two children were empty-named, role-WINDOW
    pass-throughs; UIA independently confirmed a ScrollBar sibling under the
    same popup, consistent with one of those two being the real scrollable
    item list). A NAMED child is always kept as a leaf and never recursed
    into, even if its role happens to match a container role.

    Each child's properties (name, role, rect, state, default_action) is
    read in its own try/except -- a single failing property (e.g. a rect
    read) must not discard the whole item, since downstream submenu-
    detection logic needs to compare items by whichever properties it DID
    get, not silently lose rows. `rect` can legitimately be None if
    accLocation fails; callers must check for that rather than assume it's
    always a 4-tuple. `default_action` is frequently absent/empty for
    ordinary items -- that's not an error, so its failure is never recorded
    in `errors`.

    `accessible_path`, recorded on every returned item, is the sequence of
    per-depth AccessibleChildren() indices from the ORIGINAL hwnd-level root
    down to this item -- not just its local `index` among its immediate
    siblings -- because a nested item's `child_id` (CHILDID_SELF, i.e. 0) is
    only meaningful against ITS OWN immediate parent's IAccessible, not
    against the hwnd-level root. Re-activating a nested item later
    (_resolve_msaa_accessible) must walk this same path to get back the
    right IAccessible before calling accDoDefaultAction -- calling it on the
    hwnd-level root with a nested item's child_id would silently act on the
    wrong (or no) element once an item came from depth > 0.
    """
    items, errors = [], []
    try:
        count = acc.accChildCount
        obtained = ctypes.c_long(0)
        child_array = (VARIANT * count)()
        ctypes.windll.oleacc.AccessibleChildren(
            acc, 0, count, byref(child_array), byref(obtained)
        )
    except Exception as exc:
        errors.append(f"depth {depth}: AccessibleChildren failed: {exc}")
        return items, errors

    for i in range(obtained.value):
        var = child_array[i]
        child_path = path + (i,)
        try:
            if var.vt == VT_DISPATCH:
                child_acc = var.value.QueryInterface(Accessibility.IAccessible)
                child_id = 0  # CHILDID_SELF
            else:
                child_acc = acc
                child_id = var.value  # simple element's integer child id
        except Exception as exc:
            errors.append(f"depth {depth} child {i} resolve failed: {exc}")
            continue

        name = role = state = default_action = None
        rect = None
        try:
            name = child_acc.accName(child_id) or ""
        except Exception as exc:
            errors.append(f"depth {depth} child {i} accName failed: {exc}")
        try:
            role = child_acc.accRole(child_id)
        except Exception as exc:
            errors.append(f"depth {depth} child {i} accRole failed: {exc}")
        try:
            # comtypes auto-marshals this method's four [out] parameters as
            # a returned tuple -- it does NOT accept byref output pointers
            # the way the raw Win32/C signature suggests. The prior version
            # passed byref() pointers here, which always raised a TypeError
            # ("call takes exactly 2 arguments (6 given)") and left rect
            # silently None for every single item, every run.
            rect = child_acc.accLocation(child_id)
        except Exception as exc:
            errors.append(f"depth {depth} child {i} accLocation failed: {exc}")
        try:
            state = child_acc.accState(child_id)
        except Exception as exc:
            errors.append(f"depth {depth} child {i} accState failed: {exc}")
        try:
            default_action = child_acc.accDefaultAction(child_id)
        except Exception:
            default_action = None  # absent for most ordinary items -- not an error

        is_container = (
            var.vt == VT_DISPATCH
            and not (name or "").strip()
            and role in MSAA_CONTAINER_ROLES
        )
        if is_container and depth < MSAA_MAX_RECURSION_DEPTH:
            try:
                nested_count = child_acc.accChildCount
            except Exception:
                nested_count = 0
            if nested_count:
                nested_items, nested_errors = _capture_msaa_children(
                    child_acc, depth + 1, child_path
                )
                errors.extend(nested_errors)
                if nested_items:
                    items.extend(nested_items)
                    continue
                # Recursed but found nothing real underneath -- fall through
                # and record this container itself, rather than silently
                # dropping the row.

        items.append({
            "index": i,
            "name": name,
            "role": role,
            "rect": rect,
            "child_id": child_id,
            "is_simple": var.vt != VT_DISPATCH,
            "state": state,
            "default_action": default_action,
            "has_popup": bool(state is not None and (state & STATE_SYSTEM_HASPOPUP)),
            "accessible_path": child_path,
        })
    return items, errors


def _resolve_msaa_accessible(hwnd, objid, accessible_path):
    """Re-walk from hwnd/objid's own root down `accessible_path` (as
    recorded by _capture_msaa_children) to get back the exact IAccessible a
    previously-captured item's child_id is valid against. Needed because a
    nested item's child_id is CHILDID_SELF (0) relative to ITS OWN immediate
    parent, not the hwnd-level root -- calling accDoDefaultAction(child_id)
    on the root for a nested item would silently target the wrong element.
    Returns the resolved IAccessible, or raises with a message naming
    exactly which path segment failed (so a stale capture -- the menu
    having changed since accessible_path was recorded -- is diagnosable
    rather than looking like a generic COM failure).
    """
    ppv = ctypes.POINTER(Accessibility.IAccessible)()
    hr = ctypes.windll.oleacc.AccessibleObjectFromWindow(
        hwnd, objid, byref(Accessibility.IAccessible._iid_), byref(ppv)
    )
    if hr != 0 or not ppv:
        raise RuntimeError(f"AccessibleObjectFromWindow failed, hresult={hr}")
    acc = ppv.value

    for depth, i in enumerate(accessible_path):
        count = acc.accChildCount
        obtained = ctypes.c_long(0)
        child_array = (VARIANT * count)()
        ctypes.windll.oleacc.AccessibleChildren(acc, 0, count, byref(child_array), byref(obtained))
        if i >= obtained.value:
            raise RuntimeError(
                f"accessible_path{accessible_path!r} invalid at depth {depth}: "
                f"only {obtained.value} children now (menu content likely changed)"
            )
        var = child_array[i]
        is_last = depth == len(accessible_path) - 1
        if var.vt == VT_DISPATCH:
            acc = var.value.QueryInterface(Accessibility.IAccessible)
        elif is_last:
            return acc  # simple leaf: caller uses the ORIGINAL child_id against this acc
        else:
            raise RuntimeError(
                f"accessible_path{accessible_path!r} invalid at depth {depth}: "
                "expected a nested IAccessible (VT_DISPATCH) but found a simple element"
            )
    return acc


def _capture_window_bitmap(hwnd):
    """Capture hwnd's own rendered content directly via PrintWindow, NOT a
    desktop screenshot -- a live comparison showed ImageGrab(all_screens=True)
    can capture whatever window happens to be drawn ON TOP at that pixel
    instead of this one (confirmed live: OCR on a desktop screenshot matched
    'Open Item Reports' text from an unrelated, overlapping VS Code window
    title, not the real menu, because that window happened to cover the
    popup's on-screen rect at the moment of capture). PW_RENDERFULLCONTENT
    asks the window to render its real content into the given DC regardless
    of on-screen occlusion/z-order.

    Returns (PIL.Image in RGB, (left, top, right, bottom) on-screen window
    rect) so callers can convert a local OCR coordinate back to an absolute
    screen click point. Raises RuntimeError on a degenerate rect or a
    PrintWindow failure -- never silently returns a blank/wrong bitmap.
    """
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        raise RuntimeError(f"hwnd {hwnd} has a degenerate rect: {(left, top, right, bottom)}")

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    save_bitmap = win32ui.CreateBitmap()
    try:
        save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(save_bitmap)
        ok = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)
        if not ok:
            raise RuntimeError(f"PrintWindow(hwnd={hwnd}, PW_RENDERFULLCONTENT) returned 0 (failed)")
        bmpinfo = save_bitmap.GetInfo()
        bmpstr = save_bitmap.GetBitmapBits(True)
        image = Image.frombuffer(
            "RGB", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), bmpstr, "raw", "BGRX", 0, 1
        )
    finally:
        win32gui.DeleteObject(save_bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
    return image, (left, top, right, bottom)


def _resolve_child_dialog_hwnd(win, title_re):
    """Find a 'Window'-type control nested inside `win` whose title matches
    `title_re`, and return ITS OWN native window handle -- not `win`'s.

    Confirmed live (export_xml_ocr_locate_failed_controltree_20261103_
    041904.txt) that '(Filter list by 1st letter)' is a genuine child
    Dialog/Window with its own HWND -- its 'ComponentContainerWindow'/
    'ComponentItemWindow' children suggest a GPU-composited custom
    control, not plain GDI-painted content -- and PrintWindow on the
    TOP-LEVEL main window's hwnd does NOT include its rendered content at
    all. This was proven to be a capture-target problem, not a timing
    one: OCR against `main_hwnd` returned the exact same "popup never
    opened" word list after both a 3s and a 10s retry budget, while a
    real desktop screenshot taken at the same moment clearly showed the
    popup open. OCR must run directly against this child dialog's own
    hwnd instead (same _capture_window_bitmap/_run_ocr pipeline, just a
    different target hwnd) -- this is the fix for that, reusable for any
    later popup that turns out to be the same kind of child Dialog/Window
    (e.g. the 'Select XML Option' popup is expected to be one too).
    """
    dlg = win.child_window(title_re=title_re, control_type="Window")
    hwnd = dlg.handle
    if not hwnd:
        raise RuntimeError(
            f"Dialog matching {title_re!r} resolved via UIA but has no native window handle"
        )
    return hwnd


def _run_ocr(image):
    """Run Windows' built-in OCR engine (winsdk/Windows.Media.Ocr) against a
    PIL image, returning a flat list of {"text", "rect": (x, y, w, h)} dicts
    in reading order -- rect is LOCAL to `image`, not screen coordinates.
    OcrEngine.recognize_async requires a Bgra8/Gray8/Nv12 SoftwareBitmap, so
    the image is converted to BGRA (PIL gives RGBA) before handing it over.
    Raises RuntimeError if no OCR engine is available for
    OCR_ENGINE_LANGUAGE (e.g. the language pack isn't installed) -- never
    silently skips OCR and returns an empty result that looks like "no text
    found" for a different reason.
    """
    async def _recognize():
        rgba = image.convert("RGBA")
        w, h = rgba.size
        arr = np.array(rgba)
        bgra = arr[:, :, [2, 1, 0, 3]].copy()

        writer = DataWriter()
        writer.write_bytes(bytearray(bgra.tobytes()))
        buffer = writer.detach_buffer()

        bitmap = SoftwareBitmap.create_copy_from_buffer(
            buffer, BitmapPixelFormat.BGRA8, w, h, BitmapAlphaMode.IGNORE
        )
        engine = OcrEngine.try_create_from_language(Language(OCR_ENGINE_LANGUAGE))
        if engine is None:
            raise RuntimeError(
                f"No OCR engine available for language {OCR_ENGINE_LANGUAGE!r} -- "
                "the Windows OCR language pack may not be installed on this machine."
            )
        return await engine.recognize_async(bitmap)

    result = asyncio.run(_recognize())

    words = []
    for line in result.lines:
        for word in line.words:
            r = word.bounding_rect
            # winsdk's bounding_rect fields are floats -- round to int here,
            # at the source, so every downstream computation (row unions,
            # arrow-point math, bounds checks, the eventual click) works in
            # consistent integer pixel coordinates instead of silently
            # carrying floats into pywinauto_mouse.click (which requires
            # ints and raised TypeError on a live run when it didn't).
            words.append({
                "text": word.text,
                "rect": (round(r.x), round(r.y), round(r.width), round(r.height)),
            })
    return words


def _find_ocr_phrase(words, tokens):
    """Find `tokens` (e.g. ("open", "item", "reports")) as a sequence of
    CONSECUTIVE OCR words, case-insensitive after stripping surrounding
    punctuation from each word -- not a substring/fuzzy match. This is the
    correctness gate in place of a numeric OCR confidence score, which
    Windows' OCR engine does not expose at all (unlike e.g. Tesseract).

    Returns the list of matched word dicts (same length as tokens) for the
    match found, or None if there's no match. Raises RuntimeError if MORE
    THAN ONE non-overlapping match is found -- an ambiguous result is
    exactly the kind of thing that must halt for human review rather than
    silently picking one.
    """
    def normalize(text):
        return text.strip(" \t\r\n.,:;\"'“”‘’").lower()

    normalized = [normalize(w["text"]) for w in words]
    target = list(tokens)
    n = len(target)
    matches = [
        words[i:i + n]
        for i in range(len(words) - n + 1)
        if normalized[i:i + n] == target
    ]

    if len(matches) > 1:
        raise RuntimeError(
            f"Found {len(matches)} non-overlapping matches for {tokens!r} in "
            f"the OCR output -- refusing to guess which is the real menu row: {matches}"
        )
    return matches[0] if matches else None


def _locate_scrollbar_left_edge_local(menu_hwnd, window_rect):
    """Read the popup's own vertical scrollbar position via UIA (confirmed
    live to work for this control, even though UIA can't see the menu's row
    TEXT -- see the Pass 5 module note) and return its left edge in LOCAL
    coordinates (relative to the popup window's own top-left). The real
    clickable row area ends there, not at the window's right edge -- the
    scrollbar occupies a real, non-trivial strip of the popup's right side,
    so using the window edge directly would land a click on the scrollbar
    instead of the row's '>' arrow. Returns None (never raises) if no
    scrollbar is found; callers must fall back to a window-edge-relative
    margin in that case.
    """
    try:
        win = Desktop(backend="uia").window(handle=menu_hwnd)
        scrollbar = win.child_window(control_type="ScrollBar")
        rect = scrollbar.rectangle()
        return rect.left - window_rect[0]
    except Exception:
        return None


def _locate_menu_row_via_ocr(menu_hwnd, tokens):
    """Find the menu row whose text matches `tokens` (e.g. 'Open Item
    Reports'), entirely from the popup's own rendered content -- MSAA/UIA
    can only see generic window furniture here (see _capture_msaa_children
    and the Pass 5 module note), never the real row text, so this is the
    only way to locate any row in this control at all.

    Returns a dict: {row_rect, matched_words, scrollbar_left_edge_local,
    content_right_edge, window_rect, bitmap_size, ocr_words}, all in LOCAL
    (popup-window-relative) coordinates except window_rect. Raises
    RuntimeError (never guesses/falls back silently) if the phrase isn't
    found or is ambiguous -- callers computing a click point must do their
    own bounds-check against bitmap_size/row_rect and raise on failure too,
    never attempt a click anyway.
    """
    image, window_rect = _capture_window_bitmap(menu_hwnd)
    bitmap_w, bitmap_h = image.size

    words = _run_ocr(image)
    matched = _find_ocr_phrase(words, tokens)
    if matched is None:
        raise RuntimeError(
            f"OCR found {len(words)} word(s) in the popup's own rendered "
            f"content, but none matched the token sequence {tokens!r}. "
            f"Recognized words: {[w['text'] for w in words]}"
        )

    xs = [w["rect"][0] for w in matched]
    ys = [w["rect"][1] for w in matched]
    x2s = [w["rect"][0] + w["rect"][2] for w in matched]
    y2s = [w["rect"][1] + w["rect"][3] for w in matched]
    row_rect = (min(xs), min(ys), max(x2s) - min(xs), max(y2s) - min(ys))

    scrollbar_left_local = _locate_scrollbar_left_edge_local(menu_hwnd, window_rect)
    if scrollbar_left_local is not None and 0 < scrollbar_left_local <= bitmap_w:
        content_right_edge = scrollbar_left_local
    else:
        content_right_edge = bitmap_w

    return {
        "row_rect": row_rect,
        "matched_words": matched,
        "scrollbar_left_edge_local": scrollbar_left_local,
        "content_right_edge": content_right_edge,
        "window_rect": window_rect,
        "bitmap_size": (bitmap_w, bitmap_h),
        "ocr_words": words,
    }


def _locate_submenu_arrow_via_ocr(menu_hwnd, tokens=OCR_OPEN_ITEM_REPORTS_TOKENS):
    """Find the on-screen click point for the '>' submenu arrow belonging to
    the menu row whose text matches `tokens`, via _locate_menu_row_via_ocr.

    Returns a dict: {screen_point, local_point, row_rect, matched_words,
    scrollbar_left_edge_local, window_rect, ocr_words}. Raises RuntimeError
    (never guesses/falls back silently) if the phrase isn't found, is
    ambiguous, or the computed point falls outside the matched row's own
    bounds or outside the captured window bitmap -- any of these must halt
    via NavigationCheckpoint with this function's own diagnostic evidence
    instead of attempting a click anyway.
    """
    located = _locate_menu_row_via_ocr(menu_hwnd, tokens)
    row_rect = located["row_rect"]
    row_left, row_top, row_width, row_height = row_rect
    row_bottom = row_top + row_height
    bitmap_w, bitmap_h = located["bitmap_size"]
    window_rect = located["window_rect"]

    arrow_x = located["content_right_edge"] - OCR_ARROW_MARGIN_PX
    arrow_y = row_top + row_height // 2

    if not (0 <= arrow_x < bitmap_w and 0 <= arrow_y < bitmap_h):
        raise RuntimeError(
            f"Computed arrow point ({arrow_x}, {arrow_y}) falls outside the "
            f"captured window bitmap ({bitmap_w}x{bitmap_h}) -- "
            f"content_right_edge={located['content_right_edge']}, row_rect={row_rect}"
        )
    if not (row_top <= arrow_y <= row_bottom):
        raise RuntimeError(
            f"Computed arrow y={arrow_y} falls outside the matched row's own "
            f"vertical bounds {row_top}-{row_bottom} -- refusing to click."
        )

    screen_point = (window_rect[0] + arrow_x, window_rect[1] + arrow_y)

    return {
        "screen_point": screen_point,
        "local_point": (arrow_x, arrow_y),
        "row_rect": row_rect,
        "matched_words": located["matched_words"],
        "scrollbar_left_edge_local": located["scrollbar_left_edge_local"],
        "window_rect": window_rect,
        "ocr_words": located["ocr_words"],
    }


def _locate_menu_row_center_via_ocr(menu_hwnd, tokens):
    """Find the on-screen click point for a LEAF menu row (no '>' submenu
    arrow to target -- e.g. 'Open Item: Charges For One or More Ins. Co',
    which selects a report directly) via _locate_menu_row_via_ocr. Clicks
    the horizontal center of the matched text itself rather than the row's
    right edge (there's no arrow glyph there to land on for a leaf row),
    which is always inside the row's own clickable bounds and well clear of
    the scrollbar.

    Returns a dict: {screen_point, local_point, row_rect, matched_words,
    window_rect, ocr_words}. Raises RuntimeError on the same conditions as
    _locate_submenu_arrow_via_ocr.
    """
    located = _locate_menu_row_via_ocr(menu_hwnd, tokens)
    row_rect = located["row_rect"]
    row_left, row_top, row_width, row_height = row_rect
    row_right = row_left + row_width
    row_bottom = row_top + row_height
    bitmap_w, bitmap_h = located["bitmap_size"]
    window_rect = located["window_rect"]

    click_x = row_left + row_width // 2
    click_y = row_top + row_height // 2

    if not (0 <= click_x < bitmap_w and 0 <= click_y < bitmap_h):
        raise RuntimeError(
            f"Computed row-center click point ({click_x}, {click_y}) falls "
            f"outside the captured window bitmap ({bitmap_w}x{bitmap_h}) -- "
            f"row_rect={row_rect}"
        )
    if not (row_left <= click_x <= row_right and row_top <= click_y <= row_bottom):
        raise RuntimeError(
            f"Computed click point ({click_x}, {click_y}) falls outside the "
            f"matched row's own bounds {row_rect} -- refusing to click."
        )

    screen_point = (window_rect[0] + click_x, window_rect[1] + click_y)

    return {
        "screen_point": screen_point,
        "local_point": (click_x, click_y),
        "row_rect": row_rect,
        "matched_words": located["matched_words"],
        "window_rect": window_rect,
        "ocr_words": located["ocr_words"],
    }


def _expand_submenu_item_via_ocr(
    main_hwnd, menu_hwnd, previous_items, previous_errors, previous_trace,
    tokens=OCR_OPEN_ITEM_REPORTS_TOKENS,
):
    """OCR-driven replacement for _expand_submenu_item's locate-and-click
    step: MSAA/UIA cannot see this control's real row text at all (see the
    Pass 5 module note), so the accDoDefaultAction/keyboard/coordinate-guess
    sequence in _expand_submenu_item has nothing real to act on for this
    menu. Locates the row via _locate_submenu_arrow_via_ocr, clicks its
    computed arrow point, then reuses the EXISTING
    _find_submenu_evidence/_interpret_submenu_evidence signal machinery
    UNCHANGED to verify whether the submenu actually opened -- a click
    "landing" is never itself treated as success. Returns the SAME result
    shape _expand_submenu_item returns (plus an extra "ocr_locate" key), so
    the caller's existing NavigationCheckpoint branches keep working as-is.
    """
    result = {
        "expanded": False,
        "verified": False,
        "method": None,
        "attempts": [],
        "parent_hwnd": menu_hwnd,
        "submenu_hwnd": None,
        "submenu_items": [],
        "evidence": {},
        "diagnostic_paths": [],
        "ocr_locate": None,
    }

    try:
        located = _locate_submenu_arrow_via_ocr(menu_hwnd, tokens)
    except Exception as exc:
        result["attempts"].append(f"ocr_locate: failed: {exc!r}")
        return result
    result["ocr_locate"] = located
    result["attempts"].append(
        f"ocr_locate: found row_rect={located['row_rect']} screen_point={located['screen_point']}"
    )

    try:
        pywinauto_mouse.click(coords=located["screen_point"])
    except Exception as exc:
        result["attempts"].append(f"ocr_click: failed: {exc!r}")
        return result
    result["attempts"].append(f"ocr_click: sent at {located['screen_point']}")

    evidence = _find_submenu_evidence(
        main_hwnd, menu_hwnd, previous_items, previous_errors, previous_trace, timeout=1.0
    )
    result["evidence"]["ocr_click"] = evidence
    result["diagnostic_paths"].append(evidence["trace_paths"])
    verdict = _interpret_submenu_evidence(evidence)

    if verdict is None:
        return result  # expanded stays False -- the click produced no detectable change at all

    result["method"] = "ocr_click"
    result["expanded"] = True
    if verdict == "verified":
        result["verified"] = True
        if evidence["new_candidate"] is not None and evidence["new_candidate_items"]:
            result["submenu_hwnd"] = evidence["new_candidate"]["hwnd"]
            result["submenu_items"] = evidence["new_candidate_items"]
        else:
            result["submenu_hwnd"] = menu_hwnd
            result["submenu_items"] = evidence["current_items"]
    return result


def _click_msaa_menu_item(hwnd, item):
    """Activate a captured MSAA menu item, in order of preference:

    1. accDoDefaultAction (MSAA's own mechanism for invoking an item with no
       HWND of its own, how screen readers activate menu items) against both
       OBJID_CLIENT and OBJID_WINDOW (cheap alternative worth trying on this
       custom/legacy control).
    2. Keyboard navigation using the item's own captured index within the
       menu: foreground the menu window, send Home (select the first item)
       then Down `item["index"]` times, then Enter. This assumes the
       on-screen traversal order matches MSAA's child order and that Home
       lands on the first item -- both unconfirmed until a live run reports
       whether this actually lands on the right item (see the post-click
       verification in navigate_to_point_ar_report).
    3. A coordinate click at the item's OWN captured MSAA accLocation rect
       (never a hardcoded/fixed coordinate, and never the parent menu's
       rect).

    Every attempt's outcome -- including the AccessibleObjectFromWindow
    HRESULT, which the prior version never checked before dereferencing --
    is recorded, so a failure is diagnosable instead of being silently
    swallowed.

    Returns (method, attempts): method is the name of whichever mechanism
    was attempted last without raising (NOT proof it actually selected the
    right item -- that's confirmed separately by polling for a screen
    change), or None if every attempt failed outright.
    """
    attempts = []
    for objid, objid_name in (
        (win32con.OBJID_CLIENT, "OBJID_CLIENT"),
        (OBJID_WINDOW, "OBJID_WINDOW"),
    ):
        try:
            ppv = ctypes.POINTER(Accessibility.IAccessible)()
            hr = ctypes.windll.oleacc.AccessibleObjectFromWindow(
                hwnd, objid, byref(Accessibility.IAccessible._iid_), byref(ppv)
            )
            if hr != 0 or not ppv:
                attempts.append(f"{objid_name}: AccessibleObjectFromWindow hresult={hr}")
                continue
            ppv.value.accDoDefaultAction(item["child_id"])
            return f"accDoDefaultAction({objid_name})", attempts
        except Exception as exc:
            attempts.append(f"{objid_name}: accDoDefaultAction failed: {exc!r}")

    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
        send_keys("{HOME}")
        if item["index"] > 0:
            send_keys("{DOWN}" * item["index"])
        send_keys("{ENTER}")
        attempts.append(
            f"keyboard: HOME + {item['index']}x DOWN + ENTER (sent, outcome unverified here)"
        )
        return "keyboard", attempts
    except Exception as exc:
        attempts.append(f"keyboard navigation failed: {exc!r}")

    try:
        l, t, w, h = item["rect"]
        pywinauto_mouse.click(coords=(l + w // 2, t + h // 2))
        return "coordinate", attempts
    except Exception as exc:
        attempts.append(f"coordinate click failed: {exc!r}")
        return None, attempts


def _item_signature(item):
    """Ordered signature for comparing menu-item captures across two points
    in time: name (lowercased), role, index (position), and rect -- not
    name alone, since two different menus can coincidentally share an item
    name. Callers compare these as an ordered SEQUENCE (not a set), so a
    mere reordering of otherwise-identical items still registers as a
    change.
    """
    return (
        (item.get("name") or "").strip().lower(),
        item.get("role"),
        item.get("index"),
        item.get("rect"),
    )


def _find_submenu_evidence(main_hwnd, parent_menu_hwnd, previous_items, previous_errors, previous_trace, timeout=1.0):
    """Gather independent signals for whether a submenu opened, without
    assuming MEDENT creates a new, distinct popup HWND -- it may instead
    reuse the parent menu's own HWND and simply change its accessible
    children, MSAA child count, or geometry. Returns raw signals only;
    _interpret_submenu_evidence decides expanded/verified from the
    combination, never from one of these alone.

    `previous_items`/`previous_errors`/`previous_trace` must be the EXACT
    pre-expansion capture of parent_menu_hwnd (never re-derived here) -- a
    capture with non-empty previous_errors means that baseline may itself
    be incomplete, which limits how much a later "content changed"
    comparison against it can prove (surfaced here as baseline_incomplete;
    see the caller's verdict logic for how it's applied).
    """
    trace, jsonl_path, txt_path = _poll_menu_signals(
        main_hwnd, "submenu_evidence_poll", timeout=timeout, interval=0.02
    )

    # Signal A: a new, distinct menu hwnd.
    new_candidate = _find_new_menu_candidate(trace, known_hwnds={parent_menu_hwnd})
    new_candidate_items, new_candidate_errors = [], []
    if new_candidate is not None:
        new_candidate_items, new_candidate_errors = _capture_msaa_menu_items(
            new_candidate["hwnd"], objid=win32con.OBJID_CLIENT
        )
        if not new_candidate_items:
            items_w, errors_w = _capture_msaa_menu_items(new_candidate["hwnd"], objid=OBJID_WINDOW)
            if items_w:
                new_candidate_items, new_candidate_errors = items_w, new_candidate_errors + errors_w

    # Signal B: the parent hwnd's own accessible children changed. Compared
    # as an ordered sequence (name, role, index, rect) -- not a name-only
    # set -- so count, role, position/order, and rect all participate.
    current_items, current_errors = _capture_msaa_menu_items(parent_menu_hwnd, objid=win32con.OBJID_CLIENT)
    if not current_items:
        items_w, errors_w = _capture_msaa_menu_items(parent_menu_hwnd, objid=OBJID_WINDOW)
        if items_w:
            current_items, current_errors = items_w, current_errors + errors_w

    content_changed = (
        [_item_signature(it) for it in current_items] != [_item_signature(it) for it in previous_items]
    )

    # Signal C: MSAA child-count change on the parent hwnd, both OBJIDs.
    def _last_counts(trace_, hwnd):
        for snap in reversed(trace_):
            for h in snap["hwnds"]:
                if h.get("hwnd") == hwnd and "msaa_client_children" in h:
                    return h["msaa_client_children"], h["msaa_window_children"]
        return None, None

    prev_counts = _last_counts(previous_trace, parent_menu_hwnd)
    curr_counts = _last_counts(trace, parent_menu_hwnd)
    child_count_changed = prev_counts != curr_counts

    # Signal D: geometry change on the parent hwnd.
    def _last_rect(trace_, hwnd):
        for snap in reversed(trace_):
            for h in snap["hwnds"]:
                if h.get("hwnd") == hwnd and "rect" in h:
                    return h["rect"]
        return None

    prev_rect = _last_rect(previous_trace, parent_menu_hwnd)
    curr_rect = _last_rect(trace, parent_menu_hwnd)
    geometry_changed = prev_rect is not None and curr_rect is not None and prev_rect != curr_rect

    # Signal E: screenshot, for human visual confirmation -- never used here
    # to assert success/failure on its own.
    screenshot_path = _capture_click_screenshot("submenu_evidence")

    return {
        "new_candidate": new_candidate,
        "new_candidate_items": new_candidate_items,
        "new_candidate_errors": new_candidate_errors,
        "content_changed": content_changed,
        "current_items": current_items,
        "current_errors": current_errors,
        "baseline_incomplete": bool(previous_errors),
        "child_count_changed": child_count_changed,
        "prev_child_counts": prev_counts,
        "curr_child_counts": curr_counts,
        "geometry_changed": geometry_changed,
        "prev_rect": prev_rect,
        "curr_rect": curr_rect,
        "screenshot_path": screenshot_path,
        "trace_paths": (jsonl_path, txt_path),
    }


def _interpret_submenu_evidence(evidence):
    """Decide a verdict from _find_submenu_evidence's independent signals.
    No single signal is ever sufficient alone:

    - "verified": Signal A found a corroborated new menu hwnd WITH captured
      items, OR Signal B shows the parent hwnd's own content substantially
      changed (non-empty new capture) AND the baseline wasn't itself known
      incomplete -- an incomplete starting list makes "it changed"
      unreliable. Signal A doesn't depend on the baseline at all, so it can
      still fully verify even when the baseline was incomplete.
    - "inconclusive": only child-count and/or geometry changed, or an
      uncorroborated/empty new hwnd appeared, or content changed against a
      known-incomplete baseline -- real evidence of SOMETHING, not proof of
      WHAT.
    - None: nothing changed at all.
    """
    new_candidate = evidence["new_candidate"]
    if new_candidate is not None and new_candidate["corroborated"] and evidence["new_candidate_items"]:
        return "verified"
    if evidence["content_changed"] and evidence["current_items"] and not evidence["baseline_incomplete"]:
        return "verified"
    if (
        evidence["child_count_changed"]
        or evidence["geometry_changed"]
        or new_candidate is not None
        or (evidence["content_changed"] and evidence["baseline_incomplete"])
    ):
        return "inconclusive"
    return None


def _snapshot_top_level_windows_for_pid(pid):
    """Enumerate current top-level windows owned by `pid` as
    {hwnd: (class_name, is_visible)}, for diffing against a later snapshot.
    Used by _find_leaf_selection_evidence because this app has already been
    confirmed (see _enumerate_menu_candidate_hwnds) to pre-create some
    windows and merely show/hide/reposition them rather than creating a
    fresh hwnd on demand -- a plain new-handle diff would miss that, so
    visibility is tracked per handle too, not just presence.
    """
    result = {}

    def _check(hwnd, _extra):
        try:
            _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
            if win_pid == pid:
                result[hwnd] = (win32gui.GetClassName(hwnd), bool(win32gui.IsWindowVisible(hwnd)))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_check, None)
    return result


def _find_leaf_selection_evidence(main_hwnd, menu_hwnd, previous_hwnds, timeout=2.0, interval=0.05):
    """Gather independent signals for whether clicking a LEAF menu row (one
    with no further submenu -- e.g. selecting an actual report) took
    effect. Unlike _find_submenu_evidence (which looks for a submenu
    appearing while the SAME popup legitimately stays open), a leaf
    selection is expected to make `menu_hwnd` itself disappear while a
    genuinely different, non-menu window belonging to MEDENT's own process
    appears OR transitions from hidden to visible (see
    _snapshot_top_level_windows_for_pid for why "appears" alone isn't
    enough here). Both signals are polled together every tick, not checked
    once after a fixed sleep, since either could lag the other.

    `previous_hwnds` must be the EXACT pre-click snapshot from
    _snapshot_top_level_windows_for_pid (never re-derived here) -- taken
    before the click so a window that was already hidden-then-shown can be
    told apart from one that was visible the whole time.

    Returns raw signals only; _interpret_leaf_selection_evidence decides a
    verdict from the combination.
    """
    _, main_pid = win32process.GetWindowThreadProcessId(main_hwnd)
    deadline = time.monotonic() + timeout
    popup_closed = False
    new_or_shown_windows = {}

    while True:
        popup_closed = not win32gui.IsWindow(menu_hwnd) or not win32gui.IsWindowVisible(menu_hwnd)
        current_hwnds = _snapshot_top_level_windows_for_pid(main_pid)
        new_or_shown_windows = {
            h: info for h, info in current_hwnds.items()
            if h != menu_hwnd
            and info[0] not in ("Touch Popup Menu", "WP Menu Class")
            and info[1]  # currently visible
            and (h not in previous_hwnds or not previous_hwnds[h][1])  # new, or was hidden before
        }
        if popup_closed and new_or_shown_windows:
            break
        if time.monotonic() >= deadline:
            break
        time.sleep(interval)

    screenshot_path = _capture_click_screenshot("leaf_selection_evidence")

    return {
        "popup_closed": popup_closed,
        "new_or_shown_windows": new_or_shown_windows,
        "screenshot_path": screenshot_path,
    }


def _interpret_leaf_selection_evidence(evidence):
    """Decide a verdict from _find_leaf_selection_evidence's signals:

    - "verified": the popup closed AND a genuinely new-or-newly-shown,
      non-menu top-level window owned by MEDENT's own process appeared --
      both halves of the expected behavior.
    - "inconclusive": only one half happened (popup closed with no such
      window yet, or such a window appeared/was shown but the popup is
      still up) -- real evidence of something, not proof the report screen
      is what actually opened.
    - None: neither happened.
    """
    if evidence["popup_closed"] and evidence["new_or_shown_windows"]:
        return "verified"
    if evidence["popup_closed"] or evidence["new_or_shown_windows"]:
        return "inconclusive"
    return None


def _select_leaf_menu_item_via_ocr(main_hwnd, menu_hwnd, tokens=OCR_CHARGES_FOR_INS_CO_TOKENS):
    """OCR-driven click for a LEAF menu row (no further submenu -- selects
    an actual report/action, e.g. 'Open Item: Charges For One or More Ins.
    Co') via _locate_menu_row_center_via_ocr, then verifies via
    _find_leaf_selection_evidence/_interpret_leaf_selection_evidence that
    the popup actually closed AND a different, non-menu MEDENT window
    appeared/was shown -- a click "landing" is never itself treated as
    success, same discipline as _expand_submenu_item_via_ocr.

    Returns {"selected", "verified", "method", "attempts", "ocr_locate",
    "evidence", "new_window_hwnd"}.
    """
    result = {
        "selected": False,
        "verified": False,
        "method": None,
        "attempts": [],
        "ocr_locate": None,
        "evidence": None,
        "new_window_hwnd": None,
    }

    try:
        located = _locate_menu_row_center_via_ocr(menu_hwnd, tokens)
    except Exception as exc:
        result["attempts"].append(f"ocr_locate: failed: {exc!r}")
        return result
    result["ocr_locate"] = located
    result["attempts"].append(
        f"ocr_locate: found row_rect={located['row_rect']} screen_point={located['screen_point']}"
    )

    _, main_pid = win32process.GetWindowThreadProcessId(main_hwnd)
    previous_hwnds = _snapshot_top_level_windows_for_pid(main_pid)

    try:
        pywinauto_mouse.click(coords=located["screen_point"])
    except Exception as exc:
        result["attempts"].append(f"ocr_click: failed: {exc!r}")
        return result
    result["attempts"].append(f"ocr_click: sent at {located['screen_point']}")

    evidence = _find_leaf_selection_evidence(main_hwnd, menu_hwnd, previous_hwnds, timeout=2.0)
    result["evidence"] = evidence
    verdict = _interpret_leaf_selection_evidence(evidence)

    if verdict is None:
        return result  # selected stays False -- the click produced no detectable change at all

    result["method"] = "ocr_click"
    result["selected"] = True
    if verdict == "verified":
        result["verified"] = True
        result["new_window_hwnd"] = next(iter(evidence["new_or_shown_windows"]))
    return result


# Estimated inset from a menu row's right edge to the '>' submenu chevron's
# center, in pixels -- unverified until a live run's screenshot/evidence
# confirms it; see _expand_submenu_item's failure-path NavigationCheckpoint
# if the coordinate-click attempt never lands on the real arrow.
ARROW_MARGIN_PX = 14


def _expand_submenu_item(main_hwnd, menu_hwnd, item, previous_items, previous_errors, previous_trace):
    """Open the submenu for an MSAA menu item that has one (e.g. 'Open Item
    Reports', with a '>' arrow drawn at its right edge), trying
    progressively more manual mechanisms and VERIFYING each one via
    _find_submenu_evidence/_interpret_submenu_evidence -- unlike
    _click_msaa_menu_item's generic activation, a mechanism "not raising" is
    never treated as proof it had the intended effect (a prior live run
    showed the row getting highlighted/selected without its submenu ever
    opening).

    Order of attempts, stopping as soon as one verifies or is inconclusive:
      1. accDoDefaultAction (MSAA's own 'invoke' mechanism).
      2. Keyboard: focus the item (Home + Down * index, like
         _click_msaa_menu_item) then Right Arrow -- the native Win32 gesture
         for expanding a submenu, never Enter (which commits/selects).
      3. A coordinate click at the arrow's estimated position, computed
         ONLY from the item's own verified accLocation rect (never a fixed
         screen coordinate) -- preceded by explicit pre-click validation:
         the rect is non-degenerate, the computed point falls inside it,
         the parent menu is still visible, and the menu's own content
         hasn't already drifted since `previous_items` was captured (which
         would mean `item["rect"]` no longer corresponds to the real
         current row). Any failed check aborts this attempt without
         clicking.

    `previous_items`/`previous_errors`/`previous_trace` must be the exact
    pre-expansion baseline already captured by the caller -- never
    re-derived here, so an incomplete initial capture stays visible to the
    verdict logic instead of being silently replaced by a fresh one.

    Returns: {expanded, verified, method, attempts, parent_hwnd,
    submenu_hwnd, submenu_items, evidence (per-attempt-label),
    diagnostic_paths}.
    """
    result = {
        "expanded": False,
        "verified": False,
        "method": None,
        "attempts": [],
        "parent_hwnd": menu_hwnd,
        "submenu_hwnd": None,
        "submenu_items": [],
        "evidence": {},
        "diagnostic_paths": [],
    }

    def _finalize(label, verdict):
        result["method"] = label
        result["expanded"] = True
        evidence = result["evidence"][label]
        if verdict == "verified":
            result["verified"] = True
            if evidence["new_candidate"] is not None and evidence["new_candidate_items"]:
                result["submenu_hwnd"] = evidence["new_candidate"]["hwnd"]
                result["submenu_items"] = evidence["new_candidate_items"]
            else:
                result["submenu_hwnd"] = menu_hwnd
                result["submenu_items"] = evidence["current_items"]
        return result

    def _attempt(label, action_fn):
        try:
            action_fn()
        except Exception as exc:
            result["attempts"].append(f"{label}: failed: {exc!r}")
            return None
        result["attempts"].append(f"{label}: sent")
        evidence = _find_submenu_evidence(
            main_hwnd, menu_hwnd, previous_items, previous_errors, previous_trace, timeout=1.0
        )
        result["evidence"][label] = evidence
        result["diagnostic_paths"].append(evidence["trace_paths"])
        return _interpret_submenu_evidence(evidence)

    # 1. accDoDefaultAction (OBJID_CLIENT, fallback OBJID_WINDOW). Must
    #    resolve via item["accessible_path"] (_resolve_msaa_accessible), not
    #    just re-fetch menu_hwnd's own root -- when the item came from a
    #    nested container (see _capture_msaa_children), its child_id
    #    (CHILDID_SELF) is only meaningful against that nested IAccessible,
    #    not the hwnd-level root; calling accDoDefaultAction on the root
    #    with a nested item's child_id would silently act on the wrong
    #    element (or CHILDID_SELF of the root itself) instead of raising.
    def _do_default_action():
        last_exc = None
        for objid, objid_name in ((win32con.OBJID_CLIENT, "OBJID_CLIENT"), (OBJID_WINDOW, "OBJID_WINDOW")):
            try:
                resolved_acc = _resolve_msaa_accessible(
                    menu_hwnd, objid, item.get("accessible_path", ())
                )
                resolved_acc.accDoDefaultAction(item["child_id"])
                return
            except Exception as exc:
                last_exc = exc
        raise RuntimeError(f"accDoDefaultAction failed for both OBJID_CLIENT and OBJID_WINDOW: {last_exc!r}")

    verdict = _attempt("accDoDefaultAction", _do_default_action)
    if verdict:
        return _finalize("accDoDefaultAction", verdict)

    # 2. Keyboard Right Arrow.
    def _keyboard_right_arrow():
        win32gui.SetForegroundWindow(menu_hwnd)
        time.sleep(0.1)
        send_keys("{HOME}")
        if item["index"] > 0:
            send_keys("{DOWN}" * item["index"])
        send_keys("{RIGHT}")

    verdict = _attempt("keyboard_right_arrow", _keyboard_right_arrow)
    if verdict:
        return _finalize("keyboard_right_arrow", verdict)

    # 3. Coordinate click at the estimated arrow position, with full
    #    pre-click validation -- never clicks on an unverified location.
    def _coordinate_click_arrow():
        rect = item.get("rect")
        if not rect or rect[2] <= 0 or rect[3] <= 0:
            raise RuntimeError(f"item rect is missing/degenerate: {rect!r}")
        left, top, width, height = rect

        arrow_x = left + max(width - ARROW_MARGIN_PX, width // 2)
        arrow_y = top + height // 2
        if not (left <= arrow_x <= left + width and top <= arrow_y <= top + height):
            raise RuntimeError(
                f"computed arrow point ({arrow_x}, {arrow_y}) falls outside row rect {rect}"
            )

        still_visible = any(
            h.get("hwnd") == menu_hwnd and h.get("is_window_visible")
            for h in _enumerate_menu_candidate_hwnds(main_hwnd)
        )
        if not still_visible:
            raise RuntimeError("parent menu is no longer visible -- refusing to click a stale location")

        current_items, _ = _capture_msaa_menu_items(menu_hwnd, objid=win32con.OBJID_CLIENT)
        if not current_items:
            current_items, _ = _capture_msaa_menu_items(menu_hwnd, objid=OBJID_WINDOW)
        if [_item_signature(it) for it in current_items] != [_item_signature(it) for it in previous_items]:
            raise RuntimeError(
                "menu content already changed since the baseline capture -- "
                "item['rect'] may no longer correspond to the current row"
            )

        pywinauto_mouse.click(coords=(arrow_x, arrow_y))

    verdict = _attempt("coordinate_click_arrow", _coordinate_click_arrow)
    if verdict:
        return _finalize("coordinate_click_arrow", verdict)

    return result


def _capture_diagnostics(label, win, menu_hwnds=None, menu_items=None, menu_errors=None, menu_trace_paths=None):
    """Consolidated failure-path diagnostics: top-level window list and main
    window control tree (existing conventions), a full-screen screenshot
    (pattern reused from find_open_item_reports.py), and -- if available --
    the raw MSAA capture results and poll-trace file paths, since a plain
    top-level window list is already known useless for WP Menu Class (it's
    never a top-level window). All artifacts from one call share a single
    `stamp` and distinct filename stems, so nothing here can silently
    overwrite anything else (the prior version let dump_top_level_windows
    and dump_control_tree clobber each other's output under the same name).
    """
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dump_top_level_windows(label, stamp=stamp)
    dump_control_tree(win, label, stamp=stamp)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    shot_path = OUTPUT_DIR / f"{label}_screenshot_{stamp}.png"
    try:
        ImageGrab.grab(all_screens=True).save(shot_path)
        audit("screenshot_captured", path=str(shot_path), label=label)
    except Exception as exc:
        audit("screenshot_capture_failed", label=label, error=str(exc))

    if menu_hwnds is not None:
        msaa_path = OUTPUT_DIR / f"{label}_msaa_{stamp}.txt"
        lines = [f"WP Menu Class hwnd(s): {menu_hwnds}"]
        if menu_trace_paths:
            lines.append(f"Multi-signal poll trace (structured): {menu_trace_paths[0]}")
            lines.append(f"Multi-signal poll trace (human-readable): {menu_trace_paths[1]}")
        if menu_items is not None:
            lines.append(f"{len(menu_items)} item(s) captured:")
            lines.extend(f"  {it}" for it in menu_items)
        if menu_errors:
            lines.append("errors:")
            lines.extend(f"  {e}" for e in menu_errors)
        msaa_path.write_text("\n".join(lines), encoding="utf-8")
        audit("msaa_menu_dumped", path=str(msaa_path), label=label)


def navigate_to_point_ar_report(app, output_path):
    """Click Reports, broaden hwnd/signal detection for the WP Menu Class
    flyout, identify the 'Open Item Reports' item via MSAA, verify it has a
    usable position, and expand its submenu (see _expand_submenu_item) via a
    verified accDoDefaultAction -> Right Arrow -> arrow-coordinate-click
    sequence, then drive the full report-filter/export/logout flow through
    to completion. Every branch halts via NavigationCheckpoint with the
    richest evidence gathered so far when something doesn't match what's
    been confirmed live -- see the module docstring.

    `output_path` is where the real report data (an XML Spreadsheet, read
    out of Excel via COM -- see below) ends up; a sibling '.pdf' is also
    written from the native Print dialog along the way.
    """
    win = app.top_window()
    main_hwnd = win.handle  # captured before the click, see _enumerate_menu_candidate_hwnds

    click_when_ready(
        win, title="Reports", control_type="Text", dump_label="reports_tile_not_found"
    )
    audit("clicked_reports_tile")
    _capture_click_screenshot("clicked_reports_tile")

    initial_hwnds = _enumerate_menu_candidate_hwnds(main_hwnd)
    audit("reports_menu_hwnds_enumerated", hwnds=initial_hwnds, count=len(initial_hwnds))
    if not initial_hwnds:
        _capture_diagnostics("reports_menu_hwnd_not_found", win)
        raise NavigationCheckpoint(
            "Clicked 'Reports' but no candidate menu window (neither a "
            "'WP Menu Class' child of the main MedentClient window, nor a "
            "'Touch Popup Menu' top-level window owned by its process) was "
            "found immediately afterward. The flyout control may use a "
            "different class name than either previously observed, or may "
            "require a different OBJID. Inspect the dumps under OUTPUT_DIR."
        )

    trace, jsonl_path, txt_path = _poll_menu_signals(
        main_hwnd, "reports_menu_poll", timeout=1.5, interval=0.02
    )
    audit("reports_menu_poll_complete", jsonl=str(jsonl_path), txt=str(txt_path), ticks=len(trace))

    candidate = _best_populated_candidate(trace)
    if candidate is None:
        all_hwnds = sorted({h["hwnd"] for snap in trace for h in snap["hwnds"] if "hwnd" in h})
        _capture_diagnostics(
            "reports_menu_never_populated", win,
            menu_hwnds=all_hwnds, menu_trace_paths=(jsonl_path, txt_path),
        )
        raise NavigationCheckpoint(
            f"Found candidate hwnd(s) {all_hwnds} but NONE ever reported "
            f"a non-zero rect across {len(trace)} poll ticks (~1.5s). Each "
            "tick also recorded visibility flags, PID match against "
            "MEDENT's own process, and cheap MSAA accChildCount probes "
            f"(OBJID_CLIENT and OBJID_WINDOW) for every candidate (both "
            "'WP Menu Class' children and any 'Touch Popup Menu' top-level "
            f"windows) -- inspect {txt_path} (human-readable) or "
            f"{jsonl_path} (structured) for whether any of those signals "
            "ever diverged from the rect; that's the evidence needed to "
            "pick a real gating signal next, not a re-confirmation that "
            "rect-polling fails again."
        )

    menu_hwnd = candidate["hwnd"]
    audit(
        "reports_menu_populated", hwnd=menu_hwnd, class_name=candidate["class_name"],
        rect=candidate["rect"], tick=candidate["tick"],
        corroborated=candidate["corroborated"], pid_mismatch=candidate["pid_mismatch"],
    )
    _dump_candidate_window_tree(menu_hwnd, "reports_menu")

    items, errors = _capture_msaa_menu_items(menu_hwnd, objid=win32con.OBJID_CLIENT)
    if not items:
        items_w, errors_w = _capture_msaa_menu_items(menu_hwnd, objid=OBJID_WINDOW)
        audit("reports_menu_captured_objid_window_retry", item_count=len(items_w), errors=errors_w)
        if items_w:
            items, errors = items_w, errors + errors_w
    audit("reports_menu_captured", hwnd=menu_hwnd, item_count=len(items), errors=errors)
    if not items:
        _capture_diagnostics(
            "reports_menu_capture_empty", win,
            menu_hwnds=[menu_hwnd], menu_errors=errors, menu_trace_paths=(jsonl_path, txt_path),
        )
        raise NavigationCheckpoint(
            f"A candidate window ({candidate['class_name']}, hwnd {menu_hwnd}) "
            f"was identified (corroborated={candidate['corroborated']}, "
            f"non-zero rect at tick {candidate['tick']}) but no MSAA child "
            "items could be read via EITHER OBJID_CLIENT or OBJID_WINDOW -- "
            "see captured errors, and check reports_menu_candidate_uia_*.txt "
            "in case UIA exposes this window's content where MSAA doesn't. "
            "Inspect the dumps before guessing further."
        )

    # Pass 5: MSAA only ever exposes generic window furniture for this
    # control (System/Application/titlebar/scrollbar/grip -- see the Pass 5
    # module note), never a real 'Open Item Reports' row, so `items` is kept
    # here only as the pre-click baseline _find_submenu_evidence compares
    # against -- locating and clicking the row itself goes through OCR.
    #
    # Deliberately NOT calling _capture_diagnostics here (as a prior version
    # did): its control-tree dump of the main window is slow enough (~7s,
    # confirmed live via a COMError-triggered retry) to sit in the gap
    # between confirming this popup is populated and OCR-clicking its
    # 'Open Item Reports' row -- long enough that the touch-optimized popup
    # auto-dismissed itself before the click, surfacing as a GetWindowRect
    # 'Invalid window handle' error on `menu_hwnd` in _locate_submenu_arrow_via_ocr.
    # The same evidence (menu_hwnds/items/trace_paths) is already captured on
    # both the failure path (open_item_reports_submenu_not_opened, below) and
    # the success path (open_item_reports_submenu_opened), so nothing here is
    # lost by not also dumping it eagerly on this hot, timing-sensitive path.
    audit(
        "reports_menu_msaa_has_no_real_rows", hwnd=menu_hwnd, item_count=len(items),
        note="MSAA/UIA expose no real row text for this control; locating "
             "'Open Item Reports' goes through OCR instead (see Pass 5).",
    )

    result = _expand_submenu_item_via_ocr(
        main_hwnd, menu_hwnd,
        previous_items=items, previous_errors=errors, previous_trace=trace,
    )
    audit(
        "open_item_reports_submenu_expand_result",
        **{k: v for k, v in result.items() if k not in ("evidence", "diagnostic_paths")},
    )
    _capture_click_screenshot("open_item_reports_submenu_expand_attempted")

    if not result["expanded"]:
        _capture_diagnostics(
            "open_item_reports_submenu_not_opened", win,
            menu_hwnds=[menu_hwnd], menu_items=items, menu_errors=result["attempts"],
            menu_trace_paths=(jsonl_path, txt_path),
        )
        raise NavigationCheckpoint(
            "OCR-based locate/click for 'Open Item Reports' did not produce "
            "any detectable change (new window, content, child-count, or "
            f"geometry) within ~1s, or failed before clicking at all. "
            f"Attempts: {result['attempts']}. ocr_locate={result['ocr_locate']!r}. "
            f"Inspect the per-attempt poll traces ({result['diagnostic_paths']}) "
            "and screenshots under OUTPUT_DIR -- either the phrase/arrow "
            "location logic is wrong (see ocr_locate/attempts above for which "
            "step failed), or the computed arrow point landed somewhere that "
            "doesn't actually expand the submenu."
        )

    if not result["verified"]:
        _capture_diagnostics(
            "open_item_reports_submenu_inconclusive", win,
            menu_hwnds=[menu_hwnd], menu_items=items, menu_errors=result["attempts"],
            menu_trace_paths=(jsonl_path, txt_path),
        )
        last_evidence = result["evidence"].get(result["method"], {})
        raise NavigationCheckpoint(
            f"Activating 'Open Item Reports' via method={result['method']!r} "
            "produced a detectable change (child-count and/or geometry "
            "shift, or an uncorroborated/empty new window), but it could "
            "not be confidently identified as the submenu actually opening "
            f"-- evidence: {last_evidence}. This is reported as "
            "INCONCLUSIVE, not success. Inspect the diagnostics/screenshots "
            f"under OUTPUT_DIR ({result['diagnostic_paths']}) before "
            "deciding whether this is real and adjusting the detection "
            "logic accordingly."
        )

    # Verified: don't assume identical item NAMES prove it's really a
    # distinct submenu -- a different menu could coincidentally share a
    # label. Compare the full ordered signature (name, role, index, rect)
    # against the ORIGINAL parent-menu capture, same as the internal
    # Signal-B comparison.
    submenu_items = result["submenu_items"]
    distinct_from_parent = (
        [_item_signature(it) for it in submenu_items] != [_item_signature(it) for it in items]
    )
    reused_parent_hwnd = result["submenu_hwnd"] == result["parent_hwnd"]

    audit(
        "open_item_reports_submenu_opened",
        method=result["method"], submenu_hwnd=result["submenu_hwnd"],
        reused_parent_hwnd=reused_parent_hwnd,
        submenu_items=submenu_items, distinct_from_parent=distinct_from_parent,
    )
    # Deliberately NOT calling _capture_diagnostics here -- same reasoning as
    # the identical omission above after reports_menu_msaa_has_no_real_rows:
    # its control-tree dump of the main window is slow enough to risk the
    # (reused) popup auto-closing again before the next OCR click, for no
    # benefit (the same evidence is captured on both outcomes below).

    # The submenu's own rows are generic MSAA furniture too (same Pass 5
    # situation as the parent Reports menu, confirmed live: submenu_items
    # above are 'System'/'Application'/'Vertical'/'Horizontal'/'', never a
    # real 'Open Item: Charges...' row), and it has no '>' arrow to expand
    # (a leaf report row, not another flyout) -- so selecting it goes
    # through the same OCR approach, just clicking the row's own center
    # instead of a submenu arrow (see _locate_menu_row_center_via_ocr).
    leaf_result = _select_leaf_menu_item_via_ocr(main_hwnd, result["submenu_hwnd"])
    audit(
        "charges_for_ins_co_select_result",
        **{k: v for k, v in leaf_result.items() if k != "evidence"},
    )
    _capture_click_screenshot("charges_for_ins_co_select_attempted")

    if not leaf_result["selected"]:
        _capture_diagnostics(
            "charges_for_ins_co_not_selected", win,
            menu_hwnds=[menu_hwnd, result["submenu_hwnd"]], menu_items=submenu_items,
            menu_errors=leaf_result["attempts"],
        )
        raise NavigationCheckpoint(
            "OCR-based locate/click for 'Open Item: Charges For One or More "
            "Ins. Co' did not produce any detectable change (the popup "
            "closing and a different MEDENT window appearing/being shown) "
            f"within ~2s, or failed before clicking at all. Attempts: "
            f"{leaf_result['attempts']}. ocr_locate={leaf_result['ocr_locate']!r}. "
            "Inspect the screenshots under OUTPUT_DIR and the recognized "
            "OCR words in ocr_locate (if present) -- either "
            "OCR_CHARGES_FOR_INS_CO_TOKENS doesn't match how OCR actually "
            "splits this row's text, or the computed row-center point "
            "landed somewhere that doesn't actually select the row."
        )

    # leaf_result["verified"] (the generic new-top-level-window heuristic
    # from _find_leaf_selection_evidence) is NOT the authoritative signal
    # here -- a live run confirmed the report filter screen replaces content
    # INSIDE the existing MedentClient window rather than opening a new
    # top-level window, so that heuristic reports "inconclusive" even on a
    # real success.
    #
    # The 'Print A/R Tasks on Report?' Yes/No dialog does NOT appear on its
    # own just from reaching this screen (an earlier assumption here that a
    # live run disproved -- confirm_report_popup() found no such window for
    # a full 5s retry budget). It only appears after clicking this screen's
    # own 'Print A/R Tasks on Report' field, which is a DIFFERENT control
    # from the dialog: the field's own title has no trailing '?' and the
    # field is a child of `win` itself, not a separate titled window.
    # Confirmed via print_ar_tasks_field_not_found_controltree_20261102_065517.txt
    # (captured by this exact failure live) that on THIS screen the field is
    # a Pane, not a Text -- unlike the Text-typed rows elsewhere in
    # configure_report_filters ('Report by Patient Name' etc.), so it does
    # not share their control_type.
    click_when_ready(
        win, title="Print A/R Tasks on Report", control_type="Pane",
        dump_label="print_ar_tasks_field_not_found",
    )
    audit("clicked_print_ar_tasks_on_report_field")

    try:
        confirm_report_popup(win)
    except (PywinautoTimeoutError, ElementNotFoundError) as exc:
        # TimeoutError's own repr is just "TimeoutError()" -- the real
        # cause pywinauto's wait_until_passes wraps lives in
        # original_exception (see its docstring), and omitting it here
        # cost two round-trips already diagnosing this exact failure.
        underlying = getattr(exc, "original_exception", None)
        _capture_diagnostics(
            "charges_for_ins_co_select_inconclusive", win,
            menu_hwnds=[menu_hwnd, result["submenu_hwnd"]], menu_items=submenu_items,
            menu_errors=leaf_result["attempts"]
            + [f"confirm_report_popup: failed: {exc!r}; original_exception={underlying!r}"],
        )
        raise NavigationCheckpoint(
            "Clicking 'Open Item: Charges For One or More Ins. Co' produced "
            "a detectable change (see charges_for_ins_co_select_result: "
            f"{leaf_result['evidence']}), and the 'Print A/R Tasks on "
            "Report' field was clicked, but confirm_report_popup() could "
            "not resolve-and-click the 'Print A/R Tasks on Report?' -> "
            f"'Yes' control within its retry budget: {exc!r} "
            f"(original_exception={underlying!r}). This is reported as "
            "INCONCLUSIVE, not success. Inspect the "
            "charges_for_ins_co_select_attempted screenshot under "
            "OUTPUT_DIR before deciding whether the field click landed "
            "somewhere else, the dialog needs a longer retry budget, or "
            "the dialog's real title/structure differs from what's assumed."
        )

    # Same atomic resolve-and-click discipline as confirm_report_popup
    # (see its docstring) -- a separate exists()-check here would be
    # exactly the stale-reference pattern click_when_ready's own docstring
    # already warns about, now confirmed live for this app's controls twice.
    #
    # Two wrong assumptions here, both confirmed via a control-tree dump
    # taken at the exact moment of a live TimeoutError
    # (task_types_window_not_found_controltree_20261102_084158.txt):
    # 'Task Types' is NOT a top-level window -- Desktop(...).window() only
    # searches top-level windows, but the dump shows 'Task Types' as a
    # child Window nested inside the same top-level Dialog as everything
    # else on screen, so it must be resolved as a descendant of `win`.
    # And 'All' is control_type="Pane", not "Button" -- same touch-UI
    # pattern already seen for 'Yes'/'No' in confirm_report_popup. Unlike
    # that earlier case, 'All' here genuinely IS a child of the 'Task
    # Types' window per the dump's nesting (not a sibling mistaken for a
    # child), so a plain chained child_window() lookup is correct.
    def _click_task_types_all():
        win.child_window(
            title="Task Types", control_type="Window",
        ).child_window(
            title="All", control_type="Pane",
        ).click_input()
        return True

    try:
        wait_until_passes(5, 0.3, _click_task_types_all)
    except (PywinautoTimeoutError, ElementNotFoundError) as exc:
        _capture_diagnostics("task_types_window_not_found", win)
        raise NavigationCheckpoint(
            "confirm_report_popup() ran (clicked 'Yes' on 'Print A/R Tasks "
            "on Report?', then sent ENTER x2), but the 'Task Types' window "
            "-- expected next per a live walkthrough -- and its 'All' "
            f"button could not be resolved-and-clicked within 5s: {exc!r}. "
            "Inspect the screenshots/control-tree dump under OUTPUT_DIR: "
            "either confirm_report_popup's own unverified assumptions "
            "(which window has focus for the ENTER keystrokes; see its "
            "docstring) are wrong, there's an intermediate prompt between "
            "'Yes' and 'Task Types' that isn't handled yet, or the 'All' "
            "button isn't a control_type=\"Button\" as assumed."
        )

    audit("clicked_task_types_all")
    _capture_click_screenshot("task_types_all_clicked")

    # 'All' alone was assumed to confirm AND dismiss 'Task Types' in one
    # action (task_types_ok_not_found_controltree_20261102_085533.txt: no
    # 'Task Types' window left, filter screen already showing). Live run
    # 20261103_071810 broke that assumption: the task_types_all_clicked
    # screenshot from that exact run shows 'Task Types' not just still
    # open but visually untouched (nothing checked), AND shows Windows'
    # own Quick Settings flyout (Wi-Fi/Bluetooth/battery panel) sitting on
    # top of the app at that same moment -- an OS-level popup unrelated to
    # Medent that can grab the foreground at any random point in a run.
    # pywinauto's click_input() only raises if the UIA element can't be
    # resolved; it does NOT raise just because an unrelated window ate the
    # physical click, so 'clicked_task_types_all' logged success above even
    # though nothing on screen changed. Confirm the window actually closed
    # before trusting that, and retry the click once -- this is almost
    # certainly a one-off focus-stealing event, not a wrong selector.
    def _task_types_closed():
        return not win.child_window(title="Task Types", control_type="Window").exists()

    closed = False
    for attempt in (1, 2):
        try:
            closed = wait_until_passes(3, 0.3, _task_types_closed)
            break
        except Exception:
            closed = False
            if attempt == 1:
                audit("task_types_not_closed_retrying")
                try:
                    wait_until_passes(5, 0.3, _click_task_types_all)
                    audit("clicked_task_types_all", attempt=2)
                    _capture_click_screenshot("task_types_all_clicked")
                except (PywinautoTimeoutError, ElementNotFoundError):
                    pass

    if not closed:
        _capture_diagnostics(
            "task_types_did_not_close", win,
            menu_errors=[
                "Clicked 'Task Types' -> 'All' twice, but the 'Task Types' "
                "window never closed. The first attempt's own screenshot "
                "showed Windows' Quick Settings flyout open over the app at "
                "the same moment -- check the task_types_all_clicked and "
                "task_types_did_not_close screenshots for the same or a "
                "similar OS-level popup stealing the click."
            ],
        )
        raise NavigationCheckpoint(
            "Clicked 'Task Types' -> 'All' twice, but the 'Task Types' "
            "window is still open and unchanged both times. This looks like "
            "an OS-level popup (Quick Settings, a notification toast, etc.) "
            "intercepting the click rather than a wrong selector -- inspect "
            "the task_types_did_not_close screenshot under OUTPUT_DIR for "
            "anything covering the screen, close it, and rerun. If nothing "
            "is covering the screen, the selector assumption itself needs "
            "re-checking."
        )

    # Confirmed live: clicking 'Type of Insurance' there (control_type=
    # "Pane", NOT "Text" as configure_report_filters's never-verified stub
    # assumed -- the task_types_window_not_found control-tree dump shows
    # every field on this screen, including this one, as a Pane) opens a
    # 'Type of Insurance Selection Criteria' popup listing 'Select Ins.
    # Co.', 'Clearinghouse', 'Ins. Co. Types', 'Except Ins. Co.', 'Groups'.
    #
    # The direct win.child_window(title="Select Ins. Co.", ...) attempt
    # below this comment (an earlier version here) timed out. Confirmed
    # via the top-level-window dump taken at that exact failure
    # (select_ins_co_not_found_toplevel_20261102_090152.txt): a NEW
    # top-level window ('Touch Popup', class 'Touch Popup Menu') appeared
    # -- this popup is NOT a Pane inside `win` at all, it's the SAME kind
    # of top-level flyout already seen for the Reports menu and Open Item
    # Reports submenu earlier in this function, which (per the Pass 5
    # module note) exposes no real row text via MSAA/UIA -- only OCR can
    # locate its rows. Reusing that exact machinery here instead of
    # guessing a new approach blind.
    def _click_type_of_insurance_and_poll():
        win.child_window(
            title="Type of Insurance", control_type="Pane",
        ).click_input()
        audit("opened_type_of_insurance")
        hwnds = _enumerate_menu_candidate_hwnds(main_hwnd)
        poll_trace, poll_jsonl_path, poll_txt_path = _poll_menu_signals(
            main_hwnd, "type_of_insurance_menu_poll", timeout=1.5, interval=0.02
        )
        audit(
            "type_of_insurance_menu_poll_complete",
            jsonl=str(poll_jsonl_path), txt=str(poll_txt_path), ticks=len(poll_trace),
        )
        return hwnds, poll_trace, poll_jsonl_path, poll_txt_path

    initial_ins_type_hwnds, trace, jsonl_path, txt_path = _click_type_of_insurance_and_poll()
    candidate = _best_populated_candidate(trace)

    if candidate is None:
        # Same focus-stealing possibility just confirmed for 'Task Types'
        # above -- retry the click once before treating this as a real
        # navigation problem rather than another one-off stolen click.
        audit("type_of_insurance_menu_not_found_retrying")
        initial_ins_type_hwnds, trace, jsonl_path, txt_path = _click_type_of_insurance_and_poll()
        candidate = _best_populated_candidate(trace)

    if candidate is None:
        all_hwnds = sorted({h["hwnd"] for snap in trace for h in snap["hwnds"] if "hwnd" in h})
        _capture_diagnostics(
            "type_of_insurance_menu_never_populated", win,
            menu_hwnds=all_hwnds or [h["hwnd"] for h in initial_ins_type_hwnds],
            menu_trace_paths=(jsonl_path, txt_path),
        )
        raise NavigationCheckpoint(
            "Clicked 'Type of Insurance' twice, but no candidate popup "
            f"window ({[h['hwnd'] for h in initial_ins_type_hwnds]!r} found "
            f"right after the click) ever reported a non-zero rect across "
            f"{len(trace)} poll ticks (~1.5s) on either attempt. Inspect "
            f"{txt_path} (human-readable) or {jsonl_path} (structured), and "
            "check for an OS-level popup (Quick Settings, a notification) "
            "covering the screen -- the same thing was just confirmed for "
            "'Task Types' earlier in this same run."
        )

    ins_type_menu_hwnd = candidate["hwnd"]
    audit(
        "type_of_insurance_menu_populated", hwnd=ins_type_menu_hwnd,
        class_name=candidate["class_name"], rect=candidate["rect"],
        tick=candidate["tick"], corroborated=candidate["corroborated"],
    )

    try:
        located = _locate_menu_row_center_via_ocr(ins_type_menu_hwnd, ("select", "ins", "co"))
    except Exception as exc:
        _capture_diagnostics(
            "select_ins_co_ocr_locate_failed", win, menu_hwnds=[ins_type_menu_hwnd],
            menu_errors=[f"ocr_locate: failed: {exc!r}"],
        )
        raise NavigationCheckpoint(
            "Clicked 'Type of Insurance' and found the selection-criteria "
            f"popup (hwnd {ins_type_menu_hwnd}), but OCR could not locate "
            f"a 'Select Ins. Co.' row in it: {exc!r}. Inspect the "
            "screenshots/control-tree dump under OUTPUT_DIR -- the token "
            "sequence ('select', 'ins', 'co') may not match how OCR "
            "actually splits this row's text."
        )

    pywinauto_mouse.click(coords=located["screen_point"])
    audit("clicked_select_ins_co", screen_point=located["screen_point"])
    _capture_click_screenshot("select_ins_co_clicked")

    # Per a live walkthrough, this opens an 'Insurance Company Selection'
    # grid (Id/Company/Icon/Insurance/T/Fee#/Street Address/Payer ID
    # columns) with 'Clear All' / 'All' / 'None' / 'Cancel' / 'Ok' buttons
    # along its bottom edge -- the same visual layout/button row as the
    # 'Task Types' picker (Clear All/All/Cancel/Ok, just with an extra
    # 'None' here), strongly suggesting it's built from the same embedded-
    # Window-with-Pane-buttons control confirmed for Task Types, rather
    # than the separate-top-level-window-with-no-real-row-text pattern
    # confirmed for the Reports menu / Type of Insurance Selection
    # Criteria popup. NOT yet confirmed live, unlike those two -- the
    # control-tree dump taken on failure here will say for certain.
    def _click_ins_co_all():
        win.child_window(
            title="Insurance Company Selection", control_type="Window",
        ).child_window(
            title="All", control_type="Pane",
        ).click_input()
        return True

    try:
        wait_until_passes(5, 0.3, _click_ins_co_all)
    except (PywinautoTimeoutError, ElementNotFoundError) as exc:
        _capture_diagnostics("ins_co_all_not_found", win)
        raise NavigationCheckpoint(
            "Clicked 'Select Ins. Co.', but the 'Insurance Company "
            "Selection' grid's 'All' button could not be resolved-and-"
            f"clicked within 5s: {exc!r}. Inspect the screenshots/control-"
            "tree dump/top-level-window dump under OUTPUT_DIR -- this "
            "screen might instead be a separate top-level window (like "
            "the Reports menu/Type of Insurance popup) rather than an "
            "embedded Window inside `win` (like Task Types), or 'All' "
            "might not be control_type=\"Pane\"."
        )

    audit("clicked_ins_co_all")
    _capture_click_screenshot("ins_co_all_clicked")

    # Per a live walkthrough, 'All' alone closes 'Insurance Company
    # Selection' too (same auto-close behavior already confirmed for
    # 'Task Types') and returns to the report filter screen, which then
    # immediately shows a 'Free Text Company Format' popup: 3 numbered
    # options plus an 'Enter Free Text Company Format' input pre-filled
    # with '1' (already highlighted/selected) -- no clickable Ok/button in
    # the UI at all, same keyboard-only pattern already confirmed for the
    # 'AR Task Options' popup in confirm_report_popup. A single ENTER
    # accepts the pre-filled default and dismisses it.
    send_keys("{ENTER}")
    audit("dismissed_free_text_company_format")
    _capture_click_screenshot("free_text_company_format_dismissed")

    # 'Select by Date/Days' opens a 'Date/Days Selection Criteria' popup
    # (same floating-card visual style as 'Type of Insurance Selection
    # Criteria' -- 'Date Range' / 'Number of Days' options, no further
    # input fields) -- per a live walkthrough, same 'Touch Popup Menu'
    # OCR-only pattern, reusing the exact same approach as 'Select Ins.
    # Co.' above rather than guessing a UIA-based one again.
    win.child_window(
        title="Select by Date/Days", control_type="Pane",
    ).click_input()
    audit("opened_select_by_date_days")

    initial_date_days_hwnds = _enumerate_menu_candidate_hwnds(main_hwnd)
    trace, jsonl_path, txt_path = _poll_menu_signals(
        main_hwnd, "date_days_menu_poll", timeout=1.5, interval=0.02
    )
    audit(
        "date_days_menu_poll_complete",
        jsonl=str(jsonl_path), txt=str(txt_path), ticks=len(trace),
    )

    candidate = _best_populated_candidate(trace)
    if candidate is None:
        all_hwnds = sorted({h["hwnd"] for snap in trace for h in snap["hwnds"] if "hwnd" in h})
        _capture_diagnostics(
            "date_days_menu_never_populated", win,
            menu_hwnds=all_hwnds or [h["hwnd"] for h in initial_date_days_hwnds],
            menu_trace_paths=(jsonl_path, txt_path),
        )
        raise NavigationCheckpoint(
            "Clicked 'Select by Date/Days', but no candidate popup window "
            f"({[h['hwnd'] for h in initial_date_days_hwnds]!r} found right "
            f"after the click) ever reported a non-zero rect across "
            f"{len(trace)} poll ticks (~1.5s). Inspect "
            f"{txt_path} (human-readable) or {jsonl_path} (structured)."
        )

    date_days_menu_hwnd = candidate["hwnd"]
    audit(
        "date_days_menu_populated", hwnd=date_days_menu_hwnd,
        class_name=candidate["class_name"], rect=candidate["rect"],
        tick=candidate["tick"], corroborated=candidate["corroborated"],
    )

    try:
        located = _locate_menu_row_center_via_ocr(date_days_menu_hwnd, ("date", "range"))
    except Exception as exc:
        _capture_diagnostics(
            "date_range_ocr_locate_failed", win, menu_hwnds=[date_days_menu_hwnd],
            menu_errors=[f"ocr_locate: failed: {exc!r}"],
        )
        raise NavigationCheckpoint(
            "Clicked 'Select by Date/Days' and found the selection-"
            f"criteria popup (hwnd {date_days_menu_hwnd}), but OCR could "
            f"not locate a 'Date Range' row in it: {exc!r}. Inspect the "
            "screenshots/control-tree dump under OUTPUT_DIR -- the token "
            "sequence ('date', 'range') may not match how OCR actually "
            "splits this row's text."
        )

    pywinauto_mouse.click(coords=located["screen_point"])
    audit("clicked_date_range", screen_point=located["screen_point"])
    _capture_click_screenshot("date_range_clicked")

    # Per the user: this screen's Date Range 'from' field is a FIXED
    # cutoff (09/01/2026), deliberately independent of this script's own
    # --date-from/--date-to CLI arguments (confirmed explicitly -- NOT an
    # oversight). The 'to' field already shows today's date with zero
    # interaction needed (same zero-interaction assumption the existing,
    # never-verified set_date_range() makes for its own 'to' side).
    #
    # The earlier direct-text-entry guess here (title="Date Range",
    # control_type="Text") failed with ElementNotFoundError. The
    # control-tree dump taken at that exact failure
    # (date_range_from_not_set_controltree_20261102_103011.txt, lines
    # 85-103) shows the real structure: 'Date Range' is a plain label
    # Pane (not an input), the 'from' VALUE is a separate Pane titled
    # '00/00/00' (with its own 'DOWN' dropdown-arrow child Pane), and the
    # 'to' value is likewise a Pane titled '11/02/26' (today, already
    # correct, matching the zero-interaction assumption). Clicking the
    # '00/00/00' Pane is the confirmed next action.
    FIXED_DATE_RANGE_FROM = datetime(2023, 9, 1)

    # Per the control-tree dump (date_range_from_not_set_controltree_
    # 20261102_103011.txt, lines 89-95), '00/00/00' has its own child Pane
    # titled 'DOWN' at (607,240)-(623,264) -- the dropdown-arrow icon on
    # the field's right edge, distinct from the plain-text area spanning
    # the field's full width (484-625). Clicking the parent '00/00/00'
    # Pane directly (an earlier version here) clicks ITS CENTER (~x=554),
    # landing on the text area rather than the arrow at ~x=615 -- likely
    # why the calendar didn't reliably/consistently open. 'DOWN' is a
    # genuine child per the dump's own nesting (unlike the Yes/No case in
    # confirm_report_popup, which turned out to be siblings), so clicking
    # it directly is safe.
    win.child_window(
        title="00/00/00", control_type="Pane",
    ).child_window(
        title="DOWN", control_type="Pane",
    ).click_input()
    audit("opened_date_range_from_calendar")

    # The year opens on the current year (2026); the user explicitly
    # confirmed this fixed cutoff needs year 2023, requiring 3 clicks on
    # the '<' back-arrow. That arrow is a plain glyph icon with zero
    # accessible content (confirmed earlier), so OCR -- which only reads
    # real text -- cannot locate it directly the way 'Sep' below is
    # located. The user also confirmed manually that typing a date
    # directly into the '00/00/00' field does NOT work, so this calendar
    # is the only path. Rather than guess a fixed screen coordinate for
    # the arrow with no live measurement of it, this anchors off the
    # OCR-readable year text itself (a bare 4-digit number, confirmed
    # unique in this window -- every other date on screen is MM/DD/YY,
    # never a lone 4-digit token) and estimates the arrow's position at a
    # margin to its left -- an unverified guess, same spirit as
    # ARROW_MARGIN_PX elsewhere in this module. It self-corrects: the
    # year is re-read via fresh OCR after every click, and this only
    # proceeds once it has actually decremented by exactly 1, rather than
    # firing 3 blind clicks and hoping all 3 landed.
    #
    # A first live attempt at 150px missed the real arrow -- the
    # diagnostic screenshot from that failure
    # (date_range_year_back_arrow_clicked_screenshot_20261103_012319.png)
    # shows the click (x=425) landed about 50px to the left of the '<'
    # glyph's real center (~x=475), in empty space, which also reset the
    # calendar back to today's month as a side effect. Corrected to 105px
    # based on that measurement.
    CALENDAR_YEAR_ARROW_LEFT_MARGIN_PX = 105

    def _normalize_ocr_text(text):
        return text.strip(" \t\r\n.,:;\"'“”‘’").lower()

    def _locate_calendar_year():
        image, window_rect = _capture_window_bitmap(main_hwnd)
        words = _run_ocr(image)
        year_words = [
            w for w in words
            if _normalize_ocr_text(w["text"]).isdigit() and len(_normalize_ocr_text(w["text"])) == 4
        ]
        if len(year_words) != 1:
            raise RuntimeError(
                f"found {len(year_words)} bare 4-digit word(s) in the main "
                f"window, not exactly 1 (expected only the calendar's year "
                f"header to match): {year_words!r}"
            )
        return int(_normalize_ocr_text(year_words[0]["text"])), year_words[0]["rect"], window_rect

    def _locate_calendar_year_retried(diag_label, context_msg):
        try:
            return wait_until_passes(3, 0.3, _locate_calendar_year)
        except Exception as exc:
            underlying = getattr(exc, "original_exception", None)
            _capture_diagnostics(
                diag_label, win,
                menu_errors=[f"ocr_locate: failed after retries: {exc!r}; original_exception={underlying!r}"],
            )
            raise NavigationCheckpoint(
                f"{context_msg}: {exc!r} (original_exception={underlying!r}). "
                "Inspect the screenshot under OUTPUT_DIR."
            )

    target_year = FIXED_DATE_RANGE_FROM.year
    for _ in range(3):
        current_year, year_rect, year_window_rect = _locate_calendar_year_retried(
            "date_range_year_ocr_locate_failed",
            "Opened the Date Range 'from' calendar, but OCR could not "
            "locate a bare 4-digit year anywhere in the main window",
        )
        if current_year == target_year:
            break

        arrow_x = year_rect[0] - CALENDAR_YEAR_ARROW_LEFT_MARGIN_PX
        arrow_y = year_rect[1] + year_rect[3] // 2
        screen_point = (year_window_rect[0] + arrow_x, year_window_rect[1] + arrow_y)
        pywinauto_mouse.click(coords=screen_point)
        audit("clicked_date_range_year_back_arrow", from_year=current_year, screen_point=screen_point)
        _capture_click_screenshot("date_range_year_back_arrow_clicked")

        def _year_decremented():
            new_year, _, _ = _locate_calendar_year()
            if new_year != current_year - 1:
                raise RuntimeError(f"year after click is {new_year}, expected {current_year - 1}")
            return True

        try:
            wait_until_passes(3, 0.3, _year_decremented)
        except Exception as exc:
            underlying = getattr(exc, "original_exception", None)
            _capture_diagnostics(
                "date_range_year_arrow_click_unconfirmed", win,
                menu_errors=[f"{exc!r}; original_exception={underlying!r}"],
            )
            raise NavigationCheckpoint(
                "Clicked an estimated left-arrow position (margin="
                f"{CALENDAR_YEAR_ARROW_LEFT_MARGIN_PX}px left of the year "
                f"text, which read {current_year}) to go back a year, but "
                f"the year did not decrement to {current_year - 1} within "
                f"3s: {exc!r} (original_exception={underlying!r}). Inspect "
                "the date_range_year_back_arrow_clicked screenshot under "
                "OUTPUT_DIR to see where the click actually landed vs "
                "where the real left-arrow icon is, and adjust "
                "CALENDAR_YEAR_ARROW_LEFT_MARGIN_PX accordingly."
            )

    final_year, _, _ = _locate_calendar_year_retried(
        "date_range_year_final_check_failed",
        "Finished the year back-arrow loop, but a final OCR check of the year failed",
    )
    if final_year != target_year:
        _capture_diagnostics("date_range_year_not_reached", win)
        raise NavigationCheckpoint(
            f"Clicked the year back-arrow but ended at {final_year}, not "
            f"{target_year}, after 3 attempts. Inspect the screenshots "
            "under OUTPUT_DIR and the clicked_date_range_year_back_arrow "
            "events in the audit log."
        )
    audit("date_range_year_set", year=final_year)

    # Confirmed via a control-tree dump taken right after opening this
    # overlay (date_range_from_calendar_opened_controltree_20261102_
    # 114052.txt): it exposes NO accessible content at all -- no '2026',
    # 'Jan'..'Dec', 'Today', or 'Exit' anywhere in win's own tree -- and no
    # new top-level window appeared either (confirmed via the matching
    # top-level-window dump), so it isn't a separate 'Touch Popup Menu'
    # like the Reports menu/Type of Insurance popup. It's drawn directly
    # onto `win` itself with zero accessible children -- same "OCR is the
    # only way to find it" situation as those popups, just painted onto
    # the main window's own surface instead of a child popup's. The
    # existing OCR helpers (_capture_window_bitmap/_run_ocr) are generic
    # over any hwnd, not specific to popup menus, so they work unchanged
    # against `main_hwnd` itself here.
    #
    # Only the month click is done here, not the day. A single-token OCR
    # match for a bare day number like '1' is a real ambiguity risk
    # (nothing yet confirms it's the only '1' rendered anywhere in the
    # window), so rather than chain that guess on top of an as-yet-
    # unconfirmed month click, this stops right after Sep to verify via a
    # fresh screenshot that the day grid actually updated to September
    # before attempting to OCR-match a lone digit.
    # The first live attempt here ran _locate_menu_row_center_via_ocr with
    # NO wait at all right after the click, and OCR found only the
    # background screen's own text (confirmed via the diagnostic
    # screenshot taken ~0.8s later via the SAME ImageGrab mechanism used
    # for every other screenshot in this module -- it also showed no
    # calendar at all, just the empty 'Select Number of Days' screen).
    # Every OTHER OCR-popup step in this function (Type of Insurance,
    # Date/Days) waited via _poll_menu_signals for the popup to actually
    # stabilize before running OCR on it; this one skipped that because
    # the calendar isn't a separate 'Touch Popup Menu' window
    # _poll_menu_signals knows how to wait on. Rather than conclude the
    # calendar vanished, retry the OCR locate itself via wait_until_passes
    # (the same retry discipline used for every click elsewhere in this
    # module) to cover plain render/input lag before giving up.
    def _locate_sep():
        return _locate_menu_row_center_via_ocr(main_hwnd, ("sep",))

    try:
        located = wait_until_passes(3, 0.3, _locate_sep)
    except Exception as exc:
        underlying = getattr(exc, "original_exception", None)
        _capture_diagnostics(
            "date_range_month_ocr_locate_failed", win,
            menu_errors=[f"ocr_locate: failed after retries: {exc!r}; original_exception={underlying!r}"],
        )
        raise NavigationCheckpoint(
            "Opened the Date Range 'from' calendar overlay, but OCR still "
            f"could not locate 'Sep' anywhere in the main window after "
            f"retrying for 3s: {exc!r} (original_exception={underlying!r}). "
            "Inspect the screenshot under OUTPUT_DIR -- if the calendar "
            "still isn't visible in it, the overlay may be closing itself "
            "on some input event click_input()/pywinauto_mouse.click() "
            "triggers (e.g. a focus change), not merely a render-timing "
            "issue; if it IS visible, OCR may be splitting 'Sep' "
            "differently than expected."
        )

    pywinauto_mouse.click(coords=located["screen_point"])
    audit("clicked_date_range_month", month="Sep", screen_point=located["screen_point"])
    _capture_click_screenshot("date_range_month_clicked")

    # Confirmed live: the day grid now shows September 2026. A bare day
    # number like '1' is a real ambiguity risk -- an earlier OCR dump in
    # this same window showed the engine mis-splitting a date like
    # '11/02/26' (from the window's own title bar, near the TOP of the
    # captured bitmap) into separate tokens, one of which was a lone '1'.
    # Rather than reuse _locate_menu_row_via_ocr (which raises outright on
    # >1 match, with no way to pick between them), this reads the raw OCR
    # words directly and disambiguates by position: the title bar is
    # always above the calendar content, so among every '1' match, the
    # one(s) BELOW the already-located 'Sep' row are the real candidates.
    # Failing loud (a clear diagnostic dump) if that doesn't narrow to
    # exactly one, rather than guessing.
    #
    # Reuses `located["row_rect"]` from the 'Sep' click above instead of
    # a second, separate _locate_menu_row_via_ocr call per retry -- a live
    # run with that second call found ZERO '1' matches across a full 3s
    # retry window, every single attempt, which is more consistent with
    # OCR being overloaded by two full capture+recognize passes every
    # 0.3s (up to ~20 calls in 3s) than with a one-off flaky read. Cutting
    # the OCR calls per attempt in half and spacing retries out further
    # (1s instead of 0.3s) gives the recognizer room to actually finish.
    sep_row_top = located["row_rect"][1]

    # Two live attempts trying to OCR-match the exact digit '1' (plus a
    # merged-digit fallback for e.g. '12345') found ZERO candidates every
    # time. The SECOND attempt's full word dump below the 'Sep' row
    # (date_range_day_ocr_locate_failed, 2026-11-03 01:40) proved this
    # isn't specific to '1' at all: that same capture also dropped day 3,
    # 13, and 26, AND the 'Th'/'Fr' column headers, while recognizing
    # every other day number fine (2,4-12,14-25,27-30). This is the
    # Windows OCR engine randomly failing to detect a handful of words
    # per capture (~15% miss rate observed), not a per-glyph or per-color
    # issue -- so no amount of retrying a literal '1' match is reliable.
    #
    # Fix: never depend on OCR actually reading the target day's own
    # digit. Day 1 always sits at grid row 0, in whichever weekday column
    # September 1st falls on -- both computable from plain date math.
    # Build a grid model (column x-centers, row y-centers) from
    # WHATEVER OTHER day numbers happen to be recognized this capture
    # (virtually always includes several in both day 1's column and its
    # row, even when day 1 itself is dropped), then read off day 1's
    # pixel position from that model instead of searching for '1' at all.
    def _ui_col_for_day(day):
        # UI column order is Su=0, Mo=1, ..., Sa=6; Python's weekday() is
        # Mon=0..Sun=6, so shift by 1 to realign Sunday to column 0.
        py_weekday = datetime(FIXED_DATE_RANGE_FROM.year, FIXED_DATE_RANGE_FROM.month, day).weekday()
        return (py_weekday + 1) % 7

    def _ui_row_for_day(day):
        first_col = _ui_col_for_day(1)
        return (day - 1 + first_col) // 7

    def _interpolate(known, target_index):
        # known: {index: [coord, ...]} -- averages duplicates, then
        # linearly extrapolates/interpolates to target_index using the
        # average spacing between consecutive known indices. Works even
        # if target_index itself has no data (e.g. day 1's own row/column
        # happened to have zero recognized neighbors this capture).
        points = sorted((idx, sum(vals) / len(vals)) for idx, vals in known.items())
        if target_index in known:
            return dict(points)[target_index]
        if len(points) < 2:
            raise RuntimeError(f"not enough reference points to locate index {target_index}: {points!r}")
        spacings = [
            (points[i + 1][1] - points[i][1]) / (points[i + 1][0] - points[i][0])
            for i in range(len(points) - 1)
        ]
        avg_spacing = sum(spacings) / len(spacings)
        nearest_index, nearest_coord = min(points, key=lambda p: abs(p[0] - target_index))
        return nearest_coord + avg_spacing * (target_index - nearest_index)

    def _locate_day():
        image, window_rect = _capture_window_bitmap(main_hwnd)
        words = _run_ocr(image)
        below_words = [w for w in words if w["rect"][1] > sep_row_top]

        day_rects = {}
        for w in below_words:
            t = _normalize_ocr_text(w["text"])
            if t.isdigit() and 1 <= int(t) <= 31:
                day_rects[int(t)] = w["rect"]

        if len(day_rects) < 2:
            raise RuntimeError(
                f"found only {len(day_rects)} day-number word(s) below the "
                f"'Sep' row -- not enough to build a grid model: {day_rects!r}. "
                f"All words below that row: {below_words!r}"
            )

        col_xs, row_ys = {}, {}
        for day, rect in day_rects.items():
            col_xs.setdefault(_ui_col_for_day(day), []).append(rect[0] + rect[2] / 2)
            row_ys.setdefault(_ui_row_for_day(day), []).append(rect[1] + rect[3] / 2)

        target_day = FIXED_DATE_RANGE_FROM.day
        click_x = _interpolate(col_xs, _ui_col_for_day(target_day))
        click_y = _interpolate(row_ys, _ui_row_for_day(target_day))

        return {
            "screen_point": (window_rect[0] + int(click_x), window_rect[1] + int(click_y)),
            "day_rects_seen": sorted(day_rects),
        }

    try:
        day_located = wait_until_passes(6, 1.0, _locate_day)
    except Exception as exc:
        underlying = getattr(exc, "original_exception", None)
        _capture_diagnostics(
            "date_range_day_ocr_locate_failed", win,
            menu_errors=[f"ocr_locate: failed after retries: {exc!r}; original_exception={underlying!r}"],
        )
        raise NavigationCheckpoint(
            "Clicked 'Sep' in the Date Range 'from' calendar, but could not "
            f"unambiguously locate day {FIXED_DATE_RANGE_FROM.day} below it: "
            f"{exc!r} (original_exception={underlying!r}). Inspect the "
            "screenshot under OUTPUT_DIR and the candidate dump in the "
            "error message above."
        )

    pywinauto_mouse.click(coords=day_located["screen_point"])
    audit(
        "clicked_date_range_day", day=FIXED_DATE_RANGE_FROM.day,
        screen_point=day_located["screen_point"],
        day_rects_seen=day_located["day_rects_seen"],
    )
    _capture_click_screenshot("date_range_day_clicked")

    # Confirmed live: the 'from' field now shows 09/01/23. The top-left
    # checkmark confirming this whole 'Select Number of Days' screen is
    # titled 'Done' (control_type="Pane", same touch-UI pattern as
    # everywhere else), per the control-tree dump captured earlier for
    # this same screen (date_range_from_not_set_controltree_20261102_
    # 103011.txt, lines 49-51) -- a genuine direct child of `win`, with
    # 'Cancel' (the 'X') and 'Reset' as siblings, not nested.
    win.child_window(title="Done", control_type="Pane").click_input()
    audit("clicked_date_range_done")
    _capture_click_screenshot("date_range_done_clicked")

    # Confirmed live: 'Done' returns to the 'Charges For One or More Ins.
    # Co' report filter screen, with 'Date Range' now correctly showing
    # '09/01/23-11/03/26'. 'Data Export File' is a confirmed Pane
    # directly under `win` (per task_types_window_not_found_controltree_
    # 20261102_085533.txt / date_range_from_not_set_controltree_
    # 20261102_103011.txt, both showing it at auto_id="22"). Clicking it
    # opens a '(Filter list by 1st letter)' popup listing 'Back', a
    # 'Search' field, 'None', 'Ascii', 'XML' -- per a live walkthrough,
    # a SEPARATE 'Touch Popup Menu' window -- a live run proved otherwise.
    # The control-tree dump captured at that failure
    # (data_export_file_menu_never_populated_controltree_20261103_023345.txt,
    # lines 7-13) shows '     (Filter list by 1st letter)       ' is an
    # embedded child Window/Dialog nested INSIDE `win` -- the same
    # pattern as 'Task Types' and the Date Range calendar overlay, not a
    # top-level 'Touch Popup Menu' -- so _enumerate_menu_candidate_hwnds/
    # _poll_menu_signals (which only look for that class/pattern) could
    # never find it. Its list rows ('None'/'Ascii'/'XML') have zero
    # accessible content in that same dump, so OCR is still required --
    # just run directly against `main_hwnd`, the exact fix already proven
    # for the calendar overlay, rather than through the popup-candidate
    # machinery.
    #
    # Two earlier automated attempts (clicking the 'Data Export File'
    # LABEL, then clicking its 'None   ' VALUE box) both failed to open
    # the '(Filter list by 1st letter)' popup. User confirmed via an
    # annotated live screenshot (arrow pointing at the 'Data Export File'
    # label text) that the label is the correct click target, so retrying
    # that -- the earlier label-click failure may have been a timing/stale-
    # coordinate issue right after the Date Range 'Done' click, not a
    # wrong-target issue.
    win.child_window(title="Data Export File", auto_id="22", control_type="Pane").click_input()
    audit("opened_data_export_file")

    # Two consecutive live runs (20261103_041407 and 20261103_041854) both
    # proved the click itself works and the popup DOES open -- the
    # diagnostic screenshot taken right after OCR gave up clearly shows
    # '(Filter list by 1st letter)' open with Back/Search/None/Ascii/XML.
    # Widening the OCR retry budget from 3s to 10s made NO difference --
    # byte-identical "popup never opened" word lists both times, proving
    # this was never a timing problem. The control-tree dump from the
    # second failure showed why: this dialog is a genuine child
    # Window/Dialog with its own HWND (ComponentContainerWindow/
    # ComponentItemWindow children -- a custom, likely GPU-composited
    # control), and PrintWindow on the TOP-LEVEL main window's hwnd simply
    # does not include it. Fix: resolve the dialog's own hwnd via UIA and
    # OCR THAT, not `main_hwnd`.
    def _locate_filter_dialog_hwnd():
        return _resolve_child_dialog_hwnd(win, r"\s*\(Filter list by 1st letter\)\s*")

    # Live run 20261103_201553 proved 10s isn't always enough here: the
    # diagnostic screenshot taken at that timeout clearly shows
    # '(Filter list by 1st letter)' already open and fully rendered, so
    # Medent was just slower than 10s to finish instantiating it this run
    # -- same "real dialog on screen, gave up too early" pattern already
    # fixed for the Print-to-PDF Save dialog, not a new capture-target bug.
    FILTER_DIALOG_TIMEOUT_S = 30

    try:
        filter_dialog_hwnd = wait_until_passes(FILTER_DIALOG_TIMEOUT_S, 0.5, _locate_filter_dialog_hwnd)
    except Exception as exc:
        underlying = getattr(exc, "original_exception", None)
        _capture_diagnostics(
            "filter_dialog_not_found", win,
            menu_errors=[f"resolve_dialog_hwnd: failed after retries: {exc!r}; original_exception={underlying!r}"],
        )
        raise NavigationCheckpoint(
            "Clicked 'Data Export File', but could not resolve the "
            f"'(Filter list by 1st letter)' dialog's own window handle "
            f"after retrying for {FILTER_DIALOG_TIMEOUT_S}s: {exc!r} (original_exception={underlying!r}). "
            "Inspect the screenshot/control-tree dump under OUTPUT_DIR."
        )

    def _locate_xml():
        return _locate_menu_row_center_via_ocr(filter_dialog_hwnd, ("xml",))

    def _xml_option_dialog_opened():
        # Confirmed live (recorder.py, two full runs, 2026-11-03): clicking
        # 'XML' closes '(Filter list by 1st letter)' and immediately opens
        # 'Select XML Option' -- it does NOT update the main screen's 'Data
        # Export File' field, which only changes once 'Select XML Option'
        # is itself later confirmed with OK. An earlier version of this
        # check looked at that field instead (auto_id=28 showing "xml"),
        # which is checking the wrong signal at the wrong time -- it
        # reported "click had no effect" on every single run since,
        # including fully successful ones where 'Select XML Option' was
        # already open on screen (confirmed by a live screenshot) at the
        # exact moment this check failed. Check for the dialog that
        # actually appears instead.
        try:
            _resolve_child_dialog_hwnd(win, r"\s*Select XML Option\s*")
            return True
        except Exception:
            return False

    # Runs 20261103_063339 and 20261103_065217 both clicked this exact same
    # OCR-computed screen_point (584, 600) and it worked -- 'Select XML
    # Option' opened within ~7s both times. Run 20261103_065642 sent the
    # identical click at the identical point and it silently did nothing:
    # 11s later (xml_option_dialog_not_found_controltree_20261103_065653.txt)
    # the '(Filter list by 1st letter)' popup was still open. Same
    # coordinates, same mechanism, intermittent miss -- not a wrong-target
    # or wrong-click-type problem. Confirm the dialog actually opened, and
    # retry the whole locate+click once -- same reasoning already used
    # above for the 'Data Export File' label click -- before giving up for
    # real.
    located = None
    confirmed = False
    for attempt in (1, 2):
        try:
            located = wait_until_passes(10, 0.5, _locate_xml)
        except Exception as exc:
            underlying = getattr(exc, "original_exception", None)
            _capture_diagnostics(
                "export_xml_ocr_locate_failed", win,
                menu_errors=[f"ocr_locate: failed after retries: {exc!r}; original_exception={underlying!r}"],
            )
            raise NavigationCheckpoint(
                "Clicked 'Data Export File', but OCR could not locate an "
                f"'XML' row in the filter dialog's own window after retrying "
                f"for 10s: {exc!r} (original_exception={underlying!r}). Inspect "
                "the screenshot under OUTPUT_DIR."
            )

        pywinauto_mouse.click(coords=located["screen_point"])
        audit("clicked_export_xml", screen_point=located["screen_point"], attempt=attempt)
        _capture_click_screenshot("export_xml_clicked")

        try:
            confirmed = wait_until_passes(5, 0.5, _xml_option_dialog_opened)
            break
        except Exception:
            confirmed = False
            audit("export_xml_click_unconfirmed", attempt=attempt, screen_point=located["screen_point"])

    if not confirmed:
        _capture_diagnostics(
            "export_xml_click_had_no_effect", win,
            menu_errors=[
                "Clicked the OCR-located 'XML' row twice (last screen_point="
                f"{located['screen_point']!r}), but 'Select XML Option' "
                "never opened."
            ],
        )
        raise NavigationCheckpoint(
            "Clicked 'XML' in the '(Filter list by 1st letter)' popup twice "
            "via the same OCR-computed coordinates that worked in two prior "
            "live runs, but 'Select XML Option' never opened. Inspect the "
            "export_xml_click_had_no_effect diagnostics before changing "
            "anything -- this may need a third retry, a short pre-click "
            "delay, or (last resort) a different input method for this one "
            "dialog."
        )

    audit("export_xml_confirmed")

    # Live run 20261103_042655 confirmed selecting 'XML' opens a 'Select
    # XML Option' popup: '1. Open with Microsoft Excel' / '2. Transfer to
    # my PC encrypted' / '3. Transfer to my PC unencrypted', plus 'OK' and
    # 'Cancel' buttons -- and it hit the exact same bug as 'Data Export
    # File's own popup: OCR against `main_hwnd` returned the byte-identical
    # "popup never opened" word list, even though the screenshot taken
    # right after shows the popup clearly open. Same root cause (a real
    # child Window/Dialog with its own hwnd that PrintWindow on the
    # top-level window doesn't capture) -- same fix: resolve this dialog's
    # own hwnd via UIA and OCR that directly, same as
    # '(Filter list by 1st letter)' above.
    def _locate_xml_option_dialog_hwnd():
        return _resolve_child_dialog_hwnd(win, r"\s*Select XML Option\s*")

    try:
        xml_option_dialog_hwnd = wait_until_passes(10, 0.5, _locate_xml_option_dialog_hwnd)
    except Exception as exc:
        underlying = getattr(exc, "original_exception", None)
        _capture_diagnostics(
            "xml_option_dialog_not_found", win,
            menu_errors=[f"resolve_dialog_hwnd: failed after retries: {exc!r}; original_exception={underlying!r}"],
        )
        raise NavigationCheckpoint(
            "Clicked 'XML', but could not resolve the 'Select XML Option' "
            f"dialog's own window handle after retrying for 10s: {exc!r} "
            f"(original_exception={underlying!r}). Inspect the "
            "screenshot/control-tree dump under OUTPUT_DIR."
        )

    # 'OK'/'Cancel' are rendered with a visible bordered-button look,
    # distinct from this app's usual flat touch-UI style, which may mean
    # they're real UIA Buttons here -- but that's unconfirmed, so this
    # stops right after the option click rather than guessing how to
    # resolve 'OK' too.
    def _locate_excel_option():
        return _locate_menu_row_center_via_ocr(xml_option_dialog_hwnd, ("1", "open", "with", "microsoft", "excel"))

    try:
        located = wait_until_passes(10, 0.5, _locate_excel_option)
    except Exception as exc:
        underlying = getattr(exc, "original_exception", None)
        _capture_diagnostics(
            "xml_option_excel_ocr_locate_failed", win,
            menu_errors=[f"ocr_locate: failed after retries: {exc!r}; original_exception={underlying!r}"],
        )
        raise NavigationCheckpoint(
            "Clicked 'XML', but OCR could not locate '1. Open with "
            f"Microsoft Excel' in the 'Select XML Option' dialog's own "
            f"window after retrying for 10s: {exc!r} "
            f"(original_exception={underlying!r}). Inspect the "
            "screenshot under OUTPUT_DIR -- the token sequence may not "
            "match how OCR actually splits this row's text (e.g. '1.' "
            "and 'Open' might not separate the way assumed)."
        )

    pywinauto_mouse.click(coords=located["screen_point"])
    audit("clicked_xml_option_excel", screen_point=located["screen_point"])
    _capture_click_screenshot("xml_option_excel_clicked")

    # Live run 20261103_142505 corrected an earlier, wrong assumption here:
    # there is NO separate 'OK' click for this dialog. By the time that
    # version of this code tried to click 'OK', xml_option_dialog_hwnd was
    # already an invalid handle ('Invalid window handle.' from both the UIA
    # and OCR attempts) and the top-level window title had already reverted
    # from 'Reports COMPASSIONATE WOMEN'S...' back to the plain filter-
    # screen title -- i.e. clicking '1. Open with Microsoft Excel' itself
    # already selected AND closed 'Select XML Option' in one action, same
    # one-click-does-everything pattern already confirmed for 'Task Types'
    # -> 'All' and 'Insurance Company Selection' -> 'All'. Just confirm the
    # dialog is gone (it already closed itself) before moving on.
    def _xml_option_dialog_closed():
        try:
            _resolve_child_dialog_hwnd(win, r"\s*Select XML Option\s*")
            return False
        except Exception:
            return True

    try:
        wait_until_passes(5, 0.3, _xml_option_dialog_closed)
    except Exception as exc:
        _capture_diagnostics("xml_option_dialog_did_not_close", win)
        raise NavigationCheckpoint(
            "Clicked '1. Open with Microsoft Excel' in 'Select XML "
            f"Option', but the dialog never closed on its own: {exc!r}. "
            "Inspect diagnostics under OUTPUT_DIR -- this may be the same "
            "OS-popup-stealing-the-click pattern already confirmed for "
            "'Task Types' and 'Type of Insurance' earlier in this run."
        )

    audit("xml_option_confirmed")

    # Confirmed live across TWO independent full runs (2026-11-03, captured
    # with recorder.py -- a separate, hands-off click/window observer built
    # specifically to settle this without guessing): the SAME 'Done'
    # checkmark already used above to confirm the date-range calendar
    # overlay is reused here to confirm the whole export-settings screen.
    # There is no separate 'Export' toolbar click in the real flow, despite
    # 'Export' being visible as a text label in the same toolbar -- both
    # runs went straight from 'Done' to the next prompt. Both runs also
    # confirmed that prompt -- "the detailed version of this report
    # requires wide '14 7/8 x 11 in.' paper -- are you sure?" -- is an
    # embedded Pane in `win`, not a separate top-level window, same pattern
    # as every other Yes/No prompt in this app (see confirm_report_popup's
    # docstring).
    win.child_window(title="Done", control_type="Pane").click_input()
    audit("clicked_done_export_settings")
    _capture_click_screenshot("done_export_settings_clicked")

    win.child_window(title="Yes", control_type="Pane").click_input()
    audit("confirmed_print_warning")
    _capture_click_screenshot("print_warning_confirmed")

    # Confirmed live: a real native Windows "Print" common dialog (class
    # '#32770') opens next, with a real UIA Button named 'OK' -- unlike
    # almost everything else in this app, this part is a genuine Win32
    # dialog, not a touch-UI Pane.
    print_dlg = Desktop(backend="uia").window(title="Print")
    try:
        wait_until_passes(
            10, 0.3,
            lambda: print_dlg.child_window(title="OK", control_type="Button").exists()
        )
    except Exception as exc:
        _capture_diagnostics("print_dialog_not_found", win)
        raise NavigationCheckpoint(
            "Clicked 'Done' then 'Yes', but the native 'Print' dialog's "
            f"'OK' button never appeared within 10s: {exc!r}. Inspect "
            "diagnostics under OUTPUT_DIR."
        )
    print_dlg.child_window(title="OK", control_type="Button").click_input()
    audit("print_dialog_ok")
    _capture_click_screenshot("print_dialog_ok_clicked")

    # Confirmed live: since the selected printer is "Microsoft Print to
    # PDF", clicking 'OK' immediately opens a real native "Save Print
    # Output As" common dialog -- this produces a PDF rendering of the
    # report (a side effect of however Medent's "run report" action works,
    # not something this script chooses).
    #
    # Live diagnostic (2026-11-03, diagnose_save_dialog.py) proved this
    # dialog cannot be resolved as a window object AT ALL from this script:
    # it is absent from both a raw win32gui.EnumWindows() dump and
    # Desktop(backend="uia").windows(), even though ImageGrab/OCR (which
    # reads the physical screen, not a specific hwnd) sees it fine and
    # fully rendered. This is a different, deeper problem than the earlier
    # '(Filter list by 1st letter)' bug (which WAS resolvable via UIA, just
    # needed OCR against its own hwnd instead of main_hwnd) -- here UIA
    # cannot find the window at all, so no amount of retrying
    # win.child_window(...)/Desktop(...).window(...) can ever succeed.
    # Likely cause: this dialog belongs to a separate process (printing
    # goes through splwow64.exe, confirmed running in that diagnostic) that
    # this script's automation layer cannot enumerate or resolve -- but
    # regardless of the exact cause, the fix is to never need to resolve
    # it: confirm it's on screen via full-screen OCR (same mechanism
    # dismissed_post_popup_prompts already relies on, see its docstring --
    # send_keys() delivers to whatever currently has keyboard focus without
    # resolving a window object at all), then act on it blind. The File
    # name box already has focus by default the instant this dialog opens
    # (every live screenshot shows a cursor in it, standard Windows common-
    # dialog behavior), so select-all + type the full destination path +
    # Enter needs no click and no window resolution.
    #
    # PDF name/location is independent of the XML export's output_path: a
    # human-facing deliverable (Downloads folder, HH-MM_YYYYMMDD stem)
    # rather than the internal XML artifact, so it gets its own fresh
    # timestamp instead of reusing output_path's stem. A numeric suffix is
    # appended on collision so the minute-level timestamp can never
    # overwrite an existing PDF.
    pdf_stem = f"Medent_raw_{datetime.now():%m%d%Y}"
    pdf_output_path = PDF_OUTPUT_DIR / f"{pdf_stem}.pdf"
    n = 1
    while pdf_output_path.exists():
        n += 1
        pdf_output_path = PDF_OUTPUT_DIR / f"{pdf_stem}_{n}.pdf"

    def _save_dialog_visible():
        # wait_until_passes only retries on a raised exception -- it does
        # NOT check a returned False -- so this must raise, not return
        # False, or the very first (too-early) call would be treated as
        # already having passed.
        words = _run_ocr(ImageGrab.grab(all_screens=True))
        if _find_ocr_phrase(words, ("save", "print", "output")) is None:
            raise RuntimeError("'Save Print Output As' not yet visible via OCR")
        return True

    SAVE_DIALOG_TIMEOUT_S = 60

    try:
        wait_until_passes(SAVE_DIALOG_TIMEOUT_S, 1.0, _save_dialog_visible)
    except Exception as exc:
        _capture_diagnostics("save_print_dialog_not_found", win)
        raise NavigationCheckpoint(
            "Clicked the Print dialog's 'OK', but the native 'Save Print "
            f"Output As' dialog never became visible on screen (via OCR) "
            f"within {SAVE_DIALOG_TIMEOUT_S}s: {exc!r}. Inspect diagnostics "
            "under OUTPUT_DIR."
        )

    send_keys("^a", pause=0.05)
    send_keys(str(pdf_output_path), with_spaces=True, pause=0.02)

    # Click the actual 'Save' button (per the live screenshot) instead of
    # blindly sending Enter. Located via full-screen OCR, same as
    # _save_dialog_visible above, since this dialog's window object can't
    # be resolved via UIA. The word "Save" also appears in the dialog's
    # OWN TITLE ("Save Print Output As"), so disambiguate by pairing it
    # with the nearby "Cancel" word on the same button row -- the title
    # bar's "Save" has no "Cancel" neighbor at the same height.
    def _locate_save_button_center():
        words = _run_ocr(ImageGrab.grab(all_screens=True))
        save_words = [w for w in words if w["text"].strip(" .,:;\"'").lower() == "save"]
        cancel_words = [w for w in words if w["text"].strip(" .,:;\"'").lower() == "cancel"]
        if not save_words:
            return None
        if cancel_words:
            def dy_to_nearest_cancel(w):
                return min(abs(w["rect"][1] - c["rect"][1]) for c in cancel_words)
            candidate = min(save_words, key=dy_to_nearest_cancel)
        else:
            candidate = max(save_words, key=lambda w: w["rect"][1])
        x, y, w_, h_ = candidate["rect"]
        return (x + w_ // 2, y + h_ // 2)

    def _click_save_button():
        point = _locate_save_button_center()
        if point is None:
            raise RuntimeError("'Save' button text not found via OCR")
        pywinauto_mouse.click(coords=point)
        return True

    try:
        wait_until_passes(10, 0.5, _click_save_button)
    except Exception as exc:
        _capture_diagnostics("save_button_not_found", win)
        raise NavigationCheckpoint(
            "Typed the filename into the Save dialog, but could not locate/"
            f"click the 'Save' button via OCR within 10s: {exc!r}. Inspect "
            "diagnostics under OUTPUT_DIR."
        )

    def _pdf_saved():
        # Same wait_until_passes caveat as _save_dialog_visible above: must
        # raise, not return False, for the retry to actually happen.
        if not (pdf_output_path.exists() and pdf_output_path.stat().st_size > 0):
            raise RuntimeError(f"{pdf_output_path} not yet present/non-empty")
        return True

    try:
        wait_until_passes(15, 0.5, _pdf_saved)
    except Exception:
        _capture_diagnostics("pdf_output_missing_or_empty", win)
        raise NavigationCheckpoint(
            "Sent the filename and Enter to the Save dialog (confirmed "
            f"visible via OCR beforehand), but no non-empty PDF appeared at "
            f"{pdf_output_path} within 15s."
        )
    audit("saved_print_pdf", path=str(pdf_output_path), size_bytes=pdf_output_path.stat().st_size)
    _capture_click_screenshot("print_pdf_saved")
    print(f"PDF saved: {pdf_output_path}")

    # Confirmed live (both runs): the export format "Open with Microsoft
    # Excel" makes Medent generate the real structured report data as an
    # "XML Spreadsheet 2003" file and hand it to Excel to open directly --
    # Excel's own Save As dialog in both live runs showed that exact format
    # already selected by default, confirming this (not something this
    # script asked for separately). That file lives wherever Medent's own
    # export generated it (observed as some auto-named temp-style path,
    # never under our control) until something saves it elsewhere -- no
    # "Printing Open Items Report..." progress window was ever seen in
    # either live run (the previous, never-verified run_and_export() stub's
    # wait for one is replaced here with waiting for Excel's own window,
    # which both runs did confirm).
    def _excel_opened():
        return Desktop(backend="uia").window(class_name="XLMAIN").exists()

    try:
        wait_until_passes(90, 1, _excel_opened)
    except Exception as exc:
        _capture_diagnostics("excel_did_not_open", win)
        raise NavigationCheckpoint(
            f"Saved the print output to {pdf_output_path}, but Excel never "
            f"opened within 90s: {exc!r}. The PDF may still be valid -- "
            "check it directly -- but the actual report DATA (opened via "
            "Excel) was never captured. Inspect diagnostics under "
            "OUTPUT_DIR."
        )
    audit("excel_opened")

    # Rather than UI-automate Excel's own File -> Save As -> Browse ->
    # change format to "Excel Workbook" -> Save (what the live walkthrough
    # did by hand -- the single riskiest, least-tested remaining part of
    # this whole flow: backstage menus, a format dropdown, and a
    # compatibility prompt that appeared in ONE of the two live runs but
    # not the other, depending on exactly when the format was changed),
    # read the real data out directly via Excel's own COM API instead.
    # Excel always registers itself in the Running Object Table, so this
    # needs no window-clicking at all: get the live Application object,
    # copy its ActiveWorkbook's real file off disk to OUTPUT_DIR exactly as
    # Medent generated it (already confirmed well-formed XML, same format
    # verify_exported_report() already checks), then close the workbook
    # without resaving -- the original temp file is left untouched, and
    # nothing here can hit an unexpected Excel dialog.
    pythoncom.CoInitialize()
    try:
        excel_app = win32com.client.GetActiveObject("Excel.Application")
        workbook = excel_app.ActiveWorkbook
        if workbook is None:
            raise RuntimeError("Excel is open but has no ActiveWorkbook")
        source_path = Path(workbook.FullName)
        shutil.copy(source_path, output_path)
        audit("copied_report_data", source=str(source_path), dest=str(output_path))
        workbook.Close(SaveChanges=False)
    except Exception as exc:
        _capture_diagnostics("excel_com_copy_failed", win)
        raise NavigationCheckpoint(
            f"Excel opened, but reading its ActiveWorkbook via COM failed: "
            f"{exc!r}. Inspect diagnostics under OUTPUT_DIR -- the report "
            "data may still be sitting open in Excel; check it directly "
            "before closing anything by hand."
        )

    # Belt-and-braces: close the Excel window too, in case the workbook
    # Close() above didn't also end the whole Application (e.g. other
    # workbooks were already open before this run started).
    try:
        excel_win = Desktop(backend="uia").window(class_name="XLMAIN")
        if excel_win.exists():
            excel_win.close()
    except Exception as exc:
        audit("excel_window_close_failed", error=repr(exc))
    audit("excel_closed")

    # Confirmed live: 'Exit' and the logout confirmation 'Yes' that follows
    # are both embedded Panes in `win`, same pattern as everything else --
    # reuses the existing logout() helper rather than duplicating its
    # selector.
    logout(app)
    audit("point_ar_report_complete", xml_path=str(output_path), pdf_path=str(pdf_output_path))


def confirm_report_popup(win, timeout=5, retry=0.3):
    """Click 'Yes' on the 'Print A/R Tasks on Report?' confirmation.

    Resolves the popup AND its 'Yes' child AND clicks it as ONE retried
    attempt (same discipline as click_when_ready, for the same reason).

    Confirmed live via a captured control-tree dump
    (charges_for_ins_co_select_inconclusive_controltree_20261102_071750.txt)
    that 'Print A/R Tasks on Report?' is NOT a separate top-level window --
    an earlier assumption here that caused a real TimeoutError, since
    Desktop(...).window(title=...) never finds it. It's a Pane nested
    inside the SAME top-level window as the 'Print A/R Tasks on Report'
    field itself (`win`, i.e. app.top_window()), several levels deep
    (dialog -> auto_id=3 pane -> auto_id=1 pane -> this popup pane), so it
    must be resolved as a descendant of `win`, not of the Desktop.

    That fix alone still timed out on a later live run even with
    visible_only=False added (see git history for that intermediate
    theory). The real cause, confirmed via a control-tree dump taken at
    the exact moment of that failure
    (charges_for_ins_co_select_inconclusive_controltree_20261102_081404.txt,
    lines 50-64): 'Yes' and 'No' are NOT descendants of the 'Print A/R
    Tasks on Report?' pane at all -- that pane is a plain text label with
    no children. All three ('Print A/R Tasks on Report?', 'Yes', 'No')
    are SIBLING panes under the same auto_id=1 parent pane. Chaining
    .child_window(title="Yes") off of .child_window(title="Print A/R
    Tasks on Report?") therefore searches for 'Yes' *inside* a childless
    label and can never resolve, regardless of retry budget or
    visibility flags -- this was a wrong ancestor relationship in the
    query, not a timing/visibility issue. Fixed by resolving 'Yes' as a
    direct child of `win` instead, same as 'Print A/R Tasks on Report?'
    itself. The exact title "Yes" (no trailing spaces) is confirmed
    unique in that dump -- the only lookalike is 'Yes  ' (two trailing
    spaces, a different control for 'Skip line between Accounts').
    """
    def attempt():
        win.child_window(
            title="Print A/R Tasks on Report?", control_type="Pane",
            visible_only=False,
        ).wrapper_object()
        win.child_window(
            title="Yes", control_type="Pane", visible_only=False,
        ).click_input()
        return True

    wait_until_passes(timeout, retry, attempt)
    audit("confirmed_print_tasks_popup")

    # Not confirmed which window has focus here -- the popup is gone by this
    # point. Desktop(backend="uia") has no top_window() method (that's an
    # Application method, not a Desktop one -- confirmed live via
    # AttributeError: "Neither GUI element (wrapper) nor wrapper method
    # 'top_window' were found"). send_keys() sends to whatever has focus
    # without needing to resolve a window object at all, sidestepping the
    # "which window" question entirely.
    send_keys("{ENTER}{ENTER}")
    audit("dismissed_post_popup_prompts")


def verify_exported_report(output_path):
    """Structural-only verification that the export actually produced a
    usable file: existence, non-zero size, and well-formed XML. Deliberately
    never reads/logs the file's actual content (it's patient/claims data --
    PHI), only structural facts (size, element count).
    """
    if not output_path.exists():
        audit("export_verification_failed", path=str(output_path), reason="file_not_found")
        raise RuntimeError(f"Export reported complete but no file exists at {output_path}.")

    size = output_path.stat().st_size
    if size == 0:
        audit("export_verification_failed", path=str(output_path), reason="empty_file")
        raise RuntimeError(f"Exported file at {output_path} is empty (0 bytes).")

    try:
        tree = ET.parse(output_path)
        element_count = sum(1 for _ in tree.getroot().iter())
    except ET.ParseError as exc:
        audit(
            "export_verification_failed", path=str(output_path),
            reason="xml_parse_error", error=str(exc),
        )
        raise RuntimeError(f"Exported file at {output_path} is not well-formed XML: {exc}")

    audit("export_verified", path=str(output_path), size_bytes=size, element_count=element_count)
    print(f"Export verified: {output_path} ({size} bytes, {element_count} XML elements)")


def logout(app):
    """Confirmed live (two independent runs, 2026-11-03): 'Exit' and the
    logout confirmation 'Yes' that follows it are both embedded Panes in
    the main window, same touch-UI pattern as every other prompt in this
    app -- no separate top-level dialog for either.
    """
    win = app.top_window()
    win.child_window(title="Exit", control_type="Pane").click_input()
    audit("clicked_exit")

    try:
        wait_until_passes(5, 0.3, lambda: win.child_window(title="Yes", control_type="Pane").exists())
        win.child_window(title="Yes", control_type="Pane").click_input()
        audit("confirmed_logout")
    except Exception as exc:
        audit("logout_confirmation_not_found", error=repr(exc))

    audit("logged_out")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Medent 'Point A/R Table Report' automation."
    )
    parser.add_argument(
        "--date-from",
        required=False,
        help="Report range start date, MM/DD/YYYY (default: today)",
    )
    parser.add_argument(
        "--date-to",
        required=False,
        help="Report range end date, MM/DD/YYYY (default: today)",
    )
    args = parser.parse_args()
    today = datetime.now()
    try:
        date_from = datetime.strptime(args.date_from, "%m/%d/%Y") if args.date_from else today
        date_to = datetime.strptime(args.date_to, "%m/%d/%Y") if args.date_to else today
    except ValueError as exc:
        parser.error(f"Dates must be in MM/DD/YYYY format: {exc}")
    return date_from, date_to


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    # NOTE (pre-existing, not touched by this change): date_from/date_to are
    # validated here but navigate_to_point_ar_report()'s date-range clicks
    # are still the original hardcoded calendar navigation -- these CLI
    # values are not yet actually wired into that logic. Dates requested on
    # the command line will not currently take effect; see the module
    # docstring's "Not yet confirmed at all" section.
    date_from, date_to = parse_args()

    user_id, password = get_credentials()

    app = launch_medent_if_needed()
    if is_logged_in(app):
        audit("already_logged_in")
    else:
        login(app, user_id, password)
    wait_for_main_menu(app)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    PDF_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"Medent_{datetime.now():%Y%m%d_%H%M%S}.xml"

    try:
        navigate_to_point_ar_report(app, output_path)
    except NavigationCheckpoint as exc:
        print(f"Diagnostic checkpoint reached: {exc}")
        return
    except Exception as exc:
        # Anything past this point that isn't a NavigationCheckpoint (an
        # ElementNotFoundError on a not-yet-verified selector, a failed
        # export, etc.) previously crashed with a bare traceback and no
        # diagnostic trail. Capture full diagnostics before re-raising so a
        # blocker is always inspectable, without trying to auto-recover
        # (e.g. no auto-logout-on-failure -- clicking "Exit" from an
        # unknown/partial dialog state could itself cause unintended
        # effects in a live healthcare app, so a failure halts for a human
        # to assess).
        audit("unhandled_exception", stage="navigate_to_point_ar_report", error=repr(exc))
        _capture_diagnostics("unhandled_exception", app.top_window())
        raise

    verify_exported_report(output_path)
    print(f"Done. Exported file: {output_path}")


if __name__ == "__main__":
    main()
