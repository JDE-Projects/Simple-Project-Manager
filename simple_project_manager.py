"""
Simple Project Manager — an IT project planner/tracker.

JDE-Projects "Simple X Tool": Python 3 + PySide6/pywebview, single-file UI.
Phases -> Steps -> Actions (+ Contacts), with owners, deadlines, status,
priority, tags, notes and progress roll-up. The .xlsx IS the save format.

Phase 3a scope: bridge, theme/debug/update plumbing, and the 4-level tree
render fed by in-memory sample data. Editing and .xlsx I/O arrive in 3b.
"""
import json
import os
import sys
import datetime
import urllib.request

import webview

APP_VERSION = "1.0.0"
GITHUB_OWNER = "JDE-Projects"
GITHUB_REPO = "Simple-Project-Manager"

# Status / priority vocabularies — kept here so backend and UI agree.
STATUSES = ["Not started", "In progress", "Done", "Blocked"]
PRIORITIES = ["Normal", "High", "Critical"]


def resource_path(rel: str) -> str:
    """Path to a bundled resource, working both from source and PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def app_dir() -> str:
    """Folder the app lives in — next to the .exe when frozen, else the script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _sample_items():
    """Seed data so the tree renders before .xlsx load exists (3b)."""
    return [
        {"id": "p1", "type": "Phase", "name": "Network refresh — Building A",
         "owner": "", "coowner": "", "deadline": "", "status": "",
         "priority": "Normal", "tags": ["infra"], "notes": "FY26 capital project."},
        {"id": "s1", "type": "Step", "name": "Site survey & cabling audit",
         "owner": "John", "coowner": "", "deadline": "2026-06-30",
         "status": "In progress", "priority": "Normal", "tags": ["survey"],
         "notes": "Confirm riser capacity."},
        {"id": "a1", "type": "Action", "name": "Photograph each IDF",
         "owner": "", "coowner": "", "deadline": "", "status": "Done",
         "priority": "Normal", "tags": [], "notes": ""},
        {"id": "a2", "type": "Action", "name": "Label patch panels",
         "owner": "", "coowner": "", "deadline": "", "status": "Not started",
         "priority": "Normal", "tags": [], "notes": ""},
        {"id": "c1", "type": "Contact", "name": "Acme Cabling",
         "owner": "", "coowner": "", "deadline": "", "status": "",
         "priority": "Normal", "tags": [], "notes": "Vendor — j.doe@acme.com / 555-0142"},
        {"id": "s2", "type": "Step", "name": "Order switches",
         "owner": "Priya", "coowner": "John", "deadline": "2026-06-10",
         "status": "Blocked", "priority": "Critical", "tags": ["procurement"],
         "notes": "Waiting on PO approval — overdue."},
        {"id": "p2", "type": "Phase", "name": "Cutover",
         "owner": "", "coowner": "", "deadline": "", "status": "",
         "priority": "Normal", "tags": [], "notes": ""},
        {"id": "s3", "type": "Step", "name": "Maintenance window cutover",
         "owner": "John", "coowner": "", "deadline": "2026-07-15",
         "status": "Not started", "priority": "High", "tags": ["change"],
         "notes": "After-hours."},
    ]


class Api:
    """Bridge exposed to the UI. Methods return JSON-able values; the UI awaits."""

    def __init__(self):
        self._window = None
        self._debug = False
        self._debug_path = None

    def set_window(self, w):
        self._window = w

    # --- document data (stubbed in 3a; real .xlsx I/O in 3b) ---------------
    def get_state(self):
        """Initial payload the UI loads on startup."""
        return {
            "version": APP_VERSION,
            "theme": self._load_theme(),
            "statuses": STATUSES,
            "priorities": PRIORITIES,
            "items": _sample_items(),
            "filename": "",
        }

    # --- theme preference (local file, not stored in the .xlsx) ------------
    def _pref_path(self) -> str:
        return os.path.join(app_dir(), "simple_project_manager.pref")

    def _load_theme(self) -> str:
        try:
            with open(self._pref_path(), "r", encoding="utf-8") as f:
                theme = json.load(f).get("theme")
            return theme if theme in ("dark", "light") else "dark"
        except Exception:
            return "dark"

    def save_theme(self, theme: str):
        if theme not in ("dark", "light"):
            return {"ok": False}
        try:
            with open(self._pref_path(), "w", encoding="utf-8") as f:
                json.dump({"theme": theme}, f)
            self.log(f"Theme set to {theme}")
            return {"ok": True}
        except Exception as e:
            self.log(f"Could not save theme pref: {e}")
            return {"ok": False}

    # --- misc bridge helpers ------------------------------------------------
    def open_url(self, url: str):
        """Open a link in the system browser — never navigate the app window."""
        import webbrowser
        webbrowser.open(url)
        return {"ok": True}

    def check_update(self):
        """Compare the latest published release to APP_VERSION. Silent on failure."""
        result = {"version": APP_VERSION, "latest": None, "update": False, "url": ""}
        try:
            url = (f"https://api.github.com/repos/{GITHUB_OWNER}/"
                   f"{GITHUB_REPO}/releases/latest")
            req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json"})
            with urllib.request.urlopen(req, timeout=4) as r:
                data = json.load(r)
            latest = (data.get("tag_name") or "").lstrip("v")
            result["latest"] = latest
            result["url"] = data.get("html_url", "")
            if latest and self._is_newer(latest, APP_VERSION):
                result["update"] = True
        except Exception:
            pass  # offline / private repo / rate-limited — stay quiet
        return result

    @staticmethod
    def _is_newer(latest: str, current: str) -> bool:
        def parts(v):
            out = []
            for p in v.split("."):
                try:
                    out.append(int(p))
                except ValueError:
                    out.append(0)
            return out
        return parts(latest) > parts(current)

    # --- debug log ----------------------------------------------------------
    def set_debug(self, on: bool):
        self._debug = bool(on)
        if self._debug and not self._debug_path:
            stamp = datetime.datetime.now().strftime("%m%d%Y_%H%M%S")
            self._debug_path = os.path.join(app_dir(), f"Debug_Log_{stamp}.txt")
            self.log("Debug log started")
        return {"ok": True}

    def log(self, msg: str):
        if not self._debug or not self._debug_path:
            return
        try:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self._debug_path, "a", encoding="utf-8") as f:
                f.write(f"[{ts}] {msg}\n")
        except Exception:
            pass


def _close_splash():
    try:
        import pyi_splash  # only present in the frozen build
        pyi_splash.close()
    except Exception:
        pass


def main():
    api = Api()
    win = webview.create_window(
        "Simple Project Manager",
        url=resource_path("simple_project_manager-UI.html"),
        js_api=api,
        width=1180, height=820, min_size=(900, 600),
        background_color="#0a0e14",
    )
    api.set_window(win)
    win.events.loaded += _close_splash
    try:
        webview.start(gui="qt", icon=resource_path("simple_project_manager.png"))
    except TypeError:
        webview.start(gui="qt")


if __name__ == "__main__":
    main()
