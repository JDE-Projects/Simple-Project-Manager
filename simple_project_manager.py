"""
Simple Project Manager — an IT project planner/tracker.

JDE-Projects "Simple X Tool": Python 3 + PySide6/pywebview, single-file UI.
Phases -> Steps -> Actions (+ Contacts), with owners, deadlines, status,
priority, tags, notes and progress roll-up. The .xlsx IS the save format.

Phase 3b scope: per-item editing, .xlsx Save/Open/Save As (the workbook IS
the save file), native dialogs, and a silent recovery copy of unsaved work.
"""
import json
import os
import sys
import datetime
import urllib.request

import webview
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

APP_VERSION = "1.0.0"
GITHUB_OWNER = "JDE-Projects"
GITHUB_REPO = "Simple-Project-Manager"

# Status / priority vocabularies — kept here so backend and UI agree.
STATUSES = ["Not started", "In progress", "Done", "Blocked"]
PRIORITIES = ["Low", "Normal", "High", "Critical"]

# .xlsx schema — one sheet, one row per item; Type drives the tree nesting.
SCHEMA_VERSION = "1"
COLUMNS = ["Type", "Name", "Owner", "Co-owner", "Deadline", "Status",
           "Progress", "Priority", "Tags", "Notes"]
LEVELS = {"Phase": 0, "Step": 1, "Action": 2, "Contact": 2}


def _to_date(value):
    """Coerce a stored value to a date for a real Excel date cell, else None."""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    s = str(value or "").strip()
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def resource_path(rel: str) -> str:
    """Path to a bundled resource, working both from source and PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def app_dir() -> str:
    """Folder the app lives in — next to the .exe when frozen, else the script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _item(id, type, name, owner="", coowner="", deadline="", status="",
          priority="Normal", tags=None, notes=""):
    return {"id": id, "type": type, "name": name, "owner": owner, "coowner": coowner,
            "deadline": deadline, "status": status, "priority": priority,
            "tags": tags or [], "notes": notes}


def _sample_items():
    """Seed data for testing — 3 phases of varying size, one step over the 3/2 cap.

    NOTE: this rich sample is for the testing phase. Before release we'll decide
    what a fresh launch shows (likely an empty 'New / Open' state).
    """
    return [
        # --- Phase 1: a busy phase, with a step that overflows the tree preview ---
        _item("p1", "Phase", "Network refresh — Building A", tags=["infra"],
              notes="FY26 capital project."),
        _item("s1", "Step", "Site survey & cabling audit", owner="John",
              deadline="2026-06-30", status="In progress", tags=["survey"],
              notes="Confirm riser capacity before ordering."),
        _item("a1", "Action", "Photograph each IDF", status="Done"),
        _item("a2", "Action", "Label patch panels", status="Not started"),
        _item("a3", "Action", "Test fiber runs", status="Not started",
              priority="High", deadline="2026-06-25"),
        _item("a4", "Action", "Update IDF floor map", status="Done"),
        _item("a5", "Action", "Get survey sign-off", status="Not started"),
        _item("c1", "Contact", "Acme Cabling", notes="Vendor — j.doe@acme.com / 555-0142"),
        _item("c2", "Contact", "Globex Networks", notes="Fiber contractor — ops@globex.com"),
        _item("c3", "Contact", "Initech Comms", notes="Backup vendor — 555-0199"),
        _item("s2", "Step", "Order switches", owner="Priya", coowner="John",
              deadline="2026-06-10", status="Blocked", priority="Critical",
              tags=["procurement"], notes="Waiting on PO approval — overdue."),
        _item("s3", "Step", "Rack & stack", owner="John", deadline="2026-07-05",
              status="Not started", priority="High", tags=["install"],
              notes="After delivery confirmed."),
        _item("a6", "Action", "Mount switches", status="Not started"),
        _item("a7", "Action", "Cable management", status="Not started"),

        # --- Phase 2: medium phase ---
        _item("p2", "Phase", "Server migration", tags=["migration"]),
        _item("s4", "Step", "Inventory current VMs", owner="Priya",
              deadline="2026-06-20", status="Done", tags=["audit"]),
        _item("a8", "Action", "Export VM list", status="Done"),
        _item("a9", "Action", "Tag critical workloads", status="Done"),
        _item("s5", "Step", "Migrate file server", owner="John",
              deadline="2026-07-01", status="In progress", priority="High",
              tags=["storage"], notes="Largest data set — stage over a weekend."),
        _item("a10", "Action", "Copy shares", status="Not started"),
        _item("a11", "Action", "Verify permissions", status="Not started"),
        _item("a12", "Action", "Cut over DNS", status="Not started", priority="High"),
        _item("c4", "Contact", "Vendor support", notes="Storage support — case #4471"),
        _item("s6", "Step", "Decommission old hardware", deadline="2026-07-20",
              status="Not started"),

        # --- Phase 3: light phase ---
        _item("p3", "Phase", "Cutover"),
        _item("s7", "Step", "Maintenance window cutover", owner="John",
              deadline="2026-07-15", status="Not started", priority="High",
              tags=["change"], notes="After-hours."),
        _item("a13", "Action", "Notify stakeholders", status="Not started"),
        _item("s8", "Step", "Post-cutover validation", owner="Priya",
              deadline="2026-07-16", status="Not started"),
        _item("c5", "Contact", "NOC on-call", notes="24/7 desk — 555-0100"),
    ]


class Api:
    """Bridge exposed to the UI. Methods return JSON-able values; the UI awaits."""

    def __init__(self):
        self._window = None
        self._debug = False
        self._debug_path = None
        self._filename = ""

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

    # --- progress roll-up (auto column; UI shows it richer in 3c) ----------
    def _progress_map(self, items):
        """id -> percent for phases/steps. Actions render as ✔/☐ instead."""
        prog, structure, phase, step = {}, [], None, None
        for it in items:
            t = it.get("type")
            if t == "Phase":
                phase = {"id": it["id"], "steps": []}
                step = None
                structure.append(phase)
            elif t == "Step":
                step = {"id": it["id"], "status": it.get("status", ""), "actions": []}
                if phase is None:
                    phase = {"id": None, "steps": []}
                    structure.append(phase)
                phase["steps"].append(step)
            elif t == "Action" and step is not None:
                step["actions"].append(it.get("status", ""))
        for ph in structure:
            pcts = []
            for st in ph["steps"]:
                if st["actions"]:
                    done = sum(1 for s in st["actions"] if s == "Done")
                    pct = round(100 * done / len(st["actions"]))
                else:
                    pct = 100 if st["status"] == "Done" else 0
                prog[st["id"]] = pct
                pcts.append(pct)
            if ph["id"] is not None:
                prog[ph["id"]] = round(sum(pcts) / len(pcts)) if pcts else 0
        return prog

    # --- .xlsx writer / reader (the workbook IS the save file) -------------
    def _write_xlsx(self, path, items, source=""):
        wb = Workbook()
        ws = wb.active
        ws.title = "Project"
        ws.append(COLUMNS)
        head_fill = PatternFill("solid", fgColor="178577")
        for c in range(1, len(COLUMNS) + 1):
            cell = ws.cell(1, c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"

        prog = self._progress_map(items)
        r = 1
        for it in items:
            r += 1
            typ = it.get("type", "")
            lvl = LEVELS.get(typ, 0)
            name_cell = ws.cell(r, 2, it.get("name", ""))
            name_cell.alignment = Alignment(indent=lvl, vertical="center")
            if typ == "Phase":
                name_cell.font = Font(bold=True)
            ws.cell(r, 1, typ)
            ws.cell(r, 3, it.get("owner", ""))
            ws.cell(r, 4, it.get("coowner", ""))
            d = _to_date(it.get("deadline", ""))
            dcell = ws.cell(r, 5)
            if d is not None:
                dcell.value = d
                dcell.number_format = "yyyy-mm-dd"
            else:
                dcell.value = it.get("deadline", "")
            ws.cell(r, 6, it.get("status", ""))
            pcell = ws.cell(r, 7)
            if typ == "Action":
                pcell.value = "✔" if it.get("status") == "Done" else "☐"
                pcell.alignment = Alignment(horizontal="center")
            elif it.get("id") in prog:
                pcell.value = prog[it["id"]] / 100.0
                pcell.number_format = "0%"
            ws.cell(r, 8, it.get("priority", "") if typ in ("Step", "Action") else "")
            ws.cell(r, 9, ",".join(it.get("tags") or []))
            ws.cell(r, 10, it.get("notes", ""))
            if lvl:
                ws.row_dimensions[r].outline_level = lvl
        ws.sheet_properties.outlinePr.summaryBelow = False

        for i, w in enumerate([9, 42, 12, 12, 13, 13, 10, 10, 20, 44], 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        last = max(r, 2)
        dv_s = DataValidation(type="list",
                              formula1='"%s"' % ",".join(STATUSES), allow_blank=True)
        ws.add_data_validation(dv_s)
        dv_s.add(f"F2:F{last}")
        dv_p = DataValidation(type="list",
                              formula1='"%s"' % ",".join(PRIORITIES), allow_blank=True)
        ws.add_data_validation(dv_p)
        dv_p.add(f"H2:H{last}")

        meta = wb.create_sheet("_spm")
        meta.sheet_state = "hidden"
        meta["A1"], meta["B1"] = "app", "Simple Project Manager"
        meta["A2"], meta["B2"] = "schema", SCHEMA_VERSION
        meta["A3"], meta["B3"] = "tag_colors", json.dumps({})
        meta["A4"], meta["B4"] = "source", source
        wb.save(path)

    def _read_xlsx(self, path):
        wb = load_workbook(path, data_only=True)
        is_ours = "_spm" in wb.sheetnames
        ws = wb["Project"] if "Project" in wb.sheetnames else wb.active
        header = [str(c.value).strip().lower() if c.value is not None else ""
                  for c in ws[1]]
        idx = {h: i for i, h in enumerate(header)}

        def get(row, *names, default=""):
            for n in names:
                i = idx.get(n)
                if i is not None and i < len(row) and row[i] is not None:
                    return row[i]
            return default

        items, n = [], 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            typ = str(get(row, "type") or "").strip()
            if typ not in ("Phase", "Step", "Action", "Contact"):
                continue
            n += 1
            dl = get(row, "deadline")
            if isinstance(dl, (datetime.datetime, datetime.date)):
                dl = dl.strftime("%Y-%m-%d")
            else:
                dl = str(dl).strip() if dl else ""
            raw_tags = str(get(row, "tags") or "")
            tags = [t.strip() for t in raw_tags.split(",") if t.strip()]
            items.append({
                "id": f"i{n}", "type": typ,
                "name": str(get(row, "name") or "").strip(),
                "owner": str(get(row, "owner") or "").strip(),
                "coowner": str(get(row, "co-owner", "coowner") or "").strip(),
                "deadline": dl,
                "status": str(get(row, "status") or "").strip(),
                "priority": (str(get(row, "priority") or "").strip() or "Normal"),
                "tags": tags,
                "notes": str(get(row, "notes") or "").strip(),
            })
        warn = "" if is_ours else ("Opened best-effort — this workbook has no "
                                   "Simple Project Manager marker.")
        return items, warn

    # --- document commands (called from the UI) ----------------------------
    def new_project(self):
        self._filename = ""
        self._clear_recovery()
        self.log("New project")
        return {"ok": True, "items": [], "filename": ""}

    def open_project(self):
        try:
            sel = self._window.create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=False,
                file_types=("Excel workbook (*.xlsx)", "All files (*.*)"))
        except Exception as e:
            self.log(f"Open dialog failed: {e}")
            return {"ok": False, "error": "Couldn't open the file picker."}
        if not sel:
            return {"cancelled": True}
        path = sel[0] if isinstance(sel, (list, tuple)) else sel
        try:
            items, warn = self._read_xlsx(path)
        except Exception as e:
            self.log(f"Open failed for {path}: {e}")
            return {"ok": False,
                    "error": "That file couldn't be opened as a project workbook."}
        self._filename = path
        self._clear_recovery()
        self.log(f"Opened {path} ({len(items)} rows)")
        return {"ok": True, "items": items, "filename": path, "warn": warn}

    def save_project(self, items):
        if not self._filename:
            return self.save_as(items)
        return self._do_save(self._filename, items)

    def save_as(self, items):
        suggested = os.path.basename(self._filename) if self._filename else "project.xlsx"
        try:
            sel = self._window.create_file_dialog(
                webview.SAVE_DIALOG, save_filename=suggested,
                file_types=("Excel workbook (*.xlsx)",))
        except Exception as e:
            self.log(f"Save dialog failed: {e}")
            return {"ok": False, "error": "Couldn't open the save dialog."}
        if not sel:
            return {"cancelled": True}
        path = sel[0] if isinstance(sel, (list, tuple)) else sel
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        return self._do_save(path, items)

    def _do_save(self, path, items):
        try:
            self._write_xlsx(path, items)
        except PermissionError:
            self.log(f"Save denied (file open in Excel?): {path}")
            return {"ok": False,
                    "error": "Couldn't save — is the file open in Excel?"}
        except Exception as e:
            self.log(f"Save failed for {path}: {e}")
            return {"ok": False, "error": "Couldn't save the project file."}
        self._filename = path
        self._clear_recovery()
        self.log(f"Saved {path}")
        return {"ok": True, "filename": path}

    # --- silent recovery copy of unsaved work ------------------------------
    def _recovery_path(self):
        return os.path.join(app_dir(), ".spm_recovery.xlsx")

    def write_recovery(self, items):
        try:
            self._write_xlsx(self._recovery_path(), items, source=self._filename)
        except Exception as e:
            self.log(f"Recovery write failed: {e}")
            return {"ok": False}
        return {"ok": True}

    def _clear_recovery(self):
        try:
            p = self._recovery_path()
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

    def clear_recovery(self):
        self._clear_recovery()
        return {"ok": True}

    def check_recovery(self):
        p = self._recovery_path()
        if not os.path.exists(p):
            return {"exists": False}
        try:
            items, _ = self._read_xlsx(p)
            wb = load_workbook(p, data_only=True)
            src = ""
            if "_spm" in wb.sheetnames:
                src = str(wb["_spm"]["B4"].value or "")
        except Exception as e:
            self.log(f"Recovery read failed: {e}")
            self._clear_recovery()
            return {"exists": False}
        self._filename = src
        return {"exists": True, "items": items, "filename": src}

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
