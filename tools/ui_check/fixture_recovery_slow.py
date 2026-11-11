"""Large recovery workbook for the ui_drive recovery timing scenario."""
from __future__ import annotations

import json
import os
import time

from fixture import COLUMNS
from openpyxl import Workbook


def write_recovery(path: str, source: str) -> None:
    """Write enough rows that recovery loading keeps the launch prompt delayed."""
    wb = Workbook()
    project = wb.active
    project.title = "Project"
    project.append(COLUMNS)
    project.append(["Phase", "Fixture phase", "", "", "", "", "", "", "", ""])
    project.append([
        "Step", "Fixture step", "John", "", "", "Not started", "", "Normal", "smoke", "",
    ])
    for number in range(12000):
        project.append([
            "Action", f"Fixture action {number}", "", "", "", "Not started", "", "Normal", "", "",
        ])
    meta = wb.create_sheet("_spm")
    meta.sheet_state = "hidden"
    meta["A1"], meta["B1"] = "app", "Simple Project Manager"
    meta["A2"], meta["B2"] = "schema", "1"
    meta["A3"], meta["B3"] = "tag_colors", "{}"
    meta["A4"], meta["B4"] = "source", source
    wb.save(path)


def main() -> None:
    app_dir = os.environ["UI_DRIVE_APP_DIR"]
    out_dir = os.environ["UI_DRIVE_OUT_DIR"]
    save_path = os.path.join(out_dir, "saved-project.xlsx")
    write_recovery(os.path.join(app_dir, ".spm_recovery.xlsx"), save_path)
    print(json.dumps({"mode": "recovery-slow-check", "save_path": save_path}), flush=True)
    while True:
        time.sleep(3600)


if __name__ == "__main__":
    main()
