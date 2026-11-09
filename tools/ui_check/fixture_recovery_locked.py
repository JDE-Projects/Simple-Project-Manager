"""Recovery copy that cannot be deleted, for the ui_drive recovery-locked scenarios.

Run directly for the restore-and-save scenario. The decline and unreadable
fixtures call hold_recovery with their own mode.
"""
from __future__ import annotations

import json
import os
import time

from fixture import write_recovery


def hold_recovery(mode: str, unreadable: bool = False) -> None:
    app_dir = os.environ["UI_DRIVE_APP_DIR"]
    out_dir = os.environ["UI_DRIVE_OUT_DIR"]
    save_path = os.path.join(out_dir, "saved-project.xlsx")
    recovery_path = os.path.join(app_dir, ".spm_recovery.xlsx")
    if unreadable:
        with open(recovery_path, "wb") as f:
            f.write(b"not a workbook")
    else:
        write_recovery(recovery_path, save_path)
    # Windows refuses to delete a file another process has open, while the app
    # can still read it.
    with open(recovery_path, "rb"):
        print(json.dumps({"mode": mode, "save_path": save_path}), flush=True)
        while True:
            time.sleep(3600)


if __name__ == "__main__":
    hold_recovery("recovery-locked-save")
