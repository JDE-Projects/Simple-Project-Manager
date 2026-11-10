"""Unreadable recovery copy that can be moved aside, for the recovery-unreadable scenario."""
from __future__ import annotations

import json
import os
import time


def main() -> None:
    app_dir = os.environ["UI_DRIVE_APP_DIR"]
    out_dir = os.environ["UI_DRIVE_OUT_DIR"]
    save_path = os.path.join(out_dir, "saved-project.xlsx")
    with open(os.path.join(app_dir, ".spm_recovery.xlsx"), "wb") as f:
        f.write(b"not a workbook")
    print(json.dumps({"mode": "recovery-unreadable", "save_path": save_path}), flush=True)
    while True:
        time.sleep(3600)


if __name__ == "__main__":
    main()
