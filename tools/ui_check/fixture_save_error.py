"""Recovery-workbook fixture for the ui_drive save-error scenario."""
from __future__ import annotations

import json
import os
import time

from fixture import write_recovery


def main() -> None:
    app_dir = os.environ["UI_DRIVE_APP_DIR"]
    out_dir = os.environ["UI_DRIVE_OUT_DIR"]
    save_path = os.path.join(out_dir, "saved-project.xlsx")
    write_recovery(os.path.join(app_dir, ".spm_recovery.xlsx"), save_path)
    # A folder at the save path makes the direct save fail with no dialog.
    os.makedirs(save_path)
    print(json.dumps({"mode": "save-error", "save_path": save_path}), flush=True)
    while True:
        time.sleep(3600)


if __name__ == "__main__":
    main()
