# Simple Project Manager

A standalone desktop tool to plan, track, and manage IT projects as a nested checklist: Phases hold Steps, Steps hold Actions and Contacts. Every item carries an owner and co-owner, a deadline, a status, a priority, tags, and notes, with progress rolling up automatically. Projects are saved as ordinary Excel `.xlsx` files. The interface is a clean web-style window.

Built by [JDE-Projects](https://github.com/JDE-Projects).

## Highlights

- **Four-level tree.** Phases -> Steps -> Actions (checklist items), with Contacts attached to a step. Each item nests under the nearest parent above it.
- **Excel is the save format.** New / Open / Save / Save As work on `.xlsx` files directly — no separate database. The workbook is human-readable, with the tree indented and Excel row grouping so phases fold up. A silent recovery copy guards unsaved work and offers to restore it on the next launch.
- **Rich per-item detail.** Owner and co-owner, deadline, status (Not started / In progress / Done / Blocked), priority (Low / Normal / High / Critical), colored tags, and notes — edited in a slide-in panel.
- **Automatic progress roll-up.** A step shows the percentage of its ticked actions; a phase averages its steps.
- **At-a-glance signals.** Overdue deadlines turn red and those due within a week turn amber; High and Critical priorities show a colored left bar and flag; a "Critical items" summary sits at the top of the project.
- **Filtering and reordering.** Filter by status, priority, or owner; reorder phases, steps, actions, and contacts with up/down controls. Completed steps gray out and collapse.
- **Dark and light teal themes.** A sun/moon toggle in the header switches between them; your choice is remembered locally.
- **Update check.** Checks the GitHub Releases page and points you to a newer version when one is published.

## How it works

- The backend is a small Python application; project files are read and written with [openpyxl](https://openpyxl.readthedocs.io/).
- The window is a [pywebview](https://pywebview.flowrl.com/) host on the Qt backend (PySide6), with the UI in `simple_project_manager-UI.html`. Fonts (Sora, JetBrains Mono) are bundled locally in `fonts/`, so the look holds with no internet.
- A `.xlsx` you save IS the project file — open it in Excel directly if you like. A hidden `_spm` sheet stores an app marker, schema version, and tag colors; the app only reads the columns it owns and confirms anything unexpected on import.

## Download and run

Grab the latest `Simple Project Manager` release from the [Releases](../../releases) page, unzip it, and run `Simple Project Manager.exe` inside the folder. Keep the folder together (the app ships as a folder, not a single loose .exe). No Python or setup required. Windows only.

Because it is unsigned, Windows SmartScreen may warn about an unknown publisher the first time. Click **More info > Run anyway**.

## Build from source (optional)

If you would rather run or build it yourself, you need:

- **Python 3** on the machine's PATH.
- Python packages: `pywebview`, `PySide6`, `openpyxl`, `pyinstaller`. Keep `PyQt6` uninstalled so PySide6 is the bundled binding.

```
pip install pywebview PySide6 openpyxl pyinstaller
```

Keep `simple_project_manager.py`, `simple_project_manager-UI.html`, the `fonts/` folder, `simple_project_manager.ico`, `simple_project_manager.png` and `simple_project_manager-splash.png` together. Then either:

- **Run from source:** `python simple_project_manager.py`
- **Build the .exe:** double-click `Build_Simple_Project_Manager.bat`, which uses PyInstaller to produce `dist\Simple Project Manager\Simple Project Manager.exe`. Distribute the whole `Simple Project Manager` folder.

## Using it

1. Click **+ Phase** to add a phase, then use the **+ Step** button in a phase header to add steps.
2. Click any item to open its edit panel. Inside a step you manage its action items and contacts, set owners, deadlines, status, priority, tags, and notes.
3. Tick actions as you finish them — progress rolls up to the step and phase automatically.
4. Use the filter bar to focus on a status, priority, or owner, and the up/down controls to reorder.
5. **Save** writes an `.xlsx` you can reopen here or in Excel. **Open** loads one back.

## Security and privacy

- Project `.xlsx` files are ordinary Excel workbooks saved wherever you choose. They may list internal hosts, vendors, and contact details, so keep them out of public source control.
- Your theme choice is stored in a small local preference file next to the app; it is not part of any project file.
- The debug log is off by default. When you turn it on, it writes one `Debug_Log_*.txt` next to the app for that run.

## A note on how this was built

This project was built with AI assistance. The design decisions, feature direction, and real-world testing were directed by me. The code was written and revised with an AI assistant against that direction. Treat it like any community tool, review and test it before relying on it.

## License

Released under the [PolyForm Noncommercial License 1.0.0](LICENSE). Personal and any noncommercial use, modification, and noncommercial redistribution are allowed; commercial use is not. Keep the copyright notice; no warranty.

This build bundles third-party code (Qt via PySide6, pywebview, openpyxl, and the Sora and JetBrains Mono fonts). Their notices are in [THIRD-PARTY-LICENSES.txt](THIRD-PARTY-LICENSES.txt).

For commercial licensing, open a [GitHub issue](https://github.com/JDE-Projects/Simple-Project-Manager/issues) with the title "Commercial License Inquiry".
