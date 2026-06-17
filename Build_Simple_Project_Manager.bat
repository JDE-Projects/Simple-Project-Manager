@echo off
rem ===========================================================================
rem  Build Simple Project Manager   JDE-Projects (https://github.com/JDE-Projects)
rem
rem  Double-click to build a windowed "Simple Project Manager" with PyInstaller.
rem  Closed / noncommercial build, so this uses --onedir (the Qt and other
rem  bundled libraries stay replaceable). Keep this .bat in the SAME folder as
rem  simple_project_manager.py, simple_project_manager-UI.html, the fonts folder,
rem  simple_project_manager.ico, simple_project_manager.png and
rem  simple_project_manager-splash.png.
rem  The finished app lands in dist\Simple Project Manager\.
rem ===========================================================================
cd /d "%~dp0"

rem --- bind to PySide6 (LGPL), not PyQt6 (GPL) ---
set QT_API=pyside6

rem --- check Python ---
where python >nul 2>&1
if not %errorlevel%==0 (
    echo Python was not found on PATH.
    echo Install Python 3 from https://www.python.org/downloads/ and tick
    echo "Add python.exe to PATH" during setup.
    echo.
    pause
    exit /b 1
)

rem --- make sure the source and assets are here ---
if not exist "simple_project_manager.py" (
    echo Could not find simple_project_manager.py next to this script.
    echo Put this .bat in the same folder as the source and asset files.
    echo.
    pause
    exit /b 1
)
if not exist "simple_project_manager.ico" echo WARNING: simple_project_manager.ico not found, the exe will use the default icon.
if not exist "simple_project_manager.png" echo WARNING: simple_project_manager.png not found, the taskbar icon may be generic.
if not exist "simple_project_manager-splash.png" echo WARNING: simple_project_manager-splash.png not found, no splash will show.
if not exist "fonts" echo WARNING: fonts folder not found, the window will fall back to system fonts.

rem --- make sure PyInstaller is available, install if missing ---
python -m PyInstaller --version >nul 2>&1
if not %errorlevel%==0 (
    echo PyInstaller not found. Installing it now...
    python -m pip install pyinstaller
    if not %errorlevel%==0 (
        echo Could not install PyInstaller. Check pip/network and try again.
        echo.
        pause
        exit /b 1
    )
)

rem --- make sure the runtime deps are present (pywebview + PySide6 + qtpy + openpyxl) ---
echo Ensuring dependencies (pywebview, PySide6, qtpy, openpyxl) are installed ...
python -m pip install pywebview PySide6 qtpy openpyxl
if not %errorlevel%==0 (
    echo Could not install pywebview / PySide6 / qtpy / openpyxl. Check pip/network and try again.
    echo.
    pause
    exit /b 1
)

rem --- clean previous output for a fresh build ---
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Simple Project Manager.spec" del /q "Simple Project Manager.spec"

echo.
echo Building Simple Project Manager ... this can take a minute.
echo.

python -m PyInstaller --noconfirm --onedir --windowed ^
    --name "Simple Project Manager" ^
    --icon "simple_project_manager.ico" ^
    --splash "simple_project_manager-splash.png" ^
    --add-data "simple_project_manager-UI.html;." ^
    --add-data "simple_project_manager.png;." ^
    --add-data "fonts;fonts" ^
    --collect-all PySide6 ^
    --collect-all qtpy ^
    --collect-all openpyxl ^
    simple_project_manager.py

if not %errorlevel%==0 (
    echo.
    echo Build failed. Read the last lines above for the cause.
    echo.
    pause
    exit /b 1
)

echo.
echo ===========================================================================
echo  Done. Your app folder is:  dist\Simple Project Manager\
echo  Run dist\Simple Project Manager\Simple Project Manager.exe to test, then
echo  zip the whole "Simple Project Manager" folder and attach it to Releases.
echo ===========================================================================
echo.
pause
