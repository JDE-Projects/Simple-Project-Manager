@echo off
setlocal
rem ===========================================================================
rem  Build Simple Project Manager   JDE-Projects (https://github.com/JDE-Projects)
rem
rem  Double-click to build a windowed "Simple Project Manager" with PyInstaller.
rem  Closed / noncommercial build, so this uses --onedir (the Qt and other
rem  bundled libraries stay replaceable). Keep this .bat in the SAME folder as
rem  simple_project_manager.py, simple_project_manager-UI.html, the fonts folder,
rem  simple_project_manager.ico and simple_project_manager.png.
rem  The build creates a private environment in build\venv, requires 64-bit
rem  Python 3.14, and installs the hash-locked requirements.txt package list.
rem  The finished app lands in dist\Simple Project Manager\.
rem ===========================================================================

cd /d "%~dp0"

rem --- skip interactive pauses when running in CI (GitHub Actions sets CI) ---
set "PAUSE=pause"
if defined CI set "PAUSE="

rem --- check Python: requirements.txt is locked for 64-bit Python 3.14 ---
python -c "import struct, sys; sys.exit(0 if sys.version_info[:2] == (3, 14) and struct.calcsize('P') == 8 else 1)" >nul 2>&1
if errorlevel 1 (
    echo ERROR: This build needs 64-bit Python 3.14 on PATH.
    echo Install it from https://python.org and tick "Add Python to PATH".
    %PAUSE%
    exit /b 1
)

rem --- make sure the source and assets are here ---
if not exist "simple_project_manager.py" (
    echo ERROR: Could not find simple_project_manager.py next to this script.
    echo Put this .bat in the same folder as the source and asset files.
    echo.
    %PAUSE%
    exit /b 1
)
if not exist "simple_project_manager.ico" echo WARNING: simple_project_manager.ico not found, the exe will use the default icon.
if not exist "simple_project_manager.png" echo WARNING: simple_project_manager.png not found, the taskbar icon may be generic.
if not exist "fonts" echo WARNING: fonts folder not found, the window will fall back to system fonts.

rem --- build inside a private, throwaway Python environment so the PC's own ---
rem --- Python is never changed. It is rebuilt from scratch every run. ---
set "VENV=build\venv"
set "VPY=%VENV%\Scripts\python.exe"
if exist "%VENV%\" (
    echo Removing the previous build environment...
    rmdir /s /q "%VENV%"
)
if exist "%VENV%\" (
    echo ERROR: Could not remove the previous build environment in %VENV%.
    echo Close any program using files in it, then run this script again.
    %PAUSE%
    exit /b 1
)
echo Creating a fresh build environment in %VENV% ...
python -m venv "%VENV%"
if errorlevel 1 (
    echo ERROR: Could not create the build environment in %VENV%.
    %PAUSE%
    exit /b 1
)
"%VPY%" --version
if errorlevel 1 (
    echo ERROR: Could not run Python from the build environment.
    %PAUSE%
    exit /b 1
)
"%VPY%" -m pip --version
if errorlevel 1 (
    echo ERROR: Could not run pip from the build environment.
    %PAUSE%
    exit /b 1
)
echo.
echo Installing the locked package list, each package checked against its hash...
"%VPY%" -m pip install --disable-pip-version-check --require-hashes --build-constraint build-constraints.txt -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install the locked packages from requirements.txt.
    %PAUSE%
    exit /b 1
)

echo.
echo Building Simple Project Manager, onedir so the bundled Qt stays replaceable...
set QT_API=pyside6
rem --noconfirm replaces the previous dist folder without asking.
"%VPY%" -m PyInstaller --noconfirm --onedir --windowed --name "Simple Project Manager" ^
  --icon "simple_project_manager.ico" ^
  --add-data "simple_project_manager-UI.html;." ^
  --add-data "simple_project_manager.png;." ^
  --add-data "fonts;fonts" ^
  --add-data "licenses;licenses" ^
  --collect-all openpyxl ^
  --hidden-import truststore ^
  simple_project_manager.py
if errorlevel 1 (
    echo ERROR: PyInstaller failed. See the messages above.
    %PAUSE%
    exit /b 1
)
"%VPY%" tools\trim_bundle.py "dist\Simple Project Manager"
if errorlevel 1 (
    echo ERROR: Qt bundle trim failed or was refused. See the messages above.
    %PAUSE%
    exit /b 1
)

echo.
echo ===========================================================================
echo  Done. Your app folder is:  dist\Simple Project Manager\
echo  Run dist\Simple Project Manager\Simple Project Manager.exe to test.
echo ===========================================================================
echo.
%PAUSE%
