@echo off
cd /d "%~dp0"
set "LOG=%~dp0build_log.txt"
set "VPY=%~dp0build_env\Scripts\python.exe"
echo === QC_PMMS build (clean environment) === > "%LOG%"

if not exist "QC_PMMS.py" goto missing
if not exist "NAMAA_PNG_main_3.ico" goto missing

python --version >> "%LOG%" 2>&1
if errorlevel 1 goto nopython

if exist "%VPY%" goto haveenv
echo [1/4] Creating clean build environment...
python -m venv "%~dp0build_env" >> "%LOG%" 2>&1
if errorlevel 1 goto failed

:haveenv
echo [2/4] Installing requirements (needs internet)...
"%VPY%" -m pip install --upgrade pip >> "%LOG%" 2>&1
"%VPY%" -m pip install pyogrio shapely pyproj numpy openpyxl pyinstaller >> "%LOG%" 2>&1
if errorlevel 1 goto failed

echo [3/4] Checking imports...
"%VPY%" -c "import tkinter, pyogrio, shapely, pyproj, openpyxl, numpy; print('imports OK')" >> "%LOG%" 2>&1
if errorlevel 1 goto failed

echo [4/4] Building QC_PMMS.exe - this can take several minutes...
"%VPY%" -m PyInstaller --noconfirm --clean --onefile --noconsole --name QC_PMMS --icon "NAMAA_PNG_main_3.ico" --add-data "NAMAA_PNG_main_3.ico;." --collect-all pyogrio --collect-all pyproj --collect-all shapely --hidden-import uuid --exclude-module PyQt5 --exclude-module PyQt6 --exclude-module PySide2 --exclude-module PySide6 --exclude-module matplotlib --exclude-module scipy --exclude-module IPython QC_PMMS.py >> "%LOG%" 2>&1
if errorlevel 1 goto failed
if not exist "dist\QC_PMMS.exe" goto failed

echo.
echo SUCCESS: dist\QC_PMMS.exe
pause
exit /b 0

:missing
echo QC_PMMS.py and NAMAA_PNG_main_3.ico must be in the same folder as this .bat file.
pause
exit /b 1

:nopython
echo Python was not found in PATH. Open this file from an Anaconda Prompt, or install Python 3.11/3.12 (64-bit).
pause
exit /b 1

:failed
echo.
echo BUILD FAILED. Last lines of build_log.txt:
echo -----------------------------------------
powershell -NoProfile -Command "Get-Content -Tail 30 -LiteralPath '%LOG%'"
echo -----------------------------------------
echo Full log: %LOG%
pause
exit /b 1
