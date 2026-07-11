@echo off
REM ============================================================
REM  Builds PDF_Extractor.exe — a standalone Windows executable
REM  that requires NO Python installation to run.
REM
REM  This script itself DOES need Python + pip, but only on the
REM  ONE machine you build it on. Once built, copy the .exe from
REM  the "dist" folder to any Windows PC and double-click it —
REM  no install, no Python needed there.
REM
REM  Usage: double-click this file, or run it from a terminal:
REM      build_exe.bat
REM ============================================================

setlocal

echo.
echo === PDF Extractor - build standalone .exe ===
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found on this machine.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo ^(check "Add python.exe to PATH" during install^), then re-run this script.
    pause
    exit /b 1
)

echo Installing build dependencies...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements-build.txt
if errorlevel 1 (
    echo ERROR: failed to install dependencies. See messages above.
    pause
    exit /b 1
)

echo.
echo Building PDF_Extractor.exe  (this can take 1-3 minutes)...
python -m PyInstaller --noconfirm --onefile --windowed --name PDF_Extractor ^
    --icon NONE ^
    --collect-all pdfplumber ^
    --hidden-import watchdog.observers.polling ^
    main.py

if errorlevel 1 (
    echo.
    echo Build FAILED. Scroll up for the error message.
    pause
    exit /b 1
)

echo.
echo Building PDF_Extractor_CLI.exe  (headless / for the background service)...
python -m PyInstaller --noconfirm --onefile --console --name PDF_Extractor_CLI ^
    --icon NONE ^
    --collect-all pdfplumber ^
    --hidden-import watchdog.observers.polling ^
    cli.py

if errorlevel 1 (
    echo.
    echo Build FAILED. Scroll up for the error message.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   SUCCESS
echo   Standalone apps are in the "dist" folder:
echo     dist\PDF_Extractor.exe       (desktop app, double-click)
echo     dist\PDF_Extractor_CLI.exe   (command-line / service)
echo   Copy either file to any Windows PC and run it directly -
echo   no Python installation needed on that PC.
echo ============================================================
echo.
pause
