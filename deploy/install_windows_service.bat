@echo off
REM ============================================================
REM  Installs "PDF Data Extractor" as a Windows background
REM  service using NSSM (Non-Sucking Service Manager).
REM
REM  Prerequisite: run build_exe.bat first so dist\PDF_Extractor_CLI.exe
REM  exists. This installer uses that standalone .exe, so the target
REM  machine does NOT need Python installed.
REM
REM  1. Download NSSM from https://nssm.cc/download and put
REM     nssm.exe on your PATH (or next to this script).
REM  2. Edit PDF_FOLDER below to point at your PDF folder.
REM  3. Run this script as Administrator.
REM ============================================================

set PDF_FOLDER=C:\path\to\pdf\folder
set APP_DIR=%~dp0..
set EXE_PATH=%APP_DIR%\dist\PDF_Extractor_CLI.exe

if not exist "%EXE_PATH%" (
    echo ERROR: %EXE_PATH% not found.
    echo Run build_exe.bat first to create the standalone .exe.
    pause
    exit /b 1
)

nssm install PdfDataExtractor "%EXE_PATH%" "--folder ""%PDF_FOLDER%"" --watch"
nssm set PdfDataExtractor AppDirectory "%APP_DIR%"
nssm set PdfDataExtractor DisplayName "PDF Data Extractor"
nssm set PdfDataExtractor Description "Watches a folder and extracts PDF data into Excel reports."

REM --- Survive a laptop restart: start automatically with Windows ---
nssm set PdfDataExtractor Start SERVICE_AUTO_START

REM --- Survive a crash: if the process ever dies, NSSM restarts it ---
nssm set PdfDataExtractor AppExit Default Restart
nssm set PdfDataExtractor AppRestartDelay 5000
nssm set PdfDataExtractor AppThrottle 5000

REM --- Also log service output to files for troubleshooting ---
nssm set PdfDataExtractor AppStdout "%APP_DIR%\logs\service_stdout.log"
nssm set PdfDataExtractor AppStderr "%APP_DIR%\logs\service_stderr.log"

echo.
echo Starting the service now...
nssm start PdfDataExtractor

echo.
echo ============================================================
echo   Service installed and started.
echo   It will now start automatically every time this PC boots,
echo   even before anyone logs in - no need to open the app.
echo ============================================================
echo   Check status:   nssm status PdfDataExtractor
echo   Stop it:        nssm stop PdfDataExtractor
echo   Remove it:      nssm remove PdfDataExtractor confirm
echo ============================================================
pause
