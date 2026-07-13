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
REM  2. Edit PDF_FOLDER (and OUTPUT_FOLDER) below.
REM  3. If your PDF folder is inside OneDrive (or any path only your
REM     Windows account can access), fill in SERVICE_USER / SERVICE_PASSWORD
REM     below too — see the OneDrive note further down.
REM  4. Run this script as Administrator.
REM
REM  The folder is saved ONCE to config.json (next to the .exe) via
REM  "--configure" below. The service itself is then installed with NO
REM  --folder argument — it just runs "--watch" and reads config.json,
REM  so restarting the laptop never requires re-entering the folder.
REM  To change the folder later, re-run the "--configure" line only
REM  (no need to reinstall the whole service).
REM ============================================================

set PDF_FOLDER=C:\path\to\pdf\folder
set OUTPUT_FOLDER=
set APP_DIR=%~dp0..
set EXE_PATH=%APP_DIR%\dist\PDF_Extractor_CLI.exe

REM --- OneDrive folders: fill these in so the service runs as YOU, not as
REM     "Local System" (which cannot see your OneDrive files at all). Leave
REM     both blank if your PDF folder is a normal local folder, not OneDrive.
REM     SERVICE_USER is your Windows sign-in name, e.g. MYPC\John or just John.
set SERVICE_USER=
set SERVICE_PASSWORD=

if not exist "%EXE_PATH%" (
    echo ERROR: %EXE_PATH% not found.
    echo Run build_exe.bat first to create the standalone .exe.
    pause
    exit /b 1
)

REM --- IMPORTANT: if PDF_FOLDER is a mapped network drive (e.g. Z:\...),
REM     switch it to a full UNC path instead (e.g. \\server\share\folder).
REM     Windows services start before drive letters are mapped for any
REM     user, so a mapped-drive path will fail at boot even though it
REM     works fine when you're logged in and test it manually.
REM
REM --- ONEDRIVE NOTE: a OneDrive folder (e.g. C:\Users\John\OneDrive\Factures)
REM     is a normal local path, BUT:
REM       1. In OneDrive settings, right-click that folder -> "Always keep
REM          on this device", so files are real local files, not on-demand
REM          cloud placeholders the service can't trigger a download for.
REM       2. Fill in SERVICE_USER / SERVICE_PASSWORD above so the service
REM          runs as your account (which has access to your OneDrive),
REM          instead of the default "Local System" account (which does not).

echo.
echo Saving folder configuration (one-time)...
if "%OUTPUT_FOLDER%"=="" (
    "%EXE_PATH%" --configure --folder "%PDF_FOLDER%"
) else (
    "%EXE_PATH%" --configure --folder "%PDF_FOLDER%" --output "%OUTPUT_FOLDER%"
)
if errorlevel 1 (
    echo ERROR: could not save configuration. Check the PDF_FOLDER path above.
    pause
    exit /b 1
)

nssm install PdfDataExtractor "%EXE_PATH%" "--watch"
nssm set PdfDataExtractor AppDirectory "%APP_DIR%"
nssm set PdfDataExtractor DisplayName "PDF Data Extractor"
nssm set PdfDataExtractor Description "Watches a folder and extracts PDF data into Excel reports."

REM --- Run as your Windows account instead of Local System, if given
REM     (required for OneDrive / any folder only your account can access) ---
if not "%SERVICE_USER%"=="" (
    echo Configuring service to run as %SERVICE_USER% ...
    nssm set PdfDataExtractor ObjectName "%SERVICE_USER%" "%SERVICE_PASSWORD%"
)

REM --- Survive a laptop restart: start automatically with Windows ---
nssm set PdfDataExtractor Start SERVICE_AUTO_START

REM --- Start a little after boot, once the system (and OneDrive sync,
REM     network drives, etc.) is more likely ready ---
nssm set PdfDataExtractor DelayedAutostart 1

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
echo   even before anyone logs in - no need to open the app, and
echo   no need to re-enter the folder.
echo ============================================================
echo   Check status:   nssm status PdfDataExtractor
echo   Stop it:        nssm stop PdfDataExtractor
echo   Remove it:      nssm remove PdfDataExtractor confirm
echo   Change folder:  "%EXE_PATH%" --configure --folder "NEW_PATH"
echo                   then: nssm restart PdfDataExtractor
echo ============================================================
pause
