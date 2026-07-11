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
nssm set PdfDataExtractor Start SERVICE_AUTO_START

echo.
echo Service installed. Start it with:  nssm start PdfDataExtractor
echo Stop it with:                      nssm stop PdfDataExtractor
echo Remove it with:                    nssm remove PdfDataExtractor confirm
pause
