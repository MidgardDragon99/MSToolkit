@echo off
setlocal EnableExtensions

REM ============================================================
REM  Start-IntuneTools.bat
REM  Launches IntuneTools.ps1 as the signed-in user. No elevation:
REM  Graph runs under your own identity and IntuneWinAppUtil only
REM  writes to your own folders, so nothing here needs admin.
REM
REM  Keep this file in the same folder as IntuneTools.ps1.
REM
REM  Run "Start-IntuneTools.bat debug" to keep the PowerShell
REM  console visible, which is how you see errors if it fails to
REM  start or closes unexpectedly.
REM ============================================================

set "SCRIPT_NAME=IntuneTools.ps1"
set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

set "BASE=%~dp0"
if "%BASE:~-1%"=="\" set "BASE=%BASE:~0,-1%"
set "PS1=%BASE%\%SCRIPT_NAME%"

if not exist "%PS1%" (
    echo.
    echo ERROR: %SCRIPT_NAME% was not found next to this batch file.
    echo Looked for:
    echo   "%PS1%"
    echo.
    pause
    exit /b 1
)

if not exist "%PSEXE%" (
    echo.
    echo ERROR: Windows PowerShell 5.1 was not found at:
    echo   "%PSEXE%"
    echo.
    pause
    exit /b 1
)

REM Clear the mark-of-the-web. A copy from a download or a share is blocked
REM even with the execution policy bypassed, and the error is not obvious.
"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -Command ^
 "Unblock-File -LiteralPath \"%PS1%\" -ErrorAction SilentlyContinue"

if /i "%~1"=="debug" goto DEBUG

REM Normal launch: hidden console, batch exits straight away so nothing
REM is left sitting behind the window.
start "" "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%"
exit /b 0

:DEBUG
echo Starting IntuneTools with the console visible...
echo.
"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -STA -File "%PS1%"
set "RC=%errorlevel%"
echo.
echo IntuneTools exited with code %RC%.
pause
exit /b %RC%
