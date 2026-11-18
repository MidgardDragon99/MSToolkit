@echo off
setlocal
rem Builds MSToolkit. Arguments are passed straight to Build-All.ps1, for example:
rem   Build-All.bat -SkipIntuneWin
rem   Build-All.bat -IntuneWinAppUtil "C:\Tools\IntuneWinAppUtil.exe"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-All.ps1" %*
set "RC=%ERRORLEVEL%"

echo.
pause
exit /b %RC%
