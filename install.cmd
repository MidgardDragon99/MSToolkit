@echo off
cd /d "%~dp0"

if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
    "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -File "%~dp0InstallMSToolkit.ps1"
) else (
    "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -File "%~dp0InstallMSToolkit.ps1"
)

exit /b %errorlevel%
