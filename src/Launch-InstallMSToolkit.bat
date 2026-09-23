@echo off
setlocal

set "ScriptDir=%~dp0"
set "PsScript=%ScriptDir%InstallMSToolkit.ps1"

if not exist "%PsScript%" (
    echo ERROR: Could not find "%PsScript%"
    pause
    exit /b 1
)

if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
    set "PowerShellExe=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
) else (
    set "PowerShellExe=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

"%PowerShellExe%" -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%PowerShellExe%' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%PsScript%\"' -Verb RunAs"

endlocal