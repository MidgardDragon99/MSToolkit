@echo off
cd /d "%~dp0"

powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File "%~dp0M365-Distribution-Group-Compare.ps1"
