@echo off
REM Launches the Teams Call Blocking tool as the Windows user signed in to this PC.
REM Do not use Run as different user - Microsoft sign-in (WAM) fails with 0x80070520 in that case.
REM conhost.exe gives the tool a classic console (not a Windows Terminal tab), the same kind
REM of console MSToolkit child tools get. The sign-in window is placed over that console.
start "" "%SystemRoot%\System32\conhost.exe" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0M365-Teams-Block-Number.ps1"
