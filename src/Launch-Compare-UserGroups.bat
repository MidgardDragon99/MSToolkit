@echo off
setlocal
cd /d "%~dp0"

set "SCRIPT=%~dp0Compare-UserGroups.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: Could not find "%SCRIPT%"
    echo.
    pause
    exit /b 1
)

rem The admin USERNAME can be remembered - never the password, which runas always
rem asks for itself. Shared by all four runas launchers. It is used exactly as
rem typed: no domain is added or removed.
set "SAVEDIR=%APPDATA%\MSToolkit"
set "SAVEFILE=%SAVEDIR%\launcher-admin-username.txt"
set "SAVEDUSER="
if exist "%SAVEFILE%" set /p SAVEDUSER=<"%SAVEFILE%"

set "ATTEMPT=0"

:askuser
set "ADMINUSER="
set "NEWUSER=1"
if not defined SAVEDUSER goto asknew

set /p ADMINUSER=Domain admin username [%SAVEDUSER%] - press Enter to use it, or type another: 
if "%ADMINUSER%"=="" set "ADMINUSER=%SAVEDUSER%"
if /i "%ADMINUSER%"=="%SAVEDUSER%" set "NEWUSER=0"
goto checkuser

:asknew
set "ADMINUSER="
set "NEWUSER=1"
set /p ADMINUSER=Enter domain admin username, as DOMAIN\username or username@domain.com: 
if "%ADMINUSER%"=="" goto nouser
goto checkuser

:nouser
echo No username entered.
echo.
goto asknew

:checkuser
rem runas needs the domain in the name. Without one it signs in against this
rem computer's local accounts instead of the domain.
set "HASDOMAIN="
if not "%ADMINUSER%"=="%ADMINUSER:\=%" set "HASDOMAIN=1"
if not "%ADMINUSER%"=="%ADMINUSER:@=%" set "HASDOMAIN=1"
if defined HASDOMAIN goto askremember

echo "%ADMINUSER%" has no domain. Enter it as DOMAIN\username or username@domain.com.
echo.
goto asknew

:askremember
rem Asked before runas, so the tool does not open until this is answered. The
rem username is only saved after the logon below has actually worked.
set "REMEMBER=0"
if "%NEWUSER%"=="0" goto runscript
choice /c YN /n /m "Remember %ADMINUSER% as the admin username for next time? [Y/N]: "
if errorlevel 2 goto runscript
set "REMEMBER=1"

:runscript
set /a ATTEMPT+=1

runas /user:%ADMINUSER% "powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File \"%SCRIPT%\""

if errorlevel 1 goto failed

rem Save the username only now that the logon has worked.
if not "%REMEMBER%"=="1" goto launched

if not exist "%SAVEDIR%" mkdir "%SAVEDIR%"
>"%SAVEFILE%" echo %ADMINUSER%
set "SAVEDUSER=%ADMINUSER%"
echo Username remembered. Next time, just press Enter at the username prompt.
goto launched

:failed
echo.
echo Logon failed for %ADMINUSER% (attempt %ATTEMPT%).
echo.

choice /c RUQ /n /m "Press R to retry the password, U to change the username, or Q to quit: "

if errorlevel 3 goto quit
if errorlevel 2 goto askuser
goto runscript

:launched
echo.
echo Compare User Groups is starting...
timeout /t 2 /nobreak >nul
exit /b 0

:quit
echo.
echo Cancelled.
exit /b 1
