@echo off
setlocal
cd /d "%~dp0"

echo Building MSToolkit install/uninstall SFX packages only...
echo.

set "SevenZip=C:\Program Files (x86)\7-Zip\7z.exe"
set "SfxModule=C:\Program Files (x86)\7-Zip\7zSD.sfx"

if not exist "%SevenZip%" goto MissingSevenZip
if not exist "%SfxModule%" goto MissingSfx

echo Validating install source files...

for %%F in (
  "MSToolkit.EXE"
  "MSToolkit.lnk"
  "Compare-UserGroups.ps1"
  "MSToolkit.ps1"
  "MSToolkit.ico"
  "install.cmd"
  "install-config.txt"
  "InstallMSToolkit.ps1"
  "IntuneTools.ps1"
  "Investigate-AccountLockout.ps1"
  "Launcher.ps1"
  "Launch-M365-Teams-Block-Number.cmd"
  "Launch-MSToolkit.bat"
  "Launch-InstallMSToolkit.bat"
  "Launch-NewADUser.bat"
  "Launch-UninstallMSToolkit.bat"
  "M365-Conditional-Access-User-Manager.ps1"
  "M365-Distribution-Group-Compare.ps1"
  "M365-Exchange-Online-Tools.ps1"
  "M365-Group-Compare.ps1"
  "M365-OneDrive-SharePoint-Tools.ps1"
  "M365-Teams-Block-Number.ps1"
  "NewADUser.ps1"
  "UninstallMSToolkit.ps1"
) do if not exist "%%~F" goto MissingInstallFile

for %%F in (
  "uninstall.cmd"
  "uninstall-config.txt"
  "UninstallMSToolkit.ps1"
) do if not exist "%%~F" goto MissingUninstallFile

echo Source validation passed.
echo.

del /f /q MSToolkit-Install.7z MSToolkit-Install.exe MSToolkit-Uninstall.7z MSToolkit-Uninstall.exe 2>nul

"%SevenZip%" a -t7z MSToolkit-Install.7z ^
  MSToolkit.EXE ^
  MSToolkit.lnk ^
  Compare-UserGroups.ps1 ^
  MSToolkit.ps1 ^
  MSToolkit.ico ^
  install.cmd ^
  install-config.txt ^
  InstallMSToolkit.ps1 ^
  IntuneTools.ps1 ^
  Investigate-AccountLockout.ps1 ^
  Launcher.ps1 ^
  Launch-M365-Teams-Block-Number.cmd ^
  Launch-MSToolkit.bat ^
  Launch-InstallMSToolkit.bat ^
  Launch-NewADUser.bat ^
  Launch-UninstallMSToolkit.bat ^
  M365-Conditional-Access-User-Manager.ps1 ^
  M365-Distribution-Group-Compare.ps1 ^
  M365-Exchange-Online-Tools.ps1 ^
  M365-Group-Compare.ps1 ^
  M365-OneDrive-SharePoint-Tools.ps1 ^
  M365-Teams-Block-Number.ps1 ^
  NewADUser.ps1 ^
  UninstallMSToolkit.ps1

if errorlevel 1 goto InstallArchiveFailed

copy /b "%SfxModule%" + install-config.txt + MSToolkit-Install.7z MSToolkit-Install.exe >nul
if errorlevel 1 goto InstallExeFailed

"%SevenZip%" a -t7z MSToolkit-Uninstall.7z ^
  uninstall.cmd ^
  uninstall-config.txt ^
  UninstallMSToolkit.ps1

if errorlevel 1 goto UninstallArchiveFailed

copy /b "%SfxModule%" + uninstall-config.txt + MSToolkit-Uninstall.7z MSToolkit-Uninstall.exe >nul
if errorlevel 1 goto UninstallExeFailed

if not exist MSToolkit-Install.exe goto FinalFailed
if not exist MSToolkit-Uninstall.exe goto FinalFailed

echo.
echo Done.
echo Created:
echo   MSToolkit-Install.exe
echo   MSToolkit-Uninstall.exe
echo.
pause
exit /b 0

:MissingSevenZip
echo ERROR: 7-Zip was not found:
echo "%SevenZip%"
goto Failed

:MissingSfx
echo ERROR: 7-Zip SFX module was not found:
echo "%SfxModule%"
goto Failed

:MissingInstallFile
echo ERROR: A required install source file is missing.
echo Check the file list above and make sure every payload file is in this folder.
goto Failed

:MissingUninstallFile
echo ERROR: A required uninstall source file is missing.
goto Failed

:InstallArchiveFailed
echo ERROR: Failed to create MSToolkit-Install.7z
goto Failed

:InstallExeFailed
echo ERROR: Failed to create MSToolkit-Install.exe
goto Failed

:UninstallArchiveFailed
echo ERROR: Failed to create MSToolkit-Uninstall.7z
goto Failed

:UninstallExeFailed
echo ERROR: Failed to create MSToolkit-Uninstall.exe
goto Failed

:FinalFailed
echo ERROR: One or both EXE files were not created.
goto Failed

:Failed
echo.
pause
exit /b 1
