$ErrorActionPreference = "Stop"

$AppDir = "C:\ProgramData\MSToolkit"

New-Item -ItemType Directory -Path $AppDir -Force | Out-Null

Remove-Item -Path "$AppDir\Launch-MSToolkit.bat" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\MSToolkit.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\NewADUser.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\MSToolkit.ico" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\Launch-NewADUser.bat" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\Compare-UserGroups.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\Investigate-AccountLockout.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\IntuneTools.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\M365-Group-Compare.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\M365-Distribution-Group-Compare.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\M365-Conditional-Access-User-Manager.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\M365-Teams-Block-Number.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\M365-Exchange-Online-Tools.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\M365-OneDrive-SharePoint-Tools.ps1" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\Launch-M365-Teams-Block-Number.cmd" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$AppDir\MSToolkit.lnk" -Force -ErrorAction SilentlyContinue

Copy-Item -Path "$PSScriptRoot\Launch-MSToolkit.bat" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\MSToolkit.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\NewADUser.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\MSToolkit.ico" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\Launch-NewADUser.bat" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\Compare-UserGroups.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\Investigate-AccountLockout.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\IntuneTools.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\M365-Group-Compare.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\M365-Distribution-Group-Compare.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\M365-Conditional-Access-User-Manager.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\M365-Teams-Block-Number.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\M365-Exchange-Online-Tools.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\M365-OneDrive-SharePoint-Tools.ps1" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\Launch-M365-Teams-Block-Number.cmd" -Destination $AppDir -Force
Copy-Item -Path "$PSScriptRoot\MSToolkit.lnk" -Destination $AppDir -Force

Start-Process -FilePath "cmd.exe" `
    -ArgumentList "/c `"$AppDir\Launch-MSToolkit.bat`"" `
    -WorkingDirectory $AppDir