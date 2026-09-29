$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$ProgramFilesX86 = ${env:ProgramFiles(x86)}
if ([string]::IsNullOrWhiteSpace($ProgramFilesX86)) {
    $ProgramFilesX86 = $env:ProgramFiles
}

$InstallDir = Join-Path $ProgramFilesX86 "MSToolkit"
$RuntimeDir = "C:\ProgramData\MSToolkit"
$ShortcutPath = Join-Path $env:PUBLIC "Desktop\MSToolkit.lnk"

$PayloadFiles = @(
    "MSToolkit.EXE",
    "MSToolkit.lnk",
    "Compare-UserGroups.ps1",
    "MSToolkit.ps1",
    "MSToolkit.ico",
    "Investigate-AccountLockout.ps1",
    "IntuneTools.ps1",
    "Launcher.ps1",
    "Launch-MSToolkit.bat",
    "Launch-InstallMSToolkit.bat",
    "Launch-NewADUser.bat",
    "Launch-UninstallMSToolkit.bat",
    "M365-Distribution-Group-Compare.ps1",
    "M365-Conditional-Access-User-Manager.ps1",
    "M365-Exchange-Online-Tools.ps1",
    "M365-Group-Compare.ps1",
    "M365-OneDrive-SharePoint-Tools.ps1",
    "M365-Teams-Block-Number.ps1",
    "Launch-M365-Teams-Block-Number.cmd",
    "NewADUser.ps1",
    "InstallMSToolkit.ps1",
    "UninstallMSToolkit.ps1"
)

$RuntimeFiles = @(
    "Compare-UserGroups.ps1",
    "MSToolkit.ps1",
    "MSToolkit.ico",
    "Investigate-AccountLockout.ps1",
    "IntuneTools.ps1",
    "Launch-MSToolkit.bat",
    "Launch-NewADUser.bat",
    "M365-Distribution-Group-Compare.ps1",
    "M365-Conditional-Access-User-Manager.ps1",
    "M365-Exchange-Online-Tools.ps1",
    "M365-Group-Compare.ps1",
    "M365-OneDrive-SharePoint-Tools.ps1",
    "M365-Teams-Block-Number.ps1",
    "Launch-M365-Teams-Block-Number.cmd",
    "NewADUser.ps1"
)

# Validate the complete extracted installer payload BEFORE changing either
# installation directory. This prevents a missing archive file from leaving
# a partially updated MSToolkit installation.
$MissingFiles = New-Object System.Collections.Generic.List[string]

foreach ($File in $PayloadFiles) {
    $SourcePath = Join-Path $PSScriptRoot $File
    if (-not (Test-Path -LiteralPath $SourcePath)) {
        $MissingFiles.Add($File)
    }
}

if ($MissingFiles.Count -gt 0) {
    $MissingText = ($MissingFiles.ToArray() -join ", ")
    throw "Installer payload is incomplete. Missing required file(s): $MissingText"
}

# Ensure the Windows RSAT components required by MSToolkit are available before
# installing the application. Already-installed capabilities are left alone.
$RequiredRsatCapabilities = @(
    [pscustomobject]@{
        Name        = "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
        DisplayName = "Active Directory Domain Services and LDS Tools"
    },
    [pscustomobject]@{
        Name        = "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0"
        DisplayName = "Group Policy Management Tools"
    },
    [pscustomobject]@{
        Name        = "Print.Management.Console~~~~0.0.1.0"
        DisplayName = "Print Management Console"
    },
    [pscustomobject]@{
        Name        = "Rsat.Dns.Tools~~~~0.0.1.0"
        DisplayName = "DNS Server Tools"
    }
)

function Ensure-RsatCapability {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$DisplayName
    )

    Write-Host ""
    Write-Host "Checking RSAT prerequisite: $DisplayName"

    try {
        $Capability = Get-WindowsCapability `
            -Online `
            -Name $Name `
            -ErrorAction Stop
    }
    catch {
        throw "Unable to check the Windows capability '$DisplayName' ($Name). $($_.Exception.Message)"
    }

    if (-not $Capability) {
        throw "Windows did not return capability information for '$DisplayName' ($Name)."
    }

    if ($Capability.State -eq "Installed") {
        Write-Host "  Already installed - skipping." -ForegroundColor Green
        return
    }

    Write-Host "  Missing - installing..." -ForegroundColor Yellow

    try {
        $InstallResult = Add-WindowsCapability `
            -Online `
            -Name $Name `
            -ErrorAction Stop

        $VerifiedCapability = Get-WindowsCapability `
            -Online `
            -Name $Name `
            -ErrorAction Stop

        if ($VerifiedCapability.State -ne "Installed") {
            $RestartNote = if ($InstallResult.RestartNeeded) {
                " A restart may be required."
            }
            else {
                ""
            }

            throw "Windows reported state '$($VerifiedCapability.State)' after installation.$RestartNote"
        }

        Write-Host "  Installed successfully." -ForegroundColor Green

        if ($InstallResult.RestartNeeded) {
            Write-Host "  Windows reports that a restart is required." -ForegroundColor Yellow
            $script:RsatRestartRequired = $true
        }
    }
    catch {
        throw "Failed to install required RSAT capability '$DisplayName' ($Name). $($_.Exception.Message)"
    }
}

$script:RsatRestartRequired = $false

foreach ($RsatCapability in $RequiredRsatCapabilities) {
    Ensure-RsatCapability `
        -Name $RsatCapability.Name `
        -DisplayName $RsatCapability.DisplayName
}

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
New-Item -ItemType Directory -Path $RuntimeDir -Force | Out-Null

# MSToolkit stages runtime files here during normal (non-elevated) launches.
# Grant BUILTIN\Users Modify with file/folder inheritance so the launcher can
# replace existing runtime files without requiring Run as administrator.
$IcaclsExe = Join-Path $env:SystemRoot "System32\icacls.exe"
& $IcaclsExe $RuntimeDir /grant '*S-1-5-32-545:(OI)(CI)M' /C | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Failed to set required permissions on $RuntimeDir. icacls exit code: $LASTEXITCODE"
}

foreach ($File in $PayloadFiles) {
    $SourcePath = Join-Path $PSScriptRoot $File
    Copy-Item -LiteralPath $SourcePath -Destination (Join-Path $InstallDir $File) -Force
}

foreach ($File in $RuntimeFiles) {
    $SourcePath = Join-Path $PSScriptRoot $File
    Copy-Item -LiteralPath $SourcePath -Destination (Join-Path $RuntimeDir $File) -Force
}

# Repair permissions on existing files as well as newly copied files.
& $IcaclsExe $RuntimeDir /grant '*S-1-5-32-545:(OI)(CI)M' /T /C | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Failed to apply required permissions to runtime files in $RuntimeDir. icacls exit code: $LASTEXITCODE"
}

$LaunchTarget = Join-Path $InstallDir "MSToolkit.exe"
$IconPath = Join-Path $InstallDir "MSToolkit.ico"

if (Test-Path -LiteralPath $ShortcutPath) {
    Remove-Item -LiteralPath $ShortcutPath -Force -ErrorAction SilentlyContinue
}

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $LaunchTarget
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.IconLocation = $IconPath
$Shortcut.Save()

# Apps & Features / Programs and Features entry. It runs the same UninstallMSToolkit.ps1
# the uninstall package uses (removing the program folder, runtime folder and desktop
# shortcut); that script asks for elevation itself and removes this entry when done.
$UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSToolkit"
$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$UninstallScript = Join-Path $InstallDir "UninstallMSToolkit.ps1"
$InstalledBytes = (Get-ChildItem -LiteralPath $InstallDir -Recurse -File | Measure-Object -Property Length -Sum).Sum
$EstimatedSizeKB = [int][math]::Ceiling([double]$InstalledBytes / 1KB)

New-Item -Path $UninstallKey -Force | Out-Null

$StringValues = [ordered]@{
    DisplayName     = "MSToolkit"
    DisplayIcon     = $IconPath
    InstallLocation = $InstallDir
    InstallDate     = (Get-Date -Format "yyyyMMdd")
    UninstallString = "`"$PowerShellExe`" -NoProfile -ExecutionPolicy Bypass -File `"$UninstallScript`" -ShowResult"
    # Used by management tools that uninstall silently: no result message box.
    QuietUninstallString = "`"$PowerShellExe`" -NoProfile -ExecutionPolicy Bypass -File `"$UninstallScript`""
}
foreach ($Name in $StringValues.Keys) {
    New-ItemProperty -Path $UninstallKey -Name $Name -Value $StringValues[$Name] -PropertyType String -Force | Out-Null
}

$DwordValues = [ordered]@{
    NoModify      = 1
    NoRepair      = 1
    EstimatedSize = $EstimatedSizeKB
}
foreach ($Name in $DwordValues.Keys) {
    New-ItemProperty -Path $UninstallKey -Name $Name -Value $DwordValues[$Name] -PropertyType DWord -Force | Out-Null
}

Write-Host ""
Write-Host "MSToolkit installed to: $InstallDir"
Write-Host "MSToolkit runtime files updated in: $RuntimeDir"
Write-Host "Shortcut created: $ShortcutPath"
Write-Host "Apps & Features entry created: MSToolkit"

if ($script:RsatRestartRequired) {
    Write-Host ""
    Write-Host "One or more RSAT components reported that a Windows restart is required." -ForegroundColor Yellow
}

Exit 0
