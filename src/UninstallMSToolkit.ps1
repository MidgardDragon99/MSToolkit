param(
    # Set by the Apps & Features entry. That launch has no console window to read, so
    # the result is shown in a message box instead.
    [switch]$ShowResult
)

$ErrorActionPreference = "SilentlyContinue"

$ProgramFilesX86 = ${env:ProgramFiles(x86)}
if ([string]::IsNullOrWhiteSpace($ProgramFilesX86)) {
    $ProgramFilesX86 = $env:ProgramFiles
}

$InstallDir = Join-Path $ProgramFilesX86 "MSToolkit"
$RuntimeDir = "C:\ProgramData\MSToolkit"
$ShortcutPath = Join-Path $env:PUBLIC "Desktop\MSToolkit.lnk"
$UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSToolkit"
$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$ResultSwitch = if ($ShowResult) { " -ShowResult" } else { "" }

function Show-UninstallResult {
    param([string]$Message, [bool]$Failed)

    if (-not $ShowResult) { return }

    # Never block a silent run: a management tool (RMM, software inventory) may start
    # the Apps & Features command as SYSTEM, where nobody could click OK.
    if (-not [Environment]::UserInteractive) { return }
    if ([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem) { return }

    Add-Type -AssemblyName System.Windows.Forms
    $Icon = if ($Failed) { "Error" } else { "Information" }
    [System.Windows.Forms.MessageBox]::Show($Message, "Uninstall MSToolkit", "OK", $Icon) | Out-Null
}

# Removing Program Files and ProgramData needs elevation. Apps & Features starts this
# without it, so ask for it here and continue in the elevated copy.
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Start-Process -FilePath $PowerShellExe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"$ResultSwitch" `
        -Verb RunAs
    Exit 0
}

# The Apps & Features entry and Launch-UninstallMSToolkit.bat both run this script from
# the install folder, which would keep that folder in use while it is being removed.
# Carry on from a copy in TEMP instead; that copy deletes itself when it finishes.
$RunningFrom = (Split-Path -Parent $PSCommandPath).TrimEnd('\')
$InsideRemovedFolder = $false
foreach ($Dir in @($InstallDir, $RuntimeDir)) {
    if (($RunningFrom -ieq $Dir) -or $RunningFrom.StartsWith("$Dir\", [System.StringComparison]::OrdinalIgnoreCase)) {
        $InsideRemovedFolder = $true
    }
}

if ($InsideRemovedFolder) {
    $TempCopy = Join-Path $env:TEMP ("MSToolkit-Uninstall-{0}.ps1" -f [guid]::NewGuid().ToString("N"))
    Copy-Item -LiteralPath $PSCommandPath -Destination $TempCopy -Force

    if (Test-Path -LiteralPath $TempCopy) {
        Start-Process -FilePath $PowerShellExe `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$TempCopy`"$ResultSwitch" `
            -WorkingDirectory $env:TEMP
        Exit 0
    }
}

Set-Location -LiteralPath $env:TEMP

# Only processes started from MSToolkit's own folders are stopped. The child tools
# share file names with other toolkits built from the same scripts (NewADUser.ps1
# and so on), so matching on file name alone could close another installation's
# windows when both are installed side by side.
# This uninstaller and the process that started it are skipped.
$PathPatterns = @(
    "$InstallDir\",
    "$RuntimeDir\"
)

$SkipProcessIds = @($PID)
$ParentProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $PID" -ErrorAction SilentlyContinue
if ($ParentProcess) { $SkipProcessIds += [int]$ParentProcess.ParentProcessId }

Get-Process -Name "MSToolkit" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object {
        if ($SkipProcessIds -contains [int]$_.ProcessId) { return $false }

        $CommandLine = $_.CommandLine
        if ([string]::IsNullOrWhiteSpace($CommandLine)) { return $false }

        foreach ($Pattern in $PathPatterns) {
            if ($CommandLine.IndexOf($Pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
        }

        return $false
    } |
    ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }

Start-Sleep -Seconds 2

if (Test-Path -LiteralPath $ShortcutPath) {
    Remove-Item -LiteralPath $ShortcutPath -Force -ErrorAction SilentlyContinue
}

if (Test-Path -LiteralPath $InstallDir) {
    Remove-Item -LiteralPath $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
}

if (Test-Path -LiteralPath $RuntimeDir) {
    Remove-Item -LiteralPath $RuntimeDir -Recurse -Force -ErrorAction SilentlyContinue
}

$FailedPaths = @()
foreach ($Path in @($ShortcutPath, $InstallDir, $RuntimeDir)) {
    if (Test-Path -LiteralPath $Path) {
        $FailedPaths += $Path
    }
}

# A temporary copy made above removes itself; the script has already been read in full.
if ((Split-Path -Leaf $PSCommandPath) -like "MSToolkit-Uninstall-*.ps1") {
    Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
}

if ($FailedPaths.Count -gt 0) {
    # The Apps & Features entry is kept, so the uninstall can be run again.
    Write-Host "Failed to remove:"
    $FailedPaths | ForEach-Object { Write-Host $_ }
    Show-UninstallResult -Message ("MSToolkit could not remove:`r`n`r`n" + ($FailedPaths -join "`r`n") + "`r`n`r`nClose any open MSToolkit windows and run the uninstall again.") -Failed $true
    Exit 1
}

# Everything is gone, so remove the Apps & Features entry last.
if (Test-Path -LiteralPath $UninstallKey) {
    Remove-Item -LiteralPath $UninstallKey -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "MSToolkit uninstall completed."
Show-UninstallResult -Message "MSToolkit was uninstalled: the program folder, runtime folder and desktop shortcut were removed." -Failed $false
Exit 0
