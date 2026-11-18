$ErrorActionPreference = "Stop"

# Starts MSToolkit from its install folder. Only administrators can change that
# folder, so nothing a standard user writes can end up running as the admin
# account MSToolkit is started with. (Earlier versions copied the scripts into a
# user-writable ProgramData folder first; that copy is no longer made.)

$ProgramFilesX86 = ${env:ProgramFiles(x86)}
if ([string]::IsNullOrWhiteSpace($ProgramFilesX86)) {
    $ProgramFilesX86 = $env:ProgramFiles
}

$InstallDir = Join-Path $ProgramFilesX86 "MSToolkit"
$LaunchBat = Join-Path $InstallDir "Launch-MSToolkit.bat"

if (-not (Test-Path -LiteralPath $LaunchBat)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "MSToolkit is not installed correctly: $LaunchBat was not found.`r`n`r`nRun MSToolkit-Install.exe again.",
        "MSToolkit", "OK", "Error") | Out-Null
    exit 1
}

Start-Process -FilePath "cmd.exe" `
    -ArgumentList "/c `"$LaunchBat`"" `
    -WorkingDirectory $InstallDir
