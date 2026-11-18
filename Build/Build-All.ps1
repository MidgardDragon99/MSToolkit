#Requires -Version 5.1
<#
.SYNOPSIS
    Builds MSToolkit: the launcher EXE, the 7-Zip SFX installers and the .intunewin.

.DESCRIPTION
    1. Launcher - compiles Build\Launchers\MSToolkitLauncher.cs to MSToolkit.EXE
       with the csc.exe that ships with Windows (.NET Framework 4.x).
    2. SFX      - runs Build-Install-Uninstall-Only.cmd unchanged, which makes
       MSToolkit-Install.exe and MSToolkit-Uninstall.exe.
    3. IntuneWin - stages the package folder in %TEMP% and runs
       IntuneWinAppUtil.exe with install.cmd as the setup file, which makes
       install.intunewin in the repo root.
    Prints OK / FAILED / SKIPPED per step and exits 1 if any step failed.

    Everything is written to the repo root, where .gitignore keeps it out of git.

.PARAMETER SkipLauncher
    Keep the existing MSToolkit.EXE instead of compiling it.

.PARAMETER SkipIntuneWin
    Don't build install.intunewin.

.PARAMETER IntuneWinAppUtil
    Path to IntuneWinAppUtil.exe. Default: IntuneWinAppUtil.exe in the repo root.
#>
[CmdletBinding()]
param(
    [switch]$SkipLauncher,
    [switch]$SkipIntuneWin,
    [string]$IntuneWinAppUtil
)

$ErrorActionPreference = "Stop"

$RepoRoot     = Split-Path -Parent $PSScriptRoot
$LauncherSrc  = Join-Path $PSScriptRoot "Launchers\MSToolkitLauncher.cs"
$LauncherExe  = Join-Path $RepoRoot "MSToolkit.EXE"
$IconFile     = Join-Path $RepoRoot "MSToolkit.ico"
$SfxBuilder   = Join-Path $RepoRoot "Build-Install-Uninstall-Only.cmd"
$SetupFile    = "install.cmd"
$IntuneOutput = Join-Path $RepoRoot "install.intunewin"
$StagingDir   = Join-Path $env:TEMP "MSToolkitBuild"
$DownloadHint = "Download IntuneWinAppUtil.exe from Microsoft's Microsoft-Win32-Content-Prep-Tool repository on GitHub, then put it in the repo root or pass -IntuneWinAppUtil <path>."

# Left out of the .intunewin staging copy: build inputs and outputs that don't
# belong in the package, and repository-only files.
$StagingExclude = @(
    "*.7z",
    "*.intunewin",
    "*.lnk",
    "*.log",
    "Build-Install-Uninstall-Only.cmd",
    "IntuneWinAppUtil.exe",
    "CLAUDE.md",
    "MSToolkit-Project-Setup-Brief.md",
    ".gitignore",
    ".gitattributes"
)
$StagingExcludeDirs = @(".git", "Build", "docs", "Logs")

$Results = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param([string]$Step, [string]$Status, [string]$Detail)
    $Results.Add([pscustomobject]@{ Step = $Step; Status = $Status; Detail = $Detail })
    $Colour = switch ($Status) { "OK" { "Green" } "SKIPPED" { "Yellow" } default { "Red" } }
    Write-Host ("  {0}: {1}" -f $Status, $Detail) -ForegroundColor $Colour
}

# ---------------------------------------------------------------------------
# 1. Launcher EXE
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[1/3] Launcher EXE"

if ($SkipLauncher) {
    if (Test-Path -LiteralPath $LauncherExe) {
        Add-Result "Launcher" "SKIPPED" "Kept the existing MSToolkit.EXE"
    }
    else {
        Add-Result "Launcher" "FAILED" "-SkipLauncher was given but $LauncherExe does not exist"
    }
}
else {
    try {
        $Csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
        if (-not (Test-Path -LiteralPath $Csc)) {
            $Csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
        }
        if (-not (Test-Path -LiteralPath $Csc)) { throw "csc.exe (.NET Framework 4.x) was not found under $env:WINDIR\Microsoft.NET." }
        if (-not (Test-Path -LiteralPath $LauncherSrc)) { throw "Launcher source not found: $LauncherSrc" }
        if (-not (Test-Path -LiteralPath $IconFile)) { throw "Icon not found: $IconFile" }

        # Compile into a temporary folder first, so a failed build never leaves an
        # old MSToolkit.EXE looking like a new one. The file keeps its real name,
        # because csc writes it into the EXE's assembly name and version info.
        $TempDir = Join-Path $env:TEMP ("MSToolkitLauncher-{0}" -f [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
        $TempExe = Join-Path $TempDir "MSToolkit.EXE"
        $CscArgs = @(
            "/nologo", "/target:winexe", "/platform:anycpu", "/optimize+",
            "/out:$TempExe",
            "/win32icon:$IconFile",
            "/reference:System.Windows.Forms.dll",
            $LauncherSrc
        )
        try {
            & $Csc @CscArgs
            $CscExit = $LASTEXITCODE

            if ($CscExit -ne 0) { throw "csc.exe exited with code $CscExit." }
            if (-not (Test-Path -LiteralPath $TempExe)) { throw "csc.exe reported success but produced no EXE." }

            Copy-Item -LiteralPath $TempExe -Destination $LauncherExe -Force
        }
        finally {
            Remove-Item -LiteralPath $TempDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        if (-not (Test-Path -LiteralPath $LauncherExe)) { throw "MSToolkit.EXE is missing after the copy." }
        Add-Result "Launcher" "OK" "Compiled $LauncherExe"
    }
    catch {
        Add-Result "Launcher" "FAILED" $_.Exception.Message
    }
}

# ---------------------------------------------------------------------------
# 2. 7-Zip SFX installers
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[2/3] 7-Zip SFX installers"

try {
    if (-not (Test-Path -LiteralPath $SfxBuilder)) { throw "Builder not found: $SfxBuilder" }

    # An empty stdin lets the builder's closing 'pause' fall through.
    $EmptyInput = Join-Path $env:TEMP ("MSToolkit-empty-{0}.txt" -f [guid]::NewGuid().ToString("N"))
    [System.IO.File]::WriteAllText($EmptyInput, "")

    try {
        $Process = Start-Process -FilePath $SfxBuilder -WorkingDirectory $RepoRoot `
            -RedirectStandardInput $EmptyInput -NoNewWindow -PassThru
        # Reading Handle before WaitForExit makes ExitCode available afterwards.
        $null = $Process.Handle
        $Process.WaitForExit()
        $SfxExit = $Process.ExitCode
    }
    finally {
        Remove-Item -LiteralPath $EmptyInput -Force -ErrorAction SilentlyContinue
    }

    if ($SfxExit -ne 0) { throw "Build-Install-Uninstall-Only.cmd exited with code $SfxExit." }
    foreach ($Out in @("MSToolkit-Install.exe", "MSToolkit-Uninstall.exe")) {
        if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot $Out))) { throw "$Out was not created." }
    }
    Add-Result "SFX" "OK" "Built MSToolkit-Install.exe and MSToolkit-Uninstall.exe"
}
catch {
    Add-Result "SFX" "FAILED" $_.Exception.Message
}

# ---------------------------------------------------------------------------
# 3. .intunewin
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[3/3] install.intunewin"

if ($SkipIntuneWin) {
    Add-Result "IntuneWin" "SKIPPED" "-SkipIntuneWin was given"
}
else {
    try {
        if ([string]::IsNullOrWhiteSpace($IntuneWinAppUtil)) {
            $IntuneWinAppUtil = Join-Path $RepoRoot "IntuneWinAppUtil.exe"
        }
        if (-not (Test-Path -LiteralPath $IntuneWinAppUtil)) {
            throw "IntuneWinAppUtil.exe not found at $IntuneWinAppUtil. $DownloadHint"
        }

        if (Test-Path -LiteralPath $StagingDir) {
            Remove-Item -LiteralPath $StagingDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null

        Get-ChildItem -LiteralPath $RepoRoot -Force | Where-Object {
            $Item = $_
            if ($Item.PSIsContainer) { return ($StagingExcludeDirs -notcontains $Item.Name) }
            foreach ($Pattern in $StagingExclude) {
                if ($Item.Name -like $Pattern) { return $false }
            }
            return $true
        } | ForEach-Object {
            Copy-Item -LiteralPath $_.FullName -Destination $StagingDir -Recurse -Force
        }

        if (-not (Test-Path -LiteralPath (Join-Path $StagingDir $SetupFile))) {
            throw "$SetupFile is missing from the staging folder."
        }

        Write-Host "  Staged in ${StagingDir}:"
        Get-ChildItem -LiteralPath $StagingDir -Recurse -File | ForEach-Object {
            Write-Host ("    " + $_.FullName.Substring($StagingDir.Length + 1))
        }

        $StartTime = Get-Date
        & $IntuneWinAppUtil -c $StagingDir -s $SetupFile -o $RepoRoot -q
        $IntuneExit = $LASTEXITCODE

        if ($IntuneExit -ne 0) { throw "IntuneWinAppUtil.exe exited with code $IntuneExit." }
        if (-not (Test-Path -LiteralPath $IntuneOutput)) { throw "install.intunewin was not created." }
        if ((Get-Item -LiteralPath $IntuneOutput).LastWriteTime -lt $StartTime) {
            throw "install.intunewin was not updated by this build."
        }
        Add-Result "IntuneWin" "OK" "Built $IntuneOutput"
    }
    catch {
        Add-Result "IntuneWin" "FAILED" $_.Exception.Message
    }
    finally {
        if (Test-Path -LiteralPath $StagingDir) {
            Remove-Item -LiteralPath $StagingDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Summary"
$Results.ToArray() | Format-Table -AutoSize Step, Status, Detail | Out-String -Width 4096 | Write-Host

$Failed = @($Results.ToArray() | Where-Object { $_.Status -eq "FAILED" })
if ($Failed.Count -gt 0) {
    Write-Host "Build FAILED." -ForegroundColor Red
    exit 1
}

Write-Host "Build completed." -ForegroundColor Green
exit 0
