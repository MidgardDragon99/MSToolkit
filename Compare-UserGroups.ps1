<#
.SYNOPSIS
    Windows Forms AD group comparison and add tool.

.DESCRIPTION
    Compares group memberships between a reference user and a target user.
    Shows groups the reference user has that the target user does not have,
    and allows selected memberships to be added to the target user.

    Designed to be launched from MSToolkit.ps1 and use the same selected
    domain controller.
#>

[CmdletBinding()]
param(
    [string]$Server,

    [ValidateSet("Light","Dark")]
    [string]$ThemeMode
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# ---------------------------------------------------------------------------
# Shared MSToolkit theme support.
# When launched by MSToolkit.ps1, -ThemeMode is passed explicitly so this
# window uses the main window's CURRENT theme even when it runs as another
# Windows account. When launched standalone, the saved MSToolkit preference is
# used when available; otherwise Light is the default.
# ---------------------------------------------------------------------------
function Save-MSToolkitThemePreference {
    param(
        [string]$ToolKey,
        [string]$Theme
    )

    # Remembered per tool, so toggling here never changes the theme MSToolkit
    # itself starts in. Every other setting in the file is preserved.
    try {
        $SettingsRoot = $env:APPDATA
        if ([string]::IsNullOrWhiteSpace($SettingsRoot)) {
            $SettingsRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
        }

        $SettingsDir = Join-Path $SettingsRoot "MSToolkit"
        $SettingsPath = Join-Path $SettingsDir "settings.json"

        if (-not (Test-Path -LiteralPath $SettingsDir)) {
            New-Item -ItemType Directory -Path $SettingsDir -Force | Out-Null
        }

        $Settings = $null
        if (Test-Path -LiteralPath $SettingsPath) {
            try {
                $Settings = Get-Content -LiteralPath $SettingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            }
            catch {
                $Settings = $null
            }
        }

        if ($null -eq $Settings) {
            $Settings = [pscustomobject]@{}
        }

        $PropertyName = "Theme_$ToolKey"

        if ($Settings.PSObject.Properties.Name -contains $PropertyName) {
            $Settings.$PropertyName = $Theme
        }
        else {
            $Settings | Add-Member -MemberType NoteProperty -Name $PropertyName -Value $Theme
        }

        $Settings | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $SettingsPath -Encoding UTF8
    }
    catch {
        # A theme preference is not worth failing the tool over.
    }
}

function Set-MSToolkitThemeToggleFace {
    param([System.Windows.Forms.Button]$Button)

    if (-not $Button) { return }

    if ($script:MSToolkitThemeMode -eq "Dark") {
        $Button.Text = [string][char]0x263C
        $Button.ForeColor = [System.Drawing.Color]::FromArgb(255,214,102)
    }
    else {
        $Button.Text = [string][char]0x263E
        $Button.ForeColor = [System.Drawing.Color]::FromArgb(226,234,245)
    }
}

function Invoke-MSToolkitThemeToggle {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.Button]$Button,
        [string]$ToolKey
    )

    if ($script:MSToolkitThemeMode -eq "Dark") {
        $script:MSToolkitThemeMode = "Light"
    }
    else {
        $script:MSToolkitThemeMode = "Dark"
    }

    Save-MSToolkitThemePreference -ToolKey $ToolKey -Theme $script:MSToolkitThemeMode
    Apply-MSToolkitSharedTheme -Root $Form
    Register-MSToolkitComboFiltersOn -Root $Form
    Set-MSToolkitThemeToggleFace -Button $Button

    # Some labels carry a live status colour that the shared theme pass cannot
    # infer from their design-time colour. Let the script restore those.
    if ($script:MSToolkitThemeRefreshHook) {
        try { & $script:MSToolkitThemeRefreshHook } catch { }
    }

    $Form.Refresh()
}

function New-MSToolkitThemeToggleButton {
    param(
        [System.Windows.Forms.Control]$HeaderPanel,
        [System.Windows.Forms.Form]$Form,
        [string]$ToolKey
    )

    $Button = New-Object System.Windows.Forms.Button
    $Button.Size = New-Object System.Drawing.Size(44,32)

    # Sit to the left of anything already anchored to the right of the header,
    # such as the Connect button on the Microsoft 365 tools.
    $RightEdge = $HeaderPanel.ClientSize.Width - 16

    foreach ($Existing in $HeaderPanel.Controls) {
        if ($Existing -eq $Button) { continue }

        if (($Existing.Anchor -band [System.Windows.Forms.AnchorStyles]::Right) -eq [System.Windows.Forms.AnchorStyles]::Right) {
            if ($Existing.Left -lt $RightEdge) {
                $RightEdge = $Existing.Left
            }
        }
    }

    $Button.Location = New-Object System.Drawing.Point(($RightEdge - $Button.Width - 12),22)
    $Button.Anchor = "Top,Right"
    $Button.FlatStyle = "Flat"
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = [System.Drawing.Color]::FromArgb(24,47,74)
    $Button.Font = New-Object System.Drawing.Font("Segoe UI Symbol",15)
    $Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $Button.TabStop = $false
    $Button.Tag = "ThemeToggle"

    $Tip = New-Object System.Windows.Forms.ToolTip
    $Tip.SetToolTip($Button, "Switch between Light and Dark mode")

    $Button.Add_Click({
        Invoke-MSToolkitThemeToggle -Form $Form -Button $Button -ToolKey $ToolKey
    }.GetNewClosure())

    $HeaderPanel.Controls.Add($Button)
    $Button.BringToFront()

    Set-MSToolkitThemeToggleFace -Button $Button

    return $Button
}

function Resolve-MSToolkitThemeMode {
    param([string]$RequestedTheme)

    if ($RequestedTheme -in @("Light","Dark")) {
        return $RequestedTheme
    }

    try {
        $SettingsRoot = $env:APPDATA
        if ([string]::IsNullOrWhiteSpace($SettingsRoot)) {
            $SettingsRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
        }

        $SettingsPath = Join-Path (Join-Path $SettingsRoot "MSToolkit") "settings.json"
        if (Test-Path -LiteralPath $SettingsPath) {
            $Settings = Get-Content -LiteralPath $SettingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            # This tool's own remembered choice wins when it is launched on its own.
            $ToolProperty = "Theme_CompareUserGroups"
            if (($Settings.PSObject.Properties.Name -contains $ToolProperty) -and ($Settings.$ToolProperty -in @("Light","Dark"))) {
                return [string]$Settings.$ToolProperty
            }

            if ($Settings.Theme -in @("Light","Dark")) {
                return [string]$Settings.Theme
            }
        }
    }
    catch {
        # Theme preference is cosmetic only. Fall back to Light if it cannot be read.
    }

    return "Light"
}

$script:MSToolkitThemeMode = Resolve-MSToolkitThemeMode -RequestedTheme $ThemeMode

function Get-MSToolkitThemePalette {
    if ($script:MSToolkitThemeMode -eq "Dark") {
        return [pscustomobject]@{
            MainBackground       = [System.Drawing.Color]::FromArgb(30,32,36)
            PanelBackground      = [System.Drawing.Color]::FromArgb(38,41,46)
            InputBackground      = [System.Drawing.Color]::FromArgb(45,48,54)
            OutputBackground     = [System.Drawing.Color]::FromArgb(24,26,29)
            Text                 = [System.Drawing.Color]::FromArgb(232,234,237)
            MutedText            = [System.Drawing.Color]::FromArgb(174,180,187)
            Border               = [System.Drawing.Color]::FromArgb(78,84,92)
            TopBar               = [System.Drawing.Color]::FromArgb(24,47,74)
            HeaderMuted          = [System.Drawing.Color]::FromArgb(190,205,220)
            Section              = [System.Drawing.Color]::FromArgb(122,181,238)
            Accent               = [System.Drawing.Color]::FromArgb(0,120,215)
            ButtonBackground     = [System.Drawing.Color]::FromArgb(49,53,59)
            ButtonHover          = [System.Drawing.Color]::FromArgb(60,65,72)
            SecondaryButton      = [System.Drawing.Color]::FromArgb(70,78,88)
            SelectionBackground  = [System.Drawing.Color]::FromArgb(55,105,155)
            Success              = [System.Drawing.Color]::FromArgb(118,210,142)
            Danger               = [System.Drawing.Color]::FromArgb(255,125,125)
            Warning              = [System.Drawing.Color]::FromArgb(255,184,92)
            Info                 = [System.Drawing.Color]::FromArgb(125,190,245)
            Separator            = [System.Drawing.Color]::FromArgb(125,132,140)
            SuccessButton        = [System.Drawing.Color]::FromArgb(30,112,74)
            DangerButton         = [System.Drawing.Color]::FromArgb(155,67,64)
            DisabledBackground   = [System.Drawing.Color]::FromArgb(52,55,60)
            DisabledText         = [System.Drawing.Color]::FromArgb(165,170,177)
            DisabledBorder       = [System.Drawing.Color]::FromArgb(92,97,105)
            SuccessBackground    = [System.Drawing.Color]::FromArgb(38,65,48)
            DangerBackground     = [System.Drawing.Color]::FromArgb(72,42,44)
            WarningBackground    = [System.Drawing.Color]::FromArgb(75,61,38)
            InfoBackground       = [System.Drawing.Color]::FromArgb(38,55,72)
        }
    }

    return [pscustomobject]@{
        MainBackground       = [System.Drawing.Color]::FromArgb(245,247,250)
        PanelBackground      = [System.Drawing.Color]::White
        InputBackground      = [System.Drawing.Color]::White
        OutputBackground     = [System.Drawing.Color]::White
        Text                 = [System.Drawing.Color]::FromArgb(35,35,35)
        MutedText            = [System.Drawing.Color]::DimGray
        Border               = [System.Drawing.Color]::FromArgb(210,215,220)
        TopBar               = [System.Drawing.Color]::FromArgb(31,58,93)
        HeaderMuted          = [System.Drawing.Color]::FromArgb(218,228,240)
        Section              = [System.Drawing.Color]::FromArgb(31,58,93)
        Accent               = [System.Drawing.Color]::FromArgb(0,120,215)
        ButtonBackground     = [System.Drawing.Color]::White
        ButtonHover          = [System.Drawing.Color]::FromArgb(242,246,250)
        SecondaryButton      = [System.Drawing.Color]::FromArgb(70,90,110)
        SelectionBackground  = [System.Drawing.Color]::FromArgb(0,120,215)
        Success              = [System.Drawing.Color]::ForestGreen
        Danger               = [System.Drawing.Color]::Firebrick
        Warning              = [System.Drawing.Color]::DarkOrange
        Info                 = [System.Drawing.Color]::FromArgb(35,90,145)
        Separator            = [System.Drawing.Color]::FromArgb(140,140,140)
        SuccessButton        = [System.Drawing.Color]::FromArgb(26,137,85)
        DangerButton         = [System.Drawing.Color]::Firebrick
        DisabledBackground   = [System.Drawing.Color]::FromArgb(235,235,235)
        DisabledText         = [System.Drawing.Color]::FromArgb(125,130,136)
        DisabledBorder       = [System.Drawing.Color]::FromArgb(195,200,206)
        SuccessBackground    = [System.Drawing.Color]::FromArgb(235,248,238)
        DangerBackground     = [System.Drawing.Color]::FromArgb(255,235,235)
        WarningBackground    = [System.Drawing.Color]::FromArgb(255,248,225)
        InfoBackground       = [System.Drawing.Color]::FromArgb(235,245,255)
    }
}


function Convert-MSToolkitThemeColor {
    param([System.Drawing.Color]$Color)

    if ($script:MSToolkitThemeMode -ne "Dark") {
        return $Color
    }

    $Palette = Get-MSToolkitThemePalette

    if (
        $Color.ToArgb() -eq ([System.Drawing.Color]::Black).ToArgb() -or
        (Test-MSToolkitThemeColor $Color 35 35 35) -or
        (Test-MSToolkitThemeColor $Color 35 40 45) -or
        (Test-MSToolkitThemeColor $Color 45 45 45) -or
        (Test-MSToolkitThemeColor $Color 45 55 65)
    ) {
        return $Palette.Text
    }

    if (
        $Color.ToArgb() -eq ([System.Drawing.Color]::DimGray).ToArgb() -or
        (Test-MSToolkitThemeColor $Color 65 65 65) -or
        (Test-MSToolkitThemeColor $Color 70 70 70) -or
        (Test-MSToolkitThemeColor $Color 120 120 120)
    ) {
        return $Palette.MutedText
    }

    if (
        $Color.ToArgb() -eq ([System.Drawing.Color]::Red).ToArgb() -or
        $Color.ToArgb() -eq ([System.Drawing.Color]::Firebrick).ToArgb()
    ) {
        return $Palette.Danger
    }

    if (
        $Color.ToArgb() -eq ([System.Drawing.Color]::Green).ToArgb() -or
        $Color.ToArgb() -eq ([System.Drawing.Color]::DarkGreen).ToArgb() -or
        $Color.ToArgb() -eq ([System.Drawing.Color]::ForestGreen).ToArgb()
    ) {
        return $Palette.Success
    }

    if ($Color.ToArgb() -eq ([System.Drawing.Color]::DarkOrange).ToArgb()) {
        return $Palette.Warning
    }

    if (
        (Test-MSToolkitThemeColor $Color 31 58 93) -or
        (Test-MSToolkitThemeColor $Color 35 90 145) -or
        (Test-MSToolkitThemeColor $Color 45 75 105)
    ) {
        return $Palette.Info
    }

    if (
        (Test-MSToolkitThemeColor $Color 120 120 120) -or
        (Test-MSToolkitThemeColor $Color 140 140 140)
    ) {
        return $Palette.Separator
    }

    return $Color
}


function Test-MSToolkitThemeColor {
    param(
        [System.Drawing.Color]$Color,
        [int]$R,
        [int]$G,
        [int]$B
    )

    return ($Color.R -eq $R -and $Color.G -eq $G -and $Color.B -eq $B)
}

function Apply-MSToolkitThemeControl {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Control]$Control,
        [bool]$ParentIsHeader = $false
    )

    $Palette = Get-MSToolkitThemePalette

    # Preserve the control's design-time colors so semantic roles (success, danger,
    # warning, accent) can be restored after a disabled control becomes enabled.
    if (-not $Control.PSObject.Properties['MSToolkitThemeOriginalBackColor']) {
        $Control | Add-Member -NotePropertyName MSToolkitThemeOriginalBackColor -NotePropertyValue $Control.BackColor
    }
    if (-not $Control.PSObject.Properties['MSToolkitThemeOriginalForeColor']) {
        $Control | Add-Member -NotePropertyName MSToolkitThemeOriginalForeColor -NotePropertyValue $Control.ForeColor
    }

    $OriginalBack = [System.Drawing.Color]$Control.MSToolkitThemeOriginalBackColor
    $OriginalFore = [System.Drawing.Color]$Control.MSToolkitThemeOriginalForeColor
    $IsHeader = $false

    # The theme toggle paints its own sun/moon colours, so leave it alone.
    if ("$($Control.Tag)" -eq "ThemeToggle") {
        $Control.BackColor = $Palette.TopBar
        return
    }

    if (
        ($Control -is [System.Windows.Forms.Button] -or
         $Control -is [System.Windows.Forms.CheckBox] -or
         $Control -is [System.Windows.Forms.RadioButton]) -and
        -not $Control.PSObject.Properties['MSToolkitThemeEnabledHooked']
    ) {
        $Control | Add-Member -NotePropertyName MSToolkitThemeEnabledHooked -NotePropertyValue $true
        $Control.Add_EnabledChanged({
            param($sender,$eventArgs)
            Apply-MSToolkitThemeControl -Control $sender
        })
    }

    if ($Control -is [System.Windows.Forms.Form]) {
        $Control.BackColor = $Palette.MainBackground
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.Panel]) {
        $IsHeader = (Test-MSToolkitThemeColor $OriginalBack 31 58 93)

        if ($IsHeader) {
            $Control.BackColor = $Palette.TopBar
            $Control.ForeColor = [System.Drawing.Color]::White
        }
        else {
            $Control.BackColor = $Palette.PanelBackground
            $Control.ForeColor = $Palette.Text
        }
    }
    elseif ($Control -is [System.Windows.Forms.GroupBox]) {
        $Control.BackColor = $Palette.PanelBackground
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.TabPage]) {
        $Control.BackColor = $Palette.PanelBackground
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.TabControl]) {
        $Control.BackColor = $Palette.PanelBackground
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.RichTextBox]) {
        $Control.BackColor = $Palette.OutputBackground
        $Control.ForeColor = $Palette.Text
        $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    elseif ($Control -is [System.Windows.Forms.TextBox]) {
        $Control.BackColor = $Palette.InputBackground
        $Control.ForeColor = $Palette.Text
        $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    elseif ($Control -is [System.Windows.Forms.ComboBox]) {
        $Control.BackColor = $Palette.InputBackground
        $Control.ForeColor = $Palette.Text
    }
    elseif (
        $Control -is [System.Windows.Forms.ListBox] -or
        $Control -is [System.Windows.Forms.CheckedListBox] -or
        $Control -is [System.Windows.Forms.ListView]
    ) {
        $Control.BackColor = $Palette.InputBackground
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.DataGridView]) {
        $Control.BackgroundColor = $Palette.PanelBackground
        $Control.GridColor = $Palette.Border
        $Control.DefaultCellStyle.BackColor = $Palette.PanelBackground
        $Control.DefaultCellStyle.ForeColor = $Palette.Text
        $Control.DefaultCellStyle.SelectionBackColor = $Palette.SelectionBackground
        $Control.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
        $Control.AlternatingRowsDefaultCellStyle.BackColor = if ($script:MSToolkitThemeMode -eq "Dark") {
            [System.Drawing.Color]::FromArgb(42,45,50)
        }
        else {
            [System.Drawing.Color]::FromArgb(250,251,252)
        }
        $Control.AlternatingRowsDefaultCellStyle.ForeColor = $Palette.Text
        $Control.ColumnHeadersDefaultCellStyle.BackColor = $Palette.ButtonBackground
        $Control.ColumnHeadersDefaultCellStyle.ForeColor = $Palette.Text
        $Control.ColumnHeadersDefaultCellStyle.SelectionBackColor = $Palette.ButtonBackground
        $Control.ColumnHeadersDefaultCellStyle.SelectionForeColor = $Palette.Text
        $Control.RowHeadersDefaultCellStyle.BackColor = $Palette.ButtonBackground
        $Control.RowHeadersDefaultCellStyle.ForeColor = $Palette.Text
        $Control.EnableHeadersVisualStyles = $false
    }
    elseif ($Control -is [System.Windows.Forms.StatusStrip]) {
        $Control.BackColor = $Palette.PanelBackground
        $Control.ForeColor = $Palette.Text
        foreach ($Item in $Control.Items) {
            $Item.BackColor = $Palette.PanelBackground
            $Item.ForeColor = $Palette.Text
        }
    }
    elseif ($Control -is [System.Windows.Forms.Button]) {
        $Control.UseVisualStyleBackColor = $false
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $Control.FlatAppearance.BorderSize = 1
        $Control.FlatAppearance.BorderColor = $Palette.Border
        $Control.FlatAppearance.MouseOverBackColor = $Palette.ButtonHover

        if (
            (Test-MSToolkitThemeColor $OriginalBack 0 120 215) -or
            (Test-MSToolkitThemeColor $OriginalBack 31 58 93)
        ) {
            $Control.BackColor = $Palette.Accent
            $Control.ForeColor = [System.Drawing.Color]::White
            $Control.FlatAppearance.BorderSize = 0
        }
        elseif (
            (Test-MSToolkitThemeColor $OriginalBack 180 75 70) -or
            $OriginalBack.ToArgb() -eq ([System.Drawing.Color]::Firebrick).ToArgb()
        ) {
            $Control.BackColor = $Palette.DangerButton
            $Control.ForeColor = [System.Drawing.Color]::White
            $Control.FlatAppearance.BorderSize = 0
        }
        elseif (
            (Test-MSToolkitThemeColor $OriginalBack 26 137 85) -or
            (Test-MSToolkitThemeColor $OriginalBack 22 130 80)
        ) {
            $Control.BackColor = $Palette.SuccessButton
            $Control.ForeColor = [System.Drawing.Color]::White
            $Control.FlatAppearance.BorderSize = 0
        }
        elseif (Test-MSToolkitThemeColor $OriginalBack 70 90 110) {
            $Control.BackColor = $Palette.SecondaryButton
            $Control.ForeColor = [System.Drawing.Color]::White
            $Control.FlatAppearance.BorderSize = 0
        }
        elseif (
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::Firebrick).ToArgb() -or
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::Red).ToArgb()
        ) {
            $Control.BackColor = $Palette.ButtonBackground
            $Control.ForeColor = $Palette.Danger
        }
        elseif (
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::ForestGreen).ToArgb() -or
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::DarkGreen).ToArgb()
        ) {
            $Control.BackColor = $Palette.ButtonBackground
            $Control.ForeColor = $Palette.Success
        }
        else {
            $Control.BackColor = $Palette.ButtonBackground
            $Control.ForeColor = $Palette.Text
        }
    }
    elseif ($Control -is [System.Windows.Forms.Label]) {
        if ($ParentIsHeader) {
            if (Test-MSToolkitThemeColor $OriginalFore 218 228 240) {
                $Control.ForeColor = $Palette.HeaderMuted
            }
            else {
                $Control.ForeColor = [System.Drawing.Color]::White
            }
        }
        elseif (
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::DimGray).ToArgb() -or
            (Test-MSToolkitThemeColor $OriginalFore 70 70 70) -or
            (Test-MSToolkitThemeColor $OriginalFore 120 120 120)
        ) {
            $Control.ForeColor = $Palette.MutedText
        }
        elseif (
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::Firebrick).ToArgb() -or
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::Red).ToArgb()
        ) {
            $Control.ForeColor = $Palette.Danger
        }
        elseif (
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::ForestGreen).ToArgb() -or
            $OriginalFore.ToArgb() -eq ([System.Drawing.Color]::DarkGreen).ToArgb()
        ) {
            $Control.ForeColor = $Palette.Success
        }
        else {
            $Control.ForeColor = $Palette.Text
        }
    }
    elseif (
        $Control -is [System.Windows.Forms.CheckBox] -or
        $Control -is [System.Windows.Forms.RadioButton]
    ) {
        $Control.BackColor = $Palette.PanelBackground
        $Control.ForeColor = $Palette.Text
    }

    # Disabled controls stay legible in both themes and automatically return to
    # their original semantic styling when they are enabled later.
    if (-not $Control.Enabled) {
        if ($Control -is [System.Windows.Forms.Button]) {
            $Control.BackColor = $Palette.DisabledBackground
            $Control.ForeColor = $Palette.DisabledText
            $Control.FlatAppearance.BorderSize = 1
            $Control.FlatAppearance.BorderColor = $Palette.DisabledBorder
            $Control.FlatAppearance.MouseOverBackColor = $Palette.DisabledBackground
        }
        elseif (
            $Control -is [System.Windows.Forms.CheckBox] -or
            $Control -is [System.Windows.Forms.RadioButton]
        ) {
            $Control.ForeColor = $Palette.DisabledText
        }
    }

    foreach ($Child in $Control.Controls) {
        Apply-MSToolkitThemeControl -Control $Child -ParentIsHeader:$IsHeader
    }
}

function Apply-MSToolkitSharedTheme {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Root)

    $Root.SuspendLayout()
    try {
        Apply-MSToolkitThemeControl -Control $Root
        $Root.Invalidate($true)
        $Root.Refresh()
    }
    finally {
        $Root.ResumeLayout($true)
    }
}


# Hide the PowerShell console while keeping the WinForms window visible.
# This matches the behavior used by NewADUser and Investigate Account Lockout.
if (-not ("MSToolkitCompareConsole.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace MSToolkitCompareConsole
{
    public static class NativeMethods
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}
"@
}

try {
    $consoleHandle = [MSToolkitCompareConsole.NativeMethods]::GetConsoleWindow()
    if ($consoleHandle -ne [IntPtr]::Zero -and $host.Name -notlike "*ISE*") {
        [MSToolkitCompareConsole.NativeMethods]::ShowWindow($consoleHandle, 0) | Out-Null
    }
}
catch {
    # Console hiding must never prevent the GUI from loading.
}

[System.Windows.Forms.Application]::EnableVisualStyles()

# Do not create the AD: PowerShell drive. This tool never uses it, and building it
# means connecting to a domain controller before anything else can run.
$env:ADPS_LoadDefaultDrive = 0
Import-Module ActiveDirectory -ErrorAction Stop

# ---------------------------------------------------------------------------
# DC selection helpers. The AD PowerShell module talks to domain controllers over
# ADWS (TCP 9389). If automatic discovery lands on a DC whose 9389 does not answer
# from this computer, every call waits out two ~21 second connection timeouts.
# These helpers find a DC that answers instead, without relying on ADWS to do it.
# ---------------------------------------------------------------------------

function Test-MSToolkitAdwsPort {
    param(
        [string]$HostName,
        [int]$TimeoutMilliseconds = 2000
    )

    $Client = New-Object System.Net.Sockets.TcpClient
    try {
        $Attempt = $Client.BeginConnect($HostName, 9389, $null, $null)

        if (-not $Attempt.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)) {
            return $false
        }

        $Client.EndConnect($Attempt)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $Client.Close()
    }
}

function Get-MSToolkitLocalDomainName {
    # Windows already knows which domain this computer belongs to; no network needed.
    try {
        $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop

        if ($ComputerSystem.PartOfDomain -and -not [string]::IsNullOrWhiteSpace($ComputerSystem.Domain)) {
            return [string]$ComputerSystem.Domain
        }
    }
    catch { }

    return $null
}

function Get-MSToolkitDomainControllerHostNames {
    param([string]$DomainName)

    # Plain LDAP (port 389) through System.DirectoryServices - never touches ADWS.
    try {
        $Context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $DomainName)
        $DomainObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($Context)

        return @($DomainObject.DomainControllers | ForEach-Object { [string]$_.Name } | Sort-Object -Unique)
    }
    catch {
        return @()
    }
}

function Resolve-MSToolkitWorkingServer {
    param(
        [string]$Preferred,
        [string]$DomainName
    )

    # Returns the host name of a DC that answers on ADWS, trying the preferred DC
    # first, then the DC locator's choice, then every DC in the domain. Each is
    # checked with a 2 second port test, so an unreachable DC costs 2 seconds, not 42.
    $Candidates = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($Preferred)) {
        if (($Preferred -notlike '*.*') -and $DomainName) {
            $Candidates.Add("$Preferred.$DomainName")
        }
        else {
            $Candidates.Add($Preferred)
        }
    }

    if ($DomainName) {
        try {
            $SeedDC = Get-ADDomainController -Discover -DomainName $DomainName -Service ADWS -ErrorAction Stop
            $SeedHost = [string]($SeedDC.HostName | Select-Object -First 1)
            if (-not [string]::IsNullOrWhiteSpace($SeedHost)) {
                $Candidates.Add($SeedHost)
            }
        }
        catch { }

        foreach ($HostName in (Get-MSToolkitDomainControllerHostNames -DomainName $DomainName)) {
            $Candidates.Add($HostName)
        }
    }

    $Tried = @{}
    foreach ($Candidate in $Candidates) {
        $Key = $Candidate.ToLower()
        if ($Tried.ContainsKey($Key)) { continue }
        $Tried[$Key] = $true

        if (Test-MSToolkitAdwsPort -HostName $Candidate) {
            return $Candidate
        }
    }

    return $null
}

$script:ReferenceUser = $null
$script:TargetUser = $null
$script:ComparisonRows = @()
$script:CurrentServer = $Server

# Use the server MSToolkit passed in if it answers on ADWS. If it does not - or no server
# was passed because the tool was launched on its own - use a DC that does answer.
$script:ServerNote = $null
$ResolvedServer = Resolve-MSToolkitWorkingServer -Preferred $Server -DomainName (Get-MSToolkitLocalDomainName)
if ($ResolvedServer) {
    if ((-not [string]::IsNullOrWhiteSpace($Server)) -and ($ResolvedServer -ine $Server) -and ($ResolvedServer -notlike "$Server.*")) {
        $script:ServerNote = "The selected AD server $Server is not answering on ADWS (port 9389), so $ResolvedServer is being used instead."
    }
    $script:CurrentServer = $ResolvedServer
}
$script:HighRiskGroupPatterns = @(
    'Domain Admins',
    'Enterprise Admins',
    'Schema Admins',
    'Administrators',
    'Account Operators',
    'Server Operators',
    'Backup Operators',
    'Print Operators',
    'DnsAdmins',
    'Group Policy Creator Owners',
    'Organization Management',
    'Global Administrator',
    'Privileged',
    'Admin',
    'VPN',
    'Wire',
    'ACH',
    'Core',
    'Domain Local Admin'
)

function Get-MSToolkitSidString {
    param($DirectoryObject)

    if ($null -eq $DirectoryObject -or $null -eq $DirectoryObject.SID) {
        return ""
    }

    try {
        if ($DirectoryObject.SID.PSObject.Properties.Name -contains "Value") {
            return [string]$DirectoryObject.SID.Value
        }
    }
    catch {
    }

    return [string]$DirectoryObject.SID
}

function Get-MSToolkitRidFromSid {
    param([string]$Sid)

    if ([string]::IsNullOrWhiteSpace($Sid)) {
        return -1
    }

    if ($Sid -match '-(\d+)$') {
        return [int64]$Matches[1]
    }

    return -1
}

function Get-MSToolkitCriticalGroupReason {
    param($Group)

    $Sid = Get-MSToolkitSidString -DirectoryObject $Group

    $BuiltinReasons = @{
        'S-1-5-32-544' = 'Built-in Administrators group'
        'S-1-5-32-548' = 'Built-in Account Operators group'
        'S-1-5-32-549' = 'Built-in Server Operators group'
        'S-1-5-32-550' = 'Built-in Print Operators group'
        'S-1-5-32-551' = 'Built-in Backup Operators group'
        'S-1-5-32-552' = 'Built-in Replicator group'
    }

    if ($BuiltinReasons.ContainsKey($Sid)) {
        return $BuiltinReasons[$Sid]
    }

    $Rid = Get-MSToolkitRidFromSid -Sid $Sid
    $DomainRidReasons = @{
        '512' = 'Domain Admins group'
        '513' = 'Domain Users group'
        '514' = 'Domain Guests group'
        '515' = 'Domain Computers group'
        '516' = 'Domain Controllers group'
        '517' = 'Cert Publishers group'
        '518' = 'Schema Admins group'
        '519' = 'Enterprise Admins group'
        '520' = 'Group Policy Creator Owners group'
        '521' = 'Read-only Domain Controllers group'
        '522' = 'Cloneable Domain Controllers group'
        '526' = 'Key Admins group'
        '527' = 'Enterprise Key Admins group'
    }

    $RidKey = [string]$Rid
    if ($DomainRidReasons.ContainsKey($RidKey)) {
        return $DomainRidReasons[$RidKey]
    }

    if (
        ([string]$Group.Name -ieq 'DnsAdmins') -or
        ([string]$Group.SamAccountName -ieq 'DnsAdmins')
    ) {
        return 'DNSAdmins group'
    }

    if ($Group.isCriticalSystemObject -eq $true) {
        return "Active Directory marks this group as a critical system object"
    }

    return $null
}

function Get-MSToolkitUserPickerLabel {
    param($ADUser)

    $Display = if (-not [string]::IsNullOrWhiteSpace($ADUser.Name)) { $ADUser.Name } else { $ADUser.SamAccountName }
    return "$Display ($($ADUser.SamAccountName))"
}

function ConvertFrom-MSToolkitUserPickerLabel {
    param([string]$Value)

    $Text = "$Value".Trim()

    # Entries loaded into the picker look like "Jane Doe (jdoe)".
    # Anything typed by hand is passed straight through unchanged.
    if ($Text -match '\(([^()]+)\)\s*$') {
        return $Matches[1].Trim()
    }

    return $Text
}

# ---------------------------------------------------------------------------
# MSToolkit settings (read-only)
# Site values come from %APPDATA%\MSToolkit\settings.json, written by the MSToolkit
# Settings window. This tool runs as the same admin account as MSToolkit, so it
# reads the same file. Nothing organization-specific is stored in this script;
# blank or missing values fall back as described where each one is used.
# ---------------------------------------------------------------------------
$script:MSToolkitSettingsPath  = Join-Path (Join-Path $env:APPDATA "MSToolkit") "settings.json"
$script:MSToolkitSettingsCache = $null

function Get-MSToolkitSetting {
    param(
        [string]$Name,
        [string]$Default = ""
    )

    if ($null -eq $script:MSToolkitSettingsCache) {
        try {
            if (Test-Path -LiteralPath $script:MSToolkitSettingsPath) {
                $script:MSToolkitSettingsCache = Get-Content -LiteralPath $script:MSToolkitSettingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            }
        }
        catch {
            # A damaged settings file must not stop the tool; defaults apply.
        }

        if ($null -eq $script:MSToolkitSettingsCache) {
            $script:MSToolkitSettingsCache = [pscustomobject]@{}
        }
    }

    $Value = ""
    if ($script:MSToolkitSettingsCache.PSObject.Properties.Name -contains $Name) {
        $Value = [string]$script:MSToolkitSettingsCache.$Name
    }

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    return $Value.Trim()
}

function Get-MSToolkitEmployeeUserList {
    param([string]$Server)

    $ServerArgs = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $ServerArgs['Server'] = $Server
    }

    $DomainDN = (Get-ADDomain @ServerArgs -ErrorAction Stop).DistinguishedName

    # User accounts plus the admin accounts, which are often the reference user when
    # comparing group membership. Both OUs come from MSToolkit Settings: a blank Users
    # OU searches the whole domain, and a blank Admin accounts OU adds nothing extra.
    $UsersOU = Get-MSToolkitSetting -Name "OUUsers" -Default $DomainDN
    $SearchBases = @($UsersOU)

    $AdminsOU = Get-MSToolkitSetting -Name "OUAdmins"
    if ($AdminsOU -and ($AdminsOU -ne $UsersOU)) {
        $SearchBases += $AdminsOU
    }

    $Found = New-Object System.Collections.Generic.List[object]

    foreach ($SearchBase in $SearchBases) {
        try {
            $Users = @(
                Get-ADUser -Filter * -SearchBase $SearchBase -SearchScope Subtree @ServerArgs -Properties Name,SamAccountName -ErrorAction Stop
            )

            foreach ($User in $Users) {
                $Found.Add($User)
            }
        }
        catch {
            # A missing or unreadable OU should not stop the rest of the list loading.
        }
    }

    if ($Found.Count -eq 0) {
        throw "No user accounts could be read from the configured search bases."
    }

    return @($Found.ToArray() | Sort-Object Name -Unique)
}

function Get-MSToolkitUserListNote {
    # The note shown beside a user picker when an OU setting leaves a gap, or $null
    # when both are set. IsWarning selects the warning colour over the muted one.
    if (-not (Get-MSToolkitSetting -Name "OUUsers")) {
        return [pscustomobject]@{ Text = "No Users OU is set in MSToolkit Settings, so the user list covers the whole domain."; IsWarning = $true }
    }
    if (-not (Get-MSToolkitSetting -Name "OUAdmins")) {
        return [pscustomobject]@{ Text = "No Admin accounts OU is set in MSToolkit Settings: admin accounts outside the Users OU aren't listed (typing a name still works)."; IsWarning = $true }
    }
    return $null
}

# Filtered user pickers: see Register-MSToolkitComboFilter for why Windows
# autocomplete is not used here.
$script:ComboFilterItems = @{}
$script:ComboFilterRegistered = @{}
$script:ComboFilterBusy = $false
$script:ComboFilterSkipHide = $false
$script:ComboFilterLists = @{}
$script:ComboFilterOwners = @{}

function Hide-MSToolkitComboMatches {
    param([System.Windows.Forms.ComboBox]$Combo)

    if (-not $Combo) { return }

    $List = $script:ComboFilterLists[$Combo.GetHashCode()]
    if ($List) { $List.Visible = $false }
}

function Set-MSToolkitComboMatch {
    # Accept whatever is highlighted in the match list.
    param([System.Windows.Forms.ComboBox]$Combo)

    $List = $script:ComboFilterLists[$Combo.GetHashCode()]
    if (-not $List -or -not $List.Visible -or $List.SelectedIndex -lt 0) { return $false }

    $script:ComboFilterBusy = $true

    try {
        $Combo.Text = [string]$List.SelectedItem
        $Combo.SelectionStart = $Combo.Text.Length
        $Combo.SelectionLength = 0
        $List.Visible = $false
    }
    finally {
        $script:ComboFilterBusy = $false
    }

    return $true
}

function Register-MSToolkitComboFilter {
    # Type-ahead that does not fight the control.
    #
    # Windows autocomplete opens a suggestion popup ON TOP OF the dropdown list,
    # so two lists are visible and a click lands on the entry behind the one being
    # read. Opening the real dropdown instead is no better: an open dropdown owns
    # the keyboard, so the next letters go to the list, which does its own prefix
    # jump and overwrites what was being typed.
    #
    # So: autocomplete off, dropdown left alone, and matches shown in a plain
    # ListBox under the box that never takes focus. What is on screen is what gets
    # clicked, and typing is never interrupted.
    param([System.Windows.Forms.ComboBox]$Combo)

    if (-not $Combo) { return }
    if ($Combo.DropDownStyle -eq [System.Windows.Forms.ComboBoxStyle]::DropDownList) { return }
    if ($script:ComboFilterRegistered[$Combo.GetHashCode()]) { return }

    $Form = $Combo.FindForm()
    if (-not $Form) { return }

    $Combo.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::None
    $Combo.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::None

    $List = New-Object System.Windows.Forms.ListBox
    $List.Font = $Combo.Font
    $List.Width = $Combo.Width
    $List.Height = 160
    $List.Visible = $false
    $List.TabStop = $false
    $List.IntegralHeight = $false
    $List.Tag = "ComboMatches"
    $Form.Controls.Add($List)
    $List.BringToFront()

    $script:ComboFilterLists[$Combo.GetHashCode()] = $List
    $script:ComboFilterOwners[$List.GetHashCode()] = $Combo

    $List.Add_Click({
        param($ListSender, $EventArgs)
        $Owner = $script:ComboFilterOwners[$ListSender.GetHashCode()]
        if ($Owner) {
            [void](Set-MSToolkitComboMatch -Combo $Owner)
            $Owner.Focus()
        }
    })

    $Combo.Add_TextUpdate({
        param($ComboSender, $EventArgs)

        if ($script:ComboFilterBusy) { return }

        $Key = $ComboSender.GetHashCode()
        $Master = $script:ComboFilterItems[$Key]

        # The tool reloads its pickers from time to time - on connect, or after a
        # DC change. Anything longer than what is stored is a fresh load.
        if ($null -eq $Master -or $ComboSender.Items.Count -gt @($Master).Count) {
            $Master = @($ComboSender.Items)
            $script:ComboFilterItems[$Key] = $Master
        }

        $List = $script:ComboFilterLists[$Key]
        if (-not $List -or @($Master).Count -eq 0) { return }

        # If the real dropdown is open it handles the keys itself - jumping to a
        # prefix match and writing it into the box - and the match list below
        # never gets used. Close it and let the match list do the work.
        if ($ComboSender.DroppedDown) {
            # DropDownClosed fires from this, and its handler hides the match
            # list - possibly after this one has just shown it. Flag it so that
            # one close is ignored.
            $script:ComboFilterSkipHide = $true
            $ComboSender.DroppedDown = $false
        }

        $script:ComboFilterBusy = $true

        try {
            $Typed = "$($ComboSender.Text)"

            # The ComboBox does its own prefix matching: type "Matt" and it rewrites
            # the box to the first matching entry with the rest selected. Cut that
            # back to what was actually typed.
            if ($ComboSender.SelectionLength -gt 0 -and $ComboSender.SelectionStart -le $Typed.Length) {
                $Typed = $Typed.Substring(0, $ComboSender.SelectionStart)
                $ComboSender.Text = $Typed
                $ComboSender.SelectionStart = $Typed.Length
                $ComboSender.SelectionLength = 0
            }

            if ([string]::IsNullOrWhiteSpace($Typed)) {
                $List.Visible = $false
                return
            }

            $Filtered = @($Master | Where-Object { "$_" -like "*$Typed*" })

            if ($Filtered.Count -eq 0) {
                $List.Visible = $false
                return
            }

            $List.BeginUpdate()
            $List.Items.Clear()
            $List.Items.AddRange([object[]]$Filtered)
            $List.EndUpdate()
            $List.SelectedIndex = 0

            # Sit it directly under the box, in the form's own coordinates.
            $Form = $ComboSender.FindForm()
            $Below = $ComboSender.PointToScreen((New-Object System.Drawing.Point(0, $ComboSender.Height)))
            $List.Location = $Form.PointToClient($Below)
            $List.Width = $ComboSender.Width

            $Rows = [math]::Min($Filtered.Count, 10)
            $List.Height = ($Rows * $List.ItemHeight) + 4

            $List.Visible = $true
            $List.BringToFront()
        }
        finally {
            $script:ComboFilterBusy = $false
        }
    })

    $Combo.Add_KeyDown({
        param($ComboSender, $KeyArgs)

        # While the real dropdown is open it receives the keystrokes, so the first
        # character typed goes to the list instead of the box - which is why it
        # used to take two presses to start typing. Close it here, on the way down,
        # so this character lands in the text box. Only printable keys: arrows,
        # Enter, Escape and Tab keep their normal behaviour in an open dropdown.
        if ($ComboSender.DroppedDown) {
            $Printable = ($KeyArgs.KeyCode -ge [System.Windows.Forms.Keys]::A -and $KeyArgs.KeyCode -le [System.Windows.Forms.Keys]::Z) -or
                         ($KeyArgs.KeyCode -ge [System.Windows.Forms.Keys]::D0 -and $KeyArgs.KeyCode -le [System.Windows.Forms.Keys]::D9) -or
                         ($KeyArgs.KeyCode -ge [System.Windows.Forms.Keys]::NumPad0 -and $KeyArgs.KeyCode -le [System.Windows.Forms.Keys]::NumPad9) -or
                         $KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Back -or
                         $KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Space -or
                         $KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::OemPeriod -or
                         $KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::OemMinus

            if ($Printable) {
                # Set before closing: the close raises DropDownClosed, whose handler
                # would otherwise hide the match list the filter is about to fill.
                $script:ComboFilterSkipHide = $true
                $ComboSender.DroppedDown = $false
                # Deliberately not handled: the character still has to reach the box.
            }
        }

        $List = $script:ComboFilterLists[$ComboSender.GetHashCode()]
        if (-not $List -or -not $List.Visible) { return }

        switch ($KeyArgs.KeyCode) {
            'Down' {
                if ($List.SelectedIndex -lt ($List.Items.Count - 1)) { $List.SelectedIndex++ }
                $KeyArgs.Handled = $true
                $KeyArgs.SuppressKeyPress = $true
            }
            'Up' {
                if ($List.SelectedIndex -gt 0) { $List.SelectedIndex-- }
                $KeyArgs.Handled = $true
                $KeyArgs.SuppressKeyPress = $true
            }
            'Enter' {
                [void](Set-MSToolkitComboMatch -Combo $ComboSender)
                $KeyArgs.Handled = $true
                $KeyArgs.SuppressKeyPress = $true
            }
            'Tab' {
                [void](Set-MSToolkitComboMatch -Combo $ComboSender)
            }
            'Escape' {
                $List.Visible = $false
                $KeyArgs.Handled = $true
                $KeyArgs.SuppressKeyPress = $true
            }
        }
    })

    $Combo.Add_Leave({
        param($ComboSender, $EventArgs)

        # Leaving for the match list itself is not leaving - the click has to be
        # allowed to land first.
        $List = $script:ComboFilterLists[$ComboSender.GetHashCode()]
        if (-not $List -or -not $List.Visible) { return }

        $Form = $ComboSender.FindForm()
        if ($Form) {
            $Cursor = $Form.PointToClient([System.Windows.Forms.Cursor]::Position)
            if ($List.Bounds.Contains($Cursor)) { return }
        }

        $List.Visible = $false
    })

    $Combo.Add_DropDownClosed({
        param($ComboSender, $EventArgs)

        # Picking from the real dropdown hides the match list - unless this close
        # was triggered by typing, in which case the match list is taking over.
        if ($script:ComboFilterSkipHide) {
            $script:ComboFilterSkipHide = $false
            return
        }

        Hide-MSToolkitComboMatches -Combo $ComboSender
    })

    $script:ComboFilterRegistered[$Combo.GetHashCode()] = $true
}

function Register-MSToolkitComboFiltersOn {
    # Walks a form and converts every editable dropdown it finds, so a picker
    # added later is covered without further wiring.
    param([System.Windows.Forms.Control]$Root)

    if (-not $Root) { return }

    foreach ($Child in @($Root.Controls)) {
        if ($Child -is [System.Windows.Forms.ComboBox]) {
            Register-MSToolkitComboFilter -Combo $Child
        }

        if ($Child.HasChildren) { Register-MSToolkitComboFiltersOn -Root $Child }
    }
}

function Initialize-MSToolkitUserPicker {
    param(
        [System.Windows.Forms.ComboBox[]]$Combos,
        [string]$Server
    )

    try {
        $Users = Get-MSToolkitEmployeeUserList -Server $Server

        foreach ($Combo in $Combos) {
            if (-not $Combo) { continue }

            $Existing = $Combo.Text
            $Combo.Items.Clear()

            foreach ($User in $Users) {
                $null = $Combo.Items.Add((Get-MSToolkitUserPickerLabel -ADUser $User))
            }

            $Combo.Text = $Existing
        }

        return $Users.Count
    }
    catch {
        # The combos stay typeable if the list cannot be loaded.
        return -1
    }
}

function Show-ErrorMessage {
    param(
        [string]$Message,
        [string]$Title = 'Error'
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Show-InfoMessage {
    param(
        [string]$Message,
        [string]$Title = 'Information'
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Write-AppLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"

    if ($script:txtLog) {
        $script:txtLog.AppendText($line + [Environment]::NewLine)
        $script:txtLog.SelectionStart = $script:txtLog.TextLength
        $script:txtLog.ScrollToCaret()
    }
}

function Set-BusyState {
    param(
        [bool]$Busy,
        [string]$StatusText
    )

    if ($script:MainForm) {
        $script:MainForm.UseWaitCursor = $Busy

    # A ComboBox owns its own window handle, so clearing UseWaitCursor on the form
    # does not always restore the pointer over it. Reset the cursor explicitly.
    if (-not $Busy) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    }

    if ($script:btnCompare) {
        $script:btnCompare.Enabled = -not $Busy
    }

    if ($script:btnManageGroups) {
        $script:btnManageGroups.Enabled = -not $Busy
    }

    if ($script:btnAddSelected -and $Busy) {
        $script:btnAddSelected.Enabled = $false
    }
    elseif ($script:btnAddSelected) {
        $script:btnAddSelected.Enabled = ($script:ComparisonRows.Count -gt 0)
    }

    if ($script:lblStatus -and $StatusText) {
        $script:lblStatus.Text = $StatusText
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Resolve-AdUser {
    param(
        [Parameter(Mandatory)][string]$Identity
    )

    $Identity = $Identity.Trim()

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        throw 'A user name was not entered.'
    }

    $params = @{
        Identity = $Identity
        Properties = @(
            'DisplayName',
            'SamAccountName',
            'DistinguishedName',
            'Enabled',
            'Title',
            'Department',
            'Manager',
            'mail'
        )
        ErrorAction = 'Stop'
    }

    if ($script:CurrentServer) {
        $params.Server = $script:CurrentServer
    }

    Get-ADUser @params
}

function Get-UserGroupsDetailed {
    param(
        [Parameter(Mandatory)][Microsoft.ActiveDirectory.Management.ADUser]$User,
        [switch]$Recursive
    )

    if ($Recursive) {
        $params = @{
            Identity = $User.DistinguishedName
            ErrorAction = 'Stop'
        }

        if ($script:CurrentServer) {
            $params.Server = $script:CurrentServer
        }

        $groups = @(Get-ADPrincipalGroupMembership @params)
    }
    else {
        $userParams = @{
            Identity = $User.DistinguishedName
            Properties = 'MemberOf'
            ErrorAction = 'Stop'
        }

        if ($script:CurrentServer) {
            $userParams.Server = $script:CurrentServer
        }

        $groupDns = @((Get-ADUser @userParams).MemberOf)

        $groups = foreach ($groupDn in $groupDns) {
            $groupParams = @{
                Identity = $groupDn
                ErrorAction = 'Stop'
            }

            if ($script:CurrentServer) {
                $groupParams.Server = $script:CurrentServer
            }

            Get-ADGroup @groupParams
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($group in ($groups | Sort-Object Name -Unique)) {
        $propsParams = @{
            Identity = $group.DistinguishedName
            Properties = @(
                'GroupCategory',
                'GroupScope',
                'mail',
                'Description',
                'ManagedBy',
                'MemberOf'
            )
            ErrorAction = 'Stop'
        }

        if ($script:CurrentServer) {
            $propsParams.Server = $script:CurrentServer
        }

        $props = Get-ADGroup @propsParams
        $isHighRisk = $false

        foreach ($pattern in $script:HighRiskGroupPatterns) {
            if ($props.Name -like "*$pattern*") {
                $isHighRisk = $true
                break
            }
        }

        $results.Add([pscustomobject]@{
            Name              = $props.Name
            SamAccountName    = $props.SamAccountName
            DistinguishedName = $props.DistinguishedName
            GroupCategory     = [string]$props.GroupCategory
            GroupScope        = [string]$props.GroupScope
            Mail              = $props.Mail
            Description       = $props.Description
            ManagedBy         = $props.ManagedBy
            NestedInCount     = @($props.MemberOf).Count
            IsHighRisk        = $isHighRisk
        })
    }

    return $results.ToArray()
}

function Update-ComparisonGrid {
    $script:dgvGroups.Rows.Clear()

    foreach ($row in $script:ComparisonRows) {
        $typeText = "$($row.GroupCategory) - $($row.GroupScope)"

        if ($row.IsHighRisk) {
            $typeText = "$typeText - HIGH RISK"
        }

        $index = $script:dgvGroups.Rows.Add(
            $false,
            $row.Name,
            $row.Description,
            $typeText,
            'Missing'
        )

        $script:dgvGroups.Rows[$index].Tag = $row

        if ($row.IsHighRisk) {
            $script:dgvGroups.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).DangerBackground
        }
    }

    $script:lblMissingCount.Text = "Missing groups: $($script:ComparisonRows.Count)"
    $script:btnSelectAll.Enabled = ($script:ComparisonRows.Count -gt 0)
    $script:btnClearSelection.Enabled = ($script:ComparisonRows.Count -gt 0)
    $script:btnAddSelected.Enabled = ($script:ComparisonRows.Count -gt 0)
}

function Compare-Users {
    if ([string]::IsNullOrWhiteSpace($script:txtReference.Text) -or
        [string]::IsNullOrWhiteSpace($script:txtTarget.Text)) {
        Show-InfoMessage 'Enter both a Reference User and a Target User.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Resolving users and comparing AD groups...'

        $script:dgvGroups.Rows.Clear()
        $script:ComparisonRows = @()
        $script:lblReferenceResolved.Text = ''
        $script:lblTargetResolved.Text = ''
        $script:lblMissingCount.Text = 'Missing groups: 0'

        Write-AppLog "Resolving reference user: $($script:txtReference.Text)"
        $script:ReferenceUser = Resolve-AdUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtReference.Text)
        $script:lblReferenceResolved.Text = "$($script:ReferenceUser.DisplayName)  |  $($script:ReferenceUser.SamAccountName)"

        Write-AppLog "Resolving target user: $($script:txtTarget.Text)"
        $script:TargetUser = Resolve-AdUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtTarget.Text)
        $script:lblTargetResolved.Text = "$($script:TargetUser.DisplayName)  |  $($script:TargetUser.SamAccountName)"

        if ($script:ReferenceUser.DistinguishedName -eq $script:TargetUser.DistinguishedName) {
            throw 'The Reference User and Target User resolve to the same AD account.'
        }

        $recursive = $script:chkIncludeNested.Checked

        Write-AppLog "Reading group memberships for $($script:ReferenceUser.SamAccountName). Include nested: $recursive"
        $referenceGroups = @(Get-UserGroupsDetailed -User $script:ReferenceUser -Recursive:$recursive)

        Write-AppLog "Reading group memberships for $($script:TargetUser.SamAccountName). Include nested: $recursive"
        $targetGroups = @(Get-UserGroupsDetailed -User $script:TargetUser -Recursive:$recursive)

        $targetDns = @{}
        foreach ($group in $targetGroups) {
            $targetDns[$group.DistinguishedName] = $true
        }

        $missing = New-Object System.Collections.Generic.List[object]

        foreach ($group in $referenceGroups) {
            if (-not $targetDns.ContainsKey($group.DistinguishedName)) {
                $missing.Add($group)
            }
        }

        $script:ComparisonRows = @($missing.ToArray() | Sort-Object Name)

        $script:lblReferenceCount.Text = "Reference groups: $($referenceGroups.Count)"
        $script:lblTargetCount.Text = "Target groups: $($targetGroups.Count)"

        Update-ComparisonGrid

        $script:lblStatus.Text = "Comparison complete. $($script:ComparisonRows.Count) missing group(s) found."
        Write-AppLog "Comparison complete: reference=$($referenceGroups.Count), target=$($targetGroups.Count), missing=$($script:ComparisonRows.Count)." 'SUCCESS'

        if ($script:ComparisonRows.Count -eq 0) {
            Show-InfoMessage "$($script:TargetUser.DisplayName) already has all compared memberships that $($script:ReferenceUser.DisplayName) has." 'No Missing Groups'
        }
    }
    catch {
        Show-ErrorMessage "Unable to complete the comparison.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Comparison failed: $($_.Exception.Message)" 'ERROR'
        $script:lblStatus.Text = 'Comparison failed.'
    }
    finally {
        Set-BusyState -Busy $false
    }
}

function Get-SelectedComparisonRows {
    if ($script:dgvGroups.IsCurrentCellDirty) {
        $script:dgvGroups.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        $script:dgvGroups.EndEdit()
    }

    $selected = New-Object System.Collections.Generic.List[object]

    foreach ($gridRow in $script:dgvGroups.Rows) {
        if ($gridRow.IsNewRow) {
            continue
        }

        if ([bool]$gridRow.Cells[0].Value -and $gridRow.Tag) {
            $selected.Add($gridRow.Tag)
        }
    }

    return $selected.ToArray()
}

function Add-SelectedGroups {
    if (-not $script:TargetUser) {
        Show-InfoMessage 'Run a comparison first.'
        return
    }

    $selected = @(Get-SelectedComparisonRows)

    if ($selected.Count -eq 0) {
        Show-InfoMessage 'Select at least one group to add.'
        return
    }

    $highRisk = @($selected | Where-Object { $_.IsHighRisk })

    $confirmText = "Add $($script:TargetUser.DisplayName) ($($script:TargetUser.SamAccountName)) to $($selected.Count) selected AD group(s)?"

    if ($script:chkIncludeNested.Checked) {
        $confirmText += "`r`n`r`nNOTE: Nested membership comparison is enabled. Selected groups will be added as DIRECT memberships."
    }

    if ($highRisk.Count -gt 0) {
        $confirmText += "`r`n`r`nWARNING: $($highRisk.Count) selected group(s) are marked HIGH RISK and require careful review."
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $confirmText,
        'Confirm Group Membership Changes',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog 'Add selected groups cancelled.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Adding selected AD group memberships...'

        $success = 0
        $failed = 0

        foreach ($group in $selected) {
            try {
                $params = @{
                    Identity = $group.DistinguishedName
                    Members = $script:TargetUser.DistinguishedName
                    ErrorAction = 'Stop'
                }

                if ($script:CurrentServer) {
                    $params.Server = $script:CurrentServer
                }

                Add-ADGroupMember @params

                Write-AppLog "Added $($script:TargetUser.SamAccountName) to $($group.Name)." 'SUCCESS'
                $success++
            }
            catch {
                Write-AppLog "Failed to add $($script:TargetUser.SamAccountName) to $($group.Name): $($_.Exception.Message)" 'ERROR'
                $failed++
            }
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Completed group membership changes.`r`n`r`nSuccessful: $success`r`nFailed: $failed",
            'Group Membership Results',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $(if ($failed -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information })
        ) | Out-Null

        Compare-Users
    }
    finally {
        Set-BusyState -Busy $false
    }
}


function Show-ADGroupManager {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Manage AD User Groups'
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object System.Drawing.Size(1000,680)
    $dialog.MinimumSize = New-Object System.Drawing.Size(900,600)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $dialog.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Top'
    $header.Height = 68
    $header.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $dialog.Controls.Add($header)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Manage AD User Groups'
    $title.AutoSize = $true
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold',17)
    $title.Location = New-Object System.Drawing.Point(18,6)
    $header.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = 'Load one user, review all AD groups, and add selected direct memberships'
    $subtitle.AutoSize = $true
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
    $subtitle.Location = New-Object System.Drawing.Point(20,42)
    $header.Controls.Add($subtitle)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'User (sAMAccountName)'
    $lblUser.AutoSize = $true
    $lblUser.Location = New-Object System.Drawing.Point(20,88)
    $dialog.Controls.Add($lblUser)

    # Same note as the main window's User Comparison section, beside the User label.
    $UserListNote = Get-MSToolkitUserListNote
    $lblUserListNote = $null
    if ($UserListNote) {
        $lblUserListNote = New-Object System.Windows.Forms.Label
        $lblUserListNote.Text = $UserListNote.Text
        $lblUserListNote.AutoSize = $true
        $lblUserListNote.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
        $lblUserListNote.Location = New-Object System.Drawing.Point((20 + $lblUser.PreferredWidth + 16), 88)
        $dialog.Controls.Add($lblUserListNote)
    }

    $txtUser = New-Object System.Windows.Forms.ComboBox
    $txtUser.Location = New-Object System.Drawing.Point(20,110)
    $txtUser.Size = New-Object System.Drawing.Size(560,24)
    $txtUser.DropDownStyle = 'DropDown'
    $txtUser.AutoCompleteMode = 'None'
    $txtUser.AutoCompleteSource = 'None'
    $txtUser.MaxDropDownItems = 20
    $null = Initialize-MSToolkitUserPicker -Combos @($txtUser) -Server $script:CurrentServer
    if ($script:TargetUser) { $txtUser.Text = $script:TargetUser.SamAccountName }
    $dialog.Controls.Add($txtUser)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load All Groups'
    $btnLoad.Location = New-Object System.Drawing.Point(595,106)
    $btnLoad.Size = New-Object System.Drawing.Size(135,32)
    $btnLoad.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $btnLoad.ForeColor = [System.Drawing.Color]::White
    $btnLoad.FlatStyle = 'Flat'
    $dialog.Controls.Add($btnLoad)

    $lblResolved = New-Object System.Windows.Forms.Label
    $lblResolved.Text = ''
    $lblResolved.AutoSize = $true
    $lblResolved.ForeColor = [System.Drawing.Color]::DimGray
    $lblResolved.Location = New-Object System.Drawing.Point(20,142)
    $dialog.Controls.Add($lblResolved)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(20,172)
    $grid.Size = New-Object System.Drawing.Size(945,390)
    $grid.Anchor = 'Top,Bottom,Left,Right'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.MultiSelect = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.AutoSizeRowsMode = 'None'
    $dialog.Controls.Add($grid)

    $colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colSelect.Name = 'Selected'
    $colSelect.HeaderText = 'Add'
    $colSelect.Width = 55
    [void]$grid.Columns.Add($colSelect)

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = 'Group Name'
    $colName.Width = 280
    $colName.ReadOnly = $true
    [void]$grid.Columns.Add($colName)

    $colDesc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDesc.HeaderText = 'Description'
    $colDesc.AutoSizeMode = 'Fill'
    $colDesc.ReadOnly = $true
    [void]$grid.Columns.Add($colDesc)

    $colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colType.HeaderText = 'Type'
    $colType.Width = 190
    $colType.ReadOnly = $true
    [void]$grid.Columns.Add($colType)

    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.HeaderText = 'Membership'
    $colStatus.Width = 135
    $colStatus.ReadOnly = $true
    [void]$grid.Columns.Add($colStatus)

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = 'Select Available'
    $btnSelectAll.Location = New-Object System.Drawing.Point(20,575)
    $btnSelectAll.Size = New-Object System.Drawing.Size(120,30)
    $btnSelectAll.Anchor = 'Bottom,Left'
    $dialog.Controls.Add($btnSelectAll)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = 'Clear'
    $btnClear.Location = New-Object System.Drawing.Point(150,575)
    $btnClear.Size = New-Object System.Drawing.Size(80,30)
    $btnClear.Anchor = 'Bottom,Left'
    $dialog.Controls.Add($btnClear)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Add Selected Groups'
    $btnAdd.Location = New-Object System.Drawing.Point(745,575)
    $btnAdd.Size = New-Object System.Drawing.Size(220,34)
    $btnAdd.Anchor = 'Bottom,Right'
    $btnAdd.BackColor = [System.Drawing.Color]::FromArgb(26,137,85)
    $btnAdd.ForeColor = [System.Drawing.Color]::White
    $btnAdd.FlatStyle = 'Flat'
    $btnAdd.Enabled = $false
    $dialog.Controls.Add($btnAdd)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'Remove Selected Groups'
    $btnRemove.Location = New-Object System.Drawing.Point(505,575)
    $btnRemove.Size = New-Object System.Drawing.Size(220,34)
    $btnRemove.Anchor = 'Bottom,Right'
    $btnRemove.BackColor = [System.Drawing.Color]::FromArgb(178,34,34)
    $btnRemove.ForeColor = [System.Drawing.Color]::White
    $btnRemove.FlatStyle = 'Flat'
    $btnRemove.Enabled = $false
    $dialog.Controls.Add($btnRemove)

    $script:ManageADUser = $null

    $loadGroups = {
        if ([string]::IsNullOrWhiteSpace($txtUser.Text)) {
            Show-InfoMessage 'Enter a user first.'
            return
        }

        try {
            $dialog.UseWaitCursor = $true
            $grid.Rows.Clear()
            $btnAdd.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()

            $script:ManageADUser = Resolve-AdUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $txtUser.Text)
            $lblResolved.Text = "$($script:ManageADUser.DisplayName)  |  $($script:ManageADUser.SamAccountName)"

            $userParams = @{
                Identity = $script:ManageADUser.DistinguishedName
                Properties = 'MemberOf'
                ErrorAction = 'Stop'
            }
            $groupParams = @{
                Filter = '*'
                Properties = @('Description','GroupCategory','GroupScope','DistinguishedName','SID','isCriticalSystemObject')
                ErrorAction = 'Stop'
            }
            if ($script:CurrentServer) {
                $userParams.Server = $script:CurrentServer
                $groupParams.Server = $script:CurrentServer
            }

            $directMemberships = @((Get-ADUser @userParams).MemberOf)
            $directSet = @{}
            foreach ($dn in $directMemberships) { $directSet[[string]$dn] = $true }

            $allGroups = @(Get-ADGroup @groupParams | Sort-Object Name)

            foreach ($groupItem in $allGroups) {
                $isMember = $directSet.ContainsKey([string]$groupItem.DistinguishedName)
                $isHighRisk = $false
                foreach ($pattern in $script:HighRiskGroupPatterns) {
                    if ($groupItem.Name -like "*$pattern*") { $isHighRisk = $true; break }
                }

                $criticalReason = Get-MSToolkitCriticalGroupReason -Group $groupItem
                $isCritical = -not [string]::IsNullOrWhiteSpace($criticalReason)

                $typeText = "$($groupItem.GroupCategory) - $($groupItem.GroupScope)"
                if ($isHighRisk) { $typeText += ' - HIGH RISK' }
                if ($isCritical) { $typeText += ' - PROTECTED' }

                $index = $grid.Rows.Add(
                    $false,
                    $groupItem.Name,
                    $groupItem.Description,
                    $typeText,
                    $(if ($isMember) { 'Already Direct Member' } else { 'Available' })
                )

                $grid.Rows[$index].Tag = [pscustomobject]@{
                    Name = $groupItem.Name
                    DistinguishedName = $groupItem.DistinguishedName
                    IsHighRisk = $isHighRisk
                    IsMember = $isMember
                    IsCritical = $isCritical
                    CriticalReason = $criticalReason
                }

                if ($isMember) {
                    # Selectable, because member rows are what the Remove button acts on.
                    $grid.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).SuccessBackground
                }
                elseif ($isHighRisk) {
                    $grid.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).DangerBackground
                }
            }

            $btnAdd.Enabled = ($grid.Rows.Count -gt 0)
            $btnRemove.Enabled = ($grid.Rows.Count -gt 0)
        }
        catch {
            Show-ErrorMessage "Unable to load AD groups.`r`n`r`n$($_.Exception.Message)"
        }
        finally {
            $dialog.UseWaitCursor = $false

            # A ComboBox owns its own window handle, so clearing UseWaitCursor on
            # the dialog does not always restore the pointer over it.
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            $dialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnLoad.Add_Click($loadGroups)
    # Enter accepts an entry from the user list; use the Load button to fetch groups.
    $txtUser.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $_.SuppressKeyPress = $true }
    })

    $btnSelectAll.Add_Click({
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow -and $row.Tag -and -not $row.Tag.IsMember) {
                $row.Cells['Selected'].Value = $true
            }
        }
    })

    $btnClear.Add_Click({
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow -and -not $row.Cells['Selected'].ReadOnly) {
                $row.Cells['Selected'].Value = $false
            }
        }
    })

    $btnAdd.Add_Click({
        if (-not $script:ManageADUser) {
            Show-InfoMessage 'Load a user first.'
            return
        }

        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            $grid.EndEdit()
        }

        $selected = New-Object System.Collections.Generic.List[object]
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow -and [bool]$row.Cells['Selected'].Value -and $row.Tag -and -not $row.Tag.IsMember) {
                $selected.Add($row.Tag)
            }
        }

        $selectedGroups = $selected.ToArray()
        if ($selectedGroups.Count -eq 0) {
            Show-InfoMessage 'Select at least one available group to add.'
            return
        }

        $highRisk = @($selectedGroups | Where-Object { $_.IsHighRisk })
        $names = @($selectedGroups | Sort-Object Name | ForEach-Object { $_.Name })
        $nameText = ($names -join "`r`n - ")
        $warning = ''
        if ($highRisk.Count -gt 0) {
            $warning = "`r`n`r`nWARNING: $($highRisk.Count) selected group(s) are marked HIGH RISK."
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($script:ManageADUser.DisplayName) ($($script:ManageADUser.SamAccountName)) to the following $($selectedGroups.Count) AD group(s)?`r`n`r`n - $nameText$warning",
            'Confirm Group Membership Changes',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $success = 0
        $failed = 0
        $dialog.UseWaitCursor = $true

        try {
            foreach ($groupItem in $selectedGroups) {
                try {
                    $params = @{
                        Identity = $groupItem.DistinguishedName
                        Members = $script:ManageADUser.DistinguishedName
                        Confirm = $false
                        ErrorAction = 'Stop'
                    }
                    if ($script:CurrentServer) { $params.Server = $script:CurrentServer }

                    Add-ADGroupMember @params
                    Write-AppLog "Added $($script:ManageADUser.SamAccountName) to $($groupItem.Name) from Manage Groups." 'SUCCESS'
                    $success++
                }
                catch {
                    Write-AppLog "Failed to add $($script:ManageADUser.SamAccountName) to $($groupItem.Name): $($_.Exception.Message)" 'ERROR'
                    $failed++
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Group additions complete.`r`n`r`nSuccessful: $success`r`nFailed: $failed`r`n`r`nReview the Activity Log for details.",
                'Group Addition Complete',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $(if ($failed -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information })
            ) | Out-Null

            & $loadGroups
        }
        finally {
            $dialog.UseWaitCursor = $false

            # A ComboBox owns its own window handle, so clearing UseWaitCursor on
            # the dialog does not always restore the pointer over it.
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            $dialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $btnRemove.Add_Click({
        if (-not $script:ManageADUser) {
            Show-InfoMessage 'Load a user first.'
            return
        }

        # Commit the checkbox edit before reading its value.
        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            $grid.EndEdit()
        }

        $selected = New-Object System.Collections.Generic.List[object]
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow -and [bool]$row.Cells['Selected'].Value -and $row.Tag -and $row.Tag.IsMember) {
                $selected.Add($row.Tag)
            }
        }

        $selectedGroups = $selected.ToArray()
        if ($selectedGroups.Count -eq 0) {
            Show-InfoMessage 'Select at least one group the user is already a direct member of.'
            return
        }

        # Critical groups are blocked outright and never silently dropped.
        $blocked = @($selectedGroups | Where-Object { $_.IsCritical })
        $removable = @($selectedGroups | Where-Object { -not $_.IsCritical })

        if ($blocked.Count -gt 0) {
            $blockedText = (@($blocked | Sort-Object Name | ForEach-Object { "$($_.Name)  -  $($_.CriticalReason)" }) -join "`r`n - ")

            [System.Windows.Forms.MessageBox]::Show(
                "OPERATION BLOCKED for $($blocked.Count) protected group(s). These cannot be removed with this tool:`r`n`r`n - $blockedText`r`n`r`nUse Active Directory Users and Computers if such a change is genuinely required.",
                'Protected Groups Blocked',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Stop
            ) | Out-Null

            foreach ($groupItem in $blocked) {
                Write-AppLog "BLOCKED: removal of $($script:ManageADUser.SamAccountName) from $($groupItem.Name) - $($groupItem.CriticalReason)" 'ERROR'
            }
        }

        if ($removable.Count -eq 0) { return }

        $highRisk = @($removable | Where-Object { $_.IsHighRisk })
        $names = @($removable | Sort-Object Name | ForEach-Object { $_.Name })
        $nameText = ($names -join "`r`n - ")
        $warning = ''
        if ($highRisk.Count -gt 0) {
            $warning = "`r`n`r`nWARNING: $($highRisk.Count) selected group(s) are marked HIGH RISK. Removing access can break sign-in, mailbox, or application permissions."
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "REMOVE $($script:ManageADUser.DisplayName) ($($script:ManageADUser.SamAccountName)) from the following $($removable.Count) AD group(s)?`r`n`r`n - $nameText$warning`r`n`r`nThis revokes whatever access those groups grant.",
            'Confirm Group Removal',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-AppLog 'Group removal cancelled.' 'WARNING'
            return
        }

        $success = 0
        $failed = 0
        $dialog.UseWaitCursor = $true

        try {
            foreach ($groupItem in $removable) {
                try {
                    $params = @{
                        Identity = $groupItem.DistinguishedName
                        Members = $script:ManageADUser.DistinguishedName
                        Confirm = $false
                        ErrorAction = 'Stop'
                    }
                    if ($script:CurrentServer) { $params.Server = $script:CurrentServer }

                    Remove-ADGroupMember @params
                    Write-AppLog "Removed $($script:ManageADUser.SamAccountName) from $($groupItem.Name) from Manage Groups." 'SUCCESS'
                    $success++
                }
                catch {
                    Write-AppLog "Failed to remove $($script:ManageADUser.SamAccountName) from $($groupItem.Name): $($_.Exception.Message)" 'ERROR'
                    $failed++
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Group removals complete.`r`n`r`nSuccessful: $success`r`nFailed: $failed`r`n`r`nReview the Activity Log for details.",
                'Group Removal Complete',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $(if ($failed -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information })
            ) | Out-Null

            & $loadGroups
        }
        finally {
            $dialog.UseWaitCursor = $false

            # A ComboBox owns its own window handle, so clearing UseWaitCursor on
            # the dialog does not always restore the pointer over it.
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            $dialog.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $dialog.Add_Shown({
        $dialog.Activate()
        $txtUser.Focus()
    })

    Apply-MSToolkitSharedTheme -Root $dialog

    # The shared theme pass cannot infer the note's colour, so set it afterwards.
    if ($lblUserListNote) {
        $NotePalette = Get-MSToolkitThemePalette
        $lblUserListNote.ForeColor = if ($UserListNote.IsWarning) { $NotePalette.Warning } else { $NotePalette.MutedText }
    }

    Register-MSToolkitComboFiltersOn -Root $dialog
    [void]$dialog.ShowDialog($script:MainForm)
    Remove-Variable ManageADUser -Scope Script -ErrorAction SilentlyContinue
}

# ---------------------------
# Main Windows Forms interface
# ---------------------------
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'AD Group Compare'
$MainForm.StartPosition = 'CenterScreen'
$MainForm.Size = New-Object System.Drawing.Size(1180, 800)
$MainForm.MinimumSize = New-Object System.Drawing.Size(1050, 700)
$MainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$MainForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$MainForm.FormBorderStyle = 'Sizable'
$MainForm.MaximizeBox = $true
$script:MainForm = $MainForm

# Header
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = 'Top'
$pnlHeader.Height = 78
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(31, 58, 93)
$MainForm.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Active Directory Group Compare'
$lblTitle.AutoSize = $true
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Regular)
$lblTitle.Location = New-Object System.Drawing.Point(25, 9)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Compare AD group memberships and add missing memberships to a target user'
$lblSubtitle.AutoSize = $true
$lblSubtitle.ForeColor = [System.Drawing.Color]::White
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblSubtitle.Location = New-Object System.Drawing.Point(28, 51)
$pnlHeader.Controls.Add($lblSubtitle)

$lblServerHeader = New-Object System.Windows.Forms.Label
$lblServerHeader.Text = if ($script:CurrentServer) { "AD Server: $($script:CurrentServer)" } else { 'AD Server: Default domain controller' }
$lblServerHeader.AutoSize = $true
$lblServerHeader.ForeColor = [System.Drawing.Color]::White
$lblServerHeader.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblServerHeader.Anchor = 'Top,Right'
$lblServerHeader.Location = New-Object System.Drawing.Point(875, 31)
$pnlHeader.Controls.Add($lblServerHeader)
$script:MSToolkitThemeToggleButton = New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "CompareUserGroups"

# User comparison section
$grpUsers = New-Object System.Windows.Forms.GroupBox
$grpUsers.Text = 'User Comparison'
$grpUsers.Location = New-Object System.Drawing.Point(20, 92)
$grpUsers.Size = New-Object System.Drawing.Size(1135, 165)
$grpUsers.Anchor = 'Top,Left,Right'
$MainForm.Controls.Add($grpUsers)

$lblUserFormatHint = New-Object System.Windows.Forms.Label
$lblUserFormatHint.Text = 'Enter the AD sAMAccountName for each user.'
$lblUserFormatHint.AutoSize = $true
$lblUserFormatHint.ForeColor = [System.Drawing.Color]::DimGray
$lblUserFormatHint.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Italic)
$lblUserFormatHint.Location = New-Object System.Drawing.Point(20, 23)
$grpUsers.Controls.Add($lblUserFormatHint)

# What the user lists cover, shown only when an OU setting leaves a gap:
#   no Users OU            -> orange: the lists cover the whole domain
#   Users OU, no Admin OU  -> orange: admin accounts outside the Users OU are not listed
# Its colour is applied after the theme pass (see $script:MSToolkitThemeRefreshHook).
$UsersOUSet  = [bool](Get-MSToolkitSetting -Name "OUUsers")
$AdminsOUSet = [bool](Get-MSToolkitSetting -Name "OUAdmins")
$script:UserListNoteIsWarning = $true

$lblUsersOUWarning = New-Object System.Windows.Forms.Label
$lblUsersOUWarning.Text = if (-not $UsersOUSet) { 'No Users OU is set in MSToolkit Settings, so the user lists below cover the whole domain.' } elseif (-not $AdminsOUSet) { "No Admin accounts OU is set in MSToolkit Settings: admin accounts outside the Users OU aren't listed (typing a name still works)." } else { '' }
$lblUsersOUWarning.AutoSize = $true
$lblUsersOUWarning.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$lblUsersOUWarning.Location = New-Object System.Drawing.Point((20 + $lblUserFormatHint.PreferredWidth + 16), 23)
$lblUsersOUWarning.Visible = (-not $UsersOUSet) -or (-not $AdminsOUSet)
$grpUsers.Controls.Add($lblUsersOUWarning)
$script:lblUsersOUWarning = $lblUsersOUWarning

$lblReference = New-Object System.Windows.Forms.Label
$lblReference.Text = 'Reference User'
$lblReference.AutoSize = $true
$lblReference.Location = New-Object System.Drawing.Point(20, 48)
$grpUsers.Controls.Add($lblReference)

$txtReference = New-Object System.Windows.Forms.ComboBox
$txtReference.Location = New-Object System.Drawing.Point(20, 68)
$txtReference.Size = New-Object System.Drawing.Size(430, 24)
$txtReference.Anchor = 'Top,Left,Right'
$txtReference.DropDownStyle = 'DropDown'
$txtReference.AutoCompleteMode = 'None'
$txtReference.AutoCompleteSource = 'None'
$txtReference.MaxDropDownItems = 20
$grpUsers.Controls.Add($txtReference)
$script:txtReference = $txtReference

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = 'Target User'
$lblTarget.AutoSize = $true
$lblTarget.Location = New-Object System.Drawing.Point(482, 48)
$grpUsers.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.ComboBox
$txtTarget.Location = New-Object System.Drawing.Point(482, 68)
$txtTarget.Size = New-Object System.Drawing.Size(430, 24)
$txtTarget.Anchor = 'Top,Left,Right'
$txtTarget.DropDownStyle = 'DropDown'
$txtTarget.AutoCompleteMode = 'None'
$txtTarget.AutoCompleteSource = 'None'
$txtTarget.MaxDropDownItems = 20
$grpUsers.Controls.Add($txtTarget)
$script:txtTarget = $txtTarget

$btnCompare = New-Object System.Windows.Forms.Button
$btnCompare.Text = 'Compare Users'
$btnCompare.Location = New-Object System.Drawing.Point(935, 62)
$btnCompare.Size = New-Object System.Drawing.Size(165, 42)
$btnCompare.Anchor = 'Top,Right'
$btnCompare.BackColor = [System.Drawing.Color]::FromArgb(31, 58, 93)
$btnCompare.ForeColor = [System.Drawing.Color]::White
$btnCompare.FlatStyle = 'Flat'
$btnCompare.Add_Click({ Compare-Users })
$grpUsers.Controls.Add($btnCompare)
$script:btnCompare = $btnCompare

$btnManageGroups = New-Object System.Windows.Forms.Button
$btnManageGroups.Text = 'Manage Groups'
$btnManageGroups.Location = New-Object System.Drawing.Point(935,108)
$btnManageGroups.Size = New-Object System.Drawing.Size(165,32)
$btnManageGroups.Anchor = 'Top,Right'
$btnManageGroups.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$btnManageGroups.ForeColor = [System.Drawing.Color]::White
$btnManageGroups.FlatStyle = 'Flat'
$btnManageGroups.Add_Click({ Show-ADGroupManager })
$grpUsers.Controls.Add($btnManageGroups)
$script:btnManageGroups = $btnManageGroups

$chkIncludeNested = New-Object System.Windows.Forms.CheckBox
$chkIncludeNested.Text = 'Include nested group memberships'
$chkIncludeNested.AutoSize = $true
$chkIncludeNested.Location = New-Object System.Drawing.Point(20, 101)
$chkIncludeNested.Checked = $false
$grpUsers.Controls.Add($chkIncludeNested)
$script:chkIncludeNested = $chkIncludeNested

$lblNestedWarning = New-Object System.Windows.Forms.Label
$lblNestedWarning.Text = 'If enabled, selected missing groups are still added as direct memberships.'
$lblNestedWarning.AutoSize = $true
$lblNestedWarning.ForeColor = [System.Drawing.Color]::DimGray
$lblNestedWarning.Location = New-Object System.Drawing.Point(245, 103)
$grpUsers.Controls.Add($lblNestedWarning)

$lblReferenceResolved = New-Object System.Windows.Forms.Label
$lblReferenceResolved.Text = ''
$lblReferenceResolved.AutoSize = $true
$lblReferenceResolved.Location = New-Object System.Drawing.Point(20, 128)
$grpUsers.Controls.Add($lblReferenceResolved)
$script:lblReferenceResolved = $lblReferenceResolved

$lblTargetResolved = New-Object System.Windows.Forms.Label
$lblTargetResolved.Text = ''
$lblTargetResolved.AutoSize = $true
$lblTargetResolved.Location = New-Object System.Drawing.Point(482, 128)
$grpUsers.Controls.Add($lblTargetResolved)
$script:lblTargetResolved = $lblTargetResolved

$lblReferenceCount = New-Object System.Windows.Forms.Label
$lblReferenceCount.Text = 'Reference groups: 0'
$lblReferenceCount.AutoSize = $true
$lblReferenceCount.Location = New-Object System.Drawing.Point(20, 146)
$grpUsers.Controls.Add($lblReferenceCount)
$script:lblReferenceCount = $lblReferenceCount

$lblTargetCount = New-Object System.Windows.Forms.Label
$lblTargetCount.Text = 'Target groups: 0'
$lblTargetCount.AutoSize = $true
$lblTargetCount.Location = New-Object System.Drawing.Point(482, 146)
$grpUsers.Controls.Add($lblTargetCount)
$script:lblTargetCount = $lblTargetCount

# Missing group action strip
$pnlActions = New-Object System.Windows.Forms.Panel
$pnlActions.Location = New-Object System.Drawing.Point(20, 268)
$pnlActions.Size = New-Object System.Drawing.Size(1135, 42)
$pnlActions.Anchor = 'Top,Left,Right'
$pnlActions.BackColor = [System.Drawing.Color]::White
$MainForm.Controls.Add($pnlActions)

$lblMissingCount = New-Object System.Windows.Forms.Label
$lblMissingCount.Text = 'Missing groups: 0'
$lblMissingCount.AutoSize = $true
$lblMissingCount.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblMissingCount.Location = New-Object System.Drawing.Point(15, 13)
$pnlActions.Controls.Add($lblMissingCount)
$script:lblMissingCount = $lblMissingCount

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = 'Select All'
$btnSelectAll.Size = New-Object System.Drawing.Size(92, 30)
$btnSelectAll.Location = New-Object System.Drawing.Point(755, 6)
$btnSelectAll.Anchor = 'Top,Right'
$btnSelectAll.Enabled = $false
$btnSelectAll.Add_Click({
    foreach ($row in $script:dgvGroups.Rows) {
        if (-not $row.IsNewRow) {
            $row.Cells[0].Value = $true
        }
    }
})
$pnlActions.Controls.Add($btnSelectAll)
$script:btnSelectAll = $btnSelectAll

$btnClearSelection = New-Object System.Windows.Forms.Button
$btnClearSelection.Text = 'Clear'
$btnClearSelection.Size = New-Object System.Drawing.Size(78, 30)
$btnClearSelection.Location = New-Object System.Drawing.Point(860, 6)
$btnClearSelection.Anchor = 'Top,Right'
$btnClearSelection.Enabled = $false
$btnClearSelection.Add_Click({
    foreach ($row in $script:dgvGroups.Rows) {
        if (-not $row.IsNewRow) {
            $row.Cells[0].Value = $false
        }
    }
})
$pnlActions.Controls.Add($btnClearSelection)
$script:btnClearSelection = $btnClearSelection

$btnAddSelected = New-Object System.Windows.Forms.Button
$btnAddSelected.Text = 'Add Selected Groups'
$btnAddSelected.Size = New-Object System.Drawing.Size(165, 30)
$btnAddSelected.Location = New-Object System.Drawing.Point(950, 6)
$btnAddSelected.Anchor = 'Top,Right'
$btnAddSelected.BackColor = [System.Drawing.Color]::FromArgb(26, 137, 85)
$btnAddSelected.ForeColor = [System.Drawing.Color]::White
$btnAddSelected.FlatStyle = 'Flat'
$btnAddSelected.Enabled = $false
$btnAddSelected.Add_Click({ Add-SelectedGroups })
$pnlActions.Controls.Add($btnAddSelected)
$script:btnAddSelected = $btnAddSelected

# Grid
$dgvGroups = New-Object System.Windows.Forms.DataGridView
$dgvGroups.Location = New-Object System.Drawing.Point(20, 320)
$dgvGroups.Size = New-Object System.Drawing.Size(1135, 295)
$dgvGroups.Anchor = 'Top,Bottom,Left,Right'
$dgvGroups.AllowUserToAddRows = $false
$dgvGroups.AllowUserToDeleteRows = $false
$dgvGroups.AllowUserToResizeRows = $false
$dgvGroups.MultiSelect = $false
$dgvGroups.RowHeadersVisible = $false
$dgvGroups.SelectionMode = 'FullRowSelect'
$dgvGroups.BackgroundColor = [System.Drawing.Color]::White
$dgvGroups.BorderStyle = 'Fixed3D'
$dgvGroups.AutoSizeRowsMode = 'None'
$dgvGroups.RowTemplate.Height = 22
$MainForm.Controls.Add($dgvGroups)
$script:dgvGroups = $dgvGroups

$colAdd = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colAdd.Name = 'Add'
$colAdd.HeaderText = 'Add'
$colAdd.Width = 50
$colAdd.AutoSizeMode = 'None'
[void]$dgvGroups.Columns.Add($colAdd)

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = 'GroupName'
$colName.HeaderText = 'Group Name'
$colName.Width = 285
$colName.AutoSizeMode = 'None'
$colName.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colName)

$colDescription = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDescription.Name = 'Description'
$colDescription.HeaderText = 'Description'
$colDescription.AutoSizeMode = 'Fill'
$colDescription.FillWeight = 45
$colDescription.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colDescription)

$colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colType.Name = 'Type'
$colType.HeaderText = 'Type'
$colType.Width = 225
$colType.AutoSizeMode = 'None'
$colType.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colType)

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = 'TargetStatus'
$colStatus.HeaderText = 'Target Status'
$colStatus.Width = 115
$colStatus.AutoSizeMode = 'None'
$colStatus.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colStatus)

# Activity log
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Activity Log'
$grpLog.Location = New-Object System.Drawing.Point(20, 625)
$grpLog.Size = New-Object System.Drawing.Size(1135, 105)
$grpLog.Anchor = 'Bottom,Left,Right'
$MainForm.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 18)
$txtLog.Size = New-Object System.Drawing.Size(1115, 77)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$grpLog.Controls.Add($txtLog)
$script:txtLog = $txtLog

# Status bar
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = if ($script:CurrentServer) { "Ready. Using AD server $($script:CurrentServer)." } else { 'Ready.' }
[void]$statusStrip.Items.Add($lblStatus)
$MainForm.Controls.Add($statusStrip)
$script:lblStatus = $lblStatus

Apply-MSToolkitSharedTheme -Root $MainForm

# The shared theme pass cannot infer the warning colour; set it now and after every
# theme toggle.
$script:MSToolkitThemeRefreshHook = {
    if ($script:lblUsersOUWarning) {
        $NotePalette = Get-MSToolkitThemePalette
        $script:lblUsersOUWarning.ForeColor = if ($script:UserListNoteIsWarning) { $NotePalette.Warning } else { $NotePalette.MutedText }
    }
}
& $script:MSToolkitThemeRefreshHook

Write-AppLog 'AD Group Compare initialized.'
if ($script:CurrentServer) {
    Write-AppLog "Using AD server: $($script:CurrentServer)"
}
if ($script:ServerNote) {
    Write-AppLog $script:ServerNote 'WARNING'
}

$MainForm.Add_Shown({
    $MainForm.Activate()

    $LoadedCount = Initialize-MSToolkitUserPicker -Combos @($txtReference, $txtTarget) -Server $script:CurrentServer

    if ($LoadedCount -ge 0) {
        Write-AppLog "Loaded $LoadedCount employee account(s) into the user lists."
        if (-not (Get-MSToolkitSetting -Name "OUUsers")) {
            Write-AppLog 'No Users OU is set in MSToolkit Settings, so the user list covers the whole domain. Set it under Settings > Organizational units to narrow it.' 'WARNING'
        }
    }
    else {
        Write-AppLog 'Could not load the employee user list. Type a username instead.' 'WARNING'
    }

    $txtReference.Focus()
})

# Type-ahead filtering on every editable dropdown. Its own handler, not folded
# into another, so it runs on every launch rather than only when the tool is
# started with parameters - and so a failure here cannot stop the rest of
# start-up. Multiple Shown handlers chain.
$MainForm.Add_Shown({ Register-MSToolkitComboFiltersOn -Root $MainForm })

[void]$MainForm.ShowDialog()
