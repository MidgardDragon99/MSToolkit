#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet("Light","Dark")]
    [string]$ThemeMode,

    # The SharePoint admin center address. MSToolkit passes it from its Settings
    # (SharePoint admin URL, or derived from the onmicrosoft domain). It only
    # pre-fills the header box; a value changed there is remembered per user.
    [string]$AdminUrl
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# ===========================================================================
#  M365-OneDrive-SharePoint-Tools.ps1
#  OneDrive administration: grant and remove access to someone's OneDrive,
#  audit who already has it, handle an offboarding handover, restore a deleted
#  OneDrive, and report storage and sharing.
#
#  Read-only apart from three actions, each behind its own confirmation:
#  granting access, removing access, and restoring a deleted site.
#
#  Sign in with a STANDARD account holding the SharePoint Administrator role.
#  This connects to SharePoint Online, not Exchange. The tenant admin URL is
#  filled in from MSToolkit Settings and can be changed in the header.
#
#  Launched from MSToolkit, -ThemeMode carries the console's current theme.
# ===========================================================================

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
            $ToolProperty = "Theme_M365DistributionGroupCompare"
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



if (-not ("MSToolkitConsole.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace MSToolkitConsole
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

function Hide-PowerShellConsole {
    try {
        $ConsoleHandle = [MSToolkitConsole.NativeMethods]::GetConsoleWindow()
        if ($ConsoleHandle -ne [IntPtr]::Zero) {
            [MSToolkitConsole.NativeMethods]::ShowWindow($ConsoleHandle, 0) | Out-Null
        }
    }
    catch {
        # Console hiding must never prevent the GUI from loading.
    }
}



if (-not ("MSToolkitAuthWindow.AuthWindowWatcher" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace MSToolkitAuthWindow
{
    public static class AuthWindowWatcher
    {
        private const uint MONITOR_DEFAULTTONEAREST = 2;
        private const uint SWP_NOSIZE = 0x0001;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

        private static Timer _timer;
        private static IntPtr _anchor = IntPtr.Zero;
        private static DateTime _expiresUtc;
        private static HashSet<IntPtr> _existingWindows = new HashSet<IntPtr>();

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int X,
            int Y,
            int cx,
            int cy,
            uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr GetShellWindow();

        public static void Start(IntPtr anchorWindow)
        {
            Stop();

            _anchor = anchorWindow;
            _expiresUtc = DateTime.UtcNow.AddSeconds(60);
            _existingWindows = new HashSet<IntPtr>();

            // Snapshot all windows that already exist before Exchange Online auth starts.
            EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
            {
                _existingWindows.Add(hWnd);
                return true;
            }, IntPtr.Zero);

            _timer = new Timer(CheckForNewWindow, null, 200, 200);
        }

        public static void Stop()
        {
            Timer timer = _timer;
            _timer = null;

            if (timer != null)
            {
                try { timer.Dispose(); }
                catch { }
            }

            _anchor = IntPtr.Zero;
            _existingWindows = new HashSet<IntPtr>();
        }

        private static void CheckForNewWindow(object state)
        {
            if (_anchor == IntPtr.Zero || DateTime.UtcNow >= _expiresUtc)
            {
                Stop();
                return;
            }

            try
            {
                IntPtr shellWindow = GetShellWindow();
                IntPtr monitor = MonitorFromWindow(_anchor, MONITOR_DEFAULTTONEAREST);

                if (monitor == IntPtr.Zero)
                    return;

                MONITORINFO mi = new MONITORINFO();
                mi.cbSize = Marshal.SizeOf(typeof(MONITORINFO));

                if (!GetMonitorInfo(monitor, ref mi))
                    return;

                IntPtr candidate = IntPtr.Zero;

                EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
                {
                    if (candidate != IntPtr.Zero)
                        return false;

                    if (hWnd == _anchor || hWnd == shellWindow)
                        return true;

                    if (_existingWindows.Contains(hWnd))
                        return true;

                    if (!IsWindowVisible(hWnd))
                        return true;

                    RECT rect;
                    if (!GetWindowRect(hWnd, out rect))
                        return true;

                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;

                    // Ignore tiny utility/tool-tip windows.
                    if (width < 300 || height < 180)
                        return true;

                    candidate = hWnd;
                    return false;
                }, IntPtr.Zero);

                if (candidate == IntPtr.Zero)
                    return;

                RECT candidateRect;
                if (!GetWindowRect(candidate, out candidateRect))
                    return;

                int candidateWidth = candidateRect.Right - candidateRect.Left;
                int candidateHeight = candidateRect.Bottom - candidateRect.Top;

                int workWidth = mi.rcWork.Right - mi.rcWork.Left;
                int workHeight = mi.rcWork.Bottom - mi.rcWork.Top;

                int x = mi.rcWork.Left + Math.Max(0, (workWidth - candidateWidth) / 2);
                int y = mi.rcWork.Top + Math.Max(0, (workHeight - candidateHeight) / 2);

                // Move exactly one newly-created auth window, foreground it once,
                // then stop so no other windows are touched.
                ShowWindow(candidate, 5);
                SetWindowPos(candidate, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);
                SetWindowPos(candidate, HWND_NOTOPMOST, x, y, 0, 0, SWP_NOSIZE);
                SetForegroundWindow(candidate);

                Stop();
            }
            catch
            {
                // Never allow window positioning to interfere with authentication.
            }
        }
    }
}
"@
}

function Start-M365AuthWindowWatcher {
    try {
        if ($script:MainForm -and -not $script:MainForm.IsDisposed) {
            $null = $script:MainForm.Handle
            [MSToolkitAuthWindow.AuthWindowWatcher]::Start($script:MainForm.Handle)
        }
    }
    catch {
        # Window positioning is best effort only.
    }
}

function Stop-M365AuthWindowWatcher {
    try {
        [MSToolkitAuthWindow.AuthWindowWatcher]::Stop()
    }
    catch {
        # Window positioning is best effort only.
    }
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:SPOConnected = $false

function Show-ErrorMessage {
    param([string]$Message, [string]$Title = 'Error')
    [System.Windows.Forms.MessageBox]::Show($Message,$Title,[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Show-InfoMessage {
    param([string]$Message, [string]$Title = 'Information')
    [System.Windows.Forms.MessageBox]::Show($Message,$Title,[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Write-AppLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARN','ERROR')][string]$Level = 'INFO'
    )

    if (-not $script:txtLog) { return }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message`r`n"

    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.SelectionLength = 0

    switch ($Level) {
        'SUCCESS' { $script:txtLog.SelectionColor = (Get-MSToolkitThemePalette).Success }
        'WARN'    { $script:txtLog.SelectionColor = (Get-MSToolkitThemePalette).Warning }
        'ERROR'   { $script:txtLog.SelectionColor = (Get-MSToolkitThemePalette).Danger }
        default   { $script:txtLog.SelectionColor = (Get-MSToolkitThemePalette).Text }
    }

    $script:txtLog.AppendText($line)
    $script:txtLog.SelectionColor = $script:txtLog.ForeColor
    $script:txtLog.ScrollToCaret()
}

function Set-BusyState {
    param([bool]$Busy,[string]$StatusText = '')

    $script:MainForm.UseWaitCursor = $Busy

    # A ComboBox owns its own window handle, so clearing UseWaitCursor on the form
    # does not always restore the pointer over it. Reset the cursor explicitly.
    if (-not $Busy) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    if ($script:btnConnect) { $script:btnConnect.Enabled = -not $Busy }

    # Every control that needs a live connection registers itself here, so a new
    # tab does not have to be added to this function.
    #
    # Deliberately a plain array walked by index: wrapping a generic List in @()
    # threw "Argument types do not match" here and aborted the caller, including
    # the sign-in. Enabling a button is cosmetic and must never do that, so the
    # whole loop is guarded as well.
    $Enable = [bool]((-not $Busy) -and $script:SPOConnected)

    try {
        $Controls = $script:ActionControls

        if ($null -ne $Controls) {
            for ($Index = 0; $Index -lt $Controls.Count; $Index++) {
                $Control = $Controls[$Index]
                if ($Control -is [System.Windows.Forms.Control]) {
                    $Control.Enabled = $Enable
                }
            }
        }
    }
    catch {
        Write-AppLog "Could not update the toolbar state: $($_.Exception.Message)" 'WARN'
    }

    if ($StatusText) { $script:lblStatus.Text = $StatusText }
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------------------
# Connection
#
# SharePoint Online needs the tenant admin URL, which cannot be discovered
# before connecting, so it is asked for once and remembered alongside the
# MSToolkit theme preference.
# ---------------------------------------------------------------------------
function Get-MSToolkitSettingsPath {
    return (Join-Path (Join-Path $env:APPDATA "MSToolkit") "settings.json")
}

# The tenant admin URL. Nothing organization-specific is stored here: the default
# comes from MSToolkit Settings through -AdminUrl, and a value changed in the
# header box is then remembered in its place for this account.
$script:DefaultAdminUrl = if ([string]::IsNullOrWhiteSpace($AdminUrl)) { '' } else { $AdminUrl.Trim() }

function Get-SPOAdminUrlPreference {
    try {
        $Path = Get-MSToolkitSettingsPath
        if (Test-Path -LiteralPath $Path) {
            $Saved = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($Saved.PSObject.Properties.Name -contains 'SPOAdminUrl') {
                $Url = [string]$Saved.SPOAdminUrl
                if (-not [string]::IsNullOrWhiteSpace($Url)) { return $Url }
            }
        }
    }
    catch { }

    return $script:DefaultAdminUrl
}

function Save-SPOAdminUrlPreference {
    param([string]$Url)

    # Read-modify-write: Theme, SelectedDC and the per-tool theme keys share this file.
    try {
        $Path = Get-MSToolkitSettingsPath
        $Folder = Split-Path -Parent $Path

        if (-not (Test-Path -LiteralPath $Folder)) {
            New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        }

        $Settings = $null
        if (Test-Path -LiteralPath $Path) {
            try { $Settings = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $Settings = $null }
        }

        if ($null -eq $Settings) { $Settings = [pscustomobject]@{} }

        if ($Settings.PSObject.Properties.Name -contains 'SPOAdminUrl') {
            $Settings.SPOAdminUrl = $Url
        }
        else {
            $Settings | Add-Member -MemberType NoteProperty -Name 'SPOAdminUrl' -Value $Url
        }

        $Settings | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
    }
    catch {
        Write-AppLog "Could not save the admin URL: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-SPOModule {
    if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
        return $true
    }

    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "The Microsoft.Online.SharePoint.PowerShell module is required and is not installed.`r`n`r`nInstall it for the current user now?",
        'SharePoint Online Module Required',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }

    try {
        Set-BusyState -Busy $true -StatusText 'Installing the SharePoint Online module...'
        Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
        Write-AppLog 'Installed Microsoft.Online.SharePoint.PowerShell successfully.' 'SUCCESS'
        return $true
    }
    catch {
        Show-ErrorMessage "Unable to install the SharePoint Online module.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Module installation failed: $($_.Exception.Message)" 'ERROR'
        return $false
    }
    finally {
        Set-BusyState -Busy $false
    }
}

function Connect-M365SharePoint {
    if (-not (Ensure-SPOModule)) { return }

    $AdminUrl = "$($script:txtAdminUrl.Text)".Trim().TrimEnd('/')

    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        Show-InfoMessage "Enter the SharePoint admin URL first, for example:`r`n`r`nhttps://contoso-admin.sharepoint.com"
        return
    }

    if ($AdminUrl -notmatch '^https://[\w-]+-admin\.sharepoint\.com$') {
        $Answer = [System.Windows.Forms.MessageBox]::Show(
            "'$AdminUrl' does not look like a tenant admin URL.`r`n`r`nThe expected form is https://<tenant>-admin.sharepoint.com`r`n`r`nTry it anyway?",
            'Admin URL',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Connecting to SharePoint Online...'
        Write-AppLog "Connecting to $AdminUrl."

        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch { }

        Start-M365AuthWindowWatcher
        Connect-SPOService -Url $AdminUrl -ErrorAction Stop

        # Verify the session with one lightweight tenant read.
        $Tenant = Get-SPOTenant -ErrorAction Stop

        $script:SPOConnected = $true
        $script:AdminUrl = $AdminUrl
        Save-SPOAdminUrlPreference -Url $AdminUrl

        # Just the tenant name - the full admin URL does not fit and adds nothing.
        $TenantName = $AdminUrl -replace '^https://', '' -replace '-admin\.sharepoint\.com/?$', ''
        $script:lblConnection.Text = "Connected: $TenantName"
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Success

        try {
            for ($Index = 0; $Index -lt $script:ActionControls.Count; $Index++) {
                $Control = $script:ActionControls[$Index]
                if ($Control -is [System.Windows.Forms.Control]) { $Control.Enabled = $true }
            }
        }
        catch {
            Write-AppLog "Could not enable the toolbar: $($_.Exception.Message)" 'WARN'
        }

        Write-AppLog "Connected to SharePoint Online." 'SUCCESS'

        if ($null -ne $Tenant.OrphanedPersonalSitesRetentionPeriod) {
            $script:RetentionDays = [int]$Tenant.OrphanedPersonalSitesRetentionPeriod
            Write-AppLog "A OneDrive is kept for $($script:RetentionDays) day(s) after its owner is deleted."
        }

        $script:OneDriveCache = $null
        $script:PickersLoaded = $false
        $script:lblStatus.Text = 'Connected. Loading the OneDrive list...'

        # The Grant Access tab is already open, so its tab-change event never fires.
        Initialize-OdPickers
    }
    catch {
        $script:SPOConnected = $false
        $script:lblConnection.Text = 'Not connected'
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Danger

        $Failure = $_
        Write-AppLog "Type: $($Failure.Exception.GetType().FullName)" 'ERROR'
        Write-AppLog "Message: $($Failure.Exception.Message)" 'ERROR'

        $Module = Get-Module Microsoft.Online.SharePoint.PowerShell | Select-Object -First 1
        if ($Module) { Write-AppLog "Module version: $($Module.Version)" 'ERROR' }

        Show-ErrorMessage "Unable to connect to SharePoint Online.`r`n`r`n$($Failure.Exception.Message)`r`n`r`nThe account needs the SharePoint Administrator role."
    }
    finally {
        Stop-M365AuthWindowWatcher
        Set-BusyState -Busy $false
    }
}

# ---------------------------------------------------------------------------
# OneDrive lookup
# ---------------------------------------------------------------------------
function Get-MSToolkitUserPickerLabel {
    param($Site)

    $Owner = [string]$Site.Owner
    $Title = [string]$Site.Title

    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = $Owner }
    return "$Title ($Owner)"
}

function ConvertFrom-MSToolkitUserPickerLabel {
    param([string]$Value)

    $Text = "$Value".Trim()

    # Picker entries look like "Jane Smith (jane.smith@example.com)". Anything typed
    # by hand is passed through unchanged.
    if ($Text -match '\(([^()]+)\)\s*$') { return $Matches[1].Trim() }
    return $Text
}

function Get-MSToolkitOneDriveSites {
    param([switch]$Refresh)

    if ($script:OneDriveCache -and -not $Refresh) { return $script:OneDriveCache }

    Set-BusyState -Busy $true -StatusText 'Loading OneDrive sites...'
    Write-AppLog 'Loading the OneDrive site list. On a large tenant this takes a moment.'

    $script:OneDriveCache = @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" -ErrorAction Stop |
        Sort-Object Title)

    Write-AppLog "Loaded $($script:OneDriveCache.Count) OneDrive site(s)." 'SUCCESS'
    Set-BusyState -Busy $false
    return $script:OneDriveCache
}

function Initialize-MSToolkitUserPicker {
    param($Combos)

    try {
        $Sites = Get-MSToolkitOneDriveSites

        foreach ($Combo in $Combos) {
            if ($Combo -isnot [System.Windows.Forms.ComboBox]) { continue }

            $Existing = $Combo.Text
            $Combo.Items.Clear()

            foreach ($Site in $Sites) {
                $null = $Combo.Items.Add((Get-MSToolkitUserPickerLabel -Site $Site))
            }

            $Combo.Text = $Existing
        }

        return @($Sites).Count
    }
    catch {
        Write-AppLog "Could not load the OneDrive list: $($_.Exception.Message). Type an address instead." 'WARN'
        return -1
    }
}

function Resolve-OneDriveSite {
    param([Parameter(Mandatory)][string]$Identity)

    $Identity = (ConvertFrom-MSToolkitUserPickerLabel -Value $Identity).Trim()
    if ([string]::IsNullOrWhiteSpace($Identity)) { throw 'A user was not entered.' }

    # A OneDrive URL can be derived from the address, but only the site list proves
    # the site exists, so try the list first and fall back to a direct read.
    $Sites = @(Get-MSToolkitOneDriveSites | Where-Object {
        "$($_.Owner)" -ieq $Identity -or "$($_.Title)" -ieq $Identity -or "$($_.Url)" -ieq $Identity
    })

    if ($Sites.Count -eq 1) { return $Sites[0] }

    if ($Sites.Count -gt 1) {
        throw "More than one OneDrive matches '$Identity'. Use the full address."
    }

    $Partial = @(Get-MSToolkitOneDriveSites | Where-Object {
        "$($_.Owner)" -like "*$Identity*" -or "$($_.Title)" -like "*$Identity*"
    })

    if ($Partial.Count -eq 1) { return $Partial[0] }
    if ($Partial.Count -gt 1) { throw "'$Identity' matches $($Partial.Count) OneDrive sites. Use the full address." }

    try {
        return (Get-SPOSite -Identity $Identity -ErrorAction Stop)
    }
    catch {
        throw "No OneDrive was found for '$Identity'. The owner may never have signed in to OneDrive."
    }
}

function Get-SPOSiteAdmins {
    param([string]$Url)

    return @(Get-SPOUser -Site $Url -Limit All -ErrorAction Stop | Where-Object { $_.IsSiteAdmin })
}

# ---------------------------------------------------------------------------
# Grant and remove access
# ---------------------------------------------------------------------------
function Show-OneDriveAccessState {
    param($Site)

    $script:gridAccess.Rows.Clear()
    $script:lblAccessUrl.Text = [string]$Site.Url
    $script:CurrentSite = $Site

    $Admins = Get-SPOSiteAdmins -Url $Site.Url
    $Owner = [string]$Site.Owner

    foreach ($Admin in $Admins) {
        $Login = [string]$Admin.LoginName
        $IsOwner = ($Login -ieq $Owner)

        $Kind = 'Granted access'
        if ($IsOwner) { $Kind = 'Owner' }

        $Index = $script:gridAccess.Rows.Add(
            $Kind,
            [string]$Admin.DisplayName,
            $Login,
            [string]$Admin.IsGroup
        )

        $script:gridAccess.Rows[$Index].Tag = $Admin

        if ($IsOwner) { Set-ExoRowTone -Row $script:gridAccess.Rows[$Index] -Tone 'Success' }
        else          { Set-ExoRowTone -Row $script:gridAccess.Rows[$Index] -Tone 'Warning' }
    }

    $Extra = @($Admins | Where-Object { "$($_.LoginName)" -ine $Owner }).Count
    $script:lblAccessSummary.Text = "$($Site.Title) - $Extra account(s) with granted access besides the owner."

    if ($Extra -gt 0) { $script:lblAccessSummary.ForeColor = (Get-MSToolkitThemePalette).Warning }
    else              { $script:lblAccessSummary.ForeColor = (Get-MSToolkitThemePalette).Success }
}

function Invoke-OneDriveLookup {
    if (-not $script:SPOConnected) { return }

    $Identity = $script:cboAccessOwner.Text

    try {
        Set-BusyState -Busy $true -StatusText 'Finding the OneDrive...'
        $Site = Resolve-OneDriveSite -Identity $Identity
        Write-AppLog "OneDrive for $($Site.Owner): $($Site.Url)"
        Show-OneDriveAccessState -Site $Site
        Set-BusyState -Busy $false -StatusText "Found $($Site.Url)"
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'OneDrive lookup failed.'
        Show-ErrorMessage $_.Exception.Message
        Write-AppLog "OneDrive lookup failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-OneDriveGrant {
    if (-not $script:SPOConnected -or -not $script:CurrentSite) {
        Show-InfoMessage 'Find a OneDrive first.'
        return
    }

    $Requester = (ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboAccessRequester.Text).Trim()

    if ([string]::IsNullOrWhiteSpace($Requester)) {
        Show-InfoMessage 'Enter the person who needs access.'
        return
    }

    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Give $Requester full access to this OneDrive?`r`n`r`nOwner: $($script:CurrentSite.Owner)`r`nSite: $($script:CurrentSite.Url)`r`n`r`nThey become a site collection administrator and can read, change and delete everything in it, including files the owner never shared. Remove the access when it is no longer needed.",
        'Grant OneDrive access',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog 'Grant cancelled.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText "Granting access to $Requester..."
        Set-SPOUser -Site $script:CurrentSite.Url -LoginName $Requester -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
        Write-AppLog "Granted $Requester access to $($script:CurrentSite.Url)." 'SUCCESS'
        Show-OneDriveAccessState -Site $script:CurrentSite
        Set-BusyState -Busy $false -StatusText "Granted access to $Requester."

        $Open = [System.Windows.Forms.MessageBox]::Show(
            "Access granted.`r`n`r`nOpen the OneDrive in a browser now?",
            'Granted',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information)

        if ($Open -eq [System.Windows.Forms.DialogResult]::Yes) { Invoke-OneDriveOpen }
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Grant failed.'
        Show-ErrorMessage "Could not grant access.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Grant failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-OneDriveRemoveAccess {
    if (-not $script:SPOConnected -or -not $script:CurrentSite) { return }

    $Selected = @($script:gridAccess.SelectedRows | Where-Object { $_.Tag })

    if ($Selected.Count -eq 0) {
        Show-InfoMessage 'Select the account whose access should be removed.'
        return
    }

    $Owner = [string]$script:CurrentSite.Owner
    $Targets = @($Selected | ForEach-Object { [string]$_.Tag.LoginName } | Where-Object { $_ -ine $Owner })

    if ($Targets.Count -eq 0) {
        Show-InfoMessage "That is the owner's own access, which this tool will not remove."
        return
    }

    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Remove access for:`r`n`r`n  $($Targets -join "`r`n  ")`r`n`r`nFrom: $($script:CurrentSite.Url)",
        'Remove OneDrive access',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    foreach ($Target in $Targets) {
        try {
            Set-BusyState -Busy $true -StatusText "Removing $Target..."
            Set-SPOUser -Site $script:CurrentSite.Url -LoginName $Target -IsSiteCollectionAdmin $false -ErrorAction Stop | Out-Null
            Write-AppLog "Removed $Target from $($script:CurrentSite.Url)." 'SUCCESS'
        }
        catch {
            Write-AppLog "Could not remove $($Target): $($_.Exception.Message)" 'ERROR'
        }
    }

    Show-OneDriveAccessState -Site $script:CurrentSite
    Set-BusyState -Busy $false -StatusText 'Access removed.'
}

function Invoke-OneDriveOpen {
    if (-not $script:CurrentSite) {
        Show-InfoMessage 'Find a OneDrive first.'
        return
    }

    try {
        Start-Process ([string]$script:CurrentSite.Url) -ErrorAction Stop
        Write-AppLog "Opened $($script:CurrentSite.Url) in the default browser."
    }
    catch {
        [System.Windows.Forms.Clipboard]::SetText([string]$script:CurrentSite.Url)
        Write-AppLog 'Could not open a browser; the URL was copied to the clipboard instead.' 'WARN'
    }
}

# ---------------------------------------------------------------------------
# Access audit across every OneDrive
# ---------------------------------------------------------------------------
function Invoke-OneDriveAccessAudit {
    if (-not $script:SPOConnected) { return }

    try {
        $Sites = Get-MSToolkitOneDriveSites -Refresh:$script:chkAuditRefresh.Checked

        $Answer = [System.Windows.Forms.MessageBox]::Show(
            "Check all $($Sites.Count) OneDrive site(s) for extra administrators?`r`n`r`nEach site is a separate query, so this takes roughly a second per site.",
            'Access audit',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $script:gridAudit.Rows.Clear()
        $Position = 0
        $Found = 0
        $Unreadable = 0

        foreach ($Site in $Sites) {
            $Position++
            Set-BusyState -Busy $true -StatusText "Checking $Position of $($Sites.Count): $($Site.Owner)"

            try {
                $Extra = @(Get-SPOSiteAdmins -Url $Site.Url | Where-Object { "$($_.LoginName)" -ine "$($Site.Owner)" })

                foreach ($Admin in $Extra) {
                    $Index = $script:gridAudit.Rows.Add(
                        [string]$Site.Title,
                        [string]$Site.Owner,
                        [string]$Admin.DisplayName,
                        [string]$Admin.LoginName,
                        [string]$Admin.IsGroup,
                        [string]$Site.Url
                    )
                    Set-ExoRowTone -Row $script:gridAudit.Rows[$Index] -Tone 'Warning'
                    $Found++
                }
            }
            catch {
                # Expected for most OneDrives: listing the users of a personal site
                # needs access to that site, which a SharePoint admin does not have
                # by default. Not an error, just a gap in the picture.
                $Reason = 'No access to read this site'
                if ("$($_.Exception.Message)" -notmatch 'Access is denied|blocked') {
                    $Reason = $_.Exception.Message
                }

                $Index = $script:gridAudit.Rows.Add([string]$Site.Title,[string]$Site.Owner,'(not readable)',$Reason,'',[string]$Site.Url)
                Set-ExoRowTone -Row $script:gridAudit.Rows[$Index] -Tone 'Normal'
                $Unreadable++
            }
        }

        $script:lblAuditSummary.Text = "$Found extra administrator entry(s) across $($Sites.Count) OneDrive site(s). $Unreadable site(s) could not be read."
        Write-AppLog "Audit complete: $Found extra administrator entry(s), $Unreadable site(s) not readable." 'SUCCESS'

        if ($Unreadable -gt 0) {
            Write-AppLog "Listing the users of a OneDrive needs access to that site, which a SharePoint administrator does not get automatically. Grant yourself access on the Grant Access tab to audit a specific one." 'WARN'
        }
        Set-BusyState -Busy $false -StatusText 'Access audit complete.'
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Access audit failed.'
        Show-ErrorMessage "The access audit failed.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Access audit failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Offboarding handover
# ---------------------------------------------------------------------------
function Add-OffboardRow {
    param([string]$Name,$Value,[string]$Tone = 'Normal')

    $Index = $script:gridOffboard.Rows.Add($Name,[string]$Value)
    if ($Tone -ne 'Normal') { Set-ExoRowTone -Row $script:gridOffboard.Rows[$Index] -Tone $Tone }
}

function Invoke-OneDriveOffboardReport {
    if (-not $script:SPOConnected) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading the OneDrive...'
        $script:gridOffboard.Rows.Clear()

        $Site = Get-SPOSite -Identity (Resolve-OneDriveSite -Identity $script:cboOffboardUser.Text).Url -Detailed -ErrorAction Stop
        $script:OffboardSite = $Site

        Add-OffboardRow -Name 'Owner' -Value $Site.Owner
        Add-OffboardRow -Name 'Site' -Value $Site.Url
        Add-OffboardRow -Name 'Status' -Value $Site.Status

        $UsedMb = [math]::Round(([double]$Site.StorageUsageCurrent),1)
        $QuotaMb = [double]$Site.StorageQuota
        $Percent = 0
        if ($QuotaMb -gt 0) { $Percent = [math]::Round(($UsedMb / $QuotaMb) * 100,1) }

        Add-OffboardRow -Name 'Storage used' -Value "$UsedMb MB of $QuotaMb MB ($Percent%)"
        Add-OffboardRow -Name 'Last content change' -Value $Site.LastContentModifiedDate

        $LockTone = 'Normal'
        if ("$($Site.LockState)" -ne 'Unlock') { $LockTone = 'Warning' }
        Add-OffboardRow -Name 'Lock state' -Value $Site.LockState -Tone $LockTone

        $SharingTone = 'Normal'
        if ("$($Site.SharingCapability)" -match 'External|Anonymous') { $SharingTone = 'Warning' }
        Add-OffboardRow -Name 'Sharing' -Value $Site.SharingCapability -Tone $SharingTone

        # Whether the owner still exists decides whether the retention clock is
        # running - but "access denied" is not the same as "account deleted", and
        # reading a personal site's users often is denied.
        try {
            $null = Get-SPOUser -Site $Site.Url -LoginName $Site.Owner -ErrorAction Stop
            Add-OffboardRow -Name 'Owner account' -Value 'Still present - the retention clock has not started'
        }
        catch {
            if ("$($_.Exception.Message)" -match 'Access is denied|blocked') {
                Add-OffboardRow -Name 'Owner account' -Value 'Could not check - this site cannot be read without access to it' -Tone 'Warning'
            }
            else {
                Add-OffboardRow -Name 'Owner account' -Value 'Not found - this OneDrive is on the retention clock' -Tone 'Danger'

                $Days = $script:RetentionDays
                if (-not $Days) { $Days = 30 }
                Add-OffboardRow -Name 'Retention period' -Value "$Days day(s) after the owner was deleted, then the OneDrive is removed" -Tone 'Danger'
            }
        }

        try {
            $Admins = @(Get-SPOSiteAdmins -Url $Site.Url | Where-Object { "$($_.LoginName)" -ine "$($Site.Owner)" })

            if ($Admins.Count -eq 0) {
                Add-OffboardRow -Name 'Granted access' -Value 'Nobody besides the owner'
            }
            else {
                foreach ($Admin in $Admins) {
                    Add-OffboardRow -Name 'Granted access' -Value "$($Admin.DisplayName) <$($Admin.LoginName)>" -Tone 'Warning'
                }
            }
        }
        catch {
            Add-OffboardRow -Name 'Granted access' -Value 'Could not read - grant yourself access to this OneDrive first' -Tone 'Warning'
        }

        $script:lblOffboardSummary.Text = "$($Site.Title) - $UsedMb MB, last changed $($Site.LastContentModifiedDate)"
        Write-AppLog "Offboarding report built for $($Site.Owner)." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText 'Offboarding report complete.'
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Offboarding report failed.'
        Show-ErrorMessage "Could not build the report.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Offboarding report failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-OneDriveHandover {
    if (-not $script:SPOConnected -or -not $script:OffboardSite) {
        Show-InfoMessage 'Build the report first.'
        return
    }

    $Manager = (ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboOffboardManager.Text).Trim()

    if ([string]::IsNullOrWhiteSpace($Manager)) {
        Show-InfoMessage 'Enter the manager who takes over the files.'
        return
    }

    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Give $Manager full access to this OneDrive?`r`n`r`nOwner: $($script:OffboardSite.Owner)`r`nSite: $($script:OffboardSite.Url)`r`n`r`nThey can read, change and delete everything in it. Remove the access once the handover is finished.",
        'Offboarding handover',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        Set-BusyState -Busy $true -StatusText "Granting access to $Manager..."
        Set-SPOUser -Site $script:OffboardSite.Url -LoginName $Manager -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
        Write-AppLog "Handover: granted $Manager access to $($script:OffboardSite.Url)." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText "Granted access to $Manager."
        Invoke-OneDriveOffboardReport
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Handover failed.'
        Show-ErrorMessage "Could not grant access.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Handover failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Deleted OneDrive restore
# ---------------------------------------------------------------------------
function Invoke-DeletedSiteSearch {
    if (-not $script:SPOConnected) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading deleted sites...'
        $script:gridDeleted.Rows.Clear()

        $Sites = @(Get-SPODeletedSite -IncludePersonalSite -Limit All -ErrorAction Stop | Sort-Object DeletionTime -Descending)
        $Filter = "$($script:txtDeletedFilter.Text)".Trim()

        if ($Filter) {
            $Sites = @($Sites | Where-Object { "$($_.Url)" -like "*$Filter*" -or "$($_.Title)" -like "*$Filter*" })
        }

        foreach ($Site in $Sites) {
            $Remaining = ''
            $Tone = 'Normal'

            if ($Site.DaysRemaining -ne $null) {
                $Remaining = [string]$Site.DaysRemaining

                if ([int]$Site.DaysRemaining -le 7)       { $Tone = 'Danger' }
                elseif ([int]$Site.DaysRemaining -le 21)  { $Tone = 'Warning' }
            }

            $Index = $script:gridDeleted.Rows.Add(
                [string]$Site.Url,
                [string]$Site.Title,
                $Site.DeletionTime,
                $Remaining,
                [string]$Site.SiteId,
                [string]$Site.Status
            )

            $script:gridDeleted.Rows[$Index].Tag = $Site
            Set-ExoRowTone -Row $script:gridDeleted.Rows[$Index] -Tone $Tone
        }

        # Two different tenant settings drive these countdowns, so show both rather
        # than leaving "days left" unexplained.
        $Retention = 'unknown'
        $DeletedRetention = ''

        try {
            $Tenant = Get-SPOTenant -ErrorAction Stop

            if ($null -ne $Tenant.OrphanedPersonalSitesRetentionPeriod) {
                $script:RetentionDays = [int]$Tenant.OrphanedPersonalSitesRetentionPeriod
                $Retention = "$($script:RetentionDays) day(s)"
            }

            if ($null -ne $Tenant.DeletedUserPersonalSiteRetentionPeriod) {
                $DeletedRetention = "  |  Deleted-user OneDrive retention: $($Tenant.DeletedUserPersonalSiteRetentionPeriod) day(s)"
            }
        }
        catch { }

        $script:lblDeletedSummary.Text = "$($Sites.Count) deleted site(s). Red: a week or less left. OneDrive kept after owner deletion: $Retention$DeletedRetention"
        Write-AppLog "Found $($Sites.Count) deleted site(s). Tenant OneDrive retention after owner deletion: $Retention." 'SUCCESS'
        Write-AppLog 'Days left comes from SharePoint itself. A Purview retention policy or a legal hold can keep content beyond this, and is not visible here - check Purview if the content matters.'
        Set-BusyState -Busy $false -StatusText "$($Sites.Count) deleted site(s)."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Deleted site search failed.'
        Show-ErrorMessage "Could not read the deleted sites.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Deleted site search failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-DeletedSiteRestore {
    if (-not $script:SPOConnected) { return }

    $Selected = @($script:gridDeleted.SelectedRows | Where-Object { $_.Tag })

    if ($Selected.Count -ne 1) {
        Show-InfoMessage 'Select exactly one deleted site to restore.'
        return
    }

    $Site = $Selected[0].Tag

    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Restore this site?`r`n`r`n$($Site.Url)`r`nDeleted: $($Site.DeletionTime)`r`n`r`nThe content comes back as it was. If the owner's account is gone, grant somebody access afterwards or it stays unreachable.",
        'Restore deleted site',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        Set-BusyState -Busy $true -StatusText "Restoring $($Site.Url)..."
        Restore-SPODeletedSite -Identity $Site.Url -NoWait -ErrorAction Stop | Out-Null
        Write-AppLog "Restore started for $($Site.Url). It can take several minutes to appear." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText 'Restore started.'
        Show-InfoMessage "Restore started for:`r`n`r`n$($Site.Url)`r`n`r`nIt can take several minutes. Search the deleted sites again to confirm it has gone from this list."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Restore failed.'
        Show-ErrorMessage "Could not restore the site.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Restore failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Storage and sharing report
# ---------------------------------------------------------------------------
function Invoke-StorageSharingReport {
    if (-not $script:SPOConnected) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Building the storage and sharing report...'
        $script:gridStorage.Rows.Clear()

        if ($script:cboReportScope.SelectedItem -eq 'All SharePoint sites') {
            $Sites = @(Get-SPOSite -Limit All -ErrorAction Stop | Sort-Object Url)
        }
        else {
            $Sites = @(Get-MSToolkitOneDriveSites -Refresh)
        }

        $OverQuota = 0
        $OpenSharing = 0

        foreach ($Site in $Sites) {
            $UsedMb = [math]::Round(([double]$Site.StorageUsageCurrent),1)
            $QuotaMb = [double]$Site.StorageQuota
            $Percent = 0
            if ($QuotaMb -gt 0) { $Percent = [math]::Round(($UsedMb / $QuotaMb) * 100,1) }

            $Sharing = [string]$Site.SharingCapability

            $Index = $script:gridStorage.Rows.Add(
                [string]$Site.Title,
                [string]$Site.Owner,
                $UsedMb,
                $QuotaMb,
                $Percent,
                $Sharing,
                $Site.LastContentModifiedDate,
                [string]$Site.Url
            )

            if ($Percent -ge 90) {
                Set-ExoRowTone -Row $script:gridStorage.Rows[$Index] -Tone 'Danger'
                $OverQuota++
            }
            elseif ($Sharing -match 'Anonymous|ExternalUserAndGuest') {
                Set-ExoRowTone -Row $script:gridStorage.Rows[$Index] -Tone 'Warning'
                $OpenSharing++
            }
        }

        $script:lblStorageSummary.Text = "$($Sites.Count) site(s). $OverQuota at 90% or more of quota, $OpenSharing allowing anonymous or guest sharing."
        Write-AppLog "Storage report: $($Sites.Count) site(s), $OverQuota near quota, $OpenSharing with open sharing." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText 'Storage and sharing report complete.'
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Storage report failed.'
        Show-ErrorMessage "The storage report failed.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Storage report failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
$script:SPOConnected     = $false
$script:OneDriveCache    = $null
$script:CurrentSite      = $null
$script:OffboardSite     = $null
$script:RetentionDays    = 0
$script:ActionControls   = @()
$script:UserPickerCombos = @()

function New-ExoGrid {
    param([System.Windows.Forms.Control]$Parent,[int]$Top,[int]$Height)

    $Grid = New-Object System.Windows.Forms.DataGridView
    $Grid.Location = New-Object System.Drawing.Point(12,$Top)
    $Grid.Size = New-Object System.Drawing.Size(($Parent.ClientSize.Width - 24),$Height)
    $Grid.Anchor = "Top,Bottom,Left,Right"
    $Grid.ReadOnly = $true
    $Grid.AllowUserToAddRows = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.AllowUserToResizeRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.SelectionMode = "FullRowSelect"
    $Grid.MultiSelect = $true
    $Grid.AutoSizeColumnsMode = "Fill"
    $Grid.ColumnHeadersHeightSizeMode = "AutoSize"
    $Parent.Controls.Add($Grid)
    return $Grid
}

function Set-ExoRowTone {
    param([System.Windows.Forms.DataGridViewRow]$Row,[ValidateSet('Danger','Warning','Success','Normal')][string]$Tone)

    $Palette = Get-MSToolkitThemePalette

    switch ($Tone) {
        'Danger'  { $Row.DefaultCellStyle.ForeColor = $Palette.Danger }
        'Warning' { $Row.DefaultCellStyle.ForeColor = $Palette.Warning }
        'Success' { $Row.DefaultCellStyle.ForeColor = $Palette.Success }
        default   { $Row.DefaultCellStyle.ForeColor = $Palette.Text }
    }
}

function Export-ExoGrid {
    param([System.Windows.Forms.DataGridView]$Grid,[string]$BaseName)

    if (-not $Grid -or $Grid.Rows.Count -eq 0) {
        Show-InfoMessage 'There is nothing to export yet.'
        return
    }

    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    $Dialog.Filter = 'CSV file (*.csv)|*.csv'
    $Dialog.FileName = "$BaseName-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

    if ($Dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $Rows = New-Object System.Collections.Generic.List[object]

        foreach ($Row in $Grid.Rows) {
            if ($Row.IsNewRow) { continue }

            $Item = [ordered]@{}
            foreach ($Column in $Grid.Columns) { $Item[$Column.HeaderText] = [string]$Row.Cells[$Column.Index].Value }
            $Rows.Add([pscustomobject]$Item)
        }

        $Rows.ToArray() | Export-Csv -LiteralPath $Dialog.FileName -NoTypeInformation -Encoding UTF8
        Write-AppLog "Exported $($Rows.Count) row(s) to $($Dialog.FileName)." 'SUCCESS'
        Show-InfoMessage "Exported $($Rows.Count) row(s) to:`r`n`r`n$($Dialog.FileName)"
    }
    catch {
        Show-ErrorMessage "Export failed.`r`n`r`n$($_.Exception.Message)"
    }
}

function New-OdLabel {
    param([System.Windows.Forms.Control]$Parent,[string]$Text,[int]$X,[int]$Y,[int]$Width = 120)

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Text
    $Label.Location = New-Object System.Drawing.Point($X,($Y + 4))
    $Label.Size = New-Object System.Drawing.Size($Width,20)
    $Parent.Controls.Add($Label)
    return $Label
}

function New-OdUserCombo {
    param([System.Windows.Forms.Control]$Parent,[int]$X,[int]$Y,[int]$Width = 320)

    $Combo = New-Object System.Windows.Forms.ComboBox
    $Combo.Location = New-Object System.Drawing.Point($X,$Y)
    $Combo.Size = New-Object System.Drawing.Size($Width,24)
    $Combo.DropDownStyle = 'DropDown'
    $Combo.AutoCompleteMode = 'SuggestAppend'
    $Combo.AutoCompleteSource = 'ListItems'
    $Combo.DropDownHeight = 320
    $Parent.Controls.Add($Combo)
    $script:UserPickerCombos += $Combo
    return $Combo
}

function New-OdButton {
    param(
        [System.Windows.Forms.Control]$Parent,
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 120,
        [scriptblock]$OnClick,
        [switch]$NeedsConnection
    )

    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = $Text
    $Button.Location = New-Object System.Drawing.Point($X,$Y)
    $Button.Size = New-Object System.Drawing.Size($Width,28)
    $Button.Add_Click($OnClick)
    $Parent.Controls.Add($Button)

    if ($NeedsConnection) {
        $Button.Enabled = $false
        $script:ActionControls += $Button
    }

    return $Button
}

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'OneDrive and SharePoint Tools'
$MainForm.Size = New-Object System.Drawing.Size(1500,940)
$MainForm.MinimumSize = New-Object System.Drawing.Size(1180,780)
$MainForm.StartPosition = 'CenterScreen'
$script:MainForm = $MainForm

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = 'Top'
$pnlHeader.Height = 78
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$pnlHeader.Tag = 'TopBar'
$MainForm.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'OneDrive and SharePoint Tools'
$lblTitle.AutoSize = $true
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',15)
$lblTitle.Location = New-Object System.Drawing.Point(18,8)
$pnlHeader.Controls.Add($lblTitle)

$lblAdminUrlCaption = New-Object System.Windows.Forms.Label
$lblAdminUrlCaption.Text = 'Admin URL:'
$lblAdminUrlCaption.AutoSize = $true
$lblAdminUrlCaption.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
$lblAdminUrlCaption.Location = New-Object System.Drawing.Point(20,46)
$pnlHeader.Controls.Add($lblAdminUrlCaption)

$script:txtAdminUrl = New-Object System.Windows.Forms.TextBox
$script:txtAdminUrl.Location = New-Object System.Drawing.Point(96,43)
$script:txtAdminUrl.Size = New-Object System.Drawing.Size(330,24)
$script:txtAdminUrl.Text = Get-SPOAdminUrlPreference
$pnlHeader.Controls.Add($script:txtAdminUrl)

$lblAdminHint = New-Object System.Windows.Forms.Label
$lblAdminHint.Text = 'Change it only for another tenant; a different value is remembered in its place'
$lblAdminHint.AutoSize = $true
$lblAdminHint.ForeColor = [System.Drawing.Color]::FromArgb(180,196,216)
$lblAdminHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
$lblAdminHint.Location = New-Object System.Drawing.Point(436,47)
$pnlHeader.Controls.Add($lblAdminHint)

# Same accent blue as the original M365 tools - this button had no styling at all
# and was inheriting the default grey.
$script:btnConnect = New-Object System.Windows.Forms.Button
$script:btnConnect.Text = 'Connect to SharePoint Online'
$script:btnConnect.Size = New-Object System.Drawing.Size(210,36)
$script:btnConnect.Location = New-Object System.Drawing.Point(($MainForm.ClientSize.Width - 226),12)
$script:btnConnect.Anchor = 'Top,Right'
$script:btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$script:btnConnect.ForeColor = [System.Drawing.Color]::White
$script:btnConnect.FlatStyle = 'Flat'
$script:btnConnect.FlatAppearance.BorderSize = 0
$script:btnConnect.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$script:btnConnect.Add_Click({ Connect-M365SharePoint })
$pnlHeader.Controls.Add($script:btnConnect)

# Fixed width and right aligned: an AutoSize label anchored right grows off the
# edge of the window, which is what a long tenant URL used to do.
$script:lblConnection = New-Object System.Windows.Forms.Label
$script:lblConnection.Text = 'Not connected'
$script:lblConnection.AutoSize = $false
$script:lblConnection.AutoEllipsis = $true
$script:lblConnection.Size = New-Object System.Drawing.Size(210,18)
$script:lblConnection.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:lblConnection.ForeColor = [System.Drawing.Color]::FromArgb(255,170,170)
$script:lblConnection.Location = New-Object System.Drawing.Point(($MainForm.ClientSize.Width - 226),52)
$script:lblConnection.Anchor = 'Top,Right'
$pnlHeader.Controls.Add($script:lblConnection)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$script:lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblStatus.Text = 'Enter the admin URL and connect.'
[void]$statusStrip.Items.Add($script:lblStatus)
$MainForm.Controls.Add($statusStrip)

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Dock = 'Bottom'
$pnlLog.Height = 160
$MainForm.Controls.Add($pnlLog)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = 'Activity'
$lblLog.Location = New-Object System.Drawing.Point(12,4)
$lblLog.Size = New-Object System.Drawing.Size(120,18)
$pnlLog.Controls.Add($lblLog)

$script:txtLog = New-Object System.Windows.Forms.RichTextBox
$script:txtLog.Location = New-Object System.Drawing.Point(12,24)
$script:txtLog.Size = New-Object System.Drawing.Size(($MainForm.ClientSize.Width - 24),125)
$script:txtLog.Anchor = 'Top,Bottom,Left,Right'
$script:txtLog.ReadOnly = $true
$script:txtLog.Font = New-Object System.Drawing.Font('Consolas',9)
$pnlLog.Controls.Add($script:txtLog)

$Tabs = New-Object System.Windows.Forms.TabControl
$Tabs.Dock = 'Fill'
$Tabs.Padding = New-Object System.Drawing.Point(14,6)
$MainForm.Controls.Add($Tabs)
$Tabs.BringToFront()

# Tab 1 - grant access
$tabAccess = New-Object System.Windows.Forms.TabPage
$tabAccess.Text = '  Grant Access  '
$Tabs.TabPages.Add($tabAccess)

New-OdLabel -Parent $tabAccess -Text 'OneDrive owner:' -X 12 -Y 16 -Width 110 | Out-Null
$script:cboAccessOwner = New-OdUserCombo -Parent $tabAccess -X 126 -Y 14 -Width 340
New-OdButton -Parent $tabAccess -Text 'Find OneDrive' -X 476 -Y 13 -Width 120 -NeedsConnection -OnClick { Invoke-OneDriveLookup } | Out-Null
New-OdButton -Parent $tabAccess -Text 'Open in Browser' -X 604 -Y 13 -Width 130 -NeedsConnection -OnClick { Invoke-OneDriveOpen } | Out-Null

New-OdLabel -Parent $tabAccess -Text 'Give access to:' -X 12 -Y 52 -Width 110 | Out-Null
$script:cboAccessRequester = New-OdUserCombo -Parent $tabAccess -X 126 -Y 50 -Width 340
New-OdButton -Parent $tabAccess -Text 'Grant Access' -X 476 -Y 49 -Width 120 -NeedsConnection -OnClick { Invoke-OneDriveGrant } | Out-Null
New-OdButton -Parent $tabAccess -Text 'Remove Selected' -X 604 -Y 49 -Width 130 -NeedsConnection -OnClick { Invoke-OneDriveRemoveAccess } | Out-Null
New-OdButton -Parent $tabAccess -Text 'Export CSV' -X 742 -Y 49 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridAccess -BaseName 'OneDriveAccess' } | Out-Null

$script:lblAccessUrl = New-Object System.Windows.Forms.Label
$script:lblAccessUrl.Location = New-Object System.Drawing.Point(12,86)
$script:lblAccessUrl.Size = New-Object System.Drawing.Size(1000,18)
$script:lblAccessUrl.Anchor = 'Top,Left,Right'
$script:lblAccessUrl.Font = New-Object System.Drawing.Font('Consolas',9)
$tabAccess.Controls.Add($script:lblAccessUrl)

$script:lblAccessSummary = New-Object System.Windows.Forms.Label
$script:lblAccessSummary.Location = New-Object System.Drawing.Point(12,108)
$script:lblAccessSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblAccessSummary.Anchor = 'Top,Left,Right'
$script:lblAccessSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabAccess.Controls.Add($script:lblAccessSummary)

$lblAccessHint = New-Object System.Windows.Forms.Label
$lblAccessHint.Text = 'Granted access makes somebody a site collection administrator: full control of every file, including anything never shared. Remove it when the reason has passed.'
$lblAccessHint.Location = New-Object System.Drawing.Point(12,130)
$lblAccessHint.Size = New-Object System.Drawing.Size(1000,18)
$lblAccessHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
$tabAccess.Controls.Add($lblAccessHint)

$script:gridAccess = New-ExoGrid -Parent $tabAccess -Top 154 -Height ($tabAccess.ClientSize.Height - 168)
[void]$script:gridAccess.Columns.Add('Kind','Kind')
[void]$script:gridAccess.Columns.Add('Name','Name')
[void]$script:gridAccess.Columns.Add('Login','Login name')
[void]$script:gridAccess.Columns.Add('IsGroup','Group')
$script:gridAccess.Columns['Kind'].FillWeight = 60
$script:gridAccess.Columns['IsGroup'].FillWeight = 45

# Tab 2 - access audit
$tabAudit = New-Object System.Windows.Forms.TabPage
$tabAudit.Text = '  Who Has Access  '
$Tabs.TabPages.Add($tabAudit)

New-OdButton -Parent $tabAudit -Text 'Run Audit' -X 12 -Y 14 -Width 120 -NeedsConnection -OnClick { Invoke-OneDriveAccessAudit } | Out-Null
New-OdButton -Parent $tabAudit -Text 'Export CSV' -X 140 -Y 14 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridAudit -BaseName 'OneDriveAccessAudit' } | Out-Null

$script:chkAuditRefresh = New-Object System.Windows.Forms.CheckBox
$script:chkAuditRefresh.Text = 'Reload the site list first'
$script:chkAuditRefresh.Location = New-Object System.Drawing.Point(262,18)
$script:chkAuditRefresh.Size = New-Object System.Drawing.Size(200,22)
$tabAudit.Controls.Add($script:chkAuditRefresh)

$script:lblAuditSummary = New-Object System.Windows.Forms.Label
$script:lblAuditSummary.Text = 'Every OneDrive with a site collection administrator other than its owner.'
$script:lblAuditSummary.Location = New-Object System.Drawing.Point(12,52)
$script:lblAuditSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblAuditSummary.Anchor = 'Top,Left,Right'
$script:lblAuditSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabAudit.Controls.Add($script:lblAuditSummary)

$script:gridAudit = New-ExoGrid -Parent $tabAudit -Top 80 -Height ($tabAudit.ClientSize.Height - 94)
[void]$script:gridAudit.Columns.Add('Title','OneDrive')
[void]$script:gridAudit.Columns.Add('Owner','Owner')
[void]$script:gridAudit.Columns.Add('AdminName','Has access')
[void]$script:gridAudit.Columns.Add('AdminLogin','Login name')
[void]$script:gridAudit.Columns.Add('IsGroup','Group')
[void]$script:gridAudit.Columns.Add('Url','Site')
$script:gridAudit.Columns['IsGroup'].FillWeight = 45
$script:gridAudit.Columns['Url'].FillWeight = 150

# Tab 3 - offboarding
$tabOffboard = New-Object System.Windows.Forms.TabPage
$tabOffboard.Text = '  Offboarding  '
$Tabs.TabPages.Add($tabOffboard)

New-OdLabel -Parent $tabOffboard -Text 'Departed user:' -X 12 -Y 16 -Width 100 | Out-Null
$script:cboOffboardUser = New-OdUserCombo -Parent $tabOffboard -X 116 -Y 14 -Width 340
New-OdButton -Parent $tabOffboard -Text 'Build Report' -X 466 -Y 13 -Width 120 -NeedsConnection -OnClick { Invoke-OneDriveOffboardReport } | Out-Null

New-OdLabel -Parent $tabOffboard -Text 'Hand over to:' -X 12 -Y 52 -Width 100 | Out-Null
$script:cboOffboardManager = New-OdUserCombo -Parent $tabOffboard -X 116 -Y 50 -Width 340
New-OdButton -Parent $tabOffboard -Text 'Grant Handover' -X 466 -Y 49 -Width 130 -NeedsConnection -OnClick { Invoke-OneDriveHandover } | Out-Null
New-OdButton -Parent $tabOffboard -Text 'Export CSV' -X 604 -Y 49 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridOffboard -BaseName 'OneDriveOffboarding' } | Out-Null

$script:lblOffboardSummary = New-Object System.Windows.Forms.Label
$script:lblOffboardSummary.Location = New-Object System.Drawing.Point(12,86)
$script:lblOffboardSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblOffboardSummary.Anchor = 'Top,Left,Right'
$script:lblOffboardSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabOffboard.Controls.Add($script:lblOffboardSummary)

$lblOffboardHint = New-Object System.Windows.Forms.Label
$lblOffboardHint.Text = 'Once the owner account is deleted the OneDrive is kept only for the tenant retention period, then removed with its contents.'
$lblOffboardHint.Location = New-Object System.Drawing.Point(12,108)
$lblOffboardHint.Size = New-Object System.Drawing.Size(1000,18)
$lblOffboardHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
$tabOffboard.Controls.Add($lblOffboardHint)

$script:gridOffboard = New-ExoGrid -Parent $tabOffboard -Top 132 -Height ($tabOffboard.ClientSize.Height - 146)
[void]$script:gridOffboard.Columns.Add('Name','Setting')
[void]$script:gridOffboard.Columns.Add('Value','Value')
$script:gridOffboard.Columns['Name'].FillWeight = 70
$script:gridOffboard.Columns['Value'].FillWeight = 220

# Tab 4 - deleted sites
$tabDeleted = New-Object System.Windows.Forms.TabPage
$tabDeleted.Text = '  Deleted OneDrives  '
$Tabs.TabPages.Add($tabDeleted)

New-OdLabel -Parent $tabDeleted -Text 'Filter:' -X 12 -Y 16 -Width 50 | Out-Null
$script:txtDeletedFilter = New-Object System.Windows.Forms.TextBox
$script:txtDeletedFilter.Location = New-Object System.Drawing.Point(66,14)
$script:txtDeletedFilter.Size = New-Object System.Drawing.Size(300,24)
$tabDeleted.Controls.Add($script:txtDeletedFilter)

New-OdButton -Parent $tabDeleted -Text 'Search' -X 376 -Y 13 -Width 100 -NeedsConnection -OnClick { Invoke-DeletedSiteSearch } | Out-Null
New-OdButton -Parent $tabDeleted -Text 'Restore Selected' -X 484 -Y 13 -Width 140 -NeedsConnection -OnClick { Invoke-DeletedSiteRestore } | Out-Null
New-OdButton -Parent $tabDeleted -Text 'Export CSV' -X 632 -Y 13 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridDeleted -BaseName 'DeletedSites' } | Out-Null

$script:lblDeletedSummary = New-Object System.Windows.Forms.Label
$script:lblDeletedSummary.Text = 'Deleted sites can be restored until the retention window closes, after which the content is gone.'
$script:lblDeletedSummary.Location = New-Object System.Drawing.Point(12,50)
$script:lblDeletedSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblDeletedSummary.Anchor = 'Top,Left,Right'
$script:lblDeletedSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabDeleted.Controls.Add($script:lblDeletedSummary)

$script:gridDeleted = New-ExoGrid -Parent $tabDeleted -Top 78 -Height ($tabDeleted.ClientSize.Height - 92)
[void]$script:gridDeleted.Columns.Add('Url','Site')
[void]$script:gridDeleted.Columns.Add('Title','Title')
[void]$script:gridDeleted.Columns.Add('Deleted','Deleted')
[void]$script:gridDeleted.Columns.Add('DaysLeft','Days left')
[void]$script:gridDeleted.Columns.Add('SiteId','Site ID')
[void]$script:gridDeleted.Columns.Add('Status','Status')
$script:gridDeleted.Columns['Url'].FillWeight = 180
$script:gridDeleted.Columns['DaysLeft'].FillWeight = 50
$script:gridDeleted.Columns['SiteId'].FillWeight = 90

# Tab 5 - storage and sharing
$tabStorage = New-Object System.Windows.Forms.TabPage
$tabStorage.Text = '  Storage and Sharing  '
$Tabs.TabPages.Add($tabStorage)

New-OdLabel -Parent $tabStorage -Text 'Scope:' -X 12 -Y 16 -Width 50 | Out-Null
$script:cboReportScope = New-Object System.Windows.Forms.ComboBox
$script:cboReportScope.Location = New-Object System.Drawing.Point(66,14)
$script:cboReportScope.Size = New-Object System.Drawing.Size(220,24)
$script:cboReportScope.DropDownStyle = 'DropDownList'
[void]$script:cboReportScope.Items.Add('OneDrive sites')
[void]$script:cboReportScope.Items.Add('All SharePoint sites')
$script:cboReportScope.SelectedIndex = 0
$tabStorage.Controls.Add($script:cboReportScope)

New-OdButton -Parent $tabStorage -Text 'Build Report' -X 296 -Y 13 -Width 120 -NeedsConnection -OnClick { Invoke-StorageSharingReport } | Out-Null
New-OdButton -Parent $tabStorage -Text 'Export CSV' -X 424 -Y 13 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridStorage -BaseName 'StorageAndSharing' } | Out-Null

$script:lblStorageSummary = New-Object System.Windows.Forms.Label
$script:lblStorageSummary.Text = 'Red: 90% or more of quota. Orange: anonymous or guest sharing is allowed on that site.'
$script:lblStorageSummary.Location = New-Object System.Drawing.Point(12,50)
$script:lblStorageSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblStorageSummary.Anchor = 'Top,Left,Right'
$script:lblStorageSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabStorage.Controls.Add($script:lblStorageSummary)

$script:gridStorage = New-ExoGrid -Parent $tabStorage -Top 78 -Height ($tabStorage.ClientSize.Height - 92)
[void]$script:gridStorage.Columns.Add('Title','Site')
[void]$script:gridStorage.Columns.Add('Owner','Owner')
[void]$script:gridStorage.Columns.Add('UsedMb','Used (MB)')
[void]$script:gridStorage.Columns.Add('QuotaMb','Quota (MB)')
[void]$script:gridStorage.Columns.Add('Percent','% used')
[void]$script:gridStorage.Columns.Add('Sharing','Sharing')
[void]$script:gridStorage.Columns.Add('LastChanged','Last change')
[void]$script:gridStorage.Columns.Add('Url','URL')
$script:gridStorage.Columns['UsedMb'].FillWeight = 55
$script:gridStorage.Columns['QuotaMb'].FillWeight = 55
$script:gridStorage.Columns['Percent'].FillWeight = 45
$script:gridStorage.Columns['Url'].FillWeight = 160

# --- grids, theme, start --------------------------------------------------
$script:AllGrids = @(
    $script:gridAccess, $script:gridAudit, $script:gridOffboard,
    $script:gridDeleted, $script:gridStorage
)

function Resize-OdGrids {
    foreach ($Grid in $script:AllGrids) {
        if (-not $Grid -or -not $Grid.Parent) { continue }

        $Width  = $Grid.Parent.ClientSize.Width - 24
        $Height = $Grid.Parent.ClientSize.Height - $Grid.Top - 14

        if ($Width -gt 100 -and $Height -gt 80) {
            $Grid.Size = New-Object System.Drawing.Size($Width,$Height)
        }
    }
}

# The user pickers list every OneDrive, which is one slow call, so it is loaded
# the first time a tab that needs it is opened rather than at connection.
function Initialize-OdPickers {
    if (-not $script:SPOConnected) { return }
    if ($script:PickersLoaded) { return }

    $Loaded = Initialize-MSToolkitUserPicker -Combos $script:UserPickerCombos
    if ($Loaded -ge 0) { $script:PickersLoaded = $true }
}

$Tabs.Add_SelectedIndexChanged({
    if ($Tabs.SelectedTab -eq $tabAccess -or $Tabs.SelectedTab -eq $tabOffboard) {
        Initialize-OdPickers
    }
})

$script:MSToolkitThemeRefreshHook = {
    if (-not $script:lblConnection) { return }

    if ($script:SPOConnected) { $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Success }
    else                      { $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Danger }
}

New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "M365OneDriveSharePointTools"
Apply-MSToolkitSharedTheme -Root $MainForm

$MainForm.Add_Shown({
    $MainForm.Activate()
    Resize-OdGrids

    Write-AppLog "Admin URL: $($script:txtAdminUrl.Text)"
    Write-AppLog 'OneDrive and SharePoint Tools ready.'
    Write-AppLog 'Sign in with a standard account holding the SharePoint Administrator role.'
    Write-AppLog 'Granting access, removing it, and restoring a deleted site all confirm first. Everything else is read-only.'

    if (-not $script:txtAdminUrl.Text) {
        Write-AppLog 'Enter the tenant admin URL - https://<tenant>-admin.sharepoint.com - then Connect.' 'WARN'
    }
    else {
        $script:lblStatus.Text = 'Click Connect to sign in to SharePoint Online.'
    }
})

$MainForm.Add_Resize({ Resize-OdGrids })

Hide-PowerShellConsole
[void]$MainForm.ShowDialog()

try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch { }
