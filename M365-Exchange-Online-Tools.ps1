#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet("Light","Dark")]
    [string]$ThemeMode
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# ===========================================================================
#  M365-Exchange-Online-Tools.ps1
#  Exchange Online troubleshooting tools: inbox rules and forwarding, message
#  trace, distribution group delivery checks, and a mailbox report.
#
#  Read-only apart from releasing quarantined mail, which confirms first.
#
#  Sign in with a STANDARD account holding an Exchange administrator role.
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

$script:ReferenceUser = $null
$script:TargetUser = $null
$script:ComparisonRows = @()
$script:ExchangeConnected = $false

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
    $Enable = [bool]((-not $Busy) -and $script:ExchangeConnected)

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

function Ensure-ExchangeOnlineModule {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        return $true
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        "The ExchangeOnlineManagement PowerShell module is required and is not installed.`r`n`r`nInstall it for the current user now?",
        'Exchange Online Module Required',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }

    try {
        Set-BusyState -Busy $true -StatusText 'Installing Exchange Online module...'
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-AppLog 'Installed ExchangeOnlineManagement successfully.' 'SUCCESS'
        return $true
    }
    catch {
        Show-ErrorMessage "Unable to install ExchangeOnlineManagement.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Exchange Online module installation failed: $($_.Exception.Message)" 'ERROR'
        return $false
    }
    finally {
        Set-BusyState -Busy $false
    }
}


function Connect-M365Exchange {
    if (-not (Ensure-ExchangeOnlineModule)) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Connecting to Exchange Online...'
        Write-AppLog 'Starting Exchange Online sign-in.'

        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Start-M365AuthWindowWatcher
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

        # Verify the session by reading one lightweight organization property.
        $org = Get-OrganizationConfig -ErrorAction Stop

        $script:ExchangeConnected = $true
        $script:lblConnection.Text = "Connected: $($org.DisplayName)"

        # Which identity the session actually used. Connect-ExchangeOnline picks an
        # account through Windows sign-in without asking, so this is not always the
        # account you expect, and what a cmdlet is allowed to do depends on the
        # rights of whoever it chose.
        try {
            $Connection = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
            if ($Connection) {
                $script:ConnectedAs = [string]$Connection.UserPrincipalName
                Write-AppLog "Signed in as: $($script:ConnectedAs)" 'SUCCESS'

                # Session shape, for comparing against a plain PowerShell window when
                # a cmdlet works there and not here. The EOP cmdlets in particular
                # depend on how the session was established, not just on who signed in.
                foreach ($Property in 'ConnectionUri','ConnectionId','TokenStatus','IsEopSession','ModuleName','Organization') {
                    if ($Connection.PSObject.Properties.Name -contains $Property) {
                        Write-AppLog "  $Property : $($Connection.$Property)"
                    }
                }

                $Module = Get-Module ExchangeOnlineManagement | Select-Object -First 1
                if ($Module) { Write-AppLog "  Module version : $($Module.Version)" }
            }
        }
        catch {
            Write-AppLog "Could not read the connection details: $($_.Exception.Message)" 'WARN'
        }
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
        $script:lblStatus.Text = 'Connected to Exchange Online. Pick a tab and run a check.'
        Write-AppLog "Connected to Exchange Online: $($org.DisplayName)." 'SUCCESS'

        $script:MSToolkitUserPickerCache = $null
        $LoadedCount = Initialize-MSToolkitUserPicker -Combos $script:UserPickerCombos

        if ($LoadedCount -ge 0) {
            Write-AppLog "Loaded $LoadedCount mailbox user(s) into the user lists."
        }
        else {
            Write-AppLog 'Could not load the mailbox user list. Type a user instead.' 'WARNING'
        }
    }
    catch {
        $script:ExchangeConnected = $false
        $script:lblConnection.Text = 'Not connected'
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Danger

        # "Argument types do not match" and friends come from .NET rather than from
        # Exchange, so record where it actually happened instead of the message alone.
        $Failure = $_
        $Detail = New-Object System.Collections.Generic.List[string]
        $Detail.Add("Type: $($Failure.Exception.GetType().FullName)")
        $Detail.Add("Message: $($Failure.Exception.Message)")

        $Inner = $Failure.Exception.InnerException
        $Depth = 0
        while ($Inner -and $Depth -lt 3) {
            $Detail.Add("Inner: $($Inner.GetType().FullName) - $($Inner.Message)")
            $Inner = $Inner.InnerException
            $Depth++
        }

        if ($Failure.InvocationInfo) {
            $Detail.Add("Command: $($Failure.InvocationInfo.MyCommand)")
            $Detail.Add("Line $($Failure.InvocationInfo.ScriptLineNumber): $("$($Failure.InvocationInfo.Line)".Trim())")
        }

        $Module = Get-Module ExchangeOnlineManagement | Select-Object -First 1
        if ($Module) { $Detail.Add("ExchangeOnlineManagement version: $($Module.Version)") }
        $Detail.Add("PowerShell: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)), apartment $([System.Threading.Thread]::CurrentThread.GetApartmentState())")

        foreach ($Line in $Detail) { Write-AppLog $Line 'ERROR' }

        if ($Failure.ScriptStackTrace) {
            foreach ($Line in ($Failure.ScriptStackTrace -split "`n" | Select-Object -First 4)) {
                Write-AppLog "  at $($Line.Trim())" 'ERROR'
            }
        }

        Show-ErrorMessage "Unable to connect to Exchange Online.`r`n`r`n$($Failure.Exception.Message)`r`n`r`nThe Activity log below has the exception type, the failing command and the module version."
    }
    finally {
        Stop-M365AuthWindowWatcher
        Set-BusyState -Busy $false
    }
}

function Show-RecipientSelectionDialog {
    param([Parameter(Mandatory)][array]$Recipients,[Parameter(Mandatory)][string]$SearchText)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Select User - $SearchText"
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object System.Drawing.Size(780,430)
    $dialog.MinimumSize = New-Object System.Drawing.Size(680,350)
    $dialog.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $dialog.BackColor = [System.Drawing.Color]::White

    $label = New-Object System.Windows.Forms.Label
    $label.Text = 'Multiple recipients matched. Select the correct user:'
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(15,15)
    $dialog.Controls.Add($label)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(15,45)
    $grid.Size = New-Object System.Drawing.Size(735,285)
    $grid.Anchor = 'Top,Bottom,Left,Right'
    $grid.ReadOnly = $true
    $grid.MultiSelect = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowHeadersVisible = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White

    [void]$grid.Columns.Add('DisplayName','Display Name')
    [void]$grid.Columns.Add('PrimarySmtpAddress','Primary SMTP Address')
    [void]$grid.Columns.Add('RecipientTypeDetails','Recipient Type')

    foreach ($recipient in $Recipients) {
        $index = $grid.Rows.Add([string]$recipient.DisplayName,[string]$recipient.PrimarySmtpAddress,[string]$recipient.RecipientTypeDetails)
        $grid.Rows[$index].Tag = $recipient
    }
    $dialog.Controls.Add($grid)

    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Text = 'Select User'
    $btnSelect.Size = New-Object System.Drawing.Size(110,32)
    $btnSelect.Anchor = 'Bottom,Right'
    $btnSelect.Location = New-Object System.Drawing.Point(520,345)
    $btnSelect.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
    $btnSelect.ForeColor = [System.Drawing.Color]::White
    $btnSelect.FlatStyle = 'Flat'
    $btnSelect.FlatAppearance.BorderSize = 0
    $dialog.Controls.Add($btnSelect)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size(110,32)
    $btnCancel.Anchor = 'Bottom,Right'
    $btnCancel.Location = New-Object System.Drawing.Point(640,345)
    $dialog.Controls.Add($btnCancel)

    $selectAction = {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $script:DialogSelectedRecipient = $grid.SelectedRows[0].Tag
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    }

    $btnSelect.Add_Click($selectAction)
    $grid.Add_CellDoubleClick($selectAction)
    $btnCancel.Add_Click({ $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dialog.Close() })

    if ($grid.Rows.Count -gt 0) { $grid.Rows[0].Selected = $true }

    $script:DialogSelectedRecipient = $null
    Apply-MSToolkitSharedTheme -Root $dialog
    Register-MSToolkitComboFiltersOn -Root $dialog
    $result = $dialog.ShowDialog($script:MainForm)
    $selected = $null
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { $selected = $script:DialogSelectedRecipient }

    Remove-Variable DialogSelectedRecipient -Scope Script -ErrorAction SilentlyContinue
    $dialog.Dispose()
    return $selected
}

function Get-MSToolkitUserPickerLabel {
    param($Recipient)

    $Display = if (-not [string]::IsNullOrWhiteSpace($Recipient.DisplayName)) { $Recipient.DisplayName } else { $Recipient.PrimarySmtpAddress }
    return "$Display ($($Recipient.PrimarySmtpAddress))"
}

function ConvertFrom-MSToolkitUserPickerLabel {
    param([string]$Value)

    $Text = "$Value".Trim()

    # Entries loaded into the picker look like "Jane Doe (jane.doe@contoso.com)".
    # Anything typed by hand is passed straight through unchanged.
    if ($Text -match '\(([^()]+)\)\s*$') {
        return $Matches[1].Trim()
    }

    return $Text
}

function Get-MSToolkitExchangeUserList {
    return @(
        Get-Recipient -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.PrimarySmtpAddress) } |
        Sort-Object DisplayName
    )
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
    param($Combos)

    try {
        if (-not $script:MSToolkitUserPickerCache) {
            $script:MSToolkitUserPickerCache = Get-MSToolkitExchangeUserList
        }

        foreach ($Combo in $Combos) {
            if ($Combo -isnot [System.Windows.Forms.ComboBox]) { continue }

            $Existing = $Combo.Text
            $Combo.Items.Clear()

            foreach ($Recipient in $script:MSToolkitUserPickerCache) {
                $null = $Combo.Items.Add((Get-MSToolkitUserPickerLabel -Recipient $Recipient))
            }

            $Combo.Text = $Existing
        }

        return @($script:MSToolkitUserPickerCache).Count
    }
    catch {
        # The combos stay typeable if the list cannot be loaded.
        return -1
    }
}

function Resolve-ExchangeUser {
    param([Parameter(Mandatory)][string]$Identity)

    $Identity = $Identity.Trim()
    if ([string]::IsNullOrWhiteSpace($Identity)) { throw 'A user name or email address was not entered.' }

    try {
        $exact = Get-Recipient -Identity $Identity -ErrorAction Stop
        if ($exact -and $exact.RecipientTypeDetails -notmatch 'Group|Contact|PublicFolder') { return $exact }
    }
    catch { }

    $matches = @(Get-Recipient -Anr $Identity -ResultSize 25 -ErrorAction Stop | Where-Object {
        $_.RecipientTypeDetails -notmatch 'Group|Contact|PublicFolder'
    })

    if ($matches.Count -eq 0) { throw "No Exchange Online user was found matching '$Identity'." }
    if ($matches.Count -eq 1) { return $matches[0] }

    $selected = Show-RecipientSelectionDialog -Recipients $matches -SearchText $Identity
    if (-not $selected) { throw 'User selection was canceled.' }
    return $selected
}

# ---------------------------------------------------------------------------
# Shared grid helpers
# ---------------------------------------------------------------------------
function New-ExoGrid {
    param(
        [System.Windows.Forms.Control]$Parent,
        [int]$Top,
        [int]$Height
    )

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

function Export-ExoGrid {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$BaseName
    )

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
            foreach ($Column in $Grid.Columns) {
                $Item[$Column.HeaderText] = [string]$Row.Cells[$Column.Index].Value
            }
            $Rows.Add([pscustomobject]$Item)
        }

        $Rows.ToArray() | Export-Csv -LiteralPath $Dialog.FileName -NoTypeInformation -Encoding UTF8
        Write-AppLog "Exported $($Rows.Count) row(s) to $($Dialog.FileName)." 'SUCCESS'
        Show-InfoMessage "Exported $($Rows.Count) row(s) to:`r`n`r`n$($Dialog.FileName)"
    }
    catch {
        Show-ErrorMessage "Export failed.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Export failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Copy-ExoGrid {
    param([System.Windows.Forms.DataGridView]$Grid)

    if (-not $Grid -or $Grid.Rows.Count -eq 0) {
        Show-InfoMessage 'There is nothing to copy yet.'
        return
    }

    $Text = New-Object System.Text.StringBuilder
    [void]$Text.AppendLine((($Grid.Columns | ForEach-Object { $_.HeaderText }) -join "`t"))

    foreach ($Row in $Grid.Rows) {
        if ($Row.IsNewRow) { continue }
        $Values = foreach ($Column in $Grid.Columns) { [string]$Row.Cells[$Column.Index].Value }
        [void]$Text.AppendLine(($Values -join "`t"))
    }

    [System.Windows.Forms.Clipboard]::SetText($Text.ToString())
    Write-AppLog 'Copied the grid to the clipboard.'
}

function Set-ExoRowTone {
    param(
        [System.Windows.Forms.DataGridViewRow]$Row,
        [ValidateSet('Danger','Warning','Success','Normal')]
        [string]$Tone
    )

    $Palette = Get-MSToolkitThemePalette

    switch ($Tone) {
        'Danger'  { $Row.DefaultCellStyle.ForeColor = $Palette.Danger }
        'Warning' { $Row.DefaultCellStyle.ForeColor = $Palette.Warning }
        'Success' { $Row.DefaultCellStyle.ForeColor = $Palette.Success }
        default   { $Row.DefaultCellStyle.ForeColor = $Palette.Text }
    }
}

# ---------------------------------------------------------------------------
# Inbox rules and forwarding
#
# The first thing to check after a suspected mailbox compromise: a rule that
# forwards or redirects outside the tenant, or one that files mail away and
# deletes it so the owner never sees the replies.
# ---------------------------------------------------------------------------
function Test-ExoExternalAddress {
    param([string]$Address)

    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    if ($Address -notmatch '@') { return $false }

    $Domain = ($Address -split '@')[-1].Trim().TrimEnd('>').ToLower()
    if (-not $Domain) { return $false }

    return (-not ($script:AcceptedDomains -contains $Domain))
}

function Get-ExoRuleTargets {
    param($Rule)

    # Every place a rule can send mail to somebody else.
    $Targets = New-Object System.Collections.Generic.List[string]

    foreach ($Property in 'ForwardTo','ForwardAsAttachmentTo','RedirectTo') {
        foreach ($Value in @($Rule.$Property)) {
            if ($null -eq $Value) { continue }

            $Text = [string]$Value
            if ($Text -match '"?([^"\[\]]+@[^"\[\]\s]+)') { $Text = $Matches[1] }
            $Targets.Add($Text.Trim())
        }
    }

    return $Targets.ToArray()
}

function Invoke-ExoRuleCheck {
    if (-not $script:ExchangeConnected) { return }

    $Identity = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboRuleUser.Text
    if ([string]::IsNullOrWhiteSpace($Identity)) {
        Show-InfoMessage 'Enter a user first.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading mailbox rules...'
        $script:gridRules.Rows.Clear()
        $script:lblRuleSummary.Text = ''

        $User = Resolve-ExchangeUser -Identity $Identity
        $Smtp = [string]$User.PrimarySmtpAddress
        Write-AppLog "Checking inbox rules and forwarding for $Smtp."

        if (-not $script:AcceptedDomains -or $script:AcceptedDomains.Count -eq 0) {
            $script:AcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop | ForEach-Object { ([string]$_.DomainName).ToLower() })
            Write-AppLog "Loaded $($script:AcceptedDomains.Count) accepted domain(s) to judge what counts as external."
        }

        $Risky = 0

        # Mailbox-level forwarding first - it is set outside the rules list and is
        # easy to miss, because Outlook never shows it to the owner.
        $Mailbox = Get-Mailbox -Identity $Smtp -ErrorAction Stop

        foreach ($Property in 'ForwardingSmtpAddress','ForwardingAddress') {
            $Value = [string]$Mailbox.$Property
            if ([string]::IsNullOrWhiteSpace($Value)) { continue }

            $Clean = $Value -replace '^smtp:', ''
            $External = Test-ExoExternalAddress -Address $Clean
            $Index = $script:gridRules.Rows.Add(
                'Mailbox',
                "Mailbox forwarding ($Property)",
                'Enabled',
                '',
                $Clean,
                '',
                "Also deliver to mailbox: $($Mailbox.DeliverToMailboxAndForward)"
            )

            if ($External) {
                Set-ExoRowTone -Row $script:gridRules.Rows[$Index] -Tone 'Danger'
                $Risky++
                Write-AppLog "Mailbox forwards to an EXTERNAL address: $Clean" 'WARN'
            }
            else {
                Set-ExoRowTone -Row $script:gridRules.Rows[$Index] -Tone 'Warning'
            }
        }

        $Rules = @(Get-InboxRule -Mailbox $Smtp -ErrorAction Stop)

        foreach ($Rule in $Rules) {
            $Targets = Get-ExoRuleTargets -Rule $Rule
            $External = @($Targets | Where-Object { Test-ExoExternalAddress -Address $_ })

            $Actions = New-Object System.Collections.Generic.List[string]
            if ($Targets.Count -gt 0)      { $Actions.Add("Sends to: $($Targets -join '; ')") }
            if ($Rule.MoveToFolder)        { $Actions.Add("Moves to: $($Rule.MoveToFolder)") }
            if ($Rule.DeleteMessage)       { $Actions.Add('Deletes the message') }
            if ($Rule.MarkAsRead)          { $Actions.Add('Marks as read') }
            if ($Rule.StopProcessingRules) { $Actions.Add('Stops processing further rules') }

            $Conditions = New-Object System.Collections.Generic.List[string]
            foreach ($Pair in @(
                @('From', $Rule.From),
                @('Subject contains', $Rule.SubjectContainsWords),
                @('Body contains', $Rule.BodyContainsWords),
                @('Sent to', $Rule.SentTo)
            )) {
                $Value = @($Pair[1]) -join '; '
                if (-not [string]::IsNullOrWhiteSpace($Value)) { $Conditions.Add("$($Pair[0]): $Value") }
            }

            $State = 'Enabled'
            if (-not $Rule.Enabled) { $State = 'Disabled' }

            $Index = $script:gridRules.Rows.Add(
                'Inbox rule',
                [string]$Rule.Name,
                $State,
                [string]$Rule.Priority,
                (($Targets) -join '; '),
                ($Conditions -join ' | '),
                ($Actions -join ' | ')
            )

            if ($External.Count -gt 0) {
                Set-ExoRowTone -Row $script:gridRules.Rows[$Index] -Tone 'Danger'
                $Risky++
                Write-AppLog "Rule '$($Rule.Name)' sends mail EXTERNALLY to: $($External -join '; ')" 'WARN'
            }
            elseif ($Rule.DeleteMessage -or $Targets.Count -gt 0) {
                Set-ExoRowTone -Row $script:gridRules.Rows[$Index] -Tone 'Warning'
            }
            elseif (-not $Rule.Enabled) {
                Set-ExoRowTone -Row $script:gridRules.Rows[$Index] -Tone 'Normal'
            }
        }

        $script:lblRuleSummary.Text = "$Smtp - $($Rules.Count) inbox rule(s), $Risky flagged"

        if ($Risky -gt 0) {
            $script:lblRuleSummary.ForeColor = (Get-MSToolkitThemePalette).Danger
            Write-AppLog "$Risky item(s) flagged for $Smtp. Red rows send mail outside the tenant." 'WARN'
        }
        else {
            $script:lblRuleSummary.ForeColor = (Get-MSToolkitThemePalette).Success
            Write-AppLog "No external forwarding found for $Smtp." 'SUCCESS'
        }

        Set-BusyState -Busy $false -StatusText "Checked $($Rules.Count) rule(s) for $Smtp."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Rule check failed.'
        Show-ErrorMessage "Could not read the rules.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Rule check failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Message trace
# ---------------------------------------------------------------------------
function Test-ExoMessageTraceV2 {
    # Get-MessageTrace was retired in favour of Get-MessageTraceV2. Cached because
    # every trace and every group member check asks.
    if ($null -eq $script:HasTraceV2) {
        $script:HasTraceV2 = [bool](Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)

        if ($script:HasTraceV2) {
            Write-AppLog 'Using Get-MessageTraceV2 (90 days of history, 10 days per query).'
        }
        else {
            Write-AppLog 'Get-MessageTraceV2 is not available in this module version; falling back to the retired Get-MessageTrace.' 'WARN'
        }
    }

    return $script:HasTraceV2
}

function Get-ExoTraceRows {
    param(
        [string]$Sender,
        [string]$Recipient,
        [datetime]$Start,
        [datetime]$End,
        [string]$Status
    )

    if (Test-ExoMessageTraceV2) {
        # V2 differences that matter: ResultSize replaces PageSize, and the range
        # can go back 90 days but cover at most 10 days per query.
        $Params = @{
            StartDate   = $Start
            EndDate     = $End
            ResultSize  = 5000
            ErrorAction = 'Stop'
        }

        if ($Sender)    { $Params['SenderAddress'] = $Sender }
        if ($Recipient) { $Params['RecipientAddress'] = $Recipient }
        if ($Status -and $Status -ne 'Any') { $Params['Status'] = $Status }

        return @(Get-MessageTraceV2 @Params | Sort-Object Received -Descending)
    }

    $Params = @{
        StartDate   = $Start
        EndDate     = $End
        PageSize    = 5000
        ErrorAction = 'Stop'
    }

    if ($Sender)    { $Params['SenderAddress'] = $Sender }
    if ($Recipient) { $Params['RecipientAddress'] = $Recipient }
    if ($Status -and $Status -ne 'Any') { $Params['Status'] = $Status }

    return @(Get-MessageTrace @Params | Sort-Object Received -Descending)
}

function Invoke-ExoMessageTrace {
    if (-not $script:ExchangeConnected) { return }

    $Sender    = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboTraceSender.Text
    $Recipient = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboTraceRecipient.Text
    $Subject   = "$($script:txtTraceSubject.Text)".Trim()

    if ([string]::IsNullOrWhiteSpace($Sender) -and [string]::IsNullOrWhiteSpace($Recipient)) {
        Show-InfoMessage 'Enter a sender, a recipient, or both.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Running message trace...'
        $script:gridTrace.Rows.Clear()

        $Start = $script:dtTraceStart.Value
        $End   = $script:dtTraceEnd.Value

        if ($Start -gt $End) {
            Show-InfoMessage 'The start date is after the end date.'
            Set-BusyState -Busy $false
            return
        }

        # V2 allows 90 days of history but only 10 days per query.
        if (($End - $Start).TotalDays -gt 10) {
            Show-InfoMessage "A single trace can cover at most 10 days.`r`n`r`nNarrow the range and run it again."
            Set-BusyState -Busy $false -StatusText 'Range too wide - 10 days maximum.'
            return
        }

        if (((Get-Date) - $Start).TotalDays -gt 90) {
            Write-AppLog 'Message trace only holds 90 days. For older mail use a historical search in the Defender portal.' 'WARN'
        }

        Write-AppLog "Tracing $($Start.ToString('yyyy-MM-dd HH:mm')) to $($End.ToString('yyyy-MM-dd HH:mm')). Sender '$Sender', recipient '$Recipient'."

        $Messages = Get-ExoTraceRows -Sender $Sender -Recipient $Recipient -Start $Start -End $End -Status $script:cboTraceStatus.SelectedItem

        if ($Subject) {
            $Messages = @($Messages | Where-Object { "$($_.Subject)" -like "*$Subject*" })
        }

        foreach ($Message in $Messages) {
            $Index = $script:gridTrace.Rows.Add(
                $Message.Received,
                [string]$Message.SenderAddress,
                [string]$Message.RecipientAddress,
                [string]$Message.Subject,
                [string]$Message.Status,
                [string]$Message.Size
            )

            $script:gridTrace.Rows[$Index].Tag = $Message

            switch -Regex ("$($Message.Status)") {
                'Failed|Quarantined'    { Set-ExoRowTone -Row $script:gridTrace.Rows[$Index] -Tone 'Danger' }
                'Pending|FilteredAsSpam'{ Set-ExoRowTone -Row $script:gridTrace.Rows[$Index] -Tone 'Warning' }
                'Delivered|Resolved'    { Set-ExoRowTone -Row $script:gridTrace.Rows[$Index] -Tone 'Success' }
                default                 { Set-ExoRowTone -Row $script:gridTrace.Rows[$Index] -Tone 'Normal' }
            }
        }

        $script:lblTraceSummary.Text = "$($Messages.Count) message(s). Double-click a row for the delivery detail."
        Write-AppLog "Trace returned $($Messages.Count) message(s)." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText "Trace complete - $($Messages.Count) message(s)."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Message trace failed.'
        Show-ErrorMessage "The message trace failed.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Message trace failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Show-ExoTraceDetail {
    param($Message)

    if (-not $Message) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading delivery detail...'
        if (Test-ExoMessageTraceV2) {
            $Detail = @(Get-MessageTraceDetailV2 -MessageTraceId $Message.MessageTraceId -RecipientAddress $Message.RecipientAddress -ErrorAction Stop | Sort-Object Date)
        }
        else {
            $Detail = @(Get-MessageTraceDetail -MessageTraceId $Message.MessageTraceId -RecipientAddress $Message.RecipientAddress -ErrorAction Stop | Sort-Object Date)
        }
    }
    catch {
        Set-BusyState -Busy $false
        Show-ErrorMessage "Could not read the delivery detail.`r`n`r`n$($_.Exception.Message)"
        return
    }
    finally {
        Set-BusyState -Busy $false
    }

    $Dialog = New-Object System.Windows.Forms.Form
    $Dialog.Text = "Delivery detail - $($Message.Subject)"
    $Dialog.Size = New-Object System.Drawing.Size(1000,560)
    $Dialog.StartPosition = 'CenterParent'
    $Dialog.MinimumSize = New-Object System.Drawing.Size(760,420)

    $Header = New-Object System.Windows.Forms.Label
    $Header.Text = "From $($Message.SenderAddress) to $($Message.RecipientAddress)   |   $($Message.Received)   |   $($Message.Status)"
    $Header.Location = New-Object System.Drawing.Point(14,12)
    $Header.Size = New-Object System.Drawing.Size(940,20)
    $Header.Anchor = "Top,Left,Right"
    $Dialog.Controls.Add($Header)

    $Grid = New-Object System.Windows.Forms.DataGridView
    $Grid.Location = New-Object System.Drawing.Point(14,40)
    $Grid.Size = New-Object System.Drawing.Size(956,430)
    $Grid.Anchor = "Top,Bottom,Left,Right"
    $Grid.ReadOnly = $true
    $Grid.AllowUserToAddRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.SelectionMode = 'FullRowSelect'
    $Grid.AutoSizeColumnsMode = 'Fill'
    [void]$Grid.Columns.Add('Date','Date')
    [void]$Grid.Columns.Add('Event','Event')
    [void]$Grid.Columns.Add('Action','Action')
    [void]$Grid.Columns.Add('Detail','Detail')
    $Grid.Columns['Detail'].FillWeight = 220
    $Dialog.Controls.Add($Grid)

    foreach ($Entry in $Detail) {
        [void]$Grid.Rows.Add($Entry.Date, [string]$Entry.Event, [string]$Entry.Action, [string]$Entry.Detail)
    }

    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Text = 'Close'
    $CloseButton.Size = New-Object System.Drawing.Size(90,30)
    $CloseButton.Location = New-Object System.Drawing.Point(880,480)
    $CloseButton.Anchor = "Bottom,Right"
    $CloseButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Dialog.Controls.Add($CloseButton)
    $Dialog.AcceptButton = $CloseButton

    Apply-MSToolkitSharedTheme -Root $Dialog
    Register-MSToolkitComboFiltersOn -Root $Dialog
    [void]$Dialog.ShowDialog()
}

# ---------------------------------------------------------------------------
# Group delivery check
#
# "It went to the group but I never got it." Expands the group and traces the
# same message per member, so a member who did not receive it is obvious.
# ---------------------------------------------------------------------------
function Invoke-ExoGroupDeliveryCheck {
    if (-not $script:ExchangeConnected) { return }

    $GroupText = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboGroup.Text
    if ([string]::IsNullOrWhiteSpace($GroupText)) {
        Show-InfoMessage 'Enter a distribution group first.'
        return
    }

    $FocusUser = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboGroupUser.Text
    $Sender    = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboGroupSender.Text
    $Subject   = "$($script:txtGroupSubject.Text)".Trim()

    try {
        Set-BusyState -Busy $true -StatusText 'Expanding the group...'
        $script:gridGroup.Rows.Clear()

        $Group = Get-DistributionGroup -Identity $GroupText -ErrorAction Stop
        $Members = @(Get-DistributionGroupMember -Identity $Group.Identity -ResultSize Unlimited -ErrorAction Stop |
            Where-Object { $_.PrimarySmtpAddress } |
            Sort-Object DisplayName)

        Write-AppLog "$($Group.DisplayName) has $($Members.Count) member(s)."

        if ($Members.Count -gt $script:GroupMemberLimit) {
            $Answer = [System.Windows.Forms.MessageBox]::Show(
                "$($Group.DisplayName) has $($Members.Count) members. Each one is a separate trace, so this will take a while.`r`n`r`nCheck all of them?",
                'Large group',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                Set-BusyState -Busy $false -StatusText 'Cancelled.'
                return
            }
        }

        $Start = $script:dtGroupStart.Value
        $End   = $script:dtGroupEnd.Value

        # What the group address itself received, which is the baseline everything
        # else is compared against.
        $GroupMessages = Get-ExoTraceRows -Sender $Sender -Recipient ([string]$Group.PrimarySmtpAddress) -Start $Start -End $End -Status 'Any'
        if ($Subject) { $GroupMessages = @($GroupMessages | Where-Object { "$($_.Subject)" -like "*$Subject*" }) }

        Write-AppLog "The group address received $($GroupMessages.Count) matching message(s) in that window."

        $Position = 0
        $Missing = 0

        foreach ($Member in $Members) {
            $Position++
            $Address = [string]$Member.PrimarySmtpAddress
            Set-BusyState -Busy $true -StatusText "Tracing $Position of $($Members.Count): $Address"

            $Delivered = 0
            $Failed = 0
            $Last = ''
            $LastStatus = ''

            try {
                $MemberMessages = Get-ExoTraceRows -Sender $Sender -Recipient $Address -Start $Start -End $End -Status 'Any'
                if ($Subject) { $MemberMessages = @($MemberMessages | Where-Object { "$($_.Subject)" -like "*$Subject*" }) }

                $Delivered = @($MemberMessages | Where-Object { "$($_.Status)" -match 'Delivered|Resolved' }).Count
                $Failed    = @($MemberMessages | Where-Object { "$($_.Status)" -match 'Failed|Quarantined' }).Count

                if ($MemberMessages.Count -gt 0) {
                    $Last = $MemberMessages[0].Received
                    $LastStatus = [string]$MemberMessages[0].Status
                }
            }
            catch {
                $LastStatus = "Trace failed: $($_.Exception.Message)"
            }

            $Index = $script:gridGroup.Rows.Add(
                [string]$Member.DisplayName,
                $Address,
                [string]$Member.RecipientType,
                $Delivered,
                $Failed,
                $Last,
                $LastStatus
            )

            if ($Delivered -eq 0) {
                Set-ExoRowTone -Row $script:gridGroup.Rows[$Index] -Tone 'Danger'
                $Missing++
            }
            elseif ($Failed -gt 0) {
                Set-ExoRowTone -Row $script:gridGroup.Rows[$Index] -Tone 'Warning'
            }
            else {
                Set-ExoRowTone -Row $script:gridGroup.Rows[$Index] -Tone 'Success'
            }

            if ($FocusUser -and ($Address -ieq $FocusUser -or "$($Member.DisplayName)" -ieq $FocusUser)) {
                $script:gridGroup.Rows[$Index].DefaultCellStyle.Font = New-Object System.Drawing.Font($script:gridGroup.Font, [System.Drawing.FontStyle]::Bold)
                $script:gridGroup.Rows[$Index].Selected = $true
                Write-AppLog "$Address received $Delivered of the $($GroupMessages.Count) message(s) sent to the group." 'INFO'
            }
        }

        $script:lblGroupSummary.Text = "$($Group.DisplayName): $($GroupMessages.Count) message(s) to the group, $($Members.Count) member(s), $Missing with nothing delivered"

        if ($Missing -gt 0) {
            $script:lblGroupSummary.ForeColor = (Get-MSToolkitThemePalette).Danger
            Write-AppLog "$Missing member(s) received none of the matching messages." 'WARN'
        }
        else {
            $script:lblGroupSummary.ForeColor = (Get-MSToolkitThemePalette).Success
            Write-AppLog 'Every member received at least one matching message.' 'SUCCESS'
        }

        Set-BusyState -Busy $false -StatusText 'Group delivery check complete.'
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Group delivery check failed.'
        Show-ErrorMessage "The group delivery check failed.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Group delivery check failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Mailbox report
# ---------------------------------------------------------------------------
function Add-ExoReportRow {
    param(
        [string]$Section,
        [string]$Name,
        $Value,
        [string]$Tone = 'Normal'
    )

    $Text = ''
    if ($null -ne $Value) { $Text = (@($Value) -join '; ') }

    $Index = $script:gridMailbox.Rows.Add($Section,$Name,$Text)
    if ($Tone -ne 'Normal') { Set-ExoRowTone -Row $script:gridMailbox.Rows[$Index] -Tone $Tone }
}

function Invoke-ExoMailboxReport {
    if (-not $script:ExchangeConnected) { return }

    $Identity = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboMailboxUser.Text
    if ([string]::IsNullOrWhiteSpace($Identity)) {
        Show-InfoMessage 'Enter a user first.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Building the mailbox report...'
        $script:gridMailbox.Rows.Clear()

        $User = Resolve-ExchangeUser -Identity $Identity
        $Smtp = [string]$User.PrimarySmtpAddress
        Write-AppLog "Building a mailbox report for $Smtp."

        $Mailbox = Get-Mailbox -Identity $Smtp -ErrorAction Stop

        Add-ExoReportRow -Section 'Mailbox' -Name 'Display name'  -Value $Mailbox.DisplayName
        Add-ExoReportRow -Section 'Mailbox' -Name 'Primary SMTP'  -Value $Mailbox.PrimarySmtpAddress
        Add-ExoReportRow -Section 'Mailbox' -Name 'Type'          -Value $Mailbox.RecipientTypeDetails
        Add-ExoReportRow -Section 'Mailbox' -Name 'Aliases'       -Value (@($Mailbox.EmailAddresses) -match '^smtp:' -replace '^smtp:','')
        Add-ExoReportRow -Section 'Mailbox' -Name 'Hidden from address lists' -Value $Mailbox.HiddenFromAddressListsEnabled
        Add-ExoReportRow -Section 'Mailbox' -Name 'Litigation hold' -Value $Mailbox.LitigationHoldEnabled
        Add-ExoReportRow -Section 'Mailbox' -Name 'Archive state'   -Value $Mailbox.ArchiveStatus

        $ForwardTone = 'Normal'
        if ($Mailbox.ForwardingSmtpAddress -or $Mailbox.ForwardingAddress) { $ForwardTone = 'Danger' }
        Add-ExoReportRow -Section 'Forwarding' -Name 'ForwardingSmtpAddress' -Value $Mailbox.ForwardingSmtpAddress -Tone $ForwardTone
        Add-ExoReportRow -Section 'Forwarding' -Name 'ForwardingAddress'     -Value $Mailbox.ForwardingAddress -Tone $ForwardTone
        Add-ExoReportRow -Section 'Forwarding' -Name 'Also deliver to mailbox' -Value $Mailbox.DeliverToMailboxAndForward

        try {
            $Statistics = Get-MailboxStatistics -Identity $Smtp -ErrorAction Stop
            Add-ExoReportRow -Section 'Size' -Name 'Items'        -Value $Statistics.ItemCount
            Add-ExoReportRow -Section 'Size' -Name 'Total size'   -Value $Statistics.TotalItemSize
            Add-ExoReportRow -Section 'Size' -Name 'Deleted size' -Value $Statistics.TotalDeletedItemSize
            Add-ExoReportRow -Section 'Size' -Name 'Last logon'   -Value $Statistics.LastLogonTime
        }
        catch {
            Add-ExoReportRow -Section 'Size' -Name 'Statistics' -Value "Unavailable: $($_.Exception.Message)" -Tone 'Warning'
        }

        Add-ExoReportRow -Section 'Quota' -Name 'Issue warning at' -Value $Mailbox.IssueWarningQuota
        Add-ExoReportRow -Section 'Quota' -Name 'Prohibit send at' -Value $Mailbox.ProhibitSendQuota

        try {
            $FullAccess = @(Get-MailboxPermission -Identity $Smtp -ErrorAction Stop |
                Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and "$($_.User)" -notmatch 'NT AUTHORITY|S-1-5' })

            if ($FullAccess.Count -eq 0) {
                Add-ExoReportRow -Section 'Access' -Name 'Full access' -Value 'None'
            }
            else {
                foreach ($Permission in $FullAccess) {
                    Add-ExoReportRow -Section 'Access' -Name 'Full access' -Value $Permission.User -Tone 'Warning'
                }
            }

            $SendAs = @(Get-RecipientPermission -Identity $Smtp -ErrorAction Stop |
                Where-Object { $_.AccessRights -contains 'SendAs' -and "$($_.Trustee)" -notmatch 'NT AUTHORITY|S-1-5' })

            if ($SendAs.Count -eq 0) {
                Add-ExoReportRow -Section 'Access' -Name 'Send as' -Value 'None'
            }
            else {
                foreach ($Permission in $SendAs) {
                    Add-ExoReportRow -Section 'Access' -Name 'Send as' -Value $Permission.Trustee -Tone 'Warning'
                }
            }

            Add-ExoReportRow -Section 'Access' -Name 'Send on behalf' -Value $Mailbox.GrantSendOnBehalfTo
        }
        catch {
            Add-ExoReportRow -Section 'Access' -Name 'Permissions' -Value "Unavailable: $($_.Exception.Message)" -Tone 'Warning'
        }

        try {
            $AutoReply = Get-MailboxAutoReplyConfiguration -Identity $Smtp -ErrorAction Stop
            $Tone = 'Normal'
            if ("$($AutoReply.AutoReplyState)" -ne 'Disabled') { $Tone = 'Warning' }

            Add-ExoReportRow -Section 'Automatic replies' -Name 'State' -Value $AutoReply.AutoReplyState -Tone $Tone
            if ("$($AutoReply.AutoReplyState)" -eq 'Scheduled') {
                Add-ExoReportRow -Section 'Automatic replies' -Name 'From' -Value $AutoReply.StartTime
                Add-ExoReportRow -Section 'Automatic replies' -Name 'Until' -Value $AutoReply.EndTime
            }
        }
        catch {
            Add-ExoReportRow -Section 'Automatic replies' -Name 'State' -Value "Unavailable: $($_.Exception.Message)" -Tone 'Warning'
        }

        try {
            $Groups = @(Get-Recipient -Filter "Members -eq '$($Mailbox.DistinguishedName)'" -RecipientTypeDetails MailUniversalDistributionGroup,MailUniversalSecurityGroup -ResultSize Unlimited -ErrorAction Stop |
                Sort-Object DisplayName)

            if ($Groups.Count -eq 0) {
                Add-ExoReportRow -Section 'Groups' -Name 'Direct membership' -Value 'None'
            }
            else {
                foreach ($Group in $Groups) {
                    Add-ExoReportRow -Section 'Groups' -Name 'Direct membership' -Value "$($Group.DisplayName) <$($Group.PrimarySmtpAddress)>"
                }
            }
        }
        catch {
            Add-ExoReportRow -Section 'Groups' -Name 'Direct membership' -Value "Unavailable: $($_.Exception.Message)" -Tone 'Warning'
        }

        $script:lblMailboxSummary.Text = "$($Mailbox.DisplayName) <$Smtp>"
        Write-AppLog "Mailbox report complete for $Smtp." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText "Mailbox report complete for $Smtp."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Mailbox report failed.'
        Show-ErrorMessage "The mailbox report failed.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Mailbox report failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Quarantine
#
# The everyday "where is my email" case. Searching is read-only; releasing is
# the one write here and always names the messages first.
# ---------------------------------------------------------------------------
function Invoke-ExoQuarantineSearch {
    if (-not $script:ExchangeConnected) { return }

    $Recipient = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboQuarRecipient.Text
    $Sender    = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboQuarSender.Text
    $Subject   = "$($script:txtQuarSubject.Text)".Trim()

    try {
        Set-BusyState -Busy $true -StatusText 'Searching quarantine...'
        $script:gridQuarantine.Rows.Clear()

        $Params = @{
            StartReceivedDate = $script:dtQuarStart.Value
            EndReceivedDate   = $script:dtQuarEnd.Value
            PageSize          = 1000
            ErrorAction       = 'Stop'
        }

        if ($Recipient) { $Params['RecipientAddress'] = $Recipient }
        if ($Sender)    { $Params['SenderAddress'] = $Sender }

        $Type = [string]$script:cboQuarType.SelectedItem
        if ($Type -and $Type -ne 'Any') { $Params['QuarantineTypes'] = $Type }

        $Messages = @(Get-QuarantineMessage @Params | Sort-Object ReceivedTime -Descending)

        if ($Subject) {
            $Messages = @($Messages | Where-Object { "$($_.Subject)" -like "*$Subject*" })
        }

        foreach ($Message in $Messages) {
            $Released = 'No'
            if ($Message.Released) { $Released = 'Yes' }

            $Index = $script:gridQuarantine.Rows.Add(
                $Message.ReceivedTime,
                [string]$Message.SenderAddress,
                (@($Message.RecipientAddress) -join '; '),
                [string]$Message.Subject,
                [string]$Message.QuarantineTypes,
                [string]$Message.PolicyName,
                $Released,
                [string]$Message.Expires
            )

            $script:gridQuarantine.Rows[$Index].Tag = $Message

            if ($Message.Released) {
                Set-ExoRowTone -Row $script:gridQuarantine.Rows[$Index] -Tone 'Success'
            }
            elseif ("$($Message.QuarantineTypes)" -match 'Malware|Phish') {
                Set-ExoRowTone -Row $script:gridQuarantine.Rows[$Index] -Tone 'Danger'
            }
            else {
                Set-ExoRowTone -Row $script:gridQuarantine.Rows[$Index] -Tone 'Warning'
            }
        }

        $script:lblQuarSummary.Text = "$($Messages.Count) quarantined message(s). Select rows and use Release to deliver them."
        Write-AppLog "Quarantine search returned $($Messages.Count) message(s)." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText "Quarantine search complete - $($Messages.Count) message(s)."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Quarantine search failed.'
        Show-ErrorMessage "The quarantine search failed.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Quarantine search failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-ExoQuarantineRelease {
    if (-not $script:ExchangeConnected) { return }

    $Selected = @($script:gridQuarantine.SelectedRows | Where-Object { $_.Tag })

    if ($Selected.Count -eq 0) {
        Show-InfoMessage 'Select one or more quarantined messages first.'
        return
    }

    $Pending = @($Selected | Where-Object { -not $_.Tag.Released })

    if ($Pending.Count -eq 0) {
        Show-InfoMessage 'Every selected message has already been released.'
        return
    }

    $Lines = foreach ($Row in $Pending) {
        "  $($Row.Tag.ReceivedTime)  $($Row.Tag.SenderAddress)  ->  $(@($Row.Tag.RecipientAddress) -join '; ')`r`n    $($Row.Tag.Subject)"
    }

    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Release $($Pending.Count) message(s) to their original recipients?`r`n`r`n$($Lines -join "`r`n")`r`n`r`nReleased mail is delivered immediately and cannot be recalled. Malware and high confidence phishing should normally stay quarantined.",
        'Release quarantined mail',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog 'Quarantine release cancelled.'
        return
    }

    $Released = 0
    $Failed = 0

    foreach ($Row in $Pending) {
        try {
            Set-BusyState -Busy $true -StatusText "Releasing: $($Row.Tag.Subject)"
            Release-QuarantineMessage -Identity $Row.Tag.Identity -ReleaseToAll -ErrorAction Stop
            $Released++
            Write-AppLog "Released: $($Row.Tag.Subject) from $($Row.Tag.SenderAddress)" 'SUCCESS'
            $Row.Cells['Released'].Value = 'Yes'
            Set-ExoRowTone -Row $Row -Tone 'Success'
        }
        catch {
            $Failed++
            Write-AppLog "Release failed for '$($Row.Tag.Subject)': $($_.Exception.Message)" 'ERROR'
        }
    }

    Set-BusyState -Busy $false -StatusText "Released $Released message(s), $Failed failed."
    Show-InfoMessage "Released $Released message(s).`r`n$Failed failed."
}





# ---------------------------------------------------------------------------
# Message header analyser - offline, no connection needed
# ---------------------------------------------------------------------------
function Add-ExoHeaderRow {
    param([string]$Section,[string]$Name,$Value,[string]$Tone = 'Normal')

    $Index = $script:gridHeaders.Rows.Add($Section,$Name,[string]$Value)
    if ($Tone -ne 'Normal') { Set-ExoRowTone -Row $script:gridHeaders.Rows[$Index] -Tone $Tone }
}

function Get-ExoUnfoldedHeaders {
    param([string]$Raw)

    # Header values wrap onto continuation lines that start with whitespace.
    $Lines = $Raw -split "`r?`n"
    $Unfolded = New-Object System.Collections.Generic.List[string]

    foreach ($Line in $Lines) {
        if ($Line -match '^\s+' -and $Unfolded.Count -gt 0) {
            $Unfolded[$Unfolded.Count - 1] = $Unfolded[$Unfolded.Count - 1] + ' ' + $Line.Trim()
        }
        else {
            $Unfolded.Add($Line)
        }
    }

    return $Unfolded.ToArray()
}

function Invoke-ExoHeaderAnalysis {
    $Raw = "$($script:txtHeaders.Text)"

    if ([string]::IsNullOrWhiteSpace($Raw)) {
        Show-InfoMessage 'Paste the message headers first. In Outlook: open the message, File, Properties, Internet headers.'
        return
    }

    $script:gridHeaders.Rows.Clear()
    $Lines = Get-ExoUnfoldedHeaders -Raw $Raw

    $Get = {
        param([string]$Name)
        $Match = $Lines | Where-Object { $_ -match "^$Name\s*:" } | Select-Object -First 1
        if ($Match) { return ($Match -replace "^$Name\s*:\s*", '').Trim() }
        return ''
    }

    foreach ($Name in 'From','To','Cc','Reply-To','Return-Path','Subject','Date','Message-ID') {
        $Value = & $Get $Name
        if ($Value) { Add-ExoHeaderRow -Section 'Message' -Name $Name -Value $Value }
    }

    # Envelope sender vs display From is the classic spoofing tell.
    $From = & $Get 'From'
    $ReturnPath = & $Get 'Return-Path'

    if ($From -and $ReturnPath) {
        $FromDomain = ''
        $PathDomain = ''
        if ($From -match '@([^>\s]+)')       { $FromDomain = $Matches[1].TrimEnd('>').ToLower() }
        if ($ReturnPath -match '@([^>\s]+)') { $PathDomain = $Matches[1].TrimEnd('>').ToLower() }

        if ($FromDomain -and $PathDomain -and $FromDomain -ne $PathDomain) {
            Add-ExoHeaderRow -Section 'Message' -Name 'Envelope mismatch' -Value "From is $FromDomain but Return-Path is $PathDomain" -Tone 'Warning'
        }
    }

    $Auth = & $Get 'Authentication-Results'

    if ($Auth) {
        foreach ($Check in 'spf','dkim','dmarc','compauth') {
            if ($Auth -match "$Check=([a-z]+)") {
                $Result = $Matches[1].ToLower()

                $Tone = 'Normal'
                switch ($Result) {
                    'pass'      { $Tone = 'Success' }
                    'bestguesspass' { $Tone = 'Warning' }
                    'none'      { $Tone = 'Warning' }
                    'neutral'   { $Tone = 'Warning' }
                    'softfail'  { $Tone = 'Warning' }
                    'fail'      { $Tone = 'Danger' }
                    'permerror' { $Tone = 'Danger' }
                    'temperror' { $Tone = 'Warning' }
                }

                Add-ExoHeaderRow -Section 'Authentication' -Name $Check.ToUpper() -Value $Result -Tone $Tone
            }
        }

        Add-ExoHeaderRow -Section 'Authentication' -Name 'Raw' -Value $Auth
    }
    else {
        Add-ExoHeaderRow -Section 'Authentication' -Name 'Authentication-Results' -Value 'Not present' -Tone 'Warning'
    }

    $Forefront = & $Get 'X-Forefront-Antispam-Report'

    if ($Forefront) {
        foreach ($Pair in @(@('SCL','Spam confidence'),@('SFV','Filter verdict'),@('CIP','Connecting IP'),@('CTRY','Country'),@('H','Sending host'))) {
            if ($Forefront -match "(?:^|;)\s*$($Pair[0]):([^;]+)") {
                $Value = $Matches[1].Trim()
                $Tone = 'Normal'

                if ($Pair[0] -eq 'SCL') {
                    $Number = 0
                    if ([int]::TryParse($Value,[ref]$Number)) {
                        if ($Number -ge 6) { $Tone = 'Danger' }
                        elseif ($Number -ge 5) { $Tone = 'Warning' }
                        elseif ($Number -eq -1) { $Tone = 'Success' }
                    }
                }

                Add-ExoHeaderRow -Section 'Filtering' -Name $Pair[1] -Value $Value -Tone $Tone
            }
        }
    }

    $Antispam = & $Get 'X-Microsoft-Antispam'
    if ($Antispam -match 'BCL:(\d+)') {
        Add-ExoHeaderRow -Section 'Filtering' -Name 'Bulk complaint level' -Value $Matches[1]
    }

    # Received chain, oldest first, with the delay introduced at each hop.
    $Received = @($Lines | Where-Object { $_ -match '^Received\s*:' })
    [array]::Reverse($Received)

    $Previous = $null
    $Hop = 0

    foreach ($Line in $Received) {
        $Hop++
        $Value = ($Line -replace '^Received\s*:\s*','').Trim()

        $Stamp = $null
        if ($Value -match ';\s*(.+)$') {
            $Text = $Matches[1].Trim()
            try { $Stamp = [datetime]::Parse($Text) } catch { $Stamp = $null }
        }

        $Delay = ''
        if ($Stamp -and $Previous) {
            $Seconds = [math]::Round(($Stamp - $Previous).TotalSeconds)
            $Delay = "+$Seconds s"
        }
        if ($Stamp) { $Previous = $Stamp }

        $Short = $Value
        if ($Short.Length -gt 180) { $Short = $Short.Substring(0,180) + '...' }

        $Tone = 'Normal'
        if ($Delay -match '\+(\d+) s' -and [int]$Matches[1] -gt 60) { $Tone = 'Warning' }

        Add-ExoHeaderRow -Section "Hop $Hop" -Name $Delay -Value $Short -Tone $Tone
    }

    $script:lblHeaderSummary.Text = "$($Received.Count) hop(s) parsed. SPF, DKIM and DMARC are shown above the hop list."
    Write-AppLog "Analysed pasted headers: $($Received.Count) hop(s)." 'SUCCESS'
}

# ---------------------------------------------------------------------------
# Transport rule matcher
#
# Honest about its limits: it evaluates the sender, recipient and subject
# conditions it can, and says which rules it could not fully judge.
# ---------------------------------------------------------------------------
function Test-ExoWordMatch {
    param($Words,[string]$Value)

    foreach ($Word in @($Words)) {
        if ([string]::IsNullOrWhiteSpace($Word)) { continue }
        if ("$Value" -like "*$Word*") { return $true }
    }
    return $false
}

function Test-ExoAddressMatch {
    param($Values,[string]$Address)

    foreach ($Value in @($Values)) {
        if ([string]::IsNullOrWhiteSpace($Value)) { continue }
        $Text = [string]$Value
        if ($Address -ieq $Text) { return $true }
        if ($Text -match '@' -and $Address -ieq $Text) { return $true }
        if ($Address -like "*$Text*") { return $true }
    }
    return $false
}

function Invoke-ExoTransportRuleMatch {
    if (-not $script:ExchangeConnected) { return }

    $Sender    = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboRuleSender.Text
    $Recipient = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboRuleRecipient.Text
    $Subject   = "$($script:txtRuleSubject.Text)".Trim()

    if (-not $Sender -and -not $Recipient -and -not $Subject) {
        Show-InfoMessage 'Enter a sender, a recipient or a subject to test.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading mail flow rules...'
        $script:gridTransport.Rows.Clear()

        $Rules = @(Get-TransportRule -ResultSize Unlimited -ErrorAction Stop | Sort-Object Priority)
        $SenderDomain = ''
        $RecipientDomain = ''
        if ($Sender -match '@(.+)$')    { $SenderDomain = $Matches[1].ToLower() }
        if ($Recipient -match '@(.+)$') { $RecipientDomain = $Matches[1].ToLower() }

        $Matched = 0

        foreach ($Rule in $Rules) {
            $Reasons = New-Object System.Collections.Generic.List[string]
            $Unevaluated = New-Object System.Collections.Generic.List[string]

            if ($Sender) {
                if (Test-ExoAddressMatch -Values $Rule.From -Address $Sender)                 { $Reasons.Add("From is $Sender") }
                if (Test-ExoWordMatch -Words $Rule.FromAddressContainsWords -Value $Sender)   { $Reasons.Add('Sender address contains a listed word') }
                if ($SenderDomain -and (@($Rule.SenderDomainIs) -contains $SenderDomain))     { $Reasons.Add("Sender domain is $SenderDomain") }
            }

            if ($Recipient) {
                if (Test-ExoAddressMatch -Values $Rule.SentTo -Address $Recipient)                  { $Reasons.Add("Sent to $Recipient") }
                if (Test-ExoWordMatch -Words $Rule.RecipientAddressContainsWords -Value $Recipient) { $Reasons.Add('Recipient address contains a listed word') }
                if ($RecipientDomain -and (@($Rule.RecipientDomainIs) -contains $RecipientDomain))  { $Reasons.Add("Recipient domain is $RecipientDomain") }
            }

            if ($Subject) {
                if (Test-ExoWordMatch -Words $Rule.SubjectContainsWords -Value $Subject)       { $Reasons.Add('Subject contains a listed word') }
                if (Test-ExoWordMatch -Words $Rule.SubjectOrBodyContainsWords -Value $Subject) { $Reasons.Add('Subject or body contains a listed word') }
            }

            # Conditions this tool cannot judge from three fields alone.
            foreach ($Property in 'AttachmentContainsWords','AttachmentExtensionMatchesWords','MessageSizeOver','ContentCharacterSetContainsWords','SenderIpRanges','HasClassification','AnyOfRecipientAddressContainsWords','ExceptIfFrom','ExceptIfSentTo','ExceptIfSubjectContainsWords') {
                if (@($Rule.$Property).Count -gt 0) { $Unevaluated.Add($Property) }
            }

            if ($Reasons.Count -eq 0) { continue }
            $Matched++

            $Actions = New-Object System.Collections.Generic.List[string]
            foreach ($Pair in @(
                @('Rejects the message', $Rule.RejectMessageReasonText),
                @('Deletes the message', $Rule.DeleteMessage),
                @('Prepends the subject', $Rule.PrependSubject),
                @('Sets SCL', $Rule.SetSCL),
                @('Redirects to', $Rule.RedirectMessageTo),
                @('Blind copies to', $Rule.BlindCopyTo),
                @('Moderated by', $Rule.ModerateMessageByUser),
                @('Adds a disclaimer', $Rule.ApplyHtmlDisclaimerText),
                @('Quarantines', $Rule.Quarantine)
            )) {
                $Value = @($Pair[1]) -join '; '
                if (-not [string]::IsNullOrWhiteSpace($Value) -and $Value -ne 'False') {
                    $Actions.Add("$($Pair[0]): $Value")
                }
            }

            $State = 'Enabled'
            if ("$($Rule.State)" -ne 'Enabled') { $State = [string]$Rule.State }

            $Index = $script:gridTransport.Rows.Add(
                [string]$Rule.Priority,
                [string]$Rule.Name,
                $State,
                [string]$Rule.Mode,
                ($Reasons -join ' | '),
                ($Actions -join ' | '),
                ($Unevaluated -join ', ')
            )

            if ($State -ne 'Enabled') {
                Set-ExoRowTone -Row $script:gridTransport.Rows[$Index] -Tone 'Normal'
            }
            elseif ($Actions.Count -gt 0) {
                Set-ExoRowTone -Row $script:gridTransport.Rows[$Index] -Tone 'Warning'
            }
        }

        $script:lblTransportSummary.Text = "$Matched of $($Rules.Count) rule(s) could act on this message. The last column lists conditions this check could not evaluate."
        Write-AppLog "$Matched of $($Rules.Count) mail flow rule(s) matched." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText "$Matched matching rule(s)."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Mail flow rule check failed.'
        Show-ErrorMessage "Could not read the mail flow rules.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Mail flow rule check failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Calendar and folder permissions
# ---------------------------------------------------------------------------
function Invoke-ExoFolderPermissions {
    if (-not $script:ExchangeConnected) { return }

    $Identity = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboFolderUser.Text
    if ([string]::IsNullOrWhiteSpace($Identity)) {
        Show-InfoMessage 'Enter a mailbox or room first.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading folder permissions...'
        $script:gridFolders.Rows.Clear()

        $Mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
        $Smtp = [string]$Mailbox.PrimarySmtpAddress
        Write-AppLog "Reading folder permissions for $Smtp."

        # Top of information store first: that is what "see my whole mailbox" needs.
        $Folders = @(
            @{ Path = "$($Smtp):\";          Label = 'Top of information store' },
            @{ Path = "$($Smtp):\Calendar";  Label = 'Calendar' },
            @{ Path = "$($Smtp):\Inbox";     Label = 'Inbox' },
            @{ Path = "$($Smtp):\Contacts";  Label = 'Contacts' },
            @{ Path = "$($Smtp):\Sent Items";Label = 'Sent Items' }
        )

        $Total = 0

        foreach ($Folder in $Folders) {
            try {
                $Permissions = @(Get-MailboxFolderPermission -Identity $Folder.Path -ErrorAction Stop |
                    Where-Object { "$($_.User)" -notmatch '^(Default|Anonymous)$' -or "$($_.AccessRights)" -notmatch '^None$' })

                foreach ($Permission in $Permissions) {
                    $Rights = (@($Permission.AccessRights) -join ', ')
                    $Index = $script:gridFolders.Rows.Add(
                        $Folder.Label,
                        [string]$Permission.User,
                        $Rights,
                        [string]$Permission.SharingPermissionFlags
                    )

                    if ("$($Permission.User)" -match '^(Default|Anonymous)$') {
                        if ($Rights -match 'Owner|Editor|Author|Reviewer') {
                            Set-ExoRowTone -Row $script:gridFolders.Rows[$Index] -Tone 'Danger'
                        }
                    }
                    elseif ("$($Permission.SharingPermissionFlags)" -match 'Delegate') {
                        Set-ExoRowTone -Row $script:gridFolders.Rows[$Index] -Tone 'Warning'
                    }

                    $Total++
                }
            }
            catch {
                $script:gridFolders.Rows.Add($Folder.Label,'(could not read)', $_.Exception.Message,'') | Out-Null
            }
        }

        $script:lblFolderSummary.Text = "$($Mailbox.DisplayName) <$Smtp> - $Total permission entry(s). Red: Default or Anonymous can read the folder."
        Write-AppLog "Found $Total folder permission entry(s) for $Smtp." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText "Folder permissions read for $Smtp."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Folder permission lookup failed.'
        Show-ErrorMessage "Could not read the folder permissions.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Folder permission lookup failed: $($_.Exception.Message)" 'ERROR'
    }
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
$script:ExchangeConnected = $false
$script:AcceptedDomains   = @()
$script:ActionControls    = @()
$script:UserPickerCombos  = @()
$script:GroupMemberLimit  = 40

function New-ExoLabel {
    param([System.Windows.Forms.Control]$Parent,[string]$Text,[int]$X,[int]$Y,[int]$Width = 120)

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Text
    $Label.Location = New-Object System.Drawing.Point($X,($Y + 4))
    $Label.Size = New-Object System.Drawing.Size($Width,20)
    $Parent.Controls.Add($Label)
    return $Label
}

function New-ExoUserCombo {
    param([System.Windows.Forms.Control]$Parent,[int]$X,[int]$Y,[int]$Width = 300)

    $Combo = New-Object System.Windows.Forms.ComboBox
    $Combo.Location = New-Object System.Drawing.Point($X,$Y)
    $Combo.Size = New-Object System.Drawing.Size($Width,24)
    $Combo.DropDownStyle = 'DropDown'
    $Combo.AutoCompleteMode = 'None'
    $Combo.AutoCompleteSource = 'None'
    $Combo.DropDownHeight = 320
    $Parent.Controls.Add($Combo)
    $script:UserPickerCombos += $Combo
    return $Combo
}

function New-ExoButton {
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
$MainForm.Text = 'Exchange Online Tools'
$MainForm.Size = New-Object System.Drawing.Size(1500,940)
$MainForm.MinimumSize = New-Object System.Drawing.Size(1180,780)
$MainForm.StartPosition = 'CenterScreen'
$script:MainForm = $MainForm

# --- header ---------------------------------------------------------------
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = 'Top'
$pnlHeader.Height = 70
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$pnlHeader.Tag = 'TopBar'
$MainForm.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Exchange Online Tools'
$lblTitle.AutoSize = $true
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',15)
$lblTitle.Location = New-Object System.Drawing.Point(18,10)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Rules, trace, quarantine, headers, mail flow rules, permissions, mailbox report.'
$lblSubtitle.AutoSize = $true
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblSubtitle.Location = New-Object System.Drawing.Point(20,42)
$pnlHeader.Controls.Add($lblSubtitle)

# Same accent blue as the original M365 tools - this button had no styling at all
# and was inheriting the default grey.
$script:btnConnect = New-Object System.Windows.Forms.Button
$script:btnConnect.Text = 'Connect to Exchange Online'
$script:btnConnect.Size = New-Object System.Drawing.Size(210,36)
$script:btnConnect.Location = New-Object System.Drawing.Point(($MainForm.ClientSize.Width - 226),14)
$script:btnConnect.Anchor = 'Top,Right'
$script:btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$script:btnConnect.ForeColor = [System.Drawing.Color]::White
$script:btnConnect.FlatStyle = 'Flat'
$script:btnConnect.FlatAppearance.BorderSize = 0
$script:btnConnect.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$script:btnConnect.Add_Click({ Connect-M365Exchange })
$pnlHeader.Controls.Add($script:btnConnect)

# Fixed width and right aligned, so a long organisation name cannot run off the
# edge of the window.
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

# --- status bar and activity log -----------------------------------------
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$script:lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblStatus.Text = 'Connect to Exchange Online to begin.'
[void]$statusStrip.Items.Add($script:lblStatus)
$MainForm.Controls.Add($statusStrip)

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Dock = 'Bottom'
$pnlLog.Height = 170
$MainForm.Controls.Add($pnlLog)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = 'Activity'
$lblLog.Location = New-Object System.Drawing.Point(12,4)
$lblLog.Size = New-Object System.Drawing.Size(120,18)
$pnlLog.Controls.Add($lblLog)

$script:txtLog = New-Object System.Windows.Forms.RichTextBox
$script:txtLog.Location = New-Object System.Drawing.Point(12,24)
$script:txtLog.Size = New-Object System.Drawing.Size(($MainForm.ClientSize.Width - 24),135)
$script:txtLog.Anchor = 'Top,Bottom,Left,Right'
$script:txtLog.ReadOnly = $true
$script:txtLog.Font = New-Object System.Drawing.Font('Consolas',9)
$pnlLog.Controls.Add($script:txtLog)

# --- tabs -----------------------------------------------------------------
$Tabs = New-Object System.Windows.Forms.TabControl
$Tabs.Dock = 'Fill'
$Tabs.Padding = New-Object System.Drawing.Point(14,6)
$MainForm.Controls.Add($Tabs)
$Tabs.BringToFront()

# Tab 1 - inbox rules and forwarding
$tabRules = New-Object System.Windows.Forms.TabPage
$tabRules.Text = '  Inbox Rules  '
$Tabs.TabPages.Add($tabRules)

New-ExoLabel -Parent $tabRules -Text 'Mailbox:' -X 12 -Y 16 -Width 70 | Out-Null
$script:cboRuleUser = New-ExoUserCombo -Parent $tabRules -X 86 -Y 14 -Width 360
New-ExoButton -Parent $tabRules -Text 'Check Rules' -X 460 -Y 13 -Width 120 -NeedsConnection -OnClick { Invoke-ExoRuleCheck } | Out-Null
New-ExoButton -Parent $tabRules -Text 'Export CSV'  -X 590 -Y 13 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridRules -BaseName 'InboxRules' } | Out-Null
New-ExoButton -Parent $tabRules -Text 'Copy'        -X 708 -Y 13 -Width 80  -NeedsConnection -OnClick { Copy-ExoGrid -Grid $script:gridRules } | Out-Null

$script:lblRuleSummary = New-Object System.Windows.Forms.Label
$script:lblRuleSummary.Location = New-Object System.Drawing.Point(12,50)
$script:lblRuleSummary.Size = New-Object System.Drawing.Size(900,20)
$script:lblRuleSummary.Anchor = 'Top,Left,Right'
$script:lblRuleSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabRules.Controls.Add($script:lblRuleSummary)

$lblRuleHint = New-Object System.Windows.Forms.Label
$lblRuleHint.Text = 'Red: sends mail outside the tenant. Orange: forwards internally, or deletes the message.'
$lblRuleHint.Location = New-Object System.Drawing.Point(12,72)
$lblRuleHint.Size = New-Object System.Drawing.Size(900,18)
$lblRuleHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
$tabRules.Controls.Add($lblRuleHint)

$script:gridRules = New-ExoGrid -Parent $tabRules -Top 96 -Height ($tabRules.ClientSize.Height - 110)
[void]$script:gridRules.Columns.Add('Source','Source')
[void]$script:gridRules.Columns.Add('Name','Rule')
[void]$script:gridRules.Columns.Add('State','State')
[void]$script:gridRules.Columns.Add('Priority','Priority')
[void]$script:gridRules.Columns.Add('Targets','Sends to')
[void]$script:gridRules.Columns.Add('Conditions','Conditions')
[void]$script:gridRules.Columns.Add('Actions','Actions')
$script:gridRules.Columns['Source'].FillWeight = 60
$script:gridRules.Columns['State'].FillWeight = 45
$script:gridRules.Columns['Priority'].FillWeight = 45
$script:gridRules.Columns['Conditions'].FillWeight = 150
$script:gridRules.Columns['Actions'].FillWeight = 150

# Tab 2 - message trace
$tabTrace = New-Object System.Windows.Forms.TabPage
$tabTrace.Text = '  Message Trace  '
$Tabs.TabPages.Add($tabTrace)

New-ExoLabel -Parent $tabTrace -Text 'Sender:' -X 12 -Y 16 -Width 60 | Out-Null
$script:cboTraceSender = New-ExoUserCombo -Parent $tabTrace -X 76 -Y 14 -Width 300

New-ExoLabel -Parent $tabTrace -Text 'Recipient:' -X 390 -Y 16 -Width 70 | Out-Null
$script:cboTraceRecipient = New-ExoUserCombo -Parent $tabTrace -X 464 -Y 14 -Width 300

New-ExoLabel -Parent $tabTrace -Text 'Subject has:' -X 778 -Y 16 -Width 80 | Out-Null
$script:txtTraceSubject = New-Object System.Windows.Forms.TextBox
$script:txtTraceSubject.Location = New-Object System.Drawing.Point(862,14)
$script:txtTraceSubject.Size = New-Object System.Drawing.Size(240,24)
$tabTrace.Controls.Add($script:txtTraceSubject)

New-ExoLabel -Parent $tabTrace -Text 'From:' -X 12 -Y 52 -Width 60 | Out-Null
$script:dtTraceStart = New-Object System.Windows.Forms.DateTimePicker
$script:dtTraceStart.Location = New-Object System.Drawing.Point(76,50)
$script:dtTraceStart.Size = New-Object System.Drawing.Size(210,24)
$script:dtTraceStart.Format = 'Custom'
$script:dtTraceStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$script:dtTraceStart.Value = (Get-Date).AddDays(-2)
$tabTrace.Controls.Add($script:dtTraceStart)

New-ExoLabel -Parent $tabTrace -Text 'To:' -X 300 -Y 52 -Width 30 | Out-Null
$script:dtTraceEnd = New-Object System.Windows.Forms.DateTimePicker
$script:dtTraceEnd.Location = New-Object System.Drawing.Point(334,50)
$script:dtTraceEnd.Size = New-Object System.Drawing.Size(210,24)
$script:dtTraceEnd.Format = 'Custom'
$script:dtTraceEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$script:dtTraceEnd.Value = (Get-Date)
$tabTrace.Controls.Add($script:dtTraceEnd)

New-ExoLabel -Parent $tabTrace -Text 'Status:' -X 560 -Y 52 -Width 50 | Out-Null
$script:cboTraceStatus = New-Object System.Windows.Forms.ComboBox
$script:cboTraceStatus.Location = New-Object System.Drawing.Point(614,50)
$script:cboTraceStatus.Size = New-Object System.Drawing.Size(150,24)
$script:cboTraceStatus.DropDownStyle = 'DropDownList'
foreach ($Status in 'Any','Delivered','Failed','Pending','Quarantined','FilteredAsSpam','Expanded') {
    [void]$script:cboTraceStatus.Items.Add($Status)
}
$script:cboTraceStatus.SelectedIndex = 0
$tabTrace.Controls.Add($script:cboTraceStatus)

New-ExoButton -Parent $tabTrace -Text 'Run Trace'  -X 778 -Y 49 -Width 120 -NeedsConnection -OnClick { Invoke-ExoMessageTrace } | Out-Null
New-ExoButton -Parent $tabTrace -Text 'Export CSV' -X 906 -Y 49 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridTrace -BaseName 'MessageTrace' } | Out-Null
New-ExoButton -Parent $tabTrace -Text 'Copy'       -X 1024 -Y 49 -Width 78 -NeedsConnection -OnClick { Copy-ExoGrid -Grid $script:gridTrace } | Out-Null

$script:lblTraceSummary = New-Object System.Windows.Forms.Label
$script:lblTraceSummary.Text = 'Up to 90 days of history, but at most 10 days per query. Older mail needs a historical search in the Defender portal.'
$script:lblTraceSummary.Location = New-Object System.Drawing.Point(12,86)
$script:lblTraceSummary.Size = New-Object System.Drawing.Size(1000,18)
$script:lblTraceSummary.Anchor = 'Top,Left,Right'
$script:lblTraceSummary.Font = New-Object System.Drawing.Font('Segoe UI',8)
$tabTrace.Controls.Add($script:lblTraceSummary)

$script:gridTrace = New-ExoGrid -Parent $tabTrace -Top 110 -Height ($tabTrace.ClientSize.Height - 124)
[void]$script:gridTrace.Columns.Add('Received','Received')
[void]$script:gridTrace.Columns.Add('Sender','Sender')
[void]$script:gridTrace.Columns.Add('Recipient','Recipient')
[void]$script:gridTrace.Columns.Add('Subject','Subject')
[void]$script:gridTrace.Columns.Add('Status','Status')
[void]$script:gridTrace.Columns.Add('Size','Size')
$script:gridTrace.Columns['Subject'].FillWeight = 180
$script:gridTrace.Columns['Size'].FillWeight = 50
$script:gridTrace.Add_CellDoubleClick({
    param($Sender,$EventArgs)
    if ($EventArgs.RowIndex -lt 0) { return }
    Show-ExoTraceDetail -Message $script:gridTrace.Rows[$EventArgs.RowIndex].Tag
})

# Tab 3 - group delivery
$tabGroup = New-Object System.Windows.Forms.TabPage
$tabGroup.Text = '  Group Delivery  '
$Tabs.TabPages.Add($tabGroup)

New-ExoLabel -Parent $tabGroup -Text 'Group:' -X 12 -Y 16 -Width 60 | Out-Null
$script:cboGroup = New-Object System.Windows.Forms.ComboBox
$script:cboGroup.Location = New-Object System.Drawing.Point(76,14)
$script:cboGroup.Size = New-Object System.Drawing.Size(300,24)
$script:cboGroup.DropDownStyle = 'DropDown'
$script:cboGroup.AutoCompleteMode = 'None'
$script:cboGroup.AutoCompleteSource = 'None'
$tabGroup.Controls.Add($script:cboGroup)

New-ExoLabel -Parent $tabGroup -Text 'Focus on member:' -X 390 -Y 16 -Width 110 | Out-Null
$script:cboGroupUser = New-ExoUserCombo -Parent $tabGroup -X 504 -Y 14 -Width 260

New-ExoLabel -Parent $tabGroup -Text 'Sender:' -X 778 -Y 16 -Width 60 | Out-Null
$script:cboGroupSender = New-ExoUserCombo -Parent $tabGroup -X 842 -Y 14 -Width 260

New-ExoLabel -Parent $tabGroup -Text 'From:' -X 12 -Y 52 -Width 60 | Out-Null
$script:dtGroupStart = New-Object System.Windows.Forms.DateTimePicker
$script:dtGroupStart.Location = New-Object System.Drawing.Point(76,50)
$script:dtGroupStart.Size = New-Object System.Drawing.Size(210,24)
$script:dtGroupStart.Format = 'Custom'
$script:dtGroupStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$script:dtGroupStart.Value = (Get-Date).AddDays(-2)
$tabGroup.Controls.Add($script:dtGroupStart)

New-ExoLabel -Parent $tabGroup -Text 'To:' -X 300 -Y 52 -Width 30 | Out-Null
$script:dtGroupEnd = New-Object System.Windows.Forms.DateTimePicker
$script:dtGroupEnd.Location = New-Object System.Drawing.Point(334,50)
$script:dtGroupEnd.Size = New-Object System.Drawing.Size(210,24)
$script:dtGroupEnd.Format = 'Custom'
$script:dtGroupEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$script:dtGroupEnd.Value = (Get-Date)
$tabGroup.Controls.Add($script:dtGroupEnd)

New-ExoLabel -Parent $tabGroup -Text 'Subject has:' -X 560 -Y 52 -Width 80 | Out-Null
$script:txtGroupSubject = New-Object System.Windows.Forms.TextBox
$script:txtGroupSubject.Location = New-Object System.Drawing.Point(644,50)
$script:txtGroupSubject.Size = New-Object System.Drawing.Size(120,24)
$tabGroup.Controls.Add($script:txtGroupSubject)

New-ExoButton -Parent $tabGroup -Text 'Check Delivery' -X 778 -Y 49 -Width 130 -NeedsConnection -OnClick { Invoke-ExoGroupDeliveryCheck } | Out-Null
New-ExoButton -Parent $tabGroup -Text 'Export CSV'     -X 916 -Y 49 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridGroup -BaseName 'GroupDelivery' } | Out-Null
New-ExoButton -Parent $tabGroup -Text 'Copy'           -X 1034 -Y 49 -Width 78 -NeedsConnection -OnClick { Copy-ExoGrid -Grid $script:gridGroup } | Out-Null

$script:lblGroupSummary = New-Object System.Windows.Forms.Label
$script:lblGroupSummary.Location = New-Object System.Drawing.Point(12,86)
$script:lblGroupSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblGroupSummary.Anchor = 'Top,Left,Right'
$script:lblGroupSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabGroup.Controls.Add($script:lblGroupSummary)

$lblGroupHint = New-Object System.Windows.Forms.Label
$lblGroupHint.Text = 'Each member is traced separately, so a large group takes a while. Keep the range to 10 days or less. Red rows received nothing matching.'
$lblGroupHint.Location = New-Object System.Drawing.Point(12,108)
$lblGroupHint.Size = New-Object System.Drawing.Size(1000,18)
$lblGroupHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
$tabGroup.Controls.Add($lblGroupHint)

$script:gridGroup = New-ExoGrid -Parent $tabGroup -Top 132 -Height ($tabGroup.ClientSize.Height - 146)
[void]$script:gridGroup.Columns.Add('Member','Member')
[void]$script:gridGroup.Columns.Add('Address','Address')
[void]$script:gridGroup.Columns.Add('Type','Type')
[void]$script:gridGroup.Columns.Add('Delivered','Delivered')
[void]$script:gridGroup.Columns.Add('Failed','Failed')
[void]$script:gridGroup.Columns.Add('Last','Last message')
[void]$script:gridGroup.Columns.Add('LastStatus','Last status')
$script:gridGroup.Columns['Delivered'].FillWeight = 55
$script:gridGroup.Columns['Failed'].FillWeight = 45
$script:gridGroup.Columns['Type'].FillWeight = 70

# Tab 4 - mailbox report
$tabMailbox = New-Object System.Windows.Forms.TabPage
$tabMailbox.Text = '  Mailbox Report  '
$Tabs.TabPages.Add($tabMailbox)

New-ExoLabel -Parent $tabMailbox -Text 'Mailbox:' -X 12 -Y 16 -Width 70 | Out-Null
$script:cboMailboxUser = New-ExoUserCombo -Parent $tabMailbox -X 86 -Y 14 -Width 360
New-ExoButton -Parent $tabMailbox -Text 'Build Report' -X 460 -Y 13 -Width 120 -NeedsConnection -OnClick { Invoke-ExoMailboxReport } | Out-Null
New-ExoButton -Parent $tabMailbox -Text 'Export CSV'   -X 590 -Y 13 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridMailbox -BaseName 'MailboxReport' } | Out-Null
New-ExoButton -Parent $tabMailbox -Text 'Copy'         -X 708 -Y 13 -Width 80  -NeedsConnection -OnClick { Copy-ExoGrid -Grid $script:gridMailbox } | Out-Null

$script:lblMailboxSummary = New-Object System.Windows.Forms.Label
$script:lblMailboxSummary.Location = New-Object System.Drawing.Point(12,50)
$script:lblMailboxSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblMailboxSummary.Anchor = 'Top,Left,Right'
$script:lblMailboxSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabMailbox.Controls.Add($script:lblMailboxSummary)

$lblMailboxHint = New-Object System.Windows.Forms.Label
$lblMailboxHint.Text = 'Forwarding, delegates and send-as are highlighted, because they are how mail leaves a mailbox without the owner noticing.'
$lblMailboxHint.Location = New-Object System.Drawing.Point(12,72)
$lblMailboxHint.Size = New-Object System.Drawing.Size(1000,18)
$lblMailboxHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
$tabMailbox.Controls.Add($lblMailboxHint)

$script:gridMailbox = New-ExoGrid -Parent $tabMailbox -Top 96 -Height ($tabMailbox.ClientSize.Height - 110)
[void]$script:gridMailbox.Columns.Add('Section','Section')
[void]$script:gridMailbox.Columns.Add('Name','Setting')
[void]$script:gridMailbox.Columns.Add('Value','Value')
$script:gridMailbox.Columns['Section'].FillWeight = 60
$script:gridMailbox.Columns['Name'].FillWeight = 90
$script:gridMailbox.Columns['Value'].FillWeight = 220

# Tab 5 - quarantine
$tabQuarantine = New-Object System.Windows.Forms.TabPage
$tabQuarantine.Text = '  Quarantine  '
$Tabs.TabPages.Add($tabQuarantine)

New-ExoLabel -Parent $tabQuarantine -Text 'Recipient:' -X 12 -Y 16 -Width 70 | Out-Null
$script:cboQuarRecipient = New-ExoUserCombo -Parent $tabQuarantine -X 86 -Y 14 -Width 300

New-ExoLabel -Parent $tabQuarantine -Text 'Sender:' -X 400 -Y 16 -Width 60 | Out-Null
$script:cboQuarSender = New-ExoUserCombo -Parent $tabQuarantine -X 464 -Y 14 -Width 300

New-ExoLabel -Parent $tabQuarantine -Text 'Subject has:' -X 778 -Y 16 -Width 80 | Out-Null
$script:txtQuarSubject = New-Object System.Windows.Forms.TextBox
$script:txtQuarSubject.Location = New-Object System.Drawing.Point(862,14)
$script:txtQuarSubject.Size = New-Object System.Drawing.Size(240,24)
$tabQuarantine.Controls.Add($script:txtQuarSubject)

New-ExoLabel -Parent $tabQuarantine -Text 'From:' -X 12 -Y 52 -Width 60 | Out-Null
$script:dtQuarStart = New-Object System.Windows.Forms.DateTimePicker
$script:dtQuarStart.Location = New-Object System.Drawing.Point(76,50)
$script:dtQuarStart.Size = New-Object System.Drawing.Size(210,24)
$script:dtQuarStart.Format = 'Custom'
$script:dtQuarStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$script:dtQuarStart.Value = (Get-Date).AddDays(-7)
$tabQuarantine.Controls.Add($script:dtQuarStart)

New-ExoLabel -Parent $tabQuarantine -Text 'To:' -X 300 -Y 52 -Width 30 | Out-Null
$script:dtQuarEnd = New-Object System.Windows.Forms.DateTimePicker
$script:dtQuarEnd.Location = New-Object System.Drawing.Point(334,50)
$script:dtQuarEnd.Size = New-Object System.Drawing.Size(210,24)
$script:dtQuarEnd.Format = 'Custom'
$script:dtQuarEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$script:dtQuarEnd.Value = (Get-Date)
$tabQuarantine.Controls.Add($script:dtQuarEnd)

New-ExoLabel -Parent $tabQuarantine -Text 'Type:' -X 560 -Y 52 -Width 40 | Out-Null
$script:cboQuarType = New-Object System.Windows.Forms.ComboBox
$script:cboQuarType.Location = New-Object System.Drawing.Point(604,50)
$script:cboQuarType.Size = New-Object System.Drawing.Size(160,24)
$script:cboQuarType.DropDownStyle = 'DropDownList'
foreach ($Type in 'Any','Spam','HighConfPhish','Phish','Malware','Bulk','TransportRule') {
    [void]$script:cboQuarType.Items.Add($Type)
}
$script:cboQuarType.SelectedIndex = 0
$tabQuarantine.Controls.Add($script:cboQuarType)

New-ExoButton -Parent $tabQuarantine -Text 'Search'           -X 778 -Y 49 -Width 100 -NeedsConnection -OnClick { Invoke-ExoQuarantineSearch } | Out-Null
New-ExoButton -Parent $tabQuarantine -Text 'Release Selected' -X 886 -Y 49 -Width 140 -NeedsConnection -OnClick { Invoke-ExoQuarantineRelease } | Out-Null
New-ExoButton -Parent $tabQuarantine -Text 'Export CSV'       -X 1034 -Y 49 -Width 100 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridQuarantine -BaseName 'Quarantine' } | Out-Null

$script:lblQuarSummary = New-Object System.Windows.Forms.Label
$script:lblQuarSummary.Text = 'Releasing delivers the message immediately and cannot be undone.'
$script:lblQuarSummary.Location = New-Object System.Drawing.Point(12,86)
$script:lblQuarSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblQuarSummary.Anchor = 'Top,Left,Right'
$script:lblQuarSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabQuarantine.Controls.Add($script:lblQuarSummary)

$script:gridQuarantine = New-ExoGrid -Parent $tabQuarantine -Top 112 -Height ($tabQuarantine.ClientSize.Height - 126)
[void]$script:gridQuarantine.Columns.Add('Received','Received')
[void]$script:gridQuarantine.Columns.Add('Sender','Sender')
[void]$script:gridQuarantine.Columns.Add('Recipients','Recipients')
[void]$script:gridQuarantine.Columns.Add('Subject','Subject')
[void]$script:gridQuarantine.Columns.Add('Type','Type')
[void]$script:gridQuarantine.Columns.Add('Policy','Policy')
[void]$script:gridQuarantine.Columns.Add('Released','Released')
[void]$script:gridQuarantine.Columns.Add('Expires','Expires')
$script:gridQuarantine.Columns['Subject'].FillWeight = 170
$script:gridQuarantine.Columns['Released'].FillWeight = 55

# Tab 6 - header analyser
$tabHeaders = New-Object System.Windows.Forms.TabPage
$tabHeaders.Text = '  Header Analyser  '
$Tabs.TabPages.Add($tabHeaders)

$lblHeaderPaste = New-Object System.Windows.Forms.Label
$lblHeaderPaste.Text = 'Paste the internet headers (Outlook: open the message, File, Properties, Internet headers). No connection needed.'
$lblHeaderPaste.Location = New-Object System.Drawing.Point(12,12)
$lblHeaderPaste.Size = New-Object System.Drawing.Size(1000,18)
$tabHeaders.Controls.Add($lblHeaderPaste)

$script:txtHeaders = New-Object System.Windows.Forms.TextBox
$script:txtHeaders.Location = New-Object System.Drawing.Point(12,34)
$script:txtHeaders.Size = New-Object System.Drawing.Size(1000,150)
$script:txtHeaders.Multiline = $true
$script:txtHeaders.ScrollBars = 'Vertical'
$script:txtHeaders.Anchor = 'Top,Left,Right'
$script:txtHeaders.Font = New-Object System.Drawing.Font('Consolas',9)
$tabHeaders.Controls.Add($script:txtHeaders)

New-ExoButton -Parent $tabHeaders -Text 'Analyse'    -X 12 -Y 192 -Width 110 -OnClick { Invoke-ExoHeaderAnalysis } | Out-Null
New-ExoButton -Parent $tabHeaders -Text 'Paste'      -X 130 -Y 192 -Width 90  -OnClick {
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $script:txtHeaders.Text = [System.Windows.Forms.Clipboard]::GetText()
        Invoke-ExoHeaderAnalysis
    }
} | Out-Null
New-ExoButton -Parent $tabHeaders -Text 'Clear'      -X 228 -Y 192 -Width 90  -OnClick {
    $script:txtHeaders.Clear()
    $script:gridHeaders.Rows.Clear()
    $script:lblHeaderSummary.Text = ''
} | Out-Null
New-ExoButton -Parent $tabHeaders -Text 'Export CSV' -X 326 -Y 192 -Width 110 -OnClick { Export-ExoGrid -Grid $script:gridHeaders -BaseName 'MessageHeaders' } | Out-Null

$script:lblHeaderSummary = New-Object System.Windows.Forms.Label
$script:lblHeaderSummary.Location = New-Object System.Drawing.Point(452,198)
$script:lblHeaderSummary.Size = New-Object System.Drawing.Size(560,20)
$script:lblHeaderSummary.Anchor = 'Top,Left,Right'
$script:lblHeaderSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabHeaders.Controls.Add($script:lblHeaderSummary)

$script:gridHeaders = New-ExoGrid -Parent $tabHeaders -Top 228 -Height ($tabHeaders.ClientSize.Height - 242)
[void]$script:gridHeaders.Columns.Add('Section','Section')
[void]$script:gridHeaders.Columns.Add('Name','Field')
[void]$script:gridHeaders.Columns.Add('Value','Value')
$script:gridHeaders.Columns['Section'].FillWeight = 55
$script:gridHeaders.Columns['Name'].FillWeight = 70
$script:gridHeaders.Columns['Value'].FillWeight = 260

# Tab 7 - mail flow rules
$tabTransport = New-Object System.Windows.Forms.TabPage
$tabTransport.Text = '  Mail Flow Rules  '
$Tabs.TabPages.Add($tabTransport)

New-ExoLabel -Parent $tabTransport -Text 'Sender:' -X 12 -Y 16 -Width 60 | Out-Null
$script:cboRuleSender = New-ExoUserCombo -Parent $tabTransport -X 76 -Y 14 -Width 300

New-ExoLabel -Parent $tabTransport -Text 'Recipient:' -X 390 -Y 16 -Width 70 | Out-Null
$script:cboRuleRecipient = New-ExoUserCombo -Parent $tabTransport -X 464 -Y 14 -Width 300

New-ExoLabel -Parent $tabTransport -Text 'Subject has:' -X 778 -Y 16 -Width 80 | Out-Null
$script:txtRuleSubject = New-Object System.Windows.Forms.TextBox
$script:txtRuleSubject.Location = New-Object System.Drawing.Point(862,14)
$script:txtRuleSubject.Size = New-Object System.Drawing.Size(240,24)
$tabTransport.Controls.Add($script:txtRuleSubject)

New-ExoButton -Parent $tabTransport -Text 'Find Rules' -X 12 -Y 50 -Width 120 -NeedsConnection -OnClick { Invoke-ExoTransportRuleMatch } | Out-Null
New-ExoButton -Parent $tabTransport -Text 'Export CSV' -X 140 -Y 50 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridTransport -BaseName 'MailFlowRules' } | Out-Null
New-ExoButton -Parent $tabTransport -Text 'Copy'       -X 258 -Y 50 -Width 80 -NeedsConnection -OnClick { Copy-ExoGrid -Grid $script:gridTransport } | Out-Null

$script:lblTransportSummary = New-Object System.Windows.Forms.Label
$script:lblTransportSummary.Text = 'Matches on sender, recipient and subject conditions. Attachment, size and IP conditions are listed but not evaluated.'
$script:lblTransportSummary.Location = New-Object System.Drawing.Point(12,86)
$script:lblTransportSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblTransportSummary.Anchor = 'Top,Left,Right'
$script:lblTransportSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabTransport.Controls.Add($script:lblTransportSummary)

$script:gridTransport = New-ExoGrid -Parent $tabTransport -Top 112 -Height ($tabTransport.ClientSize.Height - 126)
[void]$script:gridTransport.Columns.Add('Priority','Priority')
[void]$script:gridTransport.Columns.Add('Name','Rule')
[void]$script:gridTransport.Columns.Add('State','State')
[void]$script:gridTransport.Columns.Add('Mode','Mode')
[void]$script:gridTransport.Columns.Add('Why','Why it matches')
[void]$script:gridTransport.Columns.Add('Actions','What it does')
[void]$script:gridTransport.Columns.Add('Unevaluated','Not evaluated')
$script:gridTransport.Columns['Priority'].FillWeight = 45
$script:gridTransport.Columns['State'].FillWeight = 50
$script:gridTransport.Columns['Mode'].FillWeight = 55
$script:gridTransport.Columns['Why'].FillWeight = 150
$script:gridTransport.Columns['Actions'].FillWeight = 150

# Tab 8 - folder permissions
$tabFolders = New-Object System.Windows.Forms.TabPage
$tabFolders.Text = '  Folder Permissions  '
$Tabs.TabPages.Add($tabFolders)

New-ExoLabel -Parent $tabFolders -Text 'Mailbox or room:' -X 12 -Y 16 -Width 110 | Out-Null
$script:cboFolderUser = New-ExoUserCombo -Parent $tabFolders -X 126 -Y 14 -Width 360
New-ExoButton -Parent $tabFolders -Text 'Load Permissions' -X 500 -Y 13 -Width 140 -NeedsConnection -OnClick { Invoke-ExoFolderPermissions } | Out-Null
New-ExoButton -Parent $tabFolders -Text 'Export CSV'       -X 650 -Y 13 -Width 110 -NeedsConnection -OnClick { Export-ExoGrid -Grid $script:gridFolders -BaseName 'FolderPermissions' } | Out-Null
New-ExoButton -Parent $tabFolders -Text 'Copy'             -X 768 -Y 13 -Width 80  -NeedsConnection -OnClick { Copy-ExoGrid -Grid $script:gridFolders } | Out-Null

$script:lblFolderSummary = New-Object System.Windows.Forms.Label
$script:lblFolderSummary.Location = New-Object System.Drawing.Point(12,50)
$script:lblFolderSummary.Size = New-Object System.Drawing.Size(1000,20)
$script:lblFolderSummary.Anchor = 'Top,Left,Right'
$script:lblFolderSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabFolders.Controls.Add($script:lblFolderSummary)

$lblFolderHint = New-Object System.Windows.Forms.Label
$lblFolderHint.Text = 'Covers the top of the information store, Calendar, Inbox, Contacts and Sent Items. Delegate entries are highlighted.'
$lblFolderHint.Location = New-Object System.Drawing.Point(12,72)
$lblFolderHint.Size = New-Object System.Drawing.Size(1000,18)
$lblFolderHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
$tabFolders.Controls.Add($lblFolderHint)

$script:gridFolders = New-ExoGrid -Parent $tabFolders -Top 96 -Height ($tabFolders.ClientSize.Height - 110)
[void]$script:gridFolders.Columns.Add('Folder','Folder')
[void]$script:gridFolders.Columns.Add('User','User')
[void]$script:gridFolders.Columns.Add('Rights','Access rights')
[void]$script:gridFolders.Columns.Add('Flags','Sharing flags')
$script:gridFolders.Columns['Folder'].FillWeight = 70

# The group list is loaded on demand: it is only needed on one tab, and a large
# tenant has far more groups than mailboxes.
function Initialize-ExoGroupPicker {
    if (-not $script:ExchangeConnected) { return }
    if ($script:cboGroup.Items.Count -gt 0) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Loading distribution groups...'
        $Groups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop | Sort-Object DisplayName)

        foreach ($Group in $Groups) {
            [void]$script:cboGroup.Items.Add("$($Group.DisplayName) ($($Group.PrimarySmtpAddress))")
        }

        Write-AppLog "Loaded $($Groups.Count) distribution group(s)."
        Set-BusyState -Busy $false -StatusText "Loaded $($Groups.Count) distribution group(s)."
    }
    catch {
        Set-BusyState -Busy $false
        Write-AppLog "Could not load the group list: $($_.Exception.Message). Type a group address instead." 'WARN'
    }
}

$Tabs.Add_SelectedIndexChanged({
    if ($Tabs.SelectedTab -eq $tabGroup) { Initialize-ExoGroupPicker }
})

# --- theme ----------------------------------------------------------------
# Re-applying the theme repaints the connection label, so restore it from the
# live connection state rather than leaving it red after a toggle.
$script:MSToolkitThemeRefreshHook = {
    if (-not $script:lblConnection) { return }

    if ($script:ExchangeConnected) {
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Success
    }
    else {
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Danger
    }
}

New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "M365ExchangeOnlineTools"
Apply-MSToolkitSharedTheme -Root $MainForm

$script:AllGrids = @(
    $script:gridRules, $script:gridTrace, $script:gridGroup, $script:gridMailbox,
    $script:gridQuarantine, $script:gridHeaders,
    $script:gridTransport, $script:gridFolders
)

function Resize-ExoGrids {
    foreach ($Grid in $script:AllGrids) {
        if (-not $Grid -or -not $Grid.Parent) { continue }

        $Width  = $Grid.Parent.ClientSize.Width - 24
        $Height = $Grid.Parent.ClientSize.Height - $Grid.Top - 14

        if ($Width -gt 100 -and $Height -gt 80) {
            $Grid.Size = New-Object System.Drawing.Size($Width,$Height)
        }
    }
}

$MainForm.Add_Shown({
    $MainForm.Activate()

    # A TabPage reports no usable size until the window is shown, so the grids are
    # sized here rather than at creation. Anchoring handles every later resize.
    Resize-ExoGrids
    Write-AppLog 'Exchange Online Tools ready. Connect with a standard account holding an Exchange admin role.'
    Write-AppLog 'Read-only apart from releasing quarantined mail, which confirms first.'
})

Hide-PowerShellConsole
# Type-ahead filtering on every editable dropdown. Its own handler, not folded
# into another, so it runs on every launch rather than only when the tool is
# started with parameters - and so a failure here cannot stop the rest of
# start-up. Multiple Shown handlers chain.
$MainForm.Add_Shown({ Register-MSToolkitComboFiltersOn -Root $MainForm })

[void]$MainForm.ShowDialog()

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
