#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet("Light","Dark")]
    [string]$ThemeMode,

    # Set by the Offboard User tool: connect on startup, open the Manage User
    # Groups window, and load this user straight away.
    [string]$ManageUser,

    [switch]$AutoConnect
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
            $ToolProperty = "Theme_M365GroupCompare"
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
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace MSToolkitAuthWindow
{
    public static class AuthWindowWatcher
    {
        private const uint MONITOR_DEFAULTTONEAREST = 2;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

        private static Timer _timer;
        private static IntPtr _anchor = IntPtr.Zero;
        private static DateTime _expiresUtc;

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

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

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

        public static void Start(IntPtr anchorWindow)
        {
            Stop();

            _anchor = anchorWindow;
            _expiresUtc = DateTime.UtcNow.AddSeconds(90);
            _timer = new Timer(CheckWindows, null, 0, 250);
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
        }

        private static bool LooksLikeMicrosoftSignIn(string title, string className)
        {
            string value = ((title ?? "") + " " + (className ?? "")).ToLowerInvariant();

            return
                value.Contains("sign in") ||
                value.Contains("signin") ||
                value.Contains("pick an account") ||
                value.Contains("choose an account") ||
                value.Contains("work or school") ||
                value.Contains("microsoft account") ||
                value.Contains("credential dialog") ||
                value.Contains("web account");
        }

        private static void CheckWindows(object state)
        {
            if (_anchor == IntPtr.Zero || DateTime.UtcNow >= _expiresUtc)
            {
                Stop();
                return;
            }

            try
            {
                IntPtr monitor = MonitorFromWindow(_anchor, MONITOR_DEFAULTTONEAREST);
                if (monitor == IntPtr.Zero)
                    return;

                MONITORINFO mi = new MONITORINFO();
                mi.cbSize = Marshal.SizeOf(typeof(MONITORINFO));

                if (!GetMonitorInfo(monitor, ref mi))
                    return;

                EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
                {
                    if (hWnd == _anchor || !IsWindowVisible(hWnd))
                        return true;

                    StringBuilder titleBuilder = new StringBuilder(512);
                    GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
                    string title = titleBuilder.ToString();

                    StringBuilder classBuilder = new StringBuilder(256);
                    GetClassName(hWnd, classBuilder, classBuilder.Capacity);
                    string className = classBuilder.ToString();

                    if (!LooksLikeMicrosoftSignIn(title, className))
                        return true;

                    RECT rect;
                    if (!GetWindowRect(hWnd, out rect))
                        return true;

                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;

                    if (width < 200 || height < 120)
                        return true;

                    int workWidth = mi.rcWork.Right - mi.rcWork.Left;
                    int workHeight = mi.rcWork.Bottom - mi.rcWork.Top;

                    int x = mi.rcWork.Left + Math.Max(0, (workWidth - width) / 2);
                    int y = mi.rcWork.Top + Math.Max(0, (workHeight - height) / 2);

                    // Move the authentication window to the monitor containing the
                    // M365 tool, briefly raise it, then return it to normal topmost state.
                    ShowWindow(hWnd, 5);
                    SetWindowPos(hWnd, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE);
                    SetWindowPos(hWnd, HWND_NOTOPMOST, x, y, 0, 0, SWP_NOSIZE);
                    SetForegroundWindow(hWnd);

                    return true;
                }, IntPtr.Zero);
            }
            catch
            {
                // Never allow focus assistance to interfere with authentication.
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
$script:GraphConnected = $false

function Export-MSToolkitGridToCsv {
    # Writes whatever is currently in a grid to CSV, using the grid's own column
    # headers. The checkbox column is skipped - it is a selection, not data.
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$BaseName,
        [string]$Subject
    )

    if (-not $Grid -or $Grid.Rows.Count -eq 0) {
        Show-InfoMessage 'There is nothing to export yet. Load the groups first.'
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
            if ($Subject) { $Item['User'] = $Subject }

            foreach ($Column in $Grid.Columns) {
                if ($Column -is [System.Windows.Forms.DataGridViewCheckBoxColumn]) { continue }
                $Item[$Column.HeaderText] = [string]$Row.Cells[$Column.Index].Value
            }

            $Item['ExportedOn'] = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            $Item['ExportedBy'] = "$env:USERDOMAIN\$env:USERNAME"

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

function Show-ErrorMessage {
    param([string]$Message, [string]$Title = 'Error')
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Show-InfoMessage {
    param([string]$Message, [string]$Title = 'Information')
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Write-AppLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARN','ERROR')]
        [string]$Level = 'INFO'
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
    param(
        [bool]$Busy,
        [string]$StatusText = ''
    )

    $script:MainForm.UseWaitCursor = $Busy

    # A ComboBox owns its own window handle, so clearing UseWaitCursor on the form
    # does not always restore the pointer over it. Reset the cursor explicitly.
    if (-not $Busy) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    $script:btnConnect.Enabled = -not $Busy
    $script:btnCompare.Enabled = (-not $Busy -and $script:GraphConnected)
    $script:btnAddSelected.Enabled = (-not $Busy -and $script:ComparisonRows.Count -gt 0)
    $script:btnSelectAll.Enabled = (-not $Busy -and $script:ComparisonRows.Count -gt 0)
    $script:btnClearSelection.Enabled = (-not $Busy -and $script:ComparisonRows.Count -gt 0)

    if ($StatusText) {
        $script:lblStatus.Text = $StatusText
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Ensure-GraphAuthenticationModule {
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        return $true
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        "The Microsoft.Graph.Authentication PowerShell module is required and is not installed.`r`n`r`nInstall it for the current user now?",
        'Microsoft Graph Module Required',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        return $false
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Installing Microsoft Graph authentication module...'
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Write-AppLog 'Installed Microsoft.Graph.Authentication successfully.' 'SUCCESS'
        return $true
    }
    catch {
        Show-ErrorMessage "Unable to install Microsoft.Graph.Authentication.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Graph module installation failed: $($_.Exception.Message)" 'ERROR'
        return $false
    }
    finally {
        Set-BusyState -Busy $false
    }
}


function Connect-M365Graph {
    if (-not (Ensure-GraphAuthenticationModule)) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Connecting to Microsoft 365...'
        Write-AppLog 'Starting Microsoft Graph sign-in.'

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

        # These delegated scopes support user lookup, group-property reads,
        # direct membership comparison, and membership changes for standard groups.
        $scopes = @(
            'User.Read.All',
            'Group.Read.All',
            'GroupMember.ReadWrite.All'
        )

        Start-M365AuthWindowWatcher
        Connect-MgGraph -Scopes $scopes -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
        $context = Get-MgContext

        if (-not $context -or -not $context.Account) {
            throw 'Microsoft Graph authentication completed without a usable signed-in context.'
        }

        $script:GraphConnected = $true
        $script:lblConnection.Text = "Connected: $($context.Account)"
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Success
        $script:btnCompare.Enabled = $true
        $script:btnManageGroups.Enabled = $true
        $script:lblStatus.Text = 'Connected to Microsoft 365. Enter two users and click Compare.'
        Write-AppLog "Connected to Microsoft Graph as $($context.Account)." 'SUCCESS'

        $script:MSToolkitUserPickerCache = $null
        $LoadedCount = Initialize-MSToolkitUserPicker -Combos @($script:txtReference, $script:txtTarget)

        if ($LoadedCount -ge 0) {
            Write-AppLog "Loaded $LoadedCount Microsoft 365 user(s) into the user lists."
        }
        else {
            Write-AppLog 'Could not load the Microsoft 365 user list. Type a user instead.' 'WARNING'
        }
    }
    catch {
        $script:GraphConnected = $false
        $script:lblConnection.Text = 'Not connected'
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Danger
        Show-ErrorMessage "Unable to connect to Microsoft 365.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Microsoft Graph connection failed: $($_.Exception.Message)" 'ERROR'
    }
    finally {
        Stop-M365AuthWindowWatcher
        Set-BusyState -Busy $false
    }
}

function Invoke-GraphGetAll {
    param([Parameter(Mandatory)][string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri

    while ($next) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop

        if ($null -ne $response.value) {
            foreach ($item in $response.value) {
                $items.Add($item)
            }
        }
        elseif ($response) {
            $items.Add($response)
        }

        $next = $response.'@odata.nextLink'
    }

    return $items.ToArray()
}

function Escape-ODataString {
    param([string]$Value)
    return $Value.Replace("'", "''")
}

function Show-UserSelectionDialog {
    param(
        [Parameter(Mandatory)][array]$Users,
        [Parameter(Mandatory)][string]$SearchText
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Select User - $SearchText"
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object System.Drawing.Size(760, 430)
    $dialog.MinimumSize = New-Object System.Drawing.Size(650, 350)
    $dialog.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $dialog.BackColor = [System.Drawing.Color]::White

    $label = New-Object System.Windows.Forms.Label
    $label.Text = 'Multiple users matched. Select the correct user:'
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(15, 15)
    $dialog.Controls.Add($label)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(15, 45)
    $grid.Size = New-Object System.Drawing.Size(715, 285)
    $grid.Anchor = 'Top,Bottom,Left,Right'
    $grid.ReadOnly = $true
    $grid.MultiSelect = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowHeadersVisible = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = 'Fixed3D'

    [void]$grid.Columns.Add('DisplayName', 'Display Name')
    [void]$grid.Columns.Add('UserPrincipalName', 'User Principal Name')
    [void]$grid.Columns.Add('Mail', 'Mail')

    foreach ($user in $Users) {
        $index = $grid.Rows.Add($user.displayName, $user.userPrincipalName, $user.mail)
        $grid.Rows[$index].Tag = $user
    }

    $dialog.Controls.Add($grid)

    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Text = 'Select User'
    $btnSelect.Size = New-Object System.Drawing.Size(110, 32)
    $btnSelect.Anchor = 'Bottom,Right'
    $btnSelect.Location = New-Object System.Drawing.Point(500, 345)
    $btnSelect.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnSelect.ForeColor = [System.Drawing.Color]::White
    $btnSelect.FlatStyle = 'Flat'
    $btnSelect.FlatAppearance.BorderSize = 0
    $dialog.Controls.Add($btnSelect)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size(110, 32)
    $btnCancel.Anchor = 'Bottom,Right'
    $btnCancel.Location = New-Object System.Drawing.Point(620, 345)
    $dialog.Controls.Add($btnCancel)

    $selected = $null

    $selectAction = {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $script:DialogSelectedUser = $grid.SelectedRows[0].Tag
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    }

    $btnSelect.Add_Click($selectAction)
    $grid.Add_CellDoubleClick($selectAction)
    $btnCancel.Add_Click({ $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dialog.Close() })

    if ($grid.Rows.Count -gt 0) {
        $grid.Rows[0].Selected = $true
    }

    $script:DialogSelectedUser = $null
    Apply-MSToolkitSharedTheme -Root $dialog
    Register-MSToolkitComboFiltersOn -Root $dialog
    $result = $dialog.ShowDialog($script:MainForm)

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selected = $script:DialogSelectedUser
    }

    Remove-Variable DialogSelectedUser -Scope Script -ErrorAction SilentlyContinue
    $dialog.Dispose()
    return $selected
}

function Get-MSToolkitUserPickerLabel {
    param($GraphUser)

    $Display = if (-not [string]::IsNullOrWhiteSpace($GraphUser.displayName)) { $GraphUser.displayName } else { $GraphUser.userPrincipalName }
    return "$Display ($($GraphUser.userPrincipalName))"
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

function Get-MSToolkitGraphUserList {
    $Users = New-Object System.Collections.Generic.List[object]
    $Uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled&$top=999'

    while (-not [string]::IsNullOrWhiteSpace($Uri)) {
        $Response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop

        foreach ($Entry in $Response.value) {
            if (-not [string]::IsNullOrWhiteSpace($Entry.userPrincipalName)) {
                $Users.Add($Entry)
            }
        }

        $Uri = $Response.'@odata.nextLink'
    }

    # Invoke-MgGraphRequest returns hashtables, so sort on the key value rather
    # than a property name, which would silently leave the list unsorted.
    return @($Users.ToArray() | Sort-Object { "$($_.displayName)" })
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
    param([System.Windows.Forms.ComboBox[]]$Combos)

    try {
        if (-not $script:MSToolkitUserPickerCache) {
            $script:MSToolkitUserPickerCache = Get-MSToolkitGraphUserList
        }

        foreach ($Combo in $Combos) {
            if (-not $Combo) { continue }

            $Existing = $Combo.Text
            $Combo.Items.Clear()

            foreach ($User in $script:MSToolkitUserPickerCache) {
                $null = $Combo.Items.Add((Get-MSToolkitUserPickerLabel -GraphUser $User))
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

function Resolve-GraphUser {
    param([Parameter(Mandatory)][string]$Identity)

    $Identity = $Identity.Trim()
    if ([string]::IsNullOrWhiteSpace($Identity)) {
        throw 'A user name or UPN was not entered.'
    }

    $select = 'id,displayName,userPrincipalName,mail,accountEnabled'

    # First try the value directly as a user ID / UPN.
    try {
        $encoded = [uri]::EscapeDataString($Identity)
        $exact = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($encoded)?`$select=$select" -ErrorAction Stop
        if ($exact -and $exact.id) {
            return $exact
        }
    }
    catch {
        # Continue to display-name / UPN prefix search.
    }

    $escaped = Escape-ODataString $Identity
    $filter = "startswith(displayName,'$escaped') or startswith(userPrincipalName,'$escaped')"
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$([uri]::EscapeDataString($filter))&`$select=$select&`$top=25"
    $matches = @(Invoke-GraphGetAll -Uri $uri)

    if ($matches.Count -eq 0) {
        throw "No Microsoft 365 user was found matching '$Identity'."
    }

    if ($matches.Count -eq 1) {
        return $matches[0]
    }

    $selected = Show-UserSelectionDialog -Users $matches -SearchText $Identity
    if (-not $selected) {
        throw 'User selection was canceled.'
    }

    return $selected
}

function Get-StaticSecurityGroupsForUser {
    param([Parameter(Mandatory)][string]$UserId)

    $select = 'id,displayName,description,groupTypes,mailEnabled,securityEnabled,isAssignableToRole,onPremisesSyncEnabled'
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/memberOf/microsoft.graph.group?`$select=$select&`$top=999"
    $groups = @(Invoke-GraphGetAll -Uri $uri)

    # User requirement: security groups only; exclude distribution/mail-enabled
    # groups and dynamic-membership groups. Direct memberships only.
    $filtered = foreach ($group in $groups) {
        $groupTypes = @($group.groupTypes)
        $isDynamic = $groupTypes -contains 'DynamicMembership'

        if (($group.securityEnabled -eq $true) -and
            ($group.mailEnabled -eq $false) -and
            (-not $isDynamic)) {
            $group
        }
    }

    # Graph returns hashtables, so sort on the key value rather than a property name.
    return @($filtered | Sort-Object { "$($_.displayName)" })
}

function Update-ComparisonGrid {
    $script:dgvGroups.Rows.Clear()

    foreach ($row in $script:ComparisonRows) {
        $syncText = if ($row.OnPremisesSyncEnabled -eq $true) {
            'On-Prem Synced'
        }
        elseif ($row.OnPremisesSyncEnabled -eq $false) {
            'Formerly Synced'
        }
        else {
            'Cloud'
        }

        $typeText = if ($row.IsAssignableToRole) {
            "Role-Assignable Security - $syncText"
        }
        else {
            "Security - $syncText"
        }
        $index = $script:dgvGroups.Rows.Add(
            $false,
            $row.DisplayName,
            $row.Description,
            $typeText,
            'Missing'
        )
        $script:dgvGroups.Rows[$index].Tag = $row

        if ($row.OnPremisesSyncEnabled -eq $true) {
            $script:dgvGroups.Rows[$index].Cells['Selected'].ReadOnly = $true
            $script:dgvGroups.Rows[$index].Cells['Selected'].Value = $false
            $script:dgvGroups.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).DisabledBackground
            $script:dgvGroups.Rows[$index].DefaultCellStyle.ForeColor = (Get-MSToolkitThemePalette).DisabledText
            $script:dgvGroups.Rows[$index].Cells['MembershipStatus'].Value = 'On-Prem Managed'
        }
        elseif ($row.IsAssignableToRole) {
            $script:dgvGroups.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).WarningBackground
        }
    }

    $script:lblMissingCount.Text = "Missing groups: $($script:ComparisonRows.Count)"
    $script:btnAddSelected.Enabled = $script:ComparisonRows.Count -gt 0
    $script:btnSelectAll.Enabled = $script:ComparisonRows.Count -gt 0
    $script:btnClearSelection.Enabled = $script:ComparisonRows.Count -gt 0
}

function Compare-Users {
    if (-not $script:GraphConnected) {
        Show-InfoMessage 'Connect to Microsoft 365 first.'
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:txtReference.Text) -or
        [string]::IsNullOrWhiteSpace($script:txtTarget.Text)) {
        Show-InfoMessage 'Enter both a Reference User and a Target User.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Resolving users and comparing security groups...'
        $script:dgvGroups.Rows.Clear()
        $script:ComparisonRows = @()
        $script:lblReferenceResolved.Text = ''
        $script:lblTargetResolved.Text = ''
        $script:lblMissingCount.Text = 'Missing groups: 0'

        Write-AppLog "Resolving reference user: $($script:txtReference.Text)"
        $script:ReferenceUser = Resolve-GraphUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtReference.Text)
        $script:lblReferenceResolved.Text = "$($script:ReferenceUser.displayName)  |  $($script:ReferenceUser.userPrincipalName)"

        Write-AppLog "Resolving target user: $($script:txtTarget.Text)"
        $script:TargetUser = Resolve-GraphUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtTarget.Text)
        $script:lblTargetResolved.Text = "$($script:TargetUser.displayName)  |  $($script:TargetUser.userPrincipalName)"

        if ($script:ReferenceUser.id -eq $script:TargetUser.id) {
            throw 'The Reference User and Target User resolve to the same Microsoft 365 account.'
        }

        Write-AppLog "Reading direct static security-group memberships for $($script:ReferenceUser.userPrincipalName)."
        $referenceGroups = @(Get-StaticSecurityGroupsForUser -UserId $script:ReferenceUser.id)

        Write-AppLog "Reading direct static security-group memberships for $($script:TargetUser.userPrincipalName)."
        $targetGroups = @(Get-StaticSecurityGroupsForUser -UserId $script:TargetUser.id)

        $targetIds = @{}
        foreach ($group in $targetGroups) {
            $targetIds[$group.id] = $true
        }

        $missing = foreach ($group in $referenceGroups) {
            if (-not $targetIds.ContainsKey($group.id)) {
                [pscustomobject]@{
                    Id                    = $group.id
                    DisplayName           = $group.displayName
                    Description           = $group.description
                    IsAssignableToRole    = [bool]$group.isAssignableToRole
                    OnPremisesSyncEnabled = $group.onPremisesSyncEnabled
                }
            }
        }

        $script:ComparisonRows = @($missing | Sort-Object DisplayName)
        Update-ComparisonGrid

        $script:lblReferenceCount.Text = "Reference security groups: $($referenceGroups.Count)"
        $script:lblTargetCount.Text = "Target security groups: $($targetGroups.Count)"
        $script:lblStatus.Text = "Comparison complete. $($script:ComparisonRows.Count) missing group(s) found."

        Write-AppLog "Comparison complete: reference=$($referenceGroups.Count), target=$($targetGroups.Count), missing=$($script:ComparisonRows.Count)." 'SUCCESS'

        if ($script:ComparisonRows.Count -eq 0) {
            Show-InfoMessage "$($script:TargetUser.displayName) already has all eligible direct static security-group memberships that $($script:ReferenceUser.displayName) has." 'No Missing Groups'
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

function Get-SelectedComparisonGroups {
    $selected = New-Object System.Collections.Generic.List[object]

    # Commit any checkbox edit that is still active before reading values.
    if ($script:dgvGroups.IsCurrentCellDirty) {
        $script:dgvGroups.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null
    }
    $script:dgvGroups.EndEdit() | Out-Null

    foreach ($gridRow in $script:dgvGroups.Rows) {
        $cell = $gridRow.Cells['Selected']
        $isChecked = $false

        if ($null -ne $cell.EditedFormattedValue) {
            $isChecked = [System.Convert]::ToBoolean($cell.EditedFormattedValue)
        }
        elseif ($null -ne $cell.Value) {
            $isChecked = [System.Convert]::ToBoolean($cell.Value)
        }

        if ($isChecked) {
            $selected.Add($gridRow.Tag)
        }
    }

    return $selected.ToArray()
}

function Add-UserToSelectedGroups {
    if (-not $script:TargetUser) {
        Show-InfoMessage 'Run a comparison first.'
        return
    }

    $selected = @(Get-SelectedComparisonGroups)
    if ($selected.Count -eq 0) {
        Show-InfoMessage 'Select at least one group to add.'
        return
    }

    $onPremSelected = @($selected | Where-Object { $_.OnPremisesSyncEnabled -eq $true })
    if ($onPremSelected.Count -gt 0) {
        $blockedNames = ($onPremSelected.DisplayName | Sort-Object) -join "`r`n - "
        Show-InfoMessage "The following selected group(s) are synchronized from on-premises Active Directory and cannot be modified through Microsoft Graph.`r`n`r`nManage these memberships in Active Directory instead:`r`n`r`n - $blockedNames" 'On-Premises Managed Groups'
        $selected = @($selected | Where-Object { $_.OnPremisesSyncEnabled -ne $true })
    }

    if ($selected.Count -eq 0) {
        return
    }

    $roleAssignableCount = @($selected | Where-Object { $_.IsAssignableToRole }).Count
    $warning = ''
    if ($roleAssignableCount -gt 0) {
        $warning = "`r`n`r`nWARNING: $roleAssignableCount selected group(s) are role-assignable security groups. Microsoft Entra requires additional privileged permissions to modify those memberships. Standard groups can still be processed if a privileged group fails."
    }

    $message = "Add $($script:TargetUser.displayName) ($($script:TargetUser.userPrincipalName)) to $($selected.Count) selected group(s)?$warning"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $message,
        'Confirm Group Membership Changes',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $success = 0
    $failed = 0

    try {
        Set-BusyState -Busy $true -StatusText 'Adding target user to selected groups...'

        foreach ($group in $selected) {
            try {
                $body = @{
                    '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($script:TargetUser.id)"
                } | ConvertTo-Json -Compress

                $uri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/members/`$ref"
                Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                $success++
                Write-AppLog "Added $($script:TargetUser.userPrincipalName) to '$($group.DisplayName)'." 'SUCCESS'
            }
            catch {
                $failed++
                Write-AppLog "Failed to add $($script:TargetUser.userPrincipalName) to '$($group.DisplayName)': $($_.Exception.Message)" 'ERROR'
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        $script:lblStatus.Text = "Group updates complete. Success: $success | Failed: $failed"

        if ($failed -eq 0) {
            Show-InfoMessage "Completed successfully.`r`n`r`nAdded: $success`r`nFailed: 0" 'Group Membership Update'
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Group updates completed with one or more failures.`r`n`r`nAdded: $success`r`nFailed: $failed`r`n`r`nReview the activity log for details.",
                'Group Membership Update',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }

        # Refresh comparison so successfully-added groups disappear from the missing list.
        Compare-Users
    }
    finally {
        Set-BusyState -Busy $false
    }
}


function Show-SecurityGroupManager {
    param([string]$PrefillUser)

    if (-not $script:GraphConnected) {
        Show-InfoMessage 'Connect to Microsoft 365 first.'
        return
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Manage User Security Groups'
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object System.Drawing.Size(980, 650)
    $dialog.MinimumSize = New-Object System.Drawing.Size(900, 575)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $dialog.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Top'
    $header.Height = 68
    $header.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $dialog.Controls.Add($header)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Manage User Security Groups'
    $title.AutoSize = $true
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold',17)
    $title.Location = New-Object System.Drawing.Point(18,10)
    $header.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = 'Load one user, select direct static security-group memberships, and remove selected memberships'
    $subtitle.AutoSize = $true
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
    $subtitle.Location = New-Object System.Drawing.Point(20,40)
    $header.Controls.Add($subtitle)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'User (UPN or name)'
    $lblUser.AutoSize = $true
    $lblUser.Location = New-Object System.Drawing.Point(20,88)
    $dialog.Controls.Add($lblUser)

    $txtUser = New-Object System.Windows.Forms.ComboBox
    $txtUser.Location = New-Object System.Drawing.Point(20,110)
    $txtUser.Size = New-Object System.Drawing.Size(560,24)
    $txtUser.DropDownStyle = 'DropDown'
    $txtUser.AutoCompleteMode = 'None'
    $txtUser.AutoCompleteSource = 'None'
    $txtUser.MaxDropDownItems = 20
    $null = Initialize-MSToolkitUserPicker -Combos @($txtUser)
    $dialog.Controls.Add($txtUser)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load Groups'
    $btnLoad.Location = New-Object System.Drawing.Point(595,106)
    $btnLoad.Size = New-Object System.Drawing.Size(120,32)
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
    $grid.Size = New-Object System.Drawing.Size(925,355)
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
    $colSelect.HeaderText = 'Remove'
    $colSelect.Width = 60
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
    $colType.Width = 220
    $colType.ReadOnly = $true
    [void]$grid.Columns.Add($colType)

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = 'Select All'
    $btnSelectAll.Location = New-Object System.Drawing.Point(20,540)
    $btnSelectAll.Size = New-Object System.Drawing.Size(90,30)
    $btnSelectAll.Anchor = 'Bottom,Left'
    $dialog.Controls.Add($btnSelectAll)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = 'Clear'
    $btnClear.Location = New-Object System.Drawing.Point(120,540)
    $btnClear.Size = New-Object System.Drawing.Size(80,30)
    $btnClear.Anchor = 'Bottom,Left'
    $dialog.Controls.Add($btnClear)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Export Security Groups'
    $btnExport.Location = New-Object System.Drawing.Point(210,540)
    $btnExport.Size = New-Object System.Drawing.Size(170,30)
    $btnExport.Anchor = 'Bottom,Left'
    $btnExport.Add_Click({
        $subject = ''
        if ($script:ManageSecurityUser) {
            $subject = "$($script:ManageSecurityUser.displayName) <$($script:ManageSecurityUser.userPrincipalName)>"
        }

        $baseName = 'SecurityGroups'
        if ($script:ManageSecurityUser) {
            $baseName = "SecurityGroups-$($script:ManageSecurityUser.userPrincipalName -replace '@.*$','')"
        }

        Export-MSToolkitGridToCsv -Grid $grid -BaseName $baseName -Subject $subject
    })
    $dialog.Controls.Add($btnExport)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'Remove Selected Groups'
    $btnRemove.Location = New-Object System.Drawing.Point(725,540)
    $btnRemove.Size = New-Object System.Drawing.Size(220,34)
    $btnRemove.Anchor = 'Bottom,Right'
    $btnRemove.BackColor = [System.Drawing.Color]::Firebrick
    $btnRemove.ForeColor = [System.Drawing.Color]::White
    $btnRemove.FlatStyle = 'Flat'
    $btnRemove.Enabled = $false
    $dialog.Controls.Add($btnRemove)

    $script:ManageSecurityUser = $null

    $loadGroups = {
        if ([string]::IsNullOrWhiteSpace($txtUser.Text)) {
            Show-InfoMessage 'Enter a user first.'
            return
        }

        try {
            $dialog.UseWaitCursor = $true
            $grid.Rows.Clear()
            $btnRemove.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()

            $script:ManageSecurityUser = Resolve-GraphUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $txtUser.Text)
            $lblResolved.Text = "$($script:ManageSecurityUser.displayName)  |  $($script:ManageSecurityUser.userPrincipalName)"

            $groups = @(Get-StaticSecurityGroupsForUser -UserId $script:ManageSecurityUser.id)

            foreach ($groupItem in $groups) {
                $syncText = if ($groupItem.onPremisesSyncEnabled -eq $true) {
                    'On-Prem Synced'
                }
                elseif ($groupItem.onPremisesSyncEnabled -eq $false) {
                    'Formerly Synced'
                }
                else {
                    'Cloud'
                }

                $typeText = if ([bool]$groupItem.isAssignableToRole) {
                    "Role-Assignable Security - $syncText"
                }
                else {
                    "Security - $syncText"
                }

                $index = $grid.Rows.Add(
                    $false,
                    $groupItem.displayName,
                    $groupItem.description,
                    $typeText
                )
                $grid.Rows[$index].Tag = $groupItem

                if ($groupItem.onPremisesSyncEnabled -eq $true) {
                    $grid.Rows[$index].Cells['Selected'].ReadOnly = $true
                    $grid.Rows[$index].Cells['Selected'].Value = $false
                    $grid.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).DisabledBackground
                    $grid.Rows[$index].DefaultCellStyle.ForeColor = (Get-MSToolkitThemePalette).DisabledText
                }
                elseif ([bool]$groupItem.isAssignableToRole) {
                    $grid.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).WarningBackground
                }
            }

            $btnRemove.Enabled = ($grid.Rows.Count -gt 0)
        }
        catch {
            Show-ErrorMessage "Unable to load security groups.`r`n`r`n$($_.Exception.Message)"
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
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
        }
    })

    $btnSelectAll.Add_Click({
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow -and -not $row.Cells['Selected'].ReadOnly) {
                $row.Cells['Selected'].Value = $true
            }
        }
    })

    $btnClear.Add_Click({
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow) { $row.Cells['Selected'].Value = $false }
        }
    })

    $btnRemove.Add_Click({
        if (-not $script:ManageSecurityUser) {
            Show-InfoMessage 'Load a user first.'
            return
        }

        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            $grid.EndEdit()
        }

        $selected = New-Object System.Collections.Generic.List[object]
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow -and [bool]$row.Cells['Selected'].Value -and $row.Tag) {
                $selected.Add($row.Tag)
            }
        }

        $selectedGroups = $selected.ToArray()
        if ($selectedGroups.Count -eq 0) {
            Show-InfoMessage 'Select at least one security group to remove.'
            return
        }

        $onPremSelected = @($selectedGroups | Where-Object { $_.onPremisesSyncEnabled -eq $true })
        if ($onPremSelected.Count -gt 0) {
            $blockedNames = ($onPremSelected.displayName | Sort-Object) -join "`r`n - "
            Show-InfoMessage "The following selected group(s) are synchronized from on-premises Active Directory and cannot be modified through Microsoft Graph.`r`n`r`nManage these memberships in Active Directory instead:`r`n`r`n - $blockedNames" 'On-Premises Managed Groups'
            $selectedGroups = @($selectedGroups | Where-Object { $_.onPremisesSyncEnabled -ne $true })
        }

        if ($selectedGroups.Count -eq 0) {
            return
        }

        $names = @($selectedGroups | Sort-Object { "$($_.displayName)" } | ForEach-Object { $_.displayName })
        $nameText = ($names -join "`r`n - ")

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($script:ManageSecurityUser.displayName) from the following $($selectedGroups.Count) security group(s)?`r`n`r`n - $nameText",
            'Confirm Security Group Removal',
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
                    $uri = "https://graph.microsoft.com/v1.0/groups/$($groupItem.id)/members/$($script:ManageSecurityUser.id)/`$ref"
                    Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop | Out-Null
                    Write-AppLog "Removed $($script:ManageSecurityUser.userPrincipalName) from '$($groupItem.displayName)'." 'SUCCESS'
                    $success++
                }
                catch {
                    Write-AppLog "Failed to remove $($script:ManageSecurityUser.userPrincipalName) from '$($groupItem.displayName)': $($_.Exception.Message)" 'ERROR'
                    $failed++
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Security group removal complete.`r`n`r`nSuccessful: $success`r`nFailed: $failed`r`n`r`nReview the main Activity Log for details.",
                'Removal Complete',
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

        # Launched from Offboard User: the departing user is already known, so
        # fill it in and load the groups rather than making it be typed again.
        if ($PrefillUser) {
            $txtUser.Text = $PrefillUser
            $btnLoad.PerformClick()
        }
        else {
            $txtUser.Focus()
        }
    })

    Apply-MSToolkitSharedTheme -Root $dialog
    Register-MSToolkitComboFiltersOn -Root $dialog
    [void]$dialog.ShowDialog($script:MainForm)
    Remove-Variable ManageSecurityUser -Scope Script -ErrorAction SilentlyContinue
}

# ---------------------------
# Main Windows Forms interface
# ---------------------------
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'Microsoft 365 Group Compare'
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
$lblTitle.Text = 'Microsoft 365 Group Compare'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20, 12)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Compare direct static security-group memberships and add missing access to a target user'
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(218, 228, 240)
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$lblSubtitle.AutoSize = $true
$lblSubtitle.Location = New-Object System.Drawing.Point(23, 49)
$pnlHeader.Controls.Add($lblSubtitle)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Connect to Microsoft 365'
$btnConnect.Size = New-Object System.Drawing.Size(190, 36)
$btnConnect.Anchor = 'Top,Right'
$btnConnect.Location = New-Object System.Drawing.Point(955, 20)
$btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.FlatStyle = 'Flat'
$btnConnect.FlatAppearance.BorderSize = 0
$btnConnect.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$pnlHeader.Controls.Add($btnConnect)
$script:MSToolkitThemeToggleButton = New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "M365GroupCompare"

# Keep the connection status colour correct after a theme switch. Its design-time
# colour is Firebrick, so the shared theme pass would otherwise always restore red.
$script:MSToolkitThemeRefreshHook = {
    if ($script:lblConnection) {
        if ($script:GraphConnected) {
            $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Success
        }
        else {
            $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Danger
        }
    }
}
$script:btnConnect = $btnConnect

# Connection bar
$pnlConnection = New-Object System.Windows.Forms.Panel
$pnlConnection.Dock = 'Top'
$pnlConnection.Height = 38
$pnlConnection.BackColor = [System.Drawing.Color]::White
$pnlConnection.Padding = New-Object System.Windows.Forms.Padding(20, 0, 20, 0)
$MainForm.Controls.Add($pnlConnection)
$pnlConnection.BringToFront()

$lblConnectionLabel = New-Object System.Windows.Forms.Label
$lblConnectionLabel.Text = 'Microsoft Graph:'
$lblConnectionLabel.AutoSize = $true
$lblConnectionLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$lblConnectionLabel.Location = New-Object System.Drawing.Point(20, 10)
$pnlConnection.Controls.Add($lblConnectionLabel)

$lblConnection = New-Object System.Windows.Forms.Label
$lblConnection.Text = 'Not connected'
$lblConnection.AutoSize = $true
$lblConnection.ForeColor = [System.Drawing.Color]::Firebrick
$lblConnection.Location = New-Object System.Drawing.Point(122, 10)
$pnlConnection.Controls.Add($lblConnection)
$script:lblConnection = $lblConnection


$btnManageGroups = New-Object System.Windows.Forms.Button
$btnManageGroups.Text = 'Manage User Groups'
$btnManageGroups.Size = New-Object System.Drawing.Size(155,28)
$btnManageGroups.Anchor = 'Top,Right'
$btnManageGroups.Location = New-Object System.Drawing.Point(970,5)
$btnManageGroups.Enabled = $false
$btnManageGroups.Add_Click({ Show-SecurityGroupManager })
$pnlConnection.Controls.Add($btnManageGroups)
$script:btnManageGroups = $btnManageGroups

# User input panel
$grpUsers = New-Object System.Windows.Forms.GroupBox
$grpUsers.Text = 'User Comparison'
$grpUsers.Location = New-Object System.Drawing.Point(20, 130)
$grpUsers.Size = New-Object System.Drawing.Size(1125, 150)
$grpUsers.Anchor = 'Top,Left,Right'
$grpUsers.BackColor = [System.Drawing.Color]::White
$grpUsers.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
$MainForm.Controls.Add($grpUsers)

$lblUserFormatHint = New-Object System.Windows.Forms.Label
$lblUserFormatHint.Text = 'Enter a user principal name, or the start of a display name or UPN to search.'
$lblUserFormatHint.AutoSize = $true
$lblUserFormatHint.ForeColor = [System.Drawing.Color]::DimGray
$lblUserFormatHint.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Italic)
$lblUserFormatHint.Location = New-Object System.Drawing.Point(20, 24)
$grpUsers.Controls.Add($lblUserFormatHint)

$lblReference = New-Object System.Windows.Forms.Label
$lblReference.Text = 'Reference User'
$lblReference.AutoSize = $true
$lblReference.Location = New-Object System.Drawing.Point(18, 30)
$grpUsers.Controls.Add($lblReference)

$txtReference = New-Object System.Windows.Forms.ComboBox
$txtReference.Location = New-Object System.Drawing.Point(20, 54)
$txtReference.Size = New-Object System.Drawing.Size(430, 24)
$txtReference.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$txtReference.DropDownStyle = 'DropDown'
$txtReference.AutoCompleteMode = 'None'
$txtReference.AutoCompleteSource = 'None'
$txtReference.MaxDropDownItems = 20
$grpUsers.Controls.Add($txtReference)
$script:txtReference = $txtReference

$lblReferenceResolved = New-Object System.Windows.Forms.Label
$lblReferenceResolved.Text = ''
$lblReferenceResolved.AutoEllipsis = $true
$lblReferenceResolved.Location = New-Object System.Drawing.Point(20, 85)
$lblReferenceResolved.Size = New-Object System.Drawing.Size(430, 22)
$lblReferenceResolved.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)
$lblReferenceResolved.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$grpUsers.Controls.Add($lblReferenceResolved)
$script:lblReferenceResolved = $lblReferenceResolved

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = 'Target User'
$lblTarget.AutoSize = $true
$lblTarget.Location = New-Object System.Drawing.Point(480, 30)
$grpUsers.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.ComboBox
$txtTarget.Location = New-Object System.Drawing.Point(482, 54)
$txtTarget.Size = New-Object System.Drawing.Size(430, 24)
$txtTarget.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$txtTarget.DropDownStyle = 'DropDown'
$txtTarget.AutoCompleteMode = 'None'
$txtTarget.AutoCompleteSource = 'None'
$txtTarget.MaxDropDownItems = 20
$grpUsers.Controls.Add($txtTarget)
$script:txtTarget = $txtTarget

$lblTargetResolved = New-Object System.Windows.Forms.Label
$lblTargetResolved.Text = ''
$lblTargetResolved.AutoEllipsis = $true
$lblTargetResolved.Location = New-Object System.Drawing.Point(482, 85)
$lblTargetResolved.Size = New-Object System.Drawing.Size(430, 22)
$lblTargetResolved.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)
$lblTargetResolved.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$grpUsers.Controls.Add($lblTargetResolved)
$script:lblTargetResolved = $lblTargetResolved

$btnCompare = New-Object System.Windows.Forms.Button
$btnCompare.Text = 'Compare Users'
$btnCompare.Size = New-Object System.Drawing.Size(165, 42)
$btnCompare.Location = New-Object System.Drawing.Point(940, 48)
$btnCompare.Anchor = 'Top,Right'
$btnCompare.BackColor = [System.Drawing.Color]::FromArgb(31, 58, 93)
$btnCompare.ForeColor = [System.Drawing.Color]::White
$btnCompare.FlatStyle = 'Flat'
$btnCompare.FlatAppearance.BorderSize = 0
$btnCompare.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$btnCompare.Enabled = $false
$grpUsers.Controls.Add($btnCompare)
$script:btnCompare = $btnCompare

$lblReferenceCount = New-Object System.Windows.Forms.Label
$lblReferenceCount.Text = 'Reference security groups: 0'
$lblReferenceCount.AutoSize = $true
$lblReferenceCount.Location = New-Object System.Drawing.Point(20, 118)
$lblReferenceCount.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$grpUsers.Controls.Add($lblReferenceCount)
$script:lblReferenceCount = $lblReferenceCount

$lblTargetCount = New-Object System.Windows.Forms.Label
$lblTargetCount.Text = 'Target security groups: 0'
$lblTargetCount.AutoSize = $true
$lblTargetCount.Location = New-Object System.Drawing.Point(482, 118)
$lblTargetCount.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$grpUsers.Controls.Add($lblTargetCount)
$script:lblTargetCount = $lblTargetCount

# Results toolbar
$pnlResultsToolbar = New-Object System.Windows.Forms.Panel
$pnlResultsToolbar.Location = New-Object System.Drawing.Point(20, 292)
$pnlResultsToolbar.Size = New-Object System.Drawing.Size(1125, 46)
$pnlResultsToolbar.Anchor = 'Top,Left,Right'
$pnlResultsToolbar.BackColor = [System.Drawing.Color]::White
$MainForm.Controls.Add($pnlResultsToolbar)

$lblMissingCount = New-Object System.Windows.Forms.Label
$lblMissingCount.Text = 'Missing groups: 0'
$lblMissingCount.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblMissingCount.AutoSize = $true
$lblMissingCount.Location = New-Object System.Drawing.Point(14, 14)
$pnlResultsToolbar.Controls.Add($lblMissingCount)
$script:lblMissingCount = $lblMissingCount

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = 'Select All'
$btnSelectAll.Size = New-Object System.Drawing.Size(95, 30)
$btnSelectAll.Anchor = 'Top,Right'
$btnSelectAll.Location = New-Object System.Drawing.Point(760, 8)
$btnSelectAll.Enabled = $false
$pnlResultsToolbar.Controls.Add($btnSelectAll)
$script:btnSelectAll = $btnSelectAll

$btnClearSelection = New-Object System.Windows.Forms.Button
$btnClearSelection.Text = 'Clear'
$btnClearSelection.Size = New-Object System.Drawing.Size(80, 30)
$btnClearSelection.Anchor = 'Top,Right'
$btnClearSelection.Location = New-Object System.Drawing.Point(865, 8)
$btnClearSelection.Enabled = $false
$pnlResultsToolbar.Controls.Add($btnClearSelection)
$script:btnClearSelection = $btnClearSelection

$btnAddSelected = New-Object System.Windows.Forms.Button
$btnAddSelected.Text = 'Add Selected Groups'
$btnAddSelected.Size = New-Object System.Drawing.Size(165, 30)
$btnAddSelected.Anchor = 'Top,Right'
$btnAddSelected.Location = New-Object System.Drawing.Point(950, 8)
$btnAddSelected.BackColor = [System.Drawing.Color]::FromArgb(22, 130, 80)
$btnAddSelected.ForeColor = [System.Drawing.Color]::White
$btnAddSelected.FlatStyle = 'Flat'
$btnAddSelected.FlatAppearance.BorderSize = 0
$btnAddSelected.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$btnAddSelected.Enabled = $false
$pnlResultsToolbar.Controls.Add($btnAddSelected)
$script:btnAddSelected = $btnAddSelected

# Results grid
$dgvGroups = New-Object System.Windows.Forms.DataGridView
$dgvGroups.Location = New-Object System.Drawing.Point(20, 342)
$dgvGroups.Size = New-Object System.Drawing.Size(1125, 285)
$dgvGroups.Anchor = 'Top,Bottom,Left,Right'
$dgvGroups.BackgroundColor = [System.Drawing.Color]::White
$dgvGroups.BorderStyle = 'Fixed3D'
$dgvGroups.AllowUserToAddRows = $false
$dgvGroups.AllowUserToDeleteRows = $false
$dgvGroups.AllowUserToResizeRows = $false
$dgvGroups.RowHeadersVisible = $false
$dgvGroups.SelectionMode = 'FullRowSelect'
$dgvGroups.MultiSelect = $false
$dgvGroups.AutoGenerateColumns = $false
$dgvGroups.AutoSizeRowsMode = 'AllCells'
$dgvGroups.DefaultCellStyle.WrapMode = 'False'
$dgvGroups.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$dgvGroups.EnableHeadersVisualStyles = $false
$dgvGroups.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 237, 243)
$dgvGroups.ColumnHeadersHeight = 34
$MainForm.Controls.Add($dgvGroups)
$script:dgvGroups = $dgvGroups

$colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelect.Name = 'Selected'
$colSelect.HeaderText = 'Add'
$colSelect.Width = 50
$colSelect.FalseValue = $false
$colSelect.TrueValue = $true
[void]$dgvGroups.Columns.Add($colSelect)

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = 'GroupName'
$colName.HeaderText = 'Group Name'
$colName.Width = 310
$colName.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colName)

$colDescription = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDescription.Name = 'Description'
$colDescription.HeaderText = 'Description'
$colDescription.AutoSizeMode = 'Fill'
$colDescription.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colDescription)

$colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colType.Name = 'GroupType'
$colType.HeaderText = 'Type'
$colType.Width = 170
$colType.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colType)

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = 'MembershipStatus'
$colStatus.HeaderText = 'Target Status'
$colStatus.Width = 110
$colStatus.ReadOnly = $true
[void]$dgvGroups.Columns.Add($colStatus)

# Activity log
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Activity Log'
$grpLog.Location = New-Object System.Drawing.Point(20, 637)
$grpLog.Size = New-Object System.Drawing.Size(1125, 90)
$grpLog.Anchor = 'Bottom,Left,Right'
$grpLog.BackColor = [System.Drawing.Color]::White
$grpLog.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$MainForm.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 22)
$txtLog.Size = New-Object System.Drawing.Size(1105, 58)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.BorderStyle = 'None'
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$grpLog.Controls.Add($txtLog)
$script:txtLog = $txtLog

# Status strip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.SizingGrip = $false
$statusStrip.BackColor = [System.Drawing.Color]::FromArgb(235, 238, 242)
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = 'Connect to Microsoft 365 to begin.'
$lblStatus.Spring = $true
$lblStatus.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($lblStatus)
$MainForm.Controls.Add($statusStrip)
$script:lblStatus = $lblStatus

# Events
$btnConnect.Add_Click({ Connect-M365Graph })
$btnCompare.Add_Click({ Compare-Users })
$btnAddSelected.Add_Click({ Add-UserToSelectedGroups })

$btnSelectAll.Add_Click({
    foreach ($row in $dgvGroups.Rows) {
        if (-not $row.IsNewRow -and -not $row.Cells['Selected'].ReadOnly) {
            $row.Cells['Selected'].Value = $true
        }
    }
})

$btnClearSelection.Add_Click({
    foreach ($row in $dgvGroups.Rows) {
        $row.Cells['Selected'].Value = $false
    }
})

$dgvGroups.Add_CurrentCellDirtyStateChanged({
    if ($dgvGroups.IsCurrentCellDirty) {
        $dgvGroups.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

# Enter is used to accept an entry from the user list, so it no longer starts a
# comparison. Use the Compare button instead.
$txtReference.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
    }
})

$txtTarget.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
    }
})

$MainForm.Add_FormClosing({
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
})

Apply-MSToolkitSharedTheme -Root $MainForm

Write-AppLog 'Microsoft 365 Group Compare started.'

# Offboard User passes -AutoConnect and -ManageUser so this opens ready to work:
# sign in, then straight into Manage User Groups with that user loaded.
if ($AutoConnect -or $ManageUser) {
    $MainForm.Add_Shown({
        try {
            if (-not ($null -ne (Get-MgContext))) {
                Write-AppLog 'Connecting automatically...'
                Connect-M365Graph
            }

            if ($ManageUser -and ($null -ne (Get-MgContext))) {
                Write-AppLog "Opening Manage User Groups for $ManageUser."
                Show-SecurityGroupManager -PrefillUser $ManageUser
            }
            elseif ($ManageUser) {
                Write-AppLog 'Not connected, so Manage User Groups was not opened. Connect, then open it yourself.' 'WARNING'
            }
        }
        catch {
            Write-AppLog "Automatic start-up failed: $($_.Exception.Message)" 'ERROR'
        }
    })
}
Write-AppLog 'Only direct, non-dynamic, non-mail-enabled security-group memberships are compared.'

Hide-PowerShellConsole
# Type-ahead filtering on every editable dropdown. Its own handler, not folded
# into another, so it runs on every launch rather than only when the tool is
# started with parameters - and so a failure here cannot stop the rest of
# start-up. Multiple Shown handlers chain.
$MainForm.Add_Shown({ Register-MSToolkitComboFiltersOn -Root $MainForm })

[void]$MainForm.ShowDialog()
