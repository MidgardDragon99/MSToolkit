[CmdletBinding()]
param(
    [ValidateSet("Light","Dark")]
    [string]$ThemeMode
)

$ErrorActionPreference = 'Stop'

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
            $ToolProperty = "Theme_M365ConditionalAccess"
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



if (-not ("MSToolkitCAConsole.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace MSToolkitCAConsole {
    public static class NativeMethods {
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}
"@
}

try {
    $h = [MSToolkitCAConsole.NativeMethods]::GetConsoleWindow()
    if ($h -ne [IntPtr]::Zero) {
        [MSToolkitCAConsole.NativeMethods]::ShowWindow($h, 0) | Out-Null
    }
} catch {}


# Keep Microsoft sign-in UI on the same monitor as this app.
if (-not ("MSToolkitCAAuthWindow.AuthWindowWatcher" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace MSToolkitCAAuthWindow
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

function Start-SignInWindowWatcher {
    try {
        if ($script:MainForm -and -not $script:MainForm.IsDisposed) {
            $null = $script:MainForm.Handle
            [MSToolkitCAAuthWindow.AuthWindowWatcher]::Start($script:MainForm.Handle)
        }
    }
    catch {
        # Window positioning is best effort only.
    }
}

function Stop-SignInWindowWatcher {
    try {
        [MSToolkitCAAuthWindow.AuthWindowWatcher]::Stop()
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
$script:CurrentUser = $null
$script:Policies = @()
$script:UserGroupMap = @{}

function Show-InfoMessage {
    param([string]$Message,[string]$Title='Conditional Access User Manager')
    [System.Windows.Forms.MessageBox]::Show(
        $Message,$Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Export-MSToolkitGridToCsv {
    # Writes whatever is currently in a grid to CSV, using the grid's own column
    # headers. The checkbox column is skipped - it is a selection, not data.
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$BaseName,
        [string]$Subject
    )

    if (-not $Grid -or $Grid.Rows.Count -eq 0) {
        Show-InfoMessage 'There is nothing to export yet. Load the policies first.'
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
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show(
        $Message,'Conditional Access User Manager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Write-AppLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level='INFO'
    )

    $color = switch ($Level) {
        'SUCCESS' { (Get-MSToolkitThemePalette).Success }
        'WARNING' { (Get-MSToolkitThemePalette).Warning }
        'ERROR'   { (Get-MSToolkitThemePalette).Danger }
        default   { [System.Drawing.Color]::FromArgb(45,55,65) }
    }

    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.SelectionLength = 0
    $script:txtLog.SelectionColor = (Convert-MSToolkitThemeColor $color)
    $script:txtLog.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n")
    $script:txtLog.SelectionColor = $script:txtLog.ForeColor
    $script:txtLog.ScrollToCaret()
}

function Set-BusyState {
    param([bool]$Busy,[string]$StatusText='')

    $script:btnConnect.Enabled = -not $Busy
    $script:btnLoad.Enabled = (-not $Busy) -and $script:GraphConnected
    $hasRows = ($script:dgvPolicies.Rows.Count -gt 0) -and $script:CurrentUser
    $script:btnInclude.Enabled = (-not $Busy) -and $hasRows
    $script:btnExclude.Enabled = (-not $Busy) -and $hasRows
    $script:btnRemoveDirect.Enabled = (-not $Busy) -and $hasRows

    if ($StatusText) { $script:lblStatus.Text = $StatusText }
    $script:MainForm.UseWaitCursor = $Busy

    # A ComboBox owns its own window handle, so clearing UseWaitCursor on the form
    # does not always restore the pointer over it. Reset the cursor explicitly.
    if (-not $Busy) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-GraphGetAll {
    param([Parameter(Mandatory)][string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri

    while ($next) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop

        if ($response.value) {
            foreach ($item in $response.value) { $items.Add($item) }
        } else {
            $items.Add($response)
        }

        $next = $response.'@odata.nextLink'
    }

    return $items.ToArray()
}

function Connect-ConditionalAccessGraph {
    try {
        Set-BusyState -Busy $true -StatusText 'Connecting to Microsoft Graph...'

        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            throw 'Microsoft.Graph.Authentication is not installed.'
        }

        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

        $scopes = @(
            'Policy.Read.All',
            'Policy.ReadWrite.ConditionalAccess',
            'Application.Read.All',
            'User.Read.All',
            'Group.Read.All',
            'Directory.Read.All'
        )

        Start-SignInWindowWatcher

        Connect-MgGraph `
            -Scopes $scopes `
            -ContextScope Process `
            -NoWelcome `
            -ErrorAction Stop | Out-Null

        Stop-SignInWindowWatcher
        $context = Get-MgContext
        $script:GraphConnected = $true
        $script:lblConnection.Text = "Connected: $($context.Account)"
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Success
        $script:btnLoad.Enabled = $true
        $script:lblStatus.Text = 'Connected to Microsoft Graph.'
        Write-AppLog "Connected to Microsoft Graph as $($context.Account)." 'SUCCESS'

        $script:MSToolkitUserPickerCache = $null
        $LoadedCount = Initialize-MSToolkitUserPicker -Combos @($script:txtUser)

        if ($LoadedCount -ge 0) {
            Write-AppLog "Loaded $LoadedCount Microsoft 365 user(s) into the user list."
        }
        else {
            Write-AppLog 'Could not load the Microsoft 365 user list. Type a user instead.' 'WARNING'
        }
    }
    catch {
        $script:GraphConnected = $false
        $script:lblConnection.Text = 'Not connected'
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Danger
        Show-ErrorMessage "Unable to connect to Microsoft Graph.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Graph connection failed: $($_.Exception.Message)" 'ERROR'
    }
    finally {
        Set-BusyState -Busy $false
    }
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
        throw 'A user principal name was not entered.'
    }

    $encodedIdentity = [uri]::EscapeDataString($Identity)

    try {
        return Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/users/$encodedIdentity" `
            -ErrorAction Stop
    }
    catch {
        $safeIdentity = $Identity.Replace("'", "''")
        $filter = [uri]::EscapeDataString("userPrincipalName eq '$safeIdentity'")
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter"

        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $matches = @($response.value)

        if ($matches.Count -eq 0) {
            throw "User not found: $Identity"
        }

        return $matches[0]
    }
}

function Get-UserTransitiveGroups {
    param([Parameter(Mandatory)][string]$UserId)

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/transitiveMemberOf/microsoft.graph.group?`$select=id,displayName&`$top=999"
    return @(Invoke-GraphGetAll -Uri $uri)
}

function Get-ConditionalAccessPolicies {
    $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
    return @(Invoke-GraphGetAll -Uri $uri)
}

function Get-Array {
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value)
}

function Get-GroupNamesFromIds {
    param([string[]]$Ids)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($id in @(Get-Array $Ids)) {
        if ($script:UserGroupMap.ContainsKey([string]$id)) {
            $names.Add([string]$script:UserGroupMap[[string]$id])
        }
    }

    return @($names.ToArray() | Sort-Object)
}

function Get-PolicyUserScope {
    param(
        [Parameter(Mandatory)]$Policy,
        [Parameter(Mandatory)][string]$UserId
    )

    $users = $Policy.conditions.users

    if ($null -eq $users) {
        return [pscustomobject]@{
            Scope='No user targeting'
            EffectiveResult='Not Targeted'
            Detail=''
            DirectInclude=$false
            DirectExclude=$false
        }
    }

    $includeUsers = @(Get-Array $users.includeUsers)
    $excludeUsers = @(Get-Array $users.excludeUsers)
    $includeGroups = @(Get-Array $users.includeGroups)
    $excludeGroups = @(Get-Array $users.excludeGroups)

    $directInclude = ($includeUsers -contains $UserId)
    $directExclude = ($excludeUsers -contains $UserId)

    $excludedGroupNames = @(Get-GroupNamesFromIds -Ids $excludeGroups)
    $includedGroupNames = @(Get-GroupNamesFromIds -Ids $includeGroups)

    if ($directExclude) {
        return [pscustomobject]@{
            Scope='DIRECT EXCLUDE'
            EffectiveResult='Excluded'
            Detail='User is explicitly excluded'
            DirectInclude=$directInclude
            DirectExclude=$true
        }
    }

    if ($excludedGroupNames.Count -gt 0) {
        return [pscustomobject]@{
            Scope='Excluded via Group'
            EffectiveResult='Excluded'
            Detail=($excludedGroupNames -join '; ')
            DirectInclude=$directInclude
            DirectExclude=$false
        }
    }

    if ($directInclude) {
        return [pscustomobject]@{
            Scope='DIRECT INCLUDE'
            EffectiveResult='Included'
            Detail='User is explicitly included'
            DirectInclude=$true
            DirectExclude=$false
        }
    }

    if ($includeUsers -contains 'All') {
        return [pscustomobject]@{
            Scope='Included via All Users'
            EffectiveResult='Included'
            Detail=''
            DirectInclude=$false
            DirectExclude=$false
        }
    }

    if ($includedGroupNames.Count -gt 0) {
        return [pscustomobject]@{
            Scope='Included via Group'
            EffectiveResult='Included'
            Detail=($includedGroupNames -join '; ')
            DirectInclude=$false
            DirectExclude=$false
        }
    }

    return [pscustomobject]@{
        Scope='Not directly targeted'
        EffectiveResult='Not Targeted'
        Detail=''
        DirectInclude=$false
        DirectExclude=$false
    }
}

function Get-GrantSummary {
    param($Policy)

    if ($null -eq $Policy.grantControls) { return '' }

    $controls = @(Get-Array $Policy.grantControls.builtInControls)
    if ($controls.Count -eq 0) { return [string]$Policy.grantControls.operator }

    $prefix = if ($Policy.grantControls.operator) { "$($Policy.grantControls.operator): " } else { '' }
    return $prefix + ($controls -join ', ')
}

function Update-PolicyGrid {
    $script:dgvPolicies.Rows.Clear()

    foreach ($policy in $script:Policies) {
        $scopeInfo = Get-PolicyUserScope -Policy $policy -UserId $script:CurrentUser.id

        $index = $script:dgvPolicies.Rows.Add(
            $false,
            [string]$policy.displayName,
            [string]$policy.state,
            $scopeInfo.EffectiveResult,
            $scopeInfo.Scope,
            $scopeInfo.Detail,
            (Get-GrantSummary -Policy $policy)
        )

        $script:dgvPolicies.Rows[$index].Tag = [pscustomobject]@{
            Policy=$policy
            ScopeInfo=$scopeInfo
        }

        # Policy state takes visual priority so disabled/report-only policies do not
        # look equivalent to enabled policies. Enabled policies retain targeting colors.
        switch ([string]$policy.state) {
            'disabled' {
                $script:dgvPolicies.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).DisabledBackground
                $script:dgvPolicies.Rows[$index].DefaultCellStyle.ForeColor = (Get-MSToolkitThemePalette).DisabledText
            }
            'enabledForReportingButNotEnforced' {
                $script:dgvPolicies.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).InfoBackground
                $script:dgvPolicies.Rows[$index].DefaultCellStyle.ForeColor = (Get-MSToolkitThemePalette).Info
            }
            default {
                switch ($scopeInfo.Scope) {
                    'DIRECT EXCLUDE' {
                        $script:dgvPolicies.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).DangerBackground
                    }
                    'DIRECT INCLUDE' {
                        $script:dgvPolicies.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).SuccessBackground
                    }
                    'Excluded via Group' {
                        $script:dgvPolicies.Rows[$index].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).WarningBackground
                    }
                }
            }
        }
    }

    $script:lblPolicyCount.Text = "Policies: $($script:Policies.Count)"
    $hasRows = $script:dgvPolicies.Rows.Count -gt 0
    $script:btnInclude.Enabled = $hasRows
    $script:btnExclude.Enabled = $hasRows
    $script:btnRemoveDirect.Enabled = $hasRows
    $script:btnSelectAll.Enabled = $hasRows
    $script:btnClear.Enabled = $hasRows

    if ($script:btnExportPolicies) { $script:btnExportPolicies.Enabled = $hasRows }
}

function Load-UserPolicies {
    if (-not $script:GraphConnected) {
        Show-InfoMessage 'Connect to Microsoft Graph first.'
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:txtUser.Text)) {
        Show-InfoMessage 'Enter a user first.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Loading user and Conditional Access policies...'
        $script:dgvPolicies.Rows.Clear()
        $script:CurrentUser = $null
        $script:Policies = @()
        $script:UserGroupMap = @{}

        Write-AppLog "Resolving user: $($script:txtUser.Text)"
        try {
            $script:CurrentUser = Resolve-GraphUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtUser.Text)
        }
        catch {
            throw "USER LOOKUP failed: $($_.Exception.Message)"
        }

        $script:lblResolved.Text = "$($script:CurrentUser.displayName)  |  $($script:CurrentUser.userPrincipalName)"
        Write-AppLog "Resolved user: $($script:CurrentUser.userPrincipalName)" 'SUCCESS'

        Write-AppLog "Reading transitive group memberships for $($script:CurrentUser.userPrincipalName)..."
        try {
            $userGroups = @(Get-UserTransitiveGroups -UserId $script:CurrentUser.id)
        }
        catch {
            throw "GROUP MEMBERSHIP LOOKUP failed: $($_.Exception.Message)"
        }

        foreach ($group in $userGroups) {
            $script:UserGroupMap[[string]$group.id] = [string]$group.displayName
        }

        Write-AppLog "Read $($userGroups.Count) transitive group membership(s)." 'SUCCESS'
        Write-AppLog 'Reading Conditional Access policies...'

        try {
            $script:Policies = @(Get-ConditionalAccessPolicies | Sort-Object { "$($_.displayName)" })
        }
        catch {
            throw "CONDITIONAL ACCESS POLICY LOOKUP failed: $($_.Exception.Message)"
        }

        Write-AppLog "Read $($script:Policies.Count) Conditional Access policy/policies." 'SUCCESS'

        Update-PolicyGrid

        $script:lblStatus.Text = "Loaded $($script:Policies.Count) Conditional Access policies."
        Write-AppLog "Loaded $($script:Policies.Count) policies and $($userGroups.Count) transitive user group memberships." 'SUCCESS'
    }
    catch {
        $script:lblStatus.Text = 'Load failed.'
        Write-AppLog "Load failed: $($_.Exception.Message)" 'ERROR'
        Show-ErrorMessage "Unable to load Conditional Access policies.`r`n`r`n$($_.Exception.Message)"
    }
    finally {
        Set-BusyState -Busy $false
    }
}

function Get-SelectedPolicyRows {
    if ($script:dgvPolicies.IsCurrentCellDirty) {
        $script:dgvPolicies.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        $script:dgvPolicies.EndEdit()
    }

    $selected = New-Object System.Collections.Generic.List[object]

    foreach ($row in $script:dgvPolicies.Rows) {
        if (-not $row.IsNewRow -and [bool]$row.Cells['Selected'].Value -and $row.Tag) {
            $selected.Add($row.Tag)
        }
    }

    return $selected.ToArray()
}

function Set-DirectPolicyAssignment {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Include','Exclude','RemoveDirect')]
        [string]$Mode
    )

    if (-not $script:CurrentUser) {
        Show-InfoMessage 'Load a user first.'
        return
    }

    $selectedRows = @(Get-SelectedPolicyRows)
    if ($selectedRows.Count -eq 0) {
        Show-InfoMessage 'Check at least one policy first.'
        return
    }

    $verb = switch ($Mode) {
        'Include'      { 'directly INCLUDE' }
        'Exclude'      { 'directly EXCLUDE' }
        'RemoveDirect' { 'remove DIRECT include/exclude assignments for' }
    }

    $names = @($selectedRows | ForEach-Object { $_.Policy.displayName } | Sort-Object)
    $nameText = ($names -join "`r`n - ")

    $warning = @"
This changes the Conditional Access policy itself.

It does NOT change group memberships.
Group-based and All Users assignments remain in place.

Proceed to $verb $($script:CurrentUser.displayName) in the following $($selectedRows.Count) policy/policies?

 - $nameText
"@

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $warning,
        'Confirm Conditional Access Changes',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $success = 0
    $failed = 0

    try {
        Set-BusyState -Busy $true -StatusText 'Updating Conditional Access policies...'

        foreach ($rowData in $selectedRows) {
            $policy = $rowData.Policy

            try {
                $includeUsers = New-Object System.Collections.Generic.List[string]
                foreach ($id in @(Get-Array $policy.conditions.users.includeUsers)) {
                    $includeUsers.Add([string]$id)
                }

                $excludeUsers = New-Object System.Collections.Generic.List[string]
                foreach ($id in @(Get-Array $policy.conditions.users.excludeUsers)) {
                    $excludeUsers.Add([string]$id)
                }

                $userId = [string]$script:CurrentUser.id

                switch ($Mode) {
                    'Include' {
                        while ($excludeUsers.Contains($userId)) {
                            [void]$excludeUsers.Remove($userId)
                        }
                        if (-not $includeUsers.Contains($userId)) {
                            $includeUsers.Add($userId)
                        }
                    }

                    'Exclude' {
                        while ($includeUsers.Contains($userId)) {
                            [void]$includeUsers.Remove($userId)
                        }
                        if (-not $excludeUsers.Contains($userId)) {
                            $excludeUsers.Add($userId)
                        }
                    }

                    'RemoveDirect' {
                        while ($includeUsers.Contains($userId)) {
                            [void]$includeUsers.Remove($userId)
                        }
                        while ($excludeUsers.Contains($userId)) {
                            [void]$excludeUsers.Remove($userId)
                        }
                    }
                }

                $body = @{
                    conditions = @{
                        users = @{
                            includeUsers = $includeUsers.ToArray()
                            excludeUsers = $excludeUsers.ToArray()
                        }
                    }
                } | ConvertTo-Json -Depth 10 -Compress

                $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$($policy.id)"
                Invoke-MgGraphRequest `
                    -Method PATCH `
                    -Uri $uri `
                    -Body $body `
                    -ContentType 'application/json' `
                    -ErrorAction Stop | Out-Null

                Write-AppLog "$Mode succeeded for '$($policy.displayName)'." 'SUCCESS'
                $success++
            }
            catch {
                Write-AppLog "$Mode failed for '$($policy.displayName)': $($_.Exception.Message)" 'ERROR'
                $failed++
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Conditional Access update complete.`r`n`r`nSuccessful: $success`r`nFailed: $failed`r`n`r`nReview the Activity Log for details.",
            'Update Complete',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $(if ($failed -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information })
        ) | Out-Null

        Load-UserPolicies
    }
    finally {
        Set-BusyState -Busy $false
    }
}

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'Conditional Access User Manager'
$MainForm.StartPosition = 'CenterScreen'
$MainForm.Size = New-Object System.Drawing.Size(1260,820)
$MainForm.MinimumSize = New-Object System.Drawing.Size(1100,700)
$MainForm.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
$MainForm.Font = New-Object System.Drawing.Font('Segoe UI',9)
$MainForm.FormBorderStyle = 'Sizable'
$MainForm.MaximizeBox = $true
$script:MainForm = $MainForm

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = 'Top'
$pnlHeader.Height = 78
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$MainForm.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Conditional Access User Manager'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',20)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20,12)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Review user targeting across Conditional Access policies and manage direct user include/exclude assignments'
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI',9.5)
$lblSubtitle.AutoSize = $true
$lblSubtitle.Location = New-Object System.Drawing.Point(23,49)
$pnlHeader.Controls.Add($lblSubtitle)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Connect to Microsoft 365'
$btnConnect.Size = New-Object System.Drawing.Size(190,36)
$btnConnect.Anchor = 'Top,Right'
$btnConnect.Location = New-Object System.Drawing.Point(1035,20)
$btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.FlatStyle = 'Flat'
$btnConnect.FlatAppearance.BorderSize = 0
$btnConnect.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$btnConnect.Add_Click({ Connect-ConditionalAccessGraph })
$pnlHeader.Controls.Add($btnConnect)
$script:MSToolkitThemeToggleButton = New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "M365ConditionalAccess"

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

$pnlConnection = New-Object System.Windows.Forms.Panel
$pnlConnection.Dock = 'Top'
$pnlConnection.Height = 38
$pnlConnection.BackColor = [System.Drawing.Color]::White
$MainForm.Controls.Add($pnlConnection)
$pnlConnection.BringToFront()

$lblConnectionLabel = New-Object System.Windows.Forms.Label
$lblConnectionLabel.Text = 'Microsoft Graph:'
$lblConnectionLabel.AutoSize = $true
$lblConnectionLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$lblConnectionLabel.Location = New-Object System.Drawing.Point(20,10)
$pnlConnection.Controls.Add($lblConnectionLabel)

$lblConnection = New-Object System.Windows.Forms.Label
$lblConnection.Text = 'Not connected'
$lblConnection.AutoSize = $true
$lblConnection.ForeColor = [System.Drawing.Color]::Firebrick
$lblConnection.Location = New-Object System.Drawing.Point(122,10)
$pnlConnection.Controls.Add($lblConnection)
$script:lblConnection = $lblConnection

$grpUser = New-Object System.Windows.Forms.GroupBox
$grpUser.Text = 'User'
$grpUser.Location = New-Object System.Drawing.Point(20,130)
$grpUser.Size = New-Object System.Drawing.Size(1205,125)
$grpUser.Anchor = 'Top,Left,Right'
$grpUser.BackColor = [System.Drawing.Color]::White
$grpUser.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9.5)
$MainForm.Controls.Add($grpUser)

$lblUserHint = New-Object System.Windows.Forms.Label
$lblUserHint.Text = 'Enter the full user principal name (sign-in name), including @domain.'
$lblUserHint.AutoSize = $true
$lblUserHint.ForeColor = [System.Drawing.Color]::DimGray
$lblUserHint.Font = New-Object System.Drawing.Font('Segoe UI',8.5,[System.Drawing.FontStyle]::Italic)
$lblUserHint.Location = New-Object System.Drawing.Point(20,24)
$grpUser.Controls.Add($lblUserHint)

$txtUser = New-Object System.Windows.Forms.ComboBox
$txtUser.Location = New-Object System.Drawing.Point(20,52)
$txtUser.Size = New-Object System.Drawing.Size(700,25)
$txtUser.Font = New-Object System.Drawing.Font('Segoe UI',10)
$txtUser.DropDownStyle = 'DropDown'
$txtUser.AutoCompleteMode = 'None'
$txtUser.AutoCompleteSource = 'None'
$txtUser.MaxDropDownItems = 20
$grpUser.Controls.Add($txtUser)
$script:txtUser = $txtUser

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = 'Load Policies'
$btnLoad.Location = New-Object System.Drawing.Point(1010,47)
$btnLoad.Size = New-Object System.Drawing.Size(165,38)
$btnLoad.Anchor = 'Top,Right'
$btnLoad.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$btnLoad.ForeColor = [System.Drawing.Color]::White
$btnLoad.FlatStyle = 'Flat'
$btnLoad.FlatAppearance.BorderSize = 0
$btnLoad.Font = New-Object System.Drawing.Font('Segoe UI Semibold',10)
$btnLoad.Enabled = $false
$btnLoad.Add_Click({ Load-UserPolicies })
$grpUser.Controls.Add($btnLoad)
$script:btnLoad = $btnLoad

$lblResolved = New-Object System.Windows.Forms.Label
$lblResolved.Text = ''
$lblResolved.Location = New-Object System.Drawing.Point(20,84)
$lblResolved.Size = New-Object System.Drawing.Size(900,22)
$lblResolved.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)
$lblResolved.AutoEllipsis = $true
$grpUser.Controls.Add($lblResolved)
$script:lblResolved = $lblResolved

$pnlToolbar = New-Object System.Windows.Forms.Panel
$pnlToolbar.Location = New-Object System.Drawing.Point(20,267)
$pnlToolbar.Size = New-Object System.Drawing.Size(1205,50)
$pnlToolbar.Anchor = 'Top,Left,Right'
$pnlToolbar.BackColor = [System.Drawing.Color]::White
$MainForm.Controls.Add($pnlToolbar)

$lblPolicyCount = New-Object System.Windows.Forms.Label
$lblPolicyCount.Text = 'Policies: 0'
$lblPolicyCount.Location = New-Object System.Drawing.Point(12,16)
$lblPolicyCount.Size = New-Object System.Drawing.Size(120,22)
$lblPolicyCount.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$pnlToolbar.Controls.Add($lblPolicyCount)
$script:lblPolicyCount = $lblPolicyCount

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = 'Select All'
$btnSelectAll.Location = New-Object System.Drawing.Point(140,10)
$btnSelectAll.Size = New-Object System.Drawing.Size(90,30)
$btnSelectAll.Enabled = $false
$btnSelectAll.Add_Click({
    foreach ($row in $script:dgvPolicies.Rows) {
        if (-not $row.IsNewRow) { $row.Cells['Selected'].Value = $true }
    }
})
$pnlToolbar.Controls.Add($btnSelectAll)
$script:btnSelectAll = $btnSelectAll

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(238,10)
$btnClear.Size = New-Object System.Drawing.Size(80,30)
$btnClear.Enabled = $false
$btnClear.Add_Click({
    foreach ($row in $script:dgvPolicies.Rows) {
        if (-not $row.IsNewRow) { $row.Cells['Selected'].Value = $false }
    }
})
$pnlToolbar.Controls.Add($btnClear)

$btnExportPolicies = New-Object System.Windows.Forms.Button
$btnExportPolicies.Text = 'Export Conditional Access'
$btnExportPolicies.Location = New-Object System.Drawing.Point(328, 10)
$btnExportPolicies.Size = New-Object System.Drawing.Size(190, 30)
$btnExportPolicies.Enabled = $false
$btnExportPolicies.Add_Click({
    $subject = ''
    $baseName = 'ConditionalAccess'

    if ($script:CurrentUser) {
        $subject = "$($script:CurrentUser.displayName) <$($script:CurrentUser.userPrincipalName)>"
        $baseName = "ConditionalAccess-$($script:CurrentUser.userPrincipalName -replace '@.*$','')"
    }

    Export-MSToolkitGridToCsv -Grid $script:dgvPolicies -BaseName $baseName -Subject $subject
})
$pnlToolbar.Controls.Add($btnExportPolicies)
$script:btnExportPolicies = $btnExportPolicies
$script:btnClear = $btnClear

$btnRemoveDirect = New-Object System.Windows.Forms.Button
$btnRemoveDirect.Text = 'Remove Direct Assignment'
$btnRemoveDirect.Location = New-Object System.Drawing.Point(615,8)
$btnRemoveDirect.Size = New-Object System.Drawing.Size(190,34)
$btnRemoveDirect.Anchor = 'Top,Right'
$btnRemoveDirect.Enabled = $false
$btnRemoveDirect.Add_Click({ Set-DirectPolicyAssignment -Mode RemoveDirect })
$pnlToolbar.Controls.Add($btnRemoveDirect)
$script:btnRemoveDirect = $btnRemoveDirect

$btnExclude = New-Object System.Windows.Forms.Button
$btnExclude.Text = 'Directly Exclude User'
$btnExclude.Location = New-Object System.Drawing.Point(815,8)
$btnExclude.Size = New-Object System.Drawing.Size(180,34)
$btnExclude.Anchor = 'Top,Right'
$btnExclude.BackColor = [System.Drawing.Color]::FromArgb(180,75,70)
$btnExclude.ForeColor = [System.Drawing.Color]::White
$btnExclude.FlatStyle = 'Flat'
$btnExclude.FlatAppearance.BorderSize = 0
$btnExclude.Enabled = $false
$btnExclude.Add_Click({ Set-DirectPolicyAssignment -Mode Exclude })
$pnlToolbar.Controls.Add($btnExclude)
$script:btnExclude = $btnExclude

$btnInclude = New-Object System.Windows.Forms.Button
$btnInclude.Text = 'Directly Include User'
$btnInclude.Location = New-Object System.Drawing.Point(1005,8)
$btnInclude.Size = New-Object System.Drawing.Size(185,34)
$btnInclude.Anchor = 'Top,Right'
$btnInclude.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$btnInclude.ForeColor = [System.Drawing.Color]::White
$btnInclude.FlatStyle = 'Flat'
$btnInclude.FlatAppearance.BorderSize = 0
$btnInclude.Enabled = $false
$btnInclude.Add_Click({ Set-DirectPolicyAssignment -Mode Include })
$pnlToolbar.Controls.Add($btnInclude)
$script:btnInclude = $btnInclude

$dgvPolicies = New-Object System.Windows.Forms.DataGridView
$dgvPolicies.Location = New-Object System.Drawing.Point(20,329)
$dgvPolicies.Size = New-Object System.Drawing.Size(1205,265)
$dgvPolicies.Anchor = 'Top,Bottom,Left,Right'
$dgvPolicies.AllowUserToAddRows = $false
$dgvPolicies.AllowUserToDeleteRows = $false
$dgvPolicies.RowHeadersVisible = $false
$dgvPolicies.MultiSelect = $false
$dgvPolicies.SelectionMode = 'FullRowSelect'
$dgvPolicies.BackgroundColor = [System.Drawing.Color]::White
$dgvPolicies.AutoSizeRowsMode = 'None'
$MainForm.Controls.Add($dgvPolicies)
$script:dgvPolicies = $dgvPolicies

$colSelected = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelected.Name = 'Selected'
$colSelected.HeaderText = 'Select'
$colSelected.Width = 55
[void]$dgvPolicies.Columns.Add($colSelected)

$colPolicy = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPolicy.Name = 'PolicyName'
$colPolicy.HeaderText = 'Policy Name'
$colPolicy.Width = 245
$colPolicy.ReadOnly = $true
[void]$dgvPolicies.Columns.Add($colPolicy)

$colState = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colState.Name = 'State'
$colState.HeaderText = 'State'
$colState.Width = 125
$colState.ReadOnly = $true
[void]$dgvPolicies.Columns.Add($colState)

$colEffective = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEffective.Name = 'EffectiveResult'
$colEffective.HeaderText = 'Effective Result'
$colEffective.Width = 115
$colEffective.ReadOnly = $true
[void]$dgvPolicies.Columns.Add($colEffective)

$colScope = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colScope.Name = 'UserScope'
$colScope.HeaderText = 'User Scope'
$colScope.Width = 160
$colScope.ReadOnly = $true
[void]$dgvPolicies.Columns.Add($colScope)

$colDetail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDetail.Name = 'ScopeDetail'
$colDetail.HeaderText = 'Scope Detail'
$colDetail.AutoSizeMode = 'Fill'
$colDetail.ReadOnly = $true
[void]$dgvPolicies.Columns.Add($colDetail)

$colGrant = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGrant.Name = 'GrantControls'
$colGrant.HeaderText = 'Grant Controls'
$colGrant.Width = 190
$colGrant.ReadOnly = $true
[void]$dgvPolicies.Columns.Add($colGrant)

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Activity Log'
$grpLog.Location = New-Object System.Drawing.Point(20,605)
$grpLog.Size = New-Object System.Drawing.Size(1205,155)
$grpLog.Anchor = 'Bottom,Left,Right'
$grpLog.BackColor = [System.Drawing.Color]::White
$grpLog.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$MainForm.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(10,20)
$txtLog.Size = New-Object System.Drawing.Size(1185,123)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.BorderStyle = 'FixedSingle'
$txtLog.Font = New-Object System.Drawing.Font('Consolas',8.5)
$txtLog.ScrollBars = 'Vertical'
$txtLog.WordWrap = $false
$grpLog.Controls.Add($txtLog)
$script:txtLog = $txtLog

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.SizingGrip = $false
$MainForm.Controls.Add($statusStrip)

$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = 'Ready. Connect to Microsoft Graph.'
$lblStatus.Spring = $true
$lblStatus.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($lblStatus)
$script:lblStatus = $lblStatus

$txtUser.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter -and $script:GraphConnected) {
        Load-UserPolicies
    }
})

$MainForm.Add_Shown({
    $MainForm.Activate()
    Write-AppLog 'Conditional Access User Manager loaded.'
})

Apply-MSToolkitSharedTheme -Root $MainForm
# Type-ahead filtering on every editable dropdown. Its own handler, not folded
# into another, so it runs on every launch rather than only when the tool is
# started with parameters - and so a failure here cannot stop the rest of
# start-up. Multiple Shown handlers chain.
$MainForm.Add_Shown({ Register-MSToolkitComboFiltersOn -Root $MainForm })

[void]$MainForm.ShowDialog()
