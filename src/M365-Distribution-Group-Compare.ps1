#requires -Version 5.1

[CmdletBinding()]
param(
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
    $script:btnConnect.Enabled = -not $Busy
    $script:btnCompare.Enabled = (-not $Busy -and $script:ExchangeConnected)
    $script:btnAddSelected.Enabled = (-not $Busy -and $script:ComparisonRows.Count -gt 0)
    $script:btnSelectAll.Enabled = (-not $Busy -and $script:ComparisonRows.Count -gt 0)
    $script:btnClearSelection.Enabled = (-not $Busy -and $script:ComparisonRows.Count -gt 0)

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
        $script:lblConnection.ForeColor = (Get-MSToolkitThemePalette).Success
        $script:btnCompare.Enabled = $true
        $script:btnManageGroups.Enabled = $true
        $script:lblStatus.Text = 'Connected to Exchange Online. Enter two users and click Compare.'
        Write-AppLog "Connected to Exchange Online: $($org.DisplayName)." 'SUCCESS'

        $script:MSToolkitUserPickerCache = $null
        $LoadedCount = Initialize-MSToolkitUserPicker -Combos @($script:txtReference, $script:txtTarget)

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
        Show-ErrorMessage "Unable to connect to Exchange Online.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "Exchange Online connection failed: $($_.Exception.Message)" 'ERROR'
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

function Initialize-MSToolkitUserPicker {
    param([System.Windows.Forms.ComboBox[]]$Combos)

    try {
        if (-not $script:MSToolkitUserPickerCache) {
            $script:MSToolkitUserPickerCache = Get-MSToolkitExchangeUserList
        }

        foreach ($Combo in $Combos) {
            if (-not $Combo) { continue }

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

function Get-DistributionGroupsForUser {
    param([Parameter(Mandatory)]$Recipient)

    $dn = [string]$Recipient.DistinguishedName
    if ([string]::IsNullOrWhiteSpace($dn)) { throw "Exchange did not return a DistinguishedName for $($Recipient.PrimarySmtpAddress)." }

    # Members is a filterable Exchange property. This returns direct memberships.
    # RecipientTypeDetails restricts results to standard distribution groups only.
    $escapedDn = $dn.Replace("'","''")
    $filter = "Members -eq '$escapedDn'"

    $groups = @(Get-DistributionGroup `
        -RecipientTypeDetails MailUniversalDistributionGroup `
        -Filter $filter `
        -ResultSize Unlimited `
        -ErrorAction Stop)

    return @($groups | Sort-Object DisplayName)
}

function Update-ComparisonGrid {
    $script:dgvGroups.Rows.Clear()

    foreach ($row in $script:ComparisonRows) {
        $index = $script:dgvGroups.Rows.Add(
            $false,
            $row.DisplayName,
            $row.PrimarySmtpAddress,
            $row.ManagedBy,
            'Distribution Group',
            'Missing'
        )
        $script:dgvGroups.Rows[$index].Tag = $row
    }

    $script:lblMissingCount.Text = "Missing groups: $($script:ComparisonRows.Count)"
    $script:btnAddSelected.Enabled = $script:ComparisonRows.Count -gt 0
    $script:btnSelectAll.Enabled = $script:ComparisonRows.Count -gt 0
    $script:btnClearSelection.Enabled = $script:ComparisonRows.Count -gt 0
}

function Compare-Users {
    if (-not $script:ExchangeConnected) { Show-InfoMessage 'Connect to Exchange Online first.'; return }

    if ([string]::IsNullOrWhiteSpace($script:txtReference.Text) -or [string]::IsNullOrWhiteSpace($script:txtTarget.Text)) {
        Show-InfoMessage 'Enter both a Reference User and a Target User.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Resolving users and comparing distribution groups...'
        $script:dgvGroups.Rows.Clear()
        $script:ComparisonRows = @()
        $script:lblReferenceResolved.Text = ''
        $script:lblTargetResolved.Text = ''
        $script:lblMissingCount.Text = 'Missing groups: 0'

        Write-AppLog "Resolving reference user: $($script:txtReference.Text)"
        $script:ReferenceUser = Resolve-ExchangeUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtReference.Text)
        $script:lblReferenceResolved.Text = "$($script:ReferenceUser.DisplayName)  |  $($script:ReferenceUser.PrimarySmtpAddress)"

        Write-AppLog "Resolving target user: $($script:txtTarget.Text)"
        $script:TargetUser = Resolve-ExchangeUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtTarget.Text)
        $script:lblTargetResolved.Text = "$($script:TargetUser.DisplayName)  |  $($script:TargetUser.PrimarySmtpAddress)"

        if ([string]$script:ReferenceUser.ExternalDirectoryObjectId -and
            ([string]$script:ReferenceUser.ExternalDirectoryObjectId -eq [string]$script:TargetUser.ExternalDirectoryObjectId)) {
            throw 'The Reference User and Target User resolve to the same Microsoft 365 account.'
        }

        Write-AppLog "Reading direct distribution-group memberships for $($script:ReferenceUser.PrimarySmtpAddress)."
        $referenceGroups = @(Get-DistributionGroupsForUser -Recipient $script:ReferenceUser)

        Write-AppLog "Reading direct distribution-group memberships for $($script:TargetUser.PrimarySmtpAddress)."
        $targetGroups = @(Get-DistributionGroupsForUser -Recipient $script:TargetUser)

        $targetKeys = @{}
        foreach ($group in $targetGroups) {
            $key = ([string]$group.PrimarySmtpAddress).ToLowerInvariant()
            $targetKeys[$key] = $true
        }

        $missing = foreach ($group in $referenceGroups) {
            $smtp = [string]$group.PrimarySmtpAddress
            $key = $smtp.ToLowerInvariant()
            if (-not $targetKeys.ContainsKey($key)) {
                $owners = @($group.ManagedBy | ForEach-Object { [string]$_ }) -join '; '
                [pscustomobject]@{
                    DisplayName        = [string]$group.DisplayName
                    PrimarySmtpAddress = $smtp
                    ManagedBy          = $owners
                    Identity           = $smtp
                }
            }
        }

        $script:ComparisonRows = @($missing | Sort-Object DisplayName)
        Update-ComparisonGrid

        $script:lblReferenceCount.Text = "Reference distribution groups: $($referenceGroups.Count)"
        $script:lblTargetCount.Text = "Target distribution groups: $($targetGroups.Count)"
        $script:lblStatus.Text = "Comparison complete. $($script:ComparisonRows.Count) missing group(s) found."
        Write-AppLog "Comparison complete: reference=$($referenceGroups.Count), target=$($targetGroups.Count), missing=$($script:ComparisonRows.Count)." 'SUCCESS'

        if ($script:ComparisonRows.Count -eq 0) {
            Show-InfoMessage "$($script:TargetUser.DisplayName) already has all direct distribution-group memberships that $($script:ReferenceUser.DisplayName) has." 'No Missing Groups'
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
    foreach ($gridRow in $script:dgvGroups.Rows) {
        if ([bool]$gridRow.Cells['Selected'].Value) { $selected.Add($gridRow.Tag) }
    }
    return $selected.ToArray()
}

function Add-UserToSelectedGroups {
    if (-not $script:TargetUser) { Show-InfoMessage 'Run a comparison first.'; return }

    $selected = @(Get-SelectedComparisonGroups)
    if ($selected.Count -eq 0) { Show-InfoMessage 'Select at least one distribution group to add.'; return }

    $message = "Add $($script:TargetUser.DisplayName) ($($script:TargetUser.PrimarySmtpAddress)) to $($selected.Count) selected distribution group(s)?`r`n`r`nReference user: $($script:ReferenceUser.DisplayName)"
    $answer = [System.Windows.Forms.MessageBox]::Show($message,'Confirm Distribution Group Changes',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $success = 0
    $failed = 0

    try {
        Set-BusyState -Busy $true -StatusText 'Adding target user to selected distribution groups...'

        foreach ($group in $selected) {
            try {
                Write-AppLog "Adding $($script:TargetUser.PrimarySmtpAddress) to $($group.DisplayName)..."
                Add-DistributionGroupMember `
                    -Identity $group.Identity `
                    -Member ([string]$script:TargetUser.PrimarySmtpAddress) `
                    -BypassSecurityGroupManagerCheck `
                    -Confirm:$false `
                    -ErrorAction Stop

                Write-AppLog "Added target user to $($group.DisplayName)." 'SUCCESS'
                $success++
            }
            catch {
                Write-AppLog "Failed to add target user to $($group.DisplayName): $($_.Exception.Message)" 'ERROR'
                $failed++
            }
        }

        $script:lblStatus.Text = "Membership update complete. Success: $success | Failed: $failed"
        [System.Windows.Forms.MessageBox]::Show(
            "Distribution group update complete.`r`n`r`nSuccessful: $success`r`nFailed: $failed`r`n`r`nSee the Activity Log for details.",
            'Update Complete',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $(if ($failed -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information })
        ) | Out-Null
    }
    finally {
        Set-BusyState -Busy $false
    }

    if ($success -gt 0) { Compare-Users }
}


function Show-DistributionGroupManager {
    if (-not $script:ExchangeConnected) {
        Show-InfoMessage 'Connect to Exchange Online first.'
        return
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Manage User Distribution Groups'
    $dialog.StartPosition = 'CenterParent'
    $dialog.Size = New-Object System.Drawing.Size(980,650)
    $dialog.MinimumSize = New-Object System.Drawing.Size(900,575)
    $dialog.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $dialog.Font = New-Object System.Drawing.Font('Segoe UI',9)

    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Top'
    $header.Height = 68
    $header.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $dialog.Controls.Add($header)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Manage User Distribution Groups'
    $title.AutoSize = $true
    $title.ForeColor = [System.Drawing.Color]::White
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold',17)
    $title.Location = New-Object System.Drawing.Point(18,10)
    $header.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = 'Load one user, select direct distribution-group memberships, and remove selected memberships'
    $subtitle.AutoSize = $true
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
    $subtitle.Location = New-Object System.Drawing.Point(20,40)
    $header.Controls.Add($subtitle)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'User (email, alias or name)'
    $lblUser.AutoSize = $true
    $lblUser.Location = New-Object System.Drawing.Point(20,88)
    $dialog.Controls.Add($lblUser)

    $txtUser = New-Object System.Windows.Forms.ComboBox
    $txtUser.Location = New-Object System.Drawing.Point(20,110)
    $txtUser.Size = New-Object System.Drawing.Size(560,24)
    $txtUser.DropDownStyle = 'DropDown'
    $txtUser.AutoCompleteMode = 'SuggestAppend'
    $txtUser.AutoCompleteSource = 'ListItems'
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
    $dialog.Controls.Add($grid)

    $colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colSelect.Name = 'Selected'
    $colSelect.HeaderText = 'Remove'
    $colSelect.Width = 60
    [void]$grid.Columns.Add($colSelect)

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = 'Group Name'
    $colName.Width = 300
    $colName.ReadOnly = $true
    [void]$grid.Columns.Add($colName)

    $colSmtp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSmtp.HeaderText = 'Primary SMTP'
    $colSmtp.Width = 260
    $colSmtp.ReadOnly = $true
    [void]$grid.Columns.Add($colSmtp)

    $colOwner = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colOwner.HeaderText = 'Managed By'
    $colOwner.AutoSizeMode = 'Fill'
    $colOwner.ReadOnly = $true
    [void]$grid.Columns.Add($colOwner)

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

    $script:ManageDistroUser = $null

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

            $script:ManageDistroUser = Resolve-ExchangeUser -Identity (ConvertFrom-MSToolkitUserPickerLabel -Value $txtUser.Text)
            $lblResolved.Text = "$($script:ManageDistroUser.DisplayName)  |  $($script:ManageDistroUser.PrimarySmtpAddress)"

            $groups = @(Get-DistributionGroupsForUser -Recipient $script:ManageDistroUser)

            foreach ($groupItem in $groups) {
                $owners = @($groupItem.ManagedBy | ForEach-Object { [string]$_ }) -join '; '
                $index = $grid.Rows.Add(
                    $false,
                    [string]$groupItem.DisplayName,
                    [string]$groupItem.PrimarySmtpAddress,
                    $owners
                )
                $grid.Rows[$index].Tag = $groupItem
            }

            $btnRemove.Enabled = ($grid.Rows.Count -gt 0)
        }
        catch {
            Show-ErrorMessage "Unable to load distribution groups.`r`n`r`n$($_.Exception.Message)"
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
            if (-not $row.IsNewRow) { $row.Cells['Selected'].Value = $true }
        }
    })

    $btnClear.Add_Click({
        foreach ($row in $grid.Rows) {
            if (-not $row.IsNewRow) { $row.Cells['Selected'].Value = $false }
        }
    })

    $btnRemove.Add_Click({
        if (-not $script:ManageDistroUser) {
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
            Show-InfoMessage 'Select at least one distribution group to remove.'
            return
        }

        $names = @($selectedGroups | Sort-Object DisplayName | ForEach-Object { $_.DisplayName })
        $nameText = ($names -join "`r`n - ")

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($script:ManageDistroUser.DisplayName) from the following $($selectedGroups.Count) distribution group(s)?`r`n`r`n - $nameText",
            'Confirm Distribution Group Removal',
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
                    Write-AppLog "Removing $($script:ManageDistroUser.PrimarySmtpAddress) from $($groupItem.DisplayName)..."

                    Remove-DistributionGroupMember `
                        -Identity ([string]$groupItem.PrimarySmtpAddress) `
                        -Member ([string]$script:ManageDistroUser.PrimarySmtpAddress) `
                        -BypassSecurityGroupManagerCheck `
                        -Confirm:$false `
                        -ErrorAction Stop

                    Write-AppLog "Removed user from $($groupItem.DisplayName)." 'SUCCESS'
                    $success++
                }
                catch {
                    Write-AppLog "Failed to remove user from $($groupItem.DisplayName): $($_.Exception.Message)" 'ERROR'
                    $failed++
                }

                [System.Windows.Forms.Application]::DoEvents()
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Distribution group removal complete.`r`n`r`nSuccessful: $success`r`nFailed: $failed`r`n`r`nReview the main Activity Log for details.",
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
        $txtUser.Focus()
    })

    Apply-MSToolkitSharedTheme -Root $dialog
    [void]$dialog.ShowDialog($script:MainForm)
    Remove-Variable ManageDistroUser -Scope Script -ErrorAction SilentlyContinue
}

# ------------------------------ GUI ------------------------------
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'Microsoft 365 Distribution Group Compare'
$MainForm.StartPosition = 'CenterScreen'
$MainForm.Size = New-Object System.Drawing.Size(1180,800)
$MainForm.MinimumSize = New-Object System.Drawing.Size(1050,700)
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
$lblTitle.Text = 'Microsoft 365 Distribution Group Compare'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',20)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20,12)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Compare direct distribution-group memberships and add missing memberships to a target user'
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI',9.5)
$lblSubtitle.AutoSize = $true
$lblSubtitle.Location = New-Object System.Drawing.Point(23,49)
$pnlHeader.Controls.Add($lblSubtitle)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Connect to Exchange Online'
$btnConnect.Size = New-Object System.Drawing.Size(205,36)
$btnConnect.Anchor = 'Top,Right'
$btnConnect.Location = New-Object System.Drawing.Point(940,20)
$btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.FlatStyle = 'Flat'
$btnConnect.FlatAppearance.BorderSize = 0
$btnConnect.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$pnlHeader.Controls.Add($btnConnect)
$script:MSToolkitThemeToggleButton = New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "M365DistributionGroupCompare"

# Keep the connection status colour correct after a theme switch. Its design-time
# colour is Firebrick, so the shared theme pass would otherwise always restore red.
$script:MSToolkitThemeRefreshHook = {
    if ($script:lblConnection) {
        if ($script:ExchangeConnected) {
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
$lblConnectionLabel.Text = 'Exchange Online:'
$lblConnectionLabel.AutoSize = $true
$lblConnectionLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$lblConnectionLabel.Location = New-Object System.Drawing.Point(20,10)
$pnlConnection.Controls.Add($lblConnectionLabel)

$lblConnection = New-Object System.Windows.Forms.Label
$lblConnection.Text = 'Not connected'
$lblConnection.AutoSize = $true
$lblConnection.ForeColor = [System.Drawing.Color]::Firebrick
$lblConnection.Location = New-Object System.Drawing.Point(125,10)
$pnlConnection.Controls.Add($lblConnection)
$script:lblConnection = $lblConnection


$btnManageGroups = New-Object System.Windows.Forms.Button
$btnManageGroups.Text = 'Manage User Groups'
$btnManageGroups.Size = New-Object System.Drawing.Size(155,28)
$btnManageGroups.Anchor = 'Top,Right'
$btnManageGroups.Location = New-Object System.Drawing.Point(970,5)
$btnManageGroups.Enabled = $false
$btnManageGroups.Add_Click({ Show-DistributionGroupManager })
$pnlConnection.Controls.Add($btnManageGroups)
$script:btnManageGroups = $btnManageGroups

$grpUsers = New-Object System.Windows.Forms.GroupBox
$grpUsers.Text = 'User Comparison'
$grpUsers.Location = New-Object System.Drawing.Point(20,130)
$grpUsers.Size = New-Object System.Drawing.Size(1125,150)
$grpUsers.Anchor = 'Top,Left,Right'
$grpUsers.BackColor = [System.Drawing.Color]::White
$grpUsers.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9.5)
$MainForm.Controls.Add($grpUsers)

$lblUserFormatHint = New-Object System.Windows.Forms.Label
$lblUserFormatHint.Text = 'Enter an email address, alias or name. Partial names are searched.'
$lblUserFormatHint.AutoSize = $true
$lblUserFormatHint.ForeColor = [System.Drawing.Color]::DimGray
$lblUserFormatHint.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Italic)
$lblUserFormatHint.Location = New-Object System.Drawing.Point(20, 24)
$grpUsers.Controls.Add($lblUserFormatHint)

$lblReference = New-Object System.Windows.Forms.Label
$lblReference.Text = 'Reference User'
$lblReference.AutoSize = $true
$lblReference.Location = New-Object System.Drawing.Point(18,30)
$grpUsers.Controls.Add($lblReference)

$txtReference = New-Object System.Windows.Forms.ComboBox
$txtReference.Location = New-Object System.Drawing.Point(20,54)
$txtReference.Size = New-Object System.Drawing.Size(430,24)
$txtReference.Font = New-Object System.Drawing.Font('Segoe UI',10)
$txtReference.DropDownStyle = 'DropDown'
$txtReference.AutoCompleteMode = 'SuggestAppend'
$txtReference.AutoCompleteSource = 'ListItems'
$txtReference.MaxDropDownItems = 20
$grpUsers.Controls.Add($txtReference)
$script:txtReference = $txtReference

$lblReferenceResolved = New-Object System.Windows.Forms.Label
$lblReferenceResolved.AutoEllipsis = $true
$lblReferenceResolved.Location = New-Object System.Drawing.Point(20,85)
$lblReferenceResolved.Size = New-Object System.Drawing.Size(430,22)
$lblReferenceResolved.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)
$lblReferenceResolved.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$grpUsers.Controls.Add($lblReferenceResolved)
$script:lblReferenceResolved = $lblReferenceResolved

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = 'Target User'
$lblTarget.AutoSize = $true
$lblTarget.Location = New-Object System.Drawing.Point(480,30)
$grpUsers.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.ComboBox
$txtTarget.Location = New-Object System.Drawing.Point(482,54)
$txtTarget.Size = New-Object System.Drawing.Size(430,24)
$txtTarget.Font = New-Object System.Drawing.Font('Segoe UI',10)
$txtTarget.DropDownStyle = 'DropDown'
$txtTarget.AutoCompleteMode = 'SuggestAppend'
$txtTarget.AutoCompleteSource = 'ListItems'
$txtTarget.MaxDropDownItems = 20
$grpUsers.Controls.Add($txtTarget)
$script:txtTarget = $txtTarget

$lblTargetResolved = New-Object System.Windows.Forms.Label
$lblTargetResolved.AutoEllipsis = $true
$lblTargetResolved.Location = New-Object System.Drawing.Point(482,85)
$lblTargetResolved.Size = New-Object System.Drawing.Size(430,22)
$lblTargetResolved.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)
$lblTargetResolved.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$grpUsers.Controls.Add($lblTargetResolved)
$script:lblTargetResolved = $lblTargetResolved

$btnCompare = New-Object System.Windows.Forms.Button
$btnCompare.Text = 'Compare Users'
$btnCompare.Size = New-Object System.Drawing.Size(165,42)
$btnCompare.Location = New-Object System.Drawing.Point(940,48)
$btnCompare.Anchor = 'Top,Right'
$btnCompare.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$btnCompare.ForeColor = [System.Drawing.Color]::White
$btnCompare.FlatStyle = 'Flat'
$btnCompare.FlatAppearance.BorderSize = 0
$btnCompare.Font = New-Object System.Drawing.Font('Segoe UI Semibold',10)
$btnCompare.Enabled = $false
$grpUsers.Controls.Add($btnCompare)
$script:btnCompare = $btnCompare

$lblReferenceCount = New-Object System.Windows.Forms.Label
$lblReferenceCount.Text = 'Reference distribution groups: 0'
$lblReferenceCount.AutoSize = $true
$lblReferenceCount.Location = New-Object System.Drawing.Point(20,118)
$lblReferenceCount.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$grpUsers.Controls.Add($lblReferenceCount)
$script:lblReferenceCount = $lblReferenceCount

$lblTargetCount = New-Object System.Windows.Forms.Label
$lblTargetCount.Text = 'Target distribution groups: 0'
$lblTargetCount.AutoSize = $true
$lblTargetCount.Location = New-Object System.Drawing.Point(482,118)
$lblTargetCount.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$grpUsers.Controls.Add($lblTargetCount)
$script:lblTargetCount = $lblTargetCount

$pnlResultsToolbar = New-Object System.Windows.Forms.Panel
$pnlResultsToolbar.Location = New-Object System.Drawing.Point(20,292)
$pnlResultsToolbar.Size = New-Object System.Drawing.Size(1125,46)
$pnlResultsToolbar.Anchor = 'Top,Left,Right'
$pnlResultsToolbar.BackColor = [System.Drawing.Color]::White
$MainForm.Controls.Add($pnlResultsToolbar)

$lblMissingCount = New-Object System.Windows.Forms.Label
$lblMissingCount.Text = 'Missing groups: 0'
$lblMissingCount.Font = New-Object System.Drawing.Font('Segoe UI Semibold',10)
$lblMissingCount.AutoSize = $true
$lblMissingCount.Location = New-Object System.Drawing.Point(14,14)
$pnlResultsToolbar.Controls.Add($lblMissingCount)
$script:lblMissingCount = $lblMissingCount

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = 'Select All'
$btnSelectAll.Size = New-Object System.Drawing.Size(95,30)
$btnSelectAll.Anchor = 'Top,Right'
$btnSelectAll.Location = New-Object System.Drawing.Point(760,8)
$btnSelectAll.Enabled = $false
$pnlResultsToolbar.Controls.Add($btnSelectAll)
$script:btnSelectAll = $btnSelectAll

$btnClearSelection = New-Object System.Windows.Forms.Button
$btnClearSelection.Text = 'Clear'
$btnClearSelection.Size = New-Object System.Drawing.Size(80,30)
$btnClearSelection.Anchor = 'Top,Right'
$btnClearSelection.Location = New-Object System.Drawing.Point(865,8)
$btnClearSelection.Enabled = $false
$pnlResultsToolbar.Controls.Add($btnClearSelection)
$script:btnClearSelection = $btnClearSelection

$btnAddSelected = New-Object System.Windows.Forms.Button
$btnAddSelected.Text = 'Add Selected Groups'
$btnAddSelected.Size = New-Object System.Drawing.Size(165,30)
$btnAddSelected.Anchor = 'Top,Right'
$btnAddSelected.Location = New-Object System.Drawing.Point(950,8)
$btnAddSelected.BackColor = [System.Drawing.Color]::FromArgb(22,130,80)
$btnAddSelected.ForeColor = [System.Drawing.Color]::White
$btnAddSelected.FlatStyle = 'Flat'
$btnAddSelected.FlatAppearance.BorderSize = 0
$btnAddSelected.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$btnAddSelected.Enabled = $false
$pnlResultsToolbar.Controls.Add($btnAddSelected)
$script:btnAddSelected = $btnAddSelected

$dgvGroups = New-Object System.Windows.Forms.DataGridView
$dgvGroups.Location = New-Object System.Drawing.Point(20,342)
$dgvGroups.Size = New-Object System.Drawing.Size(1125,285)
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
$dgvGroups.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$dgvGroups.EnableHeadersVisualStyles = $false
$dgvGroups.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232,237,243)
$dgvGroups.ColumnHeadersHeight = 34
$MainForm.Controls.Add($dgvGroups)
$script:dgvGroups = $dgvGroups

$colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelect.Name = 'Selected'
$colSelect.HeaderText = 'Add'
$colSelect.Width = 50
$dgvGroups.Columns.Add($colSelect) | Out-Null

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = 'GroupName'
$colName.HeaderText = 'Group Name'
$colName.Width = 280
$colName.ReadOnly = $true
$dgvGroups.Columns.Add($colName) | Out-Null

$colSmtp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSmtp.Name = 'PrimarySmtpAddress'
$colSmtp.HeaderText = 'Primary SMTP Address'
$colSmtp.Width = 285
$colSmtp.ReadOnly = $true
$dgvGroups.Columns.Add($colSmtp) | Out-Null

$colOwner = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOwner.Name = 'ManagedBy'
$colOwner.HeaderText = 'Managed By'
$colOwner.AutoSizeMode = 'Fill'
$colOwner.ReadOnly = $true
$dgvGroups.Columns.Add($colOwner) | Out-Null

$colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colType.Name = 'Type'
$colType.HeaderText = 'Type'
$colType.Width = 135
$colType.ReadOnly = $true
$dgvGroups.Columns.Add($colType) | Out-Null

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = 'TargetStatus'
$colStatus.HeaderText = 'Target Status'
$colStatus.Width = 100
$colStatus.ReadOnly = $true
$dgvGroups.Columns.Add($colStatus) | Out-Null

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Activity Log'
$grpLog.Location = New-Object System.Drawing.Point(20,640)
$grpLog.Size = New-Object System.Drawing.Size(1125,100)
$grpLog.Anchor = 'Bottom,Left,Right'
$grpLog.BackColor = [System.Drawing.Color]::White
$MainForm.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(10,20)
$txtLog.Size = New-Object System.Drawing.Size(1105,70)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.ReadOnly = $true
$txtLog.BorderStyle = 'None'
$txtLog.Font = New-Object System.Drawing.Font('Consolas',8.5)
$txtLog.BackColor = [System.Drawing.Color]::White
$grpLog.Controls.Add($txtLog)
$script:txtLog = $txtLog

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = 'Ready. Connect to Exchange Online to begin.'
$lblStatus.Spring = $true
$lblStatus.TextAlign = 'MiddleLeft'
$statusStrip.Items.Add($lblStatus) | Out-Null
$MainForm.Controls.Add($statusStrip)
$script:lblStatus = $lblStatus

$btnConnect.Add_Click({ Connect-M365Exchange })
$btnCompare.Add_Click({ Compare-Users })
$btnAddSelected.Add_Click({ Add-UserToSelectedGroups })

$btnSelectAll.Add_Click({
    foreach ($row in $script:dgvGroups.Rows) { $row.Cells['Selected'].Value = $true }
})

$btnClearSelection.Add_Click({
    foreach ($row in $script:dgvGroups.Rows) { $row.Cells['Selected'].Value = $false }
})

# Enter is used to accept an entry from the user list, so it no longer starts a
# comparison. Use the Compare button instead.
$txtReference.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $_.SuppressKeyPress = $true } })
$txtTarget.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $_.SuppressKeyPress = $true } })

$MainForm.Add_FormClosing({
    if ($script:ExchangeConnected) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
})

Apply-MSToolkitSharedTheme -Root $MainForm

Write-AppLog 'Distribution Group Compare initialized.'
Hide-PowerShellConsole
[void]$MainForm.ShowDialog()
