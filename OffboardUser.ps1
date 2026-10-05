<#
.SYNOPSIS
    Windows Forms offboarding checklist for a departing employee.

.DESCRIPTION
    Walks an offboarding process in order, with the Active Directory steps run
    from this window and the rest linked out to the system that owns them.
    Everything organisation-specific - the sync server, MFA and access systems,
    line-of-business systems, documents and export folder - comes from the
    Offboarding section of MSToolkit Settings, and a step whose setting is blank
    is left out.

    Steps that this tool performs - disabling the account, clearing the phone
    number, replicating, and the Entra delta sync - tick themselves off only
    when they actually succeed. The steps done elsewhere are ticked by hand.

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

# Bringing a browser or RDP window to the front needs these three calls; one
# compile at startup, as in MSToolkit.
if (-not ("MSToolkitWindow" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public class MSToolkitWindow
{
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    // Keeps typed text clear of the show/hide icon inside the password field.
    public static void SetRightMargin(IntPtr handle, int margin) {
        const int EM_SETMARGINS = 0xD3;
        const int EC_RIGHTMARGIN = 0x2;
        SendMessage(handle, EM_SETMARGINS, (IntPtr)EC_RIGHTMARGIN, (IntPtr)(margin << 16));
    }
}
"@
}


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

    # The sub-step rail, the red warning lines and the loaded-user line keep
    # their own colours.
    if ("$($Control.Tag)" -eq "SubStepRail") { return }
    if ("$($Control.Tag)" -eq "StepWarning") { return }
    if ("$($Control.Tag)" -eq "LoadedUser") { return }

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

# Offboarding settings, read once. Every one is optional.
$script:OffboardSyncServer      = Get-MSToolkitSetting -Name "EntraSyncServer"
$script:OffboardMfaName         = Get-MSToolkitSetting -Name "OffboardMfaName"
$script:OffboardMfaUrl          = Get-MSToolkitSetting -Name "OffboardMfaUrl"
$script:OffboardAccessName      = Get-MSToolkitSetting -Name "OffboardAccessName"
$script:OffboardAccessHost      = Get-MSToolkitSetting -Name "OffboardAccessHost"
$script:OffboardRmmName         = Get-MSToolkitSetting -Name "OffboardRmmName"
$script:OffboardRmmReport       = Get-MSToolkitSetting -Name "OffboardRmmReport"
$script:OffboardArchiveName     = Get-MSToolkitSetting -Name "OffboardArchiveName"
$script:OffboardLobSystems      = Get-MSToolkitSetting -Name "OffboardLobSystems"
$script:OffboardDeviceDocument  = Get-MSToolkitSetting -Name "OffboardDeviceDocument"
$script:OffboardChecklistDocument = Get-MSToolkitSetting -Name "OffboardChecklistDocument"
$script:OffboardChecklistLocation = Get-MSToolkitSetting -Name "OffboardChecklistLocation"
$script:OffboardDisabledOU      = Get-MSToolkitSetting -Name "OUDisabled"

# Where group exports go: the setting, or MSToolkit's shared Logs folder, which
# the signed-in user can also open with the Logs button.
$script:OffboardExportFolder = Get-MSToolkitSetting -Name "OffboardExportFolder" -Default (Join-Path $env:ProgramData "MSToolkit\Logs")

function Get-OffboardLobList {
    # One line per system: Name | optional URL | optional note
    $List = @()
    foreach ($Line in ($script:OffboardLobSystems -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($Line)) { continue }
        $Parts = @($Line -split '\|' | ForEach-Object { $_.Trim() })
        if (-not $Parts[0]) { continue }
        $List += ,([pscustomobject]@{
            Name = $Parts[0]
            Url  = if ($Parts.Count -gt 1) { $Parts[1] } else { '' }
            Note = if ($Parts.Count -gt 2) { ($Parts[2..($Parts.Count - 1)] -join ' | ') } else { '' }
        })
    }
    return ,$List
}

function Get-OffboardDocumentLabel {
    # A friendly name for a configured document: the file name without its
    # extension, or "the document" for a web address.
    param([string]$Value)
    if ($Value -match '^https?://') { return 'the document' }
    return [System.IO.Path]::GetFileNameWithoutExtension($Value)
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

function Get-MSToolkitEmployeeUserList {
    param([string]$Server)

    $ServerArgs = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $ServerArgs['Server'] = $Server
    }

    $DomainDN = (Get-ADDomain @ServerArgs -ErrorAction Stop).DistinguishedName

    # Employee accounts plus the admin accounts, which are often the reference
    # user when comparing group membership.
    # Users OU plus the Admin accounts OU from MSToolkit Settings: a blank Users
    # OU searches the whole domain, and a blank Admin accounts OU adds nothing.
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

# Filtered user pickers: see Register-MSToolkitComboFilter for why neither Windows
# autocomplete nor the real dropdown is used here.
$script:ComboFilterItems = @{}
$script:ComboFilterRegistered = @{}
$script:ComboFilterLists = @{}
$script:ComboFilterOwners = @{}
$script:ComboFilterBusy = $false

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

        # The master list is kept so the box can be filtered as it is typed in.
        # Windows autocomplete is NOT used: it opens its own suggestion popup on
        # top of the dropdown list, and a click then lands on whatever entry sits
        # behind the suggestion - which is how the wrong user gets selected.

        foreach ($Combo in $Combos) {
            if (-not $Combo) { continue }

            $Existing = $Combo.Text
            $Combo.Items.Clear()
            foreach ($User in $Users) { $null = $Combo.Items.Add((Get-MSToolkitUserPickerLabel -ADUser $User)) }
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

    try {
        if ($script:MainForm) {
            $script:MainForm.UseWaitCursor = $Busy

            # A ComboBox owns its own window handle, so clearing UseWaitCursor on
            # the form does not always restore the pointer over it.
            if (-not $Busy) {
                [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
                $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }

        # Walked by index over a plain array. This is only cosmetic, so it is
        # wrapped as well: a failure here must never take the operation with it.
        $Buttons = $script:StepButtons

        if ($null -ne $Buttons) {
            for ($Index = 0; $Index -lt $Buttons.Count; $Index++) {
                $Button = $Buttons[$Index]
                if ($Button -is [System.Windows.Forms.Control]) { $Button.Enabled = -not $Busy }
            }
        }

        if ($script:lblStatus -and $StatusText) {
            $script:lblStatus.Text = $StatusText
        }

        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {
        # Deliberately silent: the caller is mid-operation and this is chrome.
    }
}

# ---------------------------------------------------------------------------
# Offboarding actions
#
# The four Active Directory steps run from here. Each ticks its own checkbox
# only after the operation actually succeeds, so a ticked box means the work
# was done, not that somebody clicked past it.
# ---------------------------------------------------------------------------
$script:WorkingServer = $null
$script:CurrentUser = $null
# Plain arrays on purpose: wrapping a generic List in @() throws
# "Argument types do not match" in PowerShell 5.1, and an exception raised in a
# WinForms event handler is swallowed silently - the button just does nothing.
$script:StepChecks = @()
$script:StepButtons = @()
$script:StepToolTip = New-Object System.Windows.Forms.ToolTip
$script:StepToolTip.AutoPopDelay = 15000
$script:StepToolTip.InitialDelay = 400
$script:StepToolTip.ReshowDelay = 100

function Get-MSToolkitCriticalUserReason {
    # Ported from MSToolkit: identify the accounts that must never be offboarded
    # by SID and RID rather than by name, since any of them can be renamed.
    param($User)

    if ($null -eq $User) { return $null }

    $Rid = Get-MSToolkitRidFromSid -Sid $User.SID

    switch ($Rid) {
        500 { return "Built-in Administrator account (RID 500)" }
        501 { return "Built-in Guest account (RID 501)" }
        502 { return "Kerberos ticket-granting account krbtgt (RID 502)" }
    }

    if ($User.PSObject.Properties.Name -contains "isCriticalSystemObject" -and $User.isCriticalSystemObject) {
        return "Flagged as a critical system object in Active Directory"
    }

    return $null
}

function Update-StepProgress {
    $Checks = $script:StepChecks
    $Total = 0
    $Done = 0

    if ($null -ne $Checks) {
        for ($Index = 0; $Index -lt $Checks.Count; $Index++) {
            $Total++
            if ($Checks[$Index].Checked) { $Done++ }
        }
    }

    $script:lblProgress.Text = "$Done of $Total step(s) complete"

    if ($Done -eq $Total -and $Total -gt 0) {
        $script:lblProgress.ForeColor = (Get-MSToolkitThemePalette).Success
    }
    else {
        $script:lblProgress.ForeColor = (Get-MSToolkitThemePalette).Text
    }
}

function Get-MSToolkitDefaultBrowserCommand {
    # ShellExecute (Start-Process <url>) hands the request to the desktop shell,
    # which belongs to the signed-in user and silently refuses this window when it
    # runs as the admin account. So resolve the browser and start the executable
    # directly instead.
    try {
        $ChoiceKey = "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice"

        if (Test-Path -LiteralPath $ChoiceKey) {
            $ProgId = (Get-ItemProperty -LiteralPath $ChoiceKey -ErrorAction Stop).ProgId

            if ($ProgId) {
                $CommandKey = "Registry::HKEY_CLASSES_ROOT\$ProgId\shell\open\command"

                if (Test-Path -LiteralPath $CommandKey) {
                    $Command = (Get-ItemProperty -LiteralPath $CommandKey -ErrorAction Stop).'(default)'

                    # Typically: "C:\Path\browser.exe" --single-argument %1
                    if ($Command -match '^\s*"([^"]+)"') { return $Matches[1] }
                    if ($Command -match '^\s*(\S+\.exe)') { return $Matches[1] }
                }
            }
        }
    }
    catch { }

    # Fall back to whichever common browser is actually installed.
    $Candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"),
        (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
        (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
        (Join-Path $env:ProgramFiles "Mozilla Firefox\firefox.exe")
    )

    foreach ($Candidate in $Candidates) {
        if ($Candidate -and (Test-Path -LiteralPath $Candidate)) { return $Candidate }
    }

    return $null
}

function Show-MSToolkitWindowForProcess {
    # A browser that is already running usually hands the URL to its existing
    # window and exits, so the new process has no window of its own. Look at the
    # process we started first, then fall back to any window owned by a process
    # with the same name.
    param(
        [System.Diagnostics.Process]$Process,
        [string]$ProcessName
    )

    $Deadline = (Get-Date).AddSeconds(6)

    while ((Get-Date) -lt $Deadline) {
        $Handle = [IntPtr]::Zero

        if ($Process -and -not $Process.HasExited) {
            $Process.Refresh()
            if ($Process.MainWindowHandle -ne [IntPtr]::Zero) { $Handle = $Process.MainWindowHandle }
        }

        if ($Handle -eq [IntPtr]::Zero -and $ProcessName) {
            $Existing = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
                Sort-Object StartTime -Descending | Select-Object -First 1

            if ($Existing) { $Handle = $Existing.MainWindowHandle }
        }

        if ($Handle -ne [IntPtr]::Zero) {
            if ([MSToolkitWindow]::IsIconic($Handle)) {
                [void][MSToolkitWindow]::ShowWindow($Handle, 9)   # SW_RESTORE
            }

            [void][MSToolkitWindow]::SetForegroundWindow($Handle)
            return $true
        }

        Start-Sleep -Milliseconds 250
        [System.Windows.Forms.Application]::DoEvents()
    }

    return $false
}

function Open-OffboardLink {
    param([string]$Url, [string]$Label)

    $Browser = Get-MSToolkitDefaultBrowserCommand

    if (-not $Browser) {
        [System.Windows.Forms.Clipboard]::SetText($Url)
        Write-AppLog "Could not find a browser to open $Label. The address was copied to the clipboard instead." 'WARNING'
        return
    }

    try {
        $Process = Start-Process -FilePath $Browser -ArgumentList $Url -PassThru -ErrorAction Stop
        $Name = [System.IO.Path]::GetFileNameWithoutExtension($Browser)

        Write-AppLog "Opened $Label in $Name."

        if (-not (Show-MSToolkitWindowForProcess -Process $Process -ProcessName $Name)) {
            Write-AppLog "$Label was opened, but the browser window could not be brought to the front - check the taskbar." 'WARNING'
        }
    }
    catch {
        [System.Windows.Forms.Clipboard]::SetText($Url)
        Write-AppLog "Could not open $($Label): $($_.Exception.Message). The address was copied to the clipboard instead." 'WARNING'
    }
}

function Open-OffboardAccessSystem {
    # The physical access system from MSToolkit Settings: a web address opens in
    # the browser, anything else is treated as a host for Remote Desktop.
    $Target = $script:OffboardAccessHost
    $Label = if ($script:OffboardAccessName) { $script:OffboardAccessName } else { 'the access control system' }

    if ($Target -match '^https?://') {
        Open-OffboardLink -Url $Target -Label $Label
        return
    }

    try {
        $Process = Start-Process "mstsc.exe" -ArgumentList "/v:$Target" -PassThru -ErrorAction Stop
        Write-AppLog "Opened a Remote Desktop connection to $Target for $Label."
        [void](Show-MSToolkitWindowForProcess -Process $Process -ProcessName "mstsc")
    }
    catch {
        Write-AppLog "Could not start Remote Desktop: $($_.Exception.Message)" 'ERROR'
    }
}

function Resolve-OffboardUser {
    $Identity = ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboUser.Text

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        Show-InfoMessage 'Enter the user who is leaving first.'
        return $null
    }

    try {
        $User = Get-ADUser `
            -Identity $Identity `
            -Server $script:WorkingServer `
            -Properties SID,Enabled,DisplayName,Description,Title,Department,Manager,`
                        telephoneNumber,mobile,DistinguishedName,isCriticalSystemObject `
            -ErrorAction Stop

        return $User
    }
    catch {
        Show-ErrorMessage "Could not find '$Identity' in Active Directory.`r`n`r`n$($_.Exception.Message)"
        Write-AppLog "User lookup failed for '$Identity': $($_.Exception.Message)" 'ERROR'
        return $null
    }
}

function Invoke-LoadOffboardUser {
    # Close the dropdown first: while it is open the text can still change under
    # the pointer, and the list sits over the buttons below the box.
    if ($script:cboUser.DroppedDown) { $script:cboUser.DroppedDown = $false }

    # Anything already loaded is dropped FIRST. A failed lookup used to return
    # early and leave the previous user in place while the panel still showed
    # them, so the next action ran against the wrong account.
    $Previous = $script:CurrentUser
    $script:CurrentUser = $null
    $script:ManagerDisplayName = $null
    $script:lblUserDetail.Text = ""
    $script:lblUserDetail.ForeColor = (Get-MSToolkitThemePalette).Text

    $User = Resolve-OffboardUser

    if (-not $User) {
        if ($Previous) {
            Write-AppLog "No user is loaded now - $($Previous.SamAccountName) was cleared because the new lookup failed." 'WARNING'
        }

        $script:lblStatus.Text = 'No user loaded.'
        return
    }

    $Critical = Get-MSToolkitCriticalUserReason -User $User

    if ($Critical) {
        Show-ErrorMessage "OPERATION BLOCKED`r`n`r`nUser: $($User.Name)`r`nReason: $Critical`r`n`r`nThis account cannot be offboarded with this tool."
        Write-AppLog "BLOCKED: $($User.Name) - $Critical" 'ERROR'
        return
    }

    # Switching to a different person mid-checklist: the ticks belong to whoever
    # was loaded, so say so rather than carrying them over silently.
    if ($Previous -and $Previous.DistinguishedName -ne $User.DistinguishedName) {
        Write-AppLog "Changed user: $($Previous.SamAccountName) -> $($User.SamAccountName)." 'WARNING'

        $Ticked = 0
        for ($Index = 0; $Index -lt $script:StepChecks.Count; $Index++) {
            if ($script:StepChecks[$Index].Checked) { $Ticked++ }
        }

        if ($Ticked -gt 0) {
            $Answer = [System.Windows.Forms.MessageBox]::Show(
                ("You have switched from $($Previous.SamAccountName) to $($User.SamAccountName), and $Ticked step(s) are ticked.`r`n`r`n" +
                 "Those ticks refer to $($Previous.SamAccountName). Clear them for the new user?`r`n`r`n" +
                 "Yes - start a clean checklist for $($User.SamAccountName)`r`nNo - keep the ticks as they are"),
                "Different user loaded",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button1)

            if ($Answer -eq [System.Windows.Forms.DialogResult]::Yes) {
                for ($Index = 0; $Index -lt $script:StepChecks.Count; $Index++) {
                    $Check = $script:StepChecks[$Index]
                    $Check.Checked = $false

                    $Subs = $Check.Tag
                    if ($null -ne $Subs) {
                        for ($SubIndex = 0; $SubIndex -lt $Subs.Count; $SubIndex++) { $Subs[$SubIndex].Checked = $false }
                    }
                }

                $script:GroupsExportPath = $null
                Write-AppLog 'Checklist reset for the new user.'
            }
        }
    }

    $Phone = @($User.telephoneNumber, $User.mobile | Where-Object { $_ }) -join " / "
    if (-not $Phone) { $Phone = "(none set)" }

    $ManagerName = "(none set)"
    if ($User.Manager) {
        try {
            $ManagerName = (Get-ADUser -Identity $User.Manager -Server $script:WorkingServer -ErrorAction Stop).Name
        }
        catch { $ManagerName = $User.Manager }
    }

    # Read back who was actually matched. The picker can hand over a different
    # person than the one that was typed, so this is the last chance to notice
    # before the destructive steps become available.
    $Answer = [System.Windows.Forms.MessageBox]::Show(
        ("Offboard this person?`r`n`r`n" +
         "$($User.DisplayName)`r`n" +
         "Username: $($User.SamAccountName)`r`n" +
         "Title: $($User.Title)`r`n" +
         "Department: $($User.Department)`r`n" +
         "Manager: $ManagerName`r`n" +
         "Enabled: $($User.Enabled)`r`n" +
         "Phone: $Phone`r`n`r`n" +
         "$($User.DistinguishedName)"),
        "Confirm the user",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog "Not the right person - $($User.SamAccountName) was not loaded." 'WARNING'
        $script:lblStatus.Text = 'No user loaded.'
        return
    }

    $script:CurrentUser = $User

    # Kept for the checklist header, which records the manager like the paper one.
    $script:ManagerDisplayName = $ManagerName

    $State = "Enabled"
    if (-not $User.Enabled) { $State = "ALREADY DISABLED" }

    $script:lblUserDetail.ForeColor = (Get-MSToolkitThemePalette).Success
    $script:lblUserDetail.Text = "LOADED: $($User.DisplayName)  |  $($User.SamAccountName)  |  $State`r`n" +
                                 "Title: $($User.Title)   Department: $($User.Department)   Manager: $ManagerName`r`n" +
                                 "Phone: $Phone`r`n" +
                                 "$($User.DistinguishedName)"

    $script:lblStatus.Text = "Loaded: $($User.DisplayName) ($($User.SamAccountName))"
    Write-AppLog "Loaded $($User.SamAccountName) for offboarding." 'SUCCESS'
    Write-AppLog "Manager to contact about the mailbox: $ManagerName"

    if (-not $User.Enabled) {
        Write-AppLog 'This account is already disabled - tick the disable step by hand if it was disabled earlier.' 'WARNING'
    }
}

function Test-OffboardSelectionMatches {
    # The box and the loaded account must agree before anything is changed. The
    # picker can be changed without clicking Load, and a failed Load leaves the
    # box reading one name while nothing - or somebody else - is loaded.
    param([string]$Operation)

    if (-not $script:CurrentUser) {
        Show-InfoMessage "Load the user who is leaving first."
        return $false
    }

    $Typed = (ConvertFrom-MSToolkitUserPickerLabel -Value $script:cboUser.Text).Trim()

    if ([string]::IsNullOrWhiteSpace($Typed)) { return $true }

    $User = $script:CurrentUser
    $Known = @($User.SamAccountName, $User.UserPrincipalName, $User.DisplayName, $User.Name, $User.DistinguishedName) |
        Where-Object { $_ }

    foreach ($Candidate in $Known) {
        if ("$Candidate" -ieq $Typed) { return $true }
    }

    Show-ErrorMessage ("The box says '$Typed' but the loaded account is " +
        "$($User.DisplayName) ($($User.SamAccountName)).`r`n`r`n" +
        "$Operation has NOT been run. Click Load User to load '$Typed', then try again.")

    Write-AppLog "Refused $($Operation): the picker says '$Typed' but $($User.SamAccountName) is loaded." 'ERROR'
    return $false
}

function Invoke-DisableOffboardUser {
    param([switch]$NoConfirm)

    if (-not (Test-OffboardSelectionMatches -Operation 'Disable account')) { return }

    $User = $script:CurrentUser

    $Answer = [System.Windows.Forms.DialogResult]::Yes

    if (-not $NoConfirm) {
    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Disable this Active Directory account?`r`n`r`n" +
        "Name: $($User.DisplayName)`r`n" +
        "Username: $($User.SamAccountName)`r`n" +
        "Enabled: $($User.Enabled)`r`n" +
        "Description: $($User.Description)`r`n" +
        "$($User.DistinguishedName)`r`n`r`n" +
        "They will not be able to sign in to anything that authenticates against AD. The account is not moved or deleted.",
        "Confirm Disable User",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    }

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog 'Disable cancelled.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText "Disabling $($User.SamAccountName)..."

        # Logged before the call so a failure shows exactly what was attempted,
        # against which DC, rather than leaving a silent dialog behind.
        Write-AppLog "Disabling $($User.DistinguishedName) on $($script:WorkingServer)..."

        if ([string]::IsNullOrWhiteSpace($script:WorkingServer)) {
            throw "No domain controller is selected. Restart the tool from MSToolkit."
        }

        Disable-ADAccount -Identity $User.DistinguishedName -Server $script:WorkingServer -Confirm:$false -ErrorAction Stop

        # Confirm from the directory rather than assuming the call worked.
        $Refreshed = Get-ADUser -Identity $User.DistinguishedName -Server $script:WorkingServer `
            -Properties SID,Enabled,DisplayName,Description,Title,Department,Manager,telephoneNumber,mobile,DistinguishedName,isCriticalSystemObject `
            -ErrorAction Stop

        $script:CurrentUser = $Refreshed

        if ($Refreshed.Enabled) {
            Write-AppLog "The account still reports Enabled on $($script:WorkingServer) after the disable. Check for a replication or permissions problem." 'ERROR'
            Show-ErrorMessage "The disable command ran, but $($User.SamAccountName) still reports as enabled on $($script:WorkingServer)."
            Set-BusyState -Busy $false -StatusText 'Disable did not take effect.'
            return
        }

        Add-OffboardUndoAction -User $User -Type 'DisableAccount' -Details ([ordered]@{
            PreviousEnabled = [bool]$User.Enabled
        })

        if ($script:chkDisable) { $script:chkDisable.Checked = $true }

        Write-AppLog "Disabled $($User.SamAccountName) on $($script:WorkingServer). Confirmed: Enabled = $($Refreshed.Enabled)." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText "Disabled $($User.SamAccountName)."
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Disable failed.'
        Write-AppLog "Disable failed: $($_.Exception.GetType().FullName)" 'ERROR'
        Write-AppLog "  $($_.Exception.Message)" 'ERROR'
        if ($_.InvocationInfo) { Write-AppLog "  at line $($_.InvocationInfo.ScriptLineNumber): $("$($_.InvocationInfo.Line)".Trim())" 'ERROR' }
        Show-ErrorMessage "Could not disable the account.`r`n`r`n$($_.Exception.Message)"
    }
}

function Invoke-ClearOffboardPhone {
    param([switch]$NoConfirm)

    if (-not (Test-OffboardSelectionMatches -Operation 'Clear phone number')) { return }

    $User = $script:CurrentUser
    $Phone = @($User.telephoneNumber, $User.mobile | Where-Object { $_ })

    if ($Phone.Count -eq 0) {
        Write-AppLog "$($User.SamAccountName) has no phone number set - nothing to clear." 'SUCCESS'
        $script:chkPhone.Checked = $true
        return
    }

    $Answer = [System.Windows.Forms.DialogResult]::Yes

    if (-not $NoConfirm) {
    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Clear the phone number(s) from this account?`r`n`r`n" +
        "Name: $($User.DisplayName)`r`n" +
        "Username: $($User.SamAccountName)`r`n" +
        "Telephone: $($User.telephoneNumber)`r`n" +
        "Mobile: $($User.mobile)`r`n`r`n" +
        "This removes the numbers from Active Directory so the desk phone and mobile are not published against a departed employee.",
        "Confirm Clear Phone Number",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    }

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog 'Clear phone cancelled.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText 'Clearing the phone number...'

        $Clear = @()
        if ($User.telephoneNumber) { $Clear += 'telephoneNumber' }
        if ($User.mobile) { $Clear += 'mobile' }

        Write-AppLog "Clearing $($Clear -join ' and ') on $($User.DistinguishedName)..."
        Set-ADUser -Identity $User.DistinguishedName -Server $script:WorkingServer -Clear $Clear -Confirm:$false -ErrorAction Stop

        $script:CurrentUser = Get-ADUser -Identity $User.DistinguishedName -Server $script:WorkingServer `
            -Properties SID,Enabled,DisplayName,Description,Title,Department,Manager,telephoneNumber,mobile,DistinguishedName,isCriticalSystemObject

        Add-OffboardUndoAction -User $User -Type 'ClearPhone' -Details ([ordered]@{
            telephoneNumber = [string]$User.telephoneNumber
            mobile          = [string]$User.mobile
        })

        if ($script:chkPhone) { $script:chkPhone.Checked = $true }
        Write-AppLog "Cleared $($Clear -join ' and ') on $($User.SamAccountName)." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText 'Phone number cleared.'
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Clear phone failed.'
        Write-AppLog "Clear phone failed: $($_.Exception.GetType().FullName)" 'ERROR'
        Write-AppLog "  $($_.Exception.Message)" 'ERROR'
        Show-ErrorMessage "Could not clear the phone number.`r`n`r`n$($_.Exception.Message)"
    }
}

function Invoke-OffboardReplicate {
    # Replicates the DC this tool is working against. /AdeP is enterprise-wide, so
    # one DC still pushes the change out across the forest - looping every DC was
    # slower without adding anything.
    try {
        if ([string]::IsNullOrWhiteSpace($script:WorkingServer)) {
            Write-AppLog 'No domain controller is selected.' 'ERROR'
            Set-BusyState -Busy $false -StatusText 'Replication could not start.'
            return
        }

        $Domain = Get-MSToolkitLocalDomainName
        $Target = $script:WorkingServer
        if ($Target -notlike '*.*' -and $Domain) { $Target = "$Target.$Domain" }

        Set-BusyState -Busy $true -StatusText "Replicating $Target..."
        Write-AppLog "repadmin /syncall $Target /AdeP ..."

        $global:LASTEXITCODE = 0

        repadmin /syncall "$Target" /AdeP 2>&1 | ForEach-Object {
            $Line = ($_ | Out-String).Trim()
            if (-not [string]::IsNullOrWhiteSpace($Line)) { Write-AppLog "  $Line" }
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($LASTEXITCODE -eq 0) {
            if ($script:chkReplicate) { $script:chkReplicate.Checked = $true }
            Write-AppLog "Replication completed for $Target." 'SUCCESS'
            Set-BusyState -Busy $false -StatusText "Replication complete on $Target."
        }
        else {
            Write-AppLog "Replication for $Target reported errors. repadmin exit code: $LASTEXITCODE" 'ERROR'
            Set-BusyState -Busy $false -StatusText 'Replication reported errors.'
        }
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Replication failed.'
        Write-AppLog "Replication failed: $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-OffboardDeltaSync {
    $EntraSyncServer = $script:OffboardSyncServer

    if (-not $EntraSyncServer) {
        Write-AppLog 'No Entra Connect server is set in MSToolkit Settings, so no delta sync was requested.' 'WARNING'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText "Starting the Entra Connect delta sync on $EntraSyncServer..."
        Write-AppLog "Starting Microsoft Entra Connect delta sync on $EntraSyncServer..."

        $Result = Invoke-Command -ComputerName $EntraSyncServer -ScriptBlock {
            Import-Module ADSync -ErrorAction Stop
            Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        } -ErrorAction Stop

        foreach ($Item in @($Result)) {
            if ($Item.PSObject.Properties.Name -contains "Result") {
                Write-AppLog "Entra delta sync result: $($Item.Result)"
            }
            else {
                Write-AppLog ($Item | Out-String).Trim()
            }
        }

        $script:chkDeltaSync.Checked = $true
        Write-AppLog "Delta sync requested on $EntraSyncServer. The cycle itself runs for a few minutes after this." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText 'Delta sync requested.'
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Delta sync failed.'

        if ("$($_.Exception.Message)" -match 'is busy') {
            Write-AppLog "A sync cycle is already running on $EntraSyncServer. Wait for it to finish, then run this again." 'WARNING'
        }
        else {
            Write-AppLog "Entra delta sync failed: $($_.Exception.Message)" 'ERROR'
        }
    }
}

function Invoke-OffboardAllLocal {
    # Everything this tool can do by itself, in the order the checklist expects:
    # export the groups first (so the list survives the removal), then disable,
    # remove groups, clear the phone, replicate, and request the delta sync.
    # The manual steps are done elsewhere and are not touched here.
    if (-not (Test-OffboardSelectionMatches -Operation 'Do All Local Actions')) { return }

    $User = $script:CurrentUser
    $ExportPath = Join-Path $script:OffboardExportFolder "$($User.SamAccountName)-groups-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

    $SyncLine = ''
    if ($script:OffboardSyncServer) {
        $SyncLine = "  6. Request an Entra Connect delta sync on $($script:OffboardSyncServer)`r`n"
    }

    $Answer = [System.Windows.Forms.MessageBox]::Show(
        "Run every local action against this account?`r`n`r`n" +
        "Name: $($User.DisplayName)`r`n" +
        "Username: $($User.SamAccountName)`r`n" +
        "Enabled: $($User.Enabled)`r`n" +
        "$($User.DistinguishedName)`r`n`r`n" +
        "In order:`r`n" +
        "  1. Export their group memberships to $ExportPath`r`n" +
        "  2. Disable the Active Directory account`r`n" +
        "  3. Remove every security group membership (protected groups are skipped)`r`n" +
        "  4. Clear the telephone and mobile numbers`r`n" +
        "  5. Replicate the selected domain controller`r`n" +
        $SyncLine + "`r`n" +
        "This runs without further prompts. The export happens first, so the group list is kept even though the memberships are about to go.`r`n`r`n" +
        "Nothing outside Active Directory is touched - the other steps stay manual.",
        "Confirm Do All Local Actions",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

    if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-AppLog 'Do All Local Actions cancelled.'
        return
    }

    Write-AppLog '--- Running all local actions ---' 'SUCCESS'

    # Each step reports and ticks itself. They are run in sequence rather than
    # stopping at the first failure, because a later step can still be valid -
    # but the summary at the end says exactly what did not tick.
    Export-OffboardGroups -Path $ExportPath
    Invoke-DisableOffboardUser -NoConfirm
    Remove-OffboardGroups -NoConfirm
    Invoke-ClearOffboardPhone -NoConfirm
    Invoke-OffboardReplicate
    if ($script:OffboardSyncServer) { Invoke-OffboardDeltaSync }

    $Pending = @()
    foreach ($Pair in @(
        @('Disable account', $script:chkDisable),
        @('Remove groups', $script:chkGroups),
        @('Clear phone', $script:chkPhone),
        @('Replicate', $script:chkReplicate),
        @('Delta sync', $(if ($script:OffboardSyncServer) { $script:chkDeltaSync } else { $null }))
    )) {
        if ($Pair[1] -and -not $Pair[1].Checked) { $Pending += $Pair[0] }
    }

    Write-AppLog '--- Local actions finished ---' 'SUCCESS'

    if ($Pending.Count -eq 0) {
        Write-AppLog "All local steps completed. Group list saved to $ExportPath." 'SUCCESS'
        Set-BusyState -Busy $false -StatusText 'All local actions complete - continue with the manual steps.'
        Show-InfoMessage "All local actions completed for $($User.SamAccountName).`r`n`r`nGroup list saved to:`r`n$ExportPath`r`n`r`nCarry on with the manual steps below."
    }
    else {
        Write-AppLog "Still outstanding: $($Pending -join ', '). See the lines above for why." 'WARNING'
        Set-BusyState -Busy $false -StatusText "$($Pending.Count) local step(s) did not complete."
        Show-ErrorMessage "Some local actions did not complete for $($User.SamAccountName):`r`n`r`n  $($Pending -join "`r`n  ")`r`n`r`nThe Activity log shows why."
    }
}

function Get-OffboardSummaryText {
    # For pasting into the ticket - what was done, what is still outstanding.
    $Lines = New-Object System.Collections.Generic.List[string]

    $Who = "(no user loaded)"
    if ($script:CurrentUser) { $Who = "$($script:CurrentUser.DisplayName) ($($script:CurrentUser.SamAccountName))" }

    $Lines.Add("Offboarding checklist - $Who")
    $Lines.Add("Owner: $env:USERDOMAIN\$env:USERNAME")
    $Lines.Add("Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")

    if ($script:CurrentUser) {
        # The same fields the paper checklist asks for at the top.
        $Lines.Add("Department: $($script:CurrentUser.Department)")
        $Lines.Add("Title: $($script:CurrentUser.Title)")
        $Lines.Add("Manager: $($script:ManagerDisplayName)")
        $Lines.Add("Account: $($script:CurrentUser.DistinguishedName)")
    }

    if ($script:GroupsExportPath) {
        $Lines.Add("Group membership exported to: $($script:GroupsExportPath)")
    }

    $Lines.Add("")

    for ($Index = 0; $Index -lt $script:StepChecks.Count; $Index++) {
        $Check = $script:StepChecks[$Index]
        $Mark = "[ ]"
        if ($Check.Checked) { $Mark = "[x]" }
        $Lines.Add("$Mark $($Check.Text)")

        # Sub-steps are listed under their step so the ticket shows exactly how
        # far a partly finished step got.
        $Subs = $Check.Tag

        if ($null -ne $Subs) {
            for ($SubIndex = 0; $SubIndex -lt $Subs.Count; $SubIndex++) {
                $Sub = $Subs[$SubIndex]
                $SubMark = "[ ]"
                if ($Sub.Checked) { $SubMark = "[x]" }
                $Lines.Add("      $SubMark $($Sub.Text)")
            }
        }
    }

    return ($Lines.ToArray() -join "`r`n")
}

function Copy-OffboardSummary {
    [System.Windows.Forms.Clipboard]::SetText((Get-OffboardSummaryText))
    Write-AppLog 'Checklist summary copied to the clipboard.' 'SUCCESS'
}

function Get-OffboardOutstandingCount {
    $Outstanding = 0

    for ($Index = 0; $Index -lt $script:StepChecks.Count; $Index++) {
        if (-not $script:StepChecks[$Index].Checked) { $Outstanding++ }
    }

    return $Outstanding
}

function Invoke-OffboardFinish {
    # Copy the checklist, write it to a file and open it, then close the tool.
    $Outstanding = Get-OffboardOutstandingCount

    if ($Outstanding -gt 0) {
        $Answer = [System.Windows.Forms.MessageBox]::Show(
            "$Outstanding step(s) are still unticked.`r`n`r`nFinish anyway? The checklist will record them as outstanding.",
            "Finish offboarding",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-AppLog 'Finish cancelled.'
            return
        }
    }

    $Summary = Get-OffboardSummaryText

    try {
        [System.Windows.Forms.Clipboard]::SetText($Summary)
        Write-AppLog 'Checklist copied to the clipboard.' 'SUCCESS'
    }
    catch {
        Write-AppLog "Could not copy the checklist to the clipboard: $($_.Exception.Message)" 'WARNING'
    }

    # Written to a file and opened, rather than pasted into an empty Notepad -
    # a paste is lost to the next thing that touches the clipboard.
    try {
        $Name = "Offboarding-checklist"
        if ($script:CurrentUser) { $Name = "Offboarding-$($script:CurrentUser.SamAccountName)" }

        $File = Join-Path $env:TEMP "$Name-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
        Set-Content -LiteralPath $File -Value $Summary -Encoding UTF8

        $Process = Start-Process "notepad.exe" -ArgumentList "`"$File`"" -PassThru -ErrorAction Stop
        Write-AppLog "Checklist written to $File and opened in Notepad." 'SUCCESS'
        if ($script:OffboardChecklistLocation) {
            Write-AppLog "File the completed checklist in: $($script:OffboardChecklistLocation)"
        }
        [void](Show-MSToolkitWindowForProcess -Process $Process -ProcessName "notepad")
    }
    catch {
        Write-AppLog "Could not open the checklist in Notepad: $($_.Exception.Message)" 'ERROR'
        Show-ErrorMessage "The checklist is on the clipboard, but Notepad could not be opened.`r`n`r`n$($_.Exception.Message)"
        return
    }

    $script:MainForm.Close()
}

function Get-OffboardUserGroups {
    # Direct membership only. The primary group (normally Domain Users) is not in
    # memberOf and cannot be removed while it is primary, so it never appears here.
    param($User)

    $Groups = New-Object System.Collections.Generic.List[object]

    foreach ($Dn in @($User.MemberOf)) {
        try {
            $Group = Get-ADGroup -Identity $Dn -Server $script:WorkingServer `
                -Properties SID,GroupCategory,GroupScope,Description,isCriticalSystemObject -ErrorAction Stop
            $Groups.Add($Group)
        }
        catch {
            Write-AppLog "Could not read group $($Dn): $($_.Exception.Message)" 'WARNING'
        }
    }

    return $Groups.ToArray()
}

function Export-OffboardGroups {
    param([string]$Path)

    if (-not (Test-OffboardSelectionMatches -Operation 'Export groups')) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading group membership...'

        $User = Get-ADUser -Identity $script:CurrentUser.DistinguishedName -Server $script:WorkingServer `
            -Properties MemberOf -ErrorAction Stop

        $Groups = Get-OffboardUserGroups -User $User

        if ($Groups.Count -eq 0) {
            Write-AppLog "$($script:CurrentUser.SamAccountName) has no direct group memberships to export." 'WARNING'
            Set-BusyState -Busy $false -StatusText 'Nothing to export.'
            return
        }

        $Target = $Path

        if (-not $Target) {
            $Dialog = New-Object System.Windows.Forms.SaveFileDialog
            $Dialog.Filter = 'CSV file (*.csv)|*.csv'
            $Dialog.FileName = "$($script:CurrentUser.SamAccountName)-groups-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

            Set-BusyState -Busy $false

            if ($Dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                Write-AppLog 'Group export cancelled.'
                return
            }

            $Target = $Dialog.FileName
        }
        else {
            $Folder = Split-Path -Parent $Target
            if ($Folder -and -not (Test-Path -LiteralPath $Folder)) {
                New-Item -ItemType Directory -Path $Folder -Force | Out-Null
            }
        }

        $Rows = foreach ($Group in $Groups) {
            [pscustomobject]@{
                User              = $script:CurrentUser.SamAccountName
                UserDN            = $script:CurrentUser.DistinguishedName
                GroupName         = $Group.Name
                GroupSamAccount   = $Group.SamAccountName
                GroupCategory     = $Group.GroupCategory
                GroupScope        = $Group.GroupScope
                GroupDescription  = $Group.Description
                GroupDN           = $Group.DistinguishedName
                ExportedOn        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                ExportedBy        = "$env:USERDOMAIN\$env:USERNAME"
            }
        }

        $Rows | Export-Csv -LiteralPath $Target -NoTypeInformation -Encoding UTF8

        $script:GroupsExportPath = $Target
        Write-AppLog "Exported $($Groups.Count) group membership(s) to $Target." 'SUCCESS'

        if (-not $Path) {
            Show-InfoMessage "Exported $($Groups.Count) group membership(s) to:`r`n`r`n$Target"
        }
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Group export failed.'
        Write-AppLog "Group export failed: $($_.Exception.Message)" 'ERROR'
        Show-ErrorMessage "Could not export the group membership.`r`n`r`n$($_.Exception.Message)"
    }
}

function Remove-OffboardGroups {
    param([switch]$NoConfirm)

    if (-not (Test-OffboardSelectionMatches -Operation 'Remove group memberships')) { return }

    try {
        Set-BusyState -Busy $true -StatusText 'Reading group membership...'

        $User = Get-ADUser -Identity $script:CurrentUser.DistinguishedName -Server $script:WorkingServer `
            -Properties MemberOf -ErrorAction Stop

        $Groups = Get-OffboardUserGroups -User $User
        Set-BusyState -Busy $false

        if ($Groups.Count -eq 0) {
            Write-AppLog "$($script:CurrentUser.SamAccountName) is not a direct member of any group." 'SUCCESS'
            if ($script:chkGroups) { $script:chkGroups.Checked = $true }
            return
        }

        # Protected groups are identified by SID and RID, never by name, and are
        # dropped from the removal rather than attempted and failed.
        $Removable = New-Object System.Collections.Generic.List[object]
        $Blocked = New-Object System.Collections.Generic.List[string]

        foreach ($Group in $Groups) {
            $Reason = Get-MSToolkitCriticalGroupReason -Group $Group

            if ($Reason) { $Blocked.Add("$($Group.Name) - $Reason") }
            else { $Removable.Add($Group) }
        }

        $ExportNote = "No export has been saved in this window yet."
        if ($script:GroupsExportPath) { $ExportNote = "Exported to: $($script:GroupsExportPath)" }

        $Listed = @($Removable | ForEach-Object { "  $($_.Name)" })
        if ($Listed.Count -gt 25) {
            $Listed = @($Listed | Select-Object -First 25) + @("  ... and $($Listed.Count - 25) more")
        }

        $BlockedText = ""
        if ($Blocked.Count -gt 0) {
            $BlockedText = "`r`n`r`nThese will NOT be removed - protected groups:`r`n  " + (($Blocked.ToArray()) -join "`r`n  ")
        }

        $Answer = [System.Windows.Forms.DialogResult]::Yes

        if (-not $NoConfirm) {
        $Answer = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($Removable.Count) group membership(s) from $($script:CurrentUser.SamAccountName)?`r`n`r`n" +
            "$($Listed -join "`r`n")$BlockedText`r`n`r`n" +
            "$ExportNote`r`n`r`n" +
            "Export the list first if there is any chance it will be needed again - this cannot be undone from here.",
            "Confirm Remove Group Memberships",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        }

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-AppLog 'Group removal cancelled.'
            return
        }

        foreach ($Reason in $Blocked) {
            Write-AppLog "BLOCKED: $Reason" 'WARNING'
        }

        $Removed = 0
        $Failed = 0
        $RemovedGroups = @()

        for ($Index = 0; $Index -lt $Removable.Count; $Index++) {
            $Group = $Removable[$Index]
            Set-BusyState -Busy $true -StatusText "Removing $($Group.Name)..."

            try {
                Remove-ADGroupMember -Identity $Group.DistinguishedName `
                    -Members $script:CurrentUser.DistinguishedName `
                    -Server $script:WorkingServer -Confirm:$false -ErrorAction Stop

                $Removed++
                $RemovedGroups += ,([ordered]@{
                    Name              = [string]$Group.Name
                    DistinguishedName = [string]$Group.DistinguishedName
                })

                Write-AppLog "Removed from $($Group.Name)." 'SUCCESS'
            }
            catch {
                $Failed++
                Write-AppLog "Could not remove from $($Group.Name): $($_.Exception.Message)" 'ERROR'
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        # Written even on a partial failure: what came off has to be restorable.
        if ($RemovedGroups.Count -gt 0) {
            Add-OffboardUndoAction -User $script:CurrentUser -Type 'RemoveGroups' -Details ([ordered]@{
                Groups = $RemovedGroups
            })
        }

        $script:CurrentUser = Get-ADUser -Identity $script:CurrentUser.DistinguishedName -Server $script:WorkingServer `
            -Properties SID,Enabled,DisplayName,Description,Title,Department,Manager,telephoneNumber,mobile,DistinguishedName,isCriticalSystemObject

        if ($Failed -eq 0) {
            if ($script:chkGroups) { $script:chkGroups.Checked = $true }
            Set-BusyState -Busy $false -StatusText "Removed $Removed group membership(s)."
        }
        else {
            Set-BusyState -Busy $false -StatusText "$Failed group removal(s) failed."
            Write-AppLog 'Not ticking this step: at least one removal failed.' 'WARNING'
        }

        Write-AppLog "Group removal complete: $Removed removed, $Failed failed, $($Blocked.Count) protected and left alone." 'SUCCESS'
    }
    catch {
        Set-BusyState -Busy $false -StatusText 'Group removal failed.'
        Write-AppLog "Group removal failed: $($_.Exception.Message)" 'ERROR'
        Show-ErrorMessage "Could not remove the group memberships.`r`n`r`n$($_.Exception.Message)"
    }
}

function Open-MSToolkitFileAsSignedInUser {
    param([string]$FilePath)

    # Opens a file in its default app (Word for a .docx) as the signed-in Windows
    # user, where Office is licensed and signed in. "start" runs inside that
    # user's own process, so it does not depend on handing off to the desktop
    # shell - which refuses a request from the admin account this window runs as.
    $FileName = Split-Path -Leaf $FilePath

    return (Invoke-MSToolkitAsSignedInUser `
        -Purpose "Open $FileName" `
        -HeaderText "Open Document" `
        -WarningText "Sign in as the account you are signed in to Windows with." `
        -ExplanationText "The document opens in Word under your own account, where Office is licensed - not the Domain Admin account this tool runs as." `
        -Launch {
            param($Credential)

            Start-Process -FilePath (Join-Path $env:SystemRoot "System32\cmd.exe") `
                -ArgumentList "/c start `"`" `"$FilePath`"" `
                -Credential $Credential `
                -WorkingDirectory $env:SystemRoot `
                -WindowStyle Hidden `
                -LoadUserProfile `
                -ErrorAction Stop

            Write-OffboardOutput "Opened $FilePath as $($Credential.UserName)."
        })
}

function Open-OffboardDocument {
    # A document from MSToolkit Settings: a web address, a full path, or a file
    # name looked for in MSToolkit's install folder. A full path is opened as the
    # signed-in user even when this admin account cannot see it - a file in the
    # user's own profile or OneDrive is only visible to them.
    param([string]$Value, [string]$Label)

    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    $Value = $Value.Trim().Trim('"')

    if ($Value -match '^https?://') {
        Open-OffboardLink -Url $Value -Label $Label
        return
    }

    if ([System.IO.Path]::IsPathRooted($Value)) {
        if (-not (Test-Path -LiteralPath $Value)) {
            Write-AppLog "$Label is not visible to the account this tool runs as - asking Windows to open it as you: $Value" 'WARNING'
        }
        else {
            Write-AppLog "Opening $Value"
        }
        [void](Open-MSToolkitFileAsSignedInUser -FilePath $Value)
        return
    }

    # Just a file name: look beside this tool, then in the install folder.
    $Folders = @($PSScriptRoot)
    $InstallFolder = Join-Path ${env:ProgramFiles(x86)} "MSToolkit"
    if ($InstallFolder -ne $PSScriptRoot) { $Folders += $InstallFolder }

    foreach ($Folder in $Folders) {
        $Candidate = Join-Path $Folder $Value
        if (Test-Path -LiteralPath $Candidate) {
            Write-AppLog "Opening $Candidate"
            [void](Open-MSToolkitFileAsSignedInUser -FilePath $Candidate)
            return
        }
    }

    # Nothing matched, so say what IS in the folder - far more use than repeating
    # the name that was not found.
    $Present = @()
    try {
        $Present = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter "*.doc*" -File -ErrorAction Stop |
            Select-Object -ExpandProperty Name)
    }
    catch { }

    $PresentText = "No .doc or .docx files are in that folder at all."
    if ($Present.Count -gt 0) {
        $PresentText = "Documents found in that folder:`r`n  " + ($Present -join "`r`n  ")
    }

    Write-AppLog "$Label not found. Looked for '$Value' in: $($Folders -join '; ')" 'ERROR'

    Show-ErrorMessage ("$Label was not found.`r`n`r`nLooked for:`r`n  $Value`r`n`r`nIn:`r`n  $($Folders -join "`r`n  ")`r`n`r`n$PresentText`r`n`r`n" +
        "Enter a full path or web address in MSToolkit Settings (Offboarding) to use a document kept somewhere else. " +
        "Check for a hidden second extension - a downloaded file is often saved as .docx.docx with file extensions hidden in Explorer.")
}

function Start-OffboardM365Tool {
    # Hands off to Start-MSToolkitM365Tool, the same function MSToolkit uses, so the
    # remembered password and "sign in automatically" setting are shared with it
    # rather than being asked for again here.
    #
    # The departing user is passed through, so the tool connects on startup and
    # opens Manage User Groups with them already loaded.
    param([string]$ScriptName, [string]$Label)

    $Script = Join-Path $PSScriptRoot $ScriptName

    if (-not (Test-Path -LiteralPath $Script)) {
        Write-AppLog "Could not find $ScriptName beside this tool." 'ERROR'
        Show-ErrorMessage "$Label is not installed alongside this tool.`r`n`r`nExpected: $Script"
        return
    }

    $ManageUser = ''

    if ($script:CurrentUser) {
        # UPN first: it is what both M365 tools resolve against.
        $ManageUser = [string]$script:CurrentUser.UserPrincipalName
        if (-not $ManageUser) { $ManageUser = [string]$script:CurrentUser.SamAccountName }
    }

    if ($ManageUser) {
        Write-AppLog "$Label will open Manage User Groups for $ManageUser."
    }
    else {
        Write-AppLog "No user is loaded, so $Label will open without one." 'WARNING'
    }

    Start-MSToolkitM365Tool -ToolName $Label -ScriptFile $ScriptName -ManageUser $ManageUser
}

# ---------------------------------------------------------------------------
# Signing in as the Windows user - ported verbatim from MSToolkit.ps1
#
# The Microsoft 365 tools must run as the standard account, not the admin
# account this window runs as. This is the same flow MSToolkit uses, reading and
# writing the SAME files, so a password remembered in MSToolkit is used here and
# a password remembered here is used by MSToolkit:
#
#   %APPDATA%\MSToolkit\m365-credential.xml   (DPAPI, admin profile)
#   %APPDATA%\MSToolkit\signin-auto.flag
#
# Hold Shift while clicking to force the sign-in box back up.
# ---------------------------------------------------------------------------
function Write-OffboardOutput {
    # MSToolkit writes to its output box; the ported code calls this instead.
    param([string]$Message, $Color)

    $Level = 'INFO'

    if ($Color -is [System.Drawing.Color]) {
        if ($Color -eq [System.Drawing.Color]::Red) { $Level = 'ERROR' }
        elseif ($Color -eq [System.Drawing.Color]::DarkOrange -or $Color -eq [System.Drawing.Color]::Orange) { $Level = 'WARNING' }
        elseif ($Color -eq [System.Drawing.Color]::Green -or $Color -eq [System.Drawing.Color]::DarkGreen) { $Level = 'SUCCESS' }
    }

    Write-AppLog $Message $Level
}

function Test-MSToolkitIconFontAvailable {
    try {
        $Installed = New-Object System.Drawing.Text.InstalledFontCollection
        return (@($Installed.Families | Where-Object { $_.Name -eq "Segoe MDL2 Assets" }).Count -gt 0)
    }
    catch {
        return $false
    }
}

function Get-MSToolkitSignedInWindowsUser {
    # MSToolkit runs as the admin account, so its own identity is not the person at the
    # desktop. Ask who owns the Explorer shell in this session - that is whoever is
    # actually signed in to Windows here - then fall back to the console user.
    try {
        $SessionId = (Get-Process -Id $PID).SessionId
        $Shells = @(
            Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop |
                Where-Object { $_.SessionId -eq $SessionId }
        )

        foreach ($Shell in $Shells) {
            $Owner = Invoke-CimMethod -InputObject $Shell -MethodName GetOwner -ErrorAction Stop
            if (($Owner.ReturnValue -eq 0) -and -not [string]::IsNullOrWhiteSpace($Owner.User)) {
                return "$($Owner.Domain)\$($Owner.User)"
            }
        }
    }
    catch { }

    try {
        $ConsoleUser = [string](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).UserName
        if (-not [string]::IsNullOrWhiteSpace($ConsoleUser)) {
            return $ConsoleUser
        }
    }
    catch { }

    return $null
}

# Remembered M365 sign-in. Export-Clixml protects the password with DPAPI, which ties
# it to the account MSToolkit runs as on this computer: copied elsewhere, or opened by
# any other account, the file cannot be decrypted.
$script:MSToolkitM365CredentialPath = Join-Path (Join-Path $env:APPDATA "MSToolkit") "m365-credential.xml"

function Get-MSToolkitRememberedM365Credential {
    if (-not (Test-Path -LiteralPath $script:MSToolkitM365CredentialPath)) {
        return $null
    }

    try {
        $Saved = Import-Clixml -LiteralPath $script:MSToolkitM365CredentialPath -ErrorAction Stop
        if ($Saved -and $Saved.UserName -and $Saved.Password) {
            return $Saved
        }
    }
    catch { }

    return $null
}

function Save-MSToolkitRememberedM365Credential {
    param([System.Management.Automation.PSCredential]$Credential)

    try {
        $Folder = Split-Path -Parent $script:MSToolkitM365CredentialPath
        if (-not (Test-Path -LiteralPath $Folder)) {
            New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        }

        $Credential | Export-Clixml -LiteralPath $script:MSToolkitM365CredentialPath -Force
        return $true
    }
    catch {
        return $false
    }
}

# "Sign in automatically next time" is a marker file beside the remembered credential.
$script:MSToolkitSignInAutoPath = Join-Path (Split-Path -Parent $script:MSToolkitM365CredentialPath) "signin-auto.flag"

function Test-MSToolkitAutoSignIn {
    return (Test-Path -LiteralPath $script:MSToolkitSignInAutoPath)
}

function Set-MSToolkitAutoSignIn {
    param([bool]$Enabled)

    try {
        if ($Enabled) {
            $Folder = Split-Path -Parent $script:MSToolkitSignInAutoPath
            if (-not (Test-Path -LiteralPath $Folder)) {
                New-Item -ItemType Directory -Path $Folder -Force | Out-Null
            }
            Set-Content -LiteralPath $script:MSToolkitSignInAutoPath -Value "1" -Force
        }
        elseif (Test-Path -LiteralPath $script:MSToolkitSignInAutoPath) {
            Remove-Item -LiteralPath $script:MSToolkitSignInAutoPath -Force -ErrorAction SilentlyContinue
        }
    }
    catch { }
}

function Remove-MSToolkitRememberedM365Credential {
    if (Test-Path -LiteralPath $script:MSToolkitM365CredentialPath) {
        Remove-Item -LiteralPath $script:MSToolkitM365CredentialPath -Force -ErrorAction SilentlyContinue
        return $true
    }

    return $false
}

function Show-MSToolkitM365CredentialPrompt {
    param(
        [string]$ToolName,
        [string]$Attempts = "",
        [string]$UserName = "",
        [string]$Password = "",
        [bool]$Remember = $false,
        [bool]$AutoSignIn = $false,
        [string]$HeaderText = "Microsoft 365 Sign-In",
        [string]$WarningText = "Use your STANDARD domain account - not a Domain Admin account.",
        [string]$ExplanationText = "Domain Admin accounts are not licensed or eligible for the Microsoft Graph and Exchange Online permissions these tools request."
    )

    $Palette = Get-MSToolkitThemePalette

    $CredForm = New-Object System.Windows.Forms.Form
    $CredForm.Text = "$HeaderText - $ToolName"
    $CredForm.Size = New-Object System.Drawing.Size(560,422)
    $CredForm.StartPosition = "CenterScreen"
    $CredForm.FormBorderStyle = "FixedDialog"
    $CredForm.MaximizeBox = $false
    $CredForm.MinimizeBox = $false

    $HeaderPanel = New-Object System.Windows.Forms.Panel
    $HeaderPanel.Dock = "Top"
    $HeaderPanel.Height = 62
    $HeaderPanel.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $HeaderPanel.Tag = "TopBar"
    $CredForm.Controls.Add($HeaderPanel)

    $HeaderTitle = New-Object System.Windows.Forms.Label
    $HeaderTitle.Text = $HeaderText
    $HeaderTitle.AutoSize = $true
    $HeaderTitle.ForeColor = [System.Drawing.Color]::White
    $HeaderTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",14)
    $HeaderTitle.Location = New-Object System.Drawing.Point(18,10)
    $HeaderPanel.Controls.Add($HeaderTitle)

    $HeaderSub = New-Object System.Windows.Forms.Label
    $HeaderSub.Text = $ToolName
    $HeaderSub.AutoSize = $true
    $HeaderSub.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
    $HeaderSub.Font = New-Object System.Drawing.Font("Segoe UI",9.5)
    $HeaderSub.Location = New-Object System.Drawing.Point(20,38)
    $HeaderPanel.Controls.Add($HeaderSub)

    $WarnLabel = New-Object System.Windows.Forms.Label
    $WarnLabel.Text = $WarningText
    $WarnLabel.Location = New-Object System.Drawing.Point(20,78)
    $WarnLabel.Size = New-Object System.Drawing.Size(510,22)
    $WarnLabel.ForeColor = $Palette.Danger
    $WarnLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold",10)
    $CredForm.Controls.Add($WarnLabel)

    $ExplainLabel = New-Object System.Windows.Forms.Label
    $ExplainLabel.Text = $ExplanationText
    $ExplainLabel.Location = New-Object System.Drawing.Point(20,102)
    $ExplainLabel.Size = New-Object System.Drawing.Size(510,34)
    $ExplainLabel.ForeColor = $Palette.MutedText
    $ExplainLabel.Font = New-Object System.Drawing.Font("Segoe UI",8.5)
    $CredForm.Controls.Add($ExplainLabel)

    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.Text = "User name:"
    $UserLabel.Location = New-Object System.Drawing.Point(20,148)
    $UserLabel.Size = New-Object System.Drawing.Size(100,22)
    $CredForm.Controls.Add($UserLabel)

    $UserBox = New-Object System.Windows.Forms.TextBox
    $UserBox.Location = New-Object System.Drawing.Point(125,145)
    $UserBox.Size = New-Object System.Drawing.Size(405,24)
    # Shown exactly as saved or typed - no domain is added.
    $UserBox.Text = $UserName
    $CredForm.Controls.Add($UserBox)

    $UserHintText = "Enter as DOMAIN\username."

    $UserHint = New-Object System.Windows.Forms.Label
    $UserHint.Text = $UserHintText
    $UserHint.Location = New-Object System.Drawing.Point(127,171)
    $UserHint.Size = New-Object System.Drawing.Size(403,16)
    $UserHint.Font = New-Object System.Drawing.Font("Segoe UI",8)
    $UserHint.ForeColor = $Palette.MutedText
    $CredForm.Controls.Add($UserHint)

    $UserTip = New-Object System.Windows.Forms.ToolTip
    $UserTip.SetToolTip($UserBox, $UserHintText)

    $PassLabel = New-Object System.Windows.Forms.Label
    $PassLabel.Text = "Password:"
    $PassLabel.Location = New-Object System.Drawing.Point(20,202)
    $PassLabel.Size = New-Object System.Drawing.Size(100,22)
    $CredForm.Controls.Add($PassLabel)

    $PassBox = New-Object System.Windows.Forms.TextBox
    $PassBox.Location = New-Object System.Drawing.Point(125,199)
    $PassBox.Size = New-Object System.Drawing.Size(405,24)
    $PassBox.UseSystemPasswordChar = $true
    $PassBox.Text = $Password
    $CredForm.Controls.Add($PassBox)

    # Reveal control, drawn inside the password field like a browser does.
    # Segoe MDL2 Assets glyphs live in the Private Use Area, so fall back to a
    # text label if that font is not present rather than render empty boxes.
    $UseIconFont = Test-MSToolkitIconFontAvailable

    # A Label rather than a Button: a Button always paints its own background and
    # focus cue, which shows as a box sitting on top of the field.
    $RevealButton = New-Object System.Windows.Forms.Label
    $RevealButton.AutoSize = $false
    $RevealButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $RevealButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $RevealButton.Tag = "RevealButton"

    if ($UseIconFont) {
        $RevealButton.Size = New-Object System.Drawing.Size(22,18)
        $RevealButton.Font = New-Object System.Drawing.Font("Segoe MDL2 Assets",9)
        $RevealButton.Text = [string][char]0xE7B3
    }
    else {
        $RevealButton.Size = New-Object System.Drawing.Size(38,18)
        $RevealButton.Font = New-Object System.Drawing.Font("Segoe UI",8)
        $RevealButton.Text = "Show"
    }

    # Sit fully inside the field: clear of the right border and vertically centred
    # against the control's real height rather than an assumed one.
    $RevealInset = 5
    $RevealButton.Location = New-Object System.Drawing.Point(
        ($PassBox.Right - $RevealButton.Width - $RevealInset),
        ($PassBox.Top + [int][Math]::Floor(($PassBox.Height - $RevealButton.Height) / 2))
    )

    $CredForm.Controls.Add($RevealButton)
    $RevealButton.BringToFront()

    $RevealTip = New-Object System.Windows.Forms.ToolTip
    $RevealTip.SetToolTip($RevealButton, "Show or hide the password")

    $RevealButton.Add_Click({
        $PassBox.UseSystemPasswordChar = -not $PassBox.UseSystemPasswordChar

        if ($UseIconFont) {
            if ($PassBox.UseSystemPasswordChar) {
                $RevealButton.Text = [string][char]0xE7B3
            }
            else {
                $RevealButton.Text = [string][char]0xED1A
            }
        }
        else {
            if ($PassBox.UseSystemPasswordChar) {
                $RevealButton.Text = "Show"
            }
            else {
                $RevealButton.Text = "Hide"
            }
        }

        $PassBox.Focus()
        $PassBox.SelectionStart = $PassBox.Text.Length
    }.GetNewClosure())

    $RememberBox = New-Object System.Windows.Forms.CheckBox
    $RememberBox.Text = "Remember password on this computer"
    $RememberBox.Location = New-Object System.Drawing.Point(125,230)
    $RememberBox.Size = New-Object System.Drawing.Size(405,22)
    $RememberBox.Checked = $Remember
    $CredForm.Controls.Add($RememberBox)

    # Only meaningful when the password is remembered, so it follows that box.
    $AutoBox = New-Object System.Windows.Forms.CheckBox
    $AutoBox.Text = "Sign in automatically next time"
    $AutoBox.Location = New-Object System.Drawing.Point(125,254)
    $AutoBox.Size = New-Object System.Drawing.Size(405,22)
    $AutoBox.Checked = ($Remember -and $AutoSignIn)
    $AutoBox.Enabled = $Remember
    $CredForm.Controls.Add($AutoBox)

    $AutoHint = New-Object System.Windows.Forms.Label
    $AutoHint.Text = "Hold Shift when clicking to show this box again."
    $AutoHint.Location = New-Object System.Drawing.Point(143,276)
    $AutoHint.Size = New-Object System.Drawing.Size(387,18)
    $AutoHint.Font = New-Object System.Drawing.Font("Segoe UI",8)
    $AutoHint.ForeColor = $Palette.MutedText
    $CredForm.Controls.Add($AutoHint)

    $RememberBox.Add_CheckedChanged({
        $AutoBox.Enabled = $RememberBox.Checked
        if (-not $RememberBox.Checked) {
            $AutoBox.Checked = $false
        }
    })

    $AttemptLabel = New-Object System.Windows.Forms.Label
    $AttemptLabel.Text = $Attempts
    $AttemptLabel.Location = New-Object System.Drawing.Point(20,300)
    $AttemptLabel.Size = New-Object System.Drawing.Size(510,22)
    $AttemptLabel.ForeColor = $Palette.Danger
    $AttemptLabel.Font = New-Object System.Drawing.Font("Segoe UI",8.5)
    $CredForm.Controls.Add($AttemptLabel)

    $OkButton = New-Object System.Windows.Forms.Button
    $OkButton.Text = "Sign In"
    $OkButton.Location = New-Object System.Drawing.Point(320,336)
    $OkButton.Size = New-Object System.Drawing.Size(100,32)
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CredForm.Controls.Add($OkButton)
    $CredForm.AcceptButton = $OkButton

    $CancelBtn = New-Object System.Windows.Forms.Button
    $CancelBtn.Text = "Cancel"
    $CancelBtn.Location = New-Object System.Drawing.Point(430,336)
    $CancelBtn.Size = New-Object System.Drawing.Size(100,32)
    $CancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CredForm.Controls.Add($CancelBtn)
    $CredForm.CancelButton = $CancelBtn

    $CredForm.Add_Shown({
        $CredForm.Activate()

        # With a remembered password filled in, one Enter signs in.
        if ($PassBox.Text.Length -gt 0) {
            $OkButton.Focus()
        }
        else {
            $UserBox.Focus()
            $UserBox.SelectionStart = $UserBox.Text.Length
        }
    })

    Apply-MSToolkitSharedTheme -Root $CredForm

    $AutoHint.ForeColor = $Palette.MutedText
    $UserHint.ForeColor = $Palette.MutedText

    # Blend the reveal control into the field and keep typed text clear of it.
    $RevealButton.BackColor = $PassBox.BackColor
    $RevealButton.ForeColor = $Palette.MutedText

    $CredForm.Add_Shown({
        try {
            [MSToolkitWindow]::SetRightMargin($PassBox.Handle, ($RevealButton.Width + 10))
        }
        catch { }
    }.GetNewClosure())

    $Result = $CredForm.ShowDialog()

    if ($Result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $UserName = "$($UserBox.Text)".Trim()

    # Start-Process -Credential needs DOMAIN\username. Without a domain it signs in
    # against this computer's local accounts, which is never what these tools want.
    $HasDomain = ($UserName -match '^[^\\@]+\\[^\\@]+$')

    if ([string]::IsNullOrWhiteSpace($UserName) -or -not $HasDomain -or [string]::IsNullOrWhiteSpace($PassBox.Text)) {
        return [pscustomobject]@{ Incomplete = $true; UserName = $UserName; Remember = $RememberBox.Checked; AutoSignIn = $AutoBox.Checked }
    }

    $SecurePassword = ConvertTo-SecureString $PassBox.Text -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($UserName, $SecurePassword)
    $Credential | Add-Member -NotePropertyName Remember -NotePropertyValue $RememberBox.Checked
    $Credential | Add-Member -NotePropertyName AutoSignIn -NotePropertyValue ($RememberBox.Checked -and $AutoBox.Checked)
    return $Credential
}

function Test-MSToolkitCredentialFailure {
    param($ErrorRecord)

    # True only when Windows rejected the sign-in itself. Anything else - access
    # denied, file in use, a bad path - is a different problem and must not be
    # treated as a wrong password, or it would clear a saved password that works.
    $Exception = $ErrorRecord.Exception
    while ($Exception) {
        if ($Exception -is [System.ComponentModel.Win32Exception]) {
            # 1326 bad user name or password, 1327 account restriction, 1328 logon hours,
            # 1329 workstation, 1330 password expired, 1331 account disabled,
            # 1909 account locked out, 1907 password must change
            if ($Exception.NativeErrorCode -in 1326,1327,1328,1329,1330,1331,1907,1909) {
                return $true
            }
        }
        $Exception = $Exception.InnerException
    }

    $Message = [string]$ErrorRecord.Exception.Message
    return ($Message -match '(?i)user name or password is incorrect|unknown user name or bad password|logon failure|account (is )?(currently )?(disabled|locked)|password (has )?expired|must be changed')
}

function Test-MSToolkitSameAccount {
    param(
        [string]$First,
        [string]$Second
    )

    # Same account however its letters are cased or its domain is written (NetBIOS or
    # DNS name), by comparing SIDs. If either name cannot be resolved, only an exact
    # match counts.
    if ([string]::IsNullOrWhiteSpace($First) -or [string]::IsNullOrWhiteSpace($Second)) {
        return $false
    }

    if ($First.Trim() -ieq $Second.Trim()) {
        return $true
    }

    try {
        $FirstSid  = ([System.Security.Principal.NTAccount]$First.Trim()).Translate([System.Security.Principal.SecurityIdentifier]).Value
        $SecondSid = ([System.Security.Principal.NTAccount]$Second.Trim()).Translate([System.Security.Principal.SecurityIdentifier]).Value
        return ($FirstSid -eq $SecondSid)
    }
    catch {
        return $false
    }
}

function Invoke-MSToolkitAsSignedInUser {
    param(
        [string]$Purpose,
        [string]$HeaderText = "Microsoft 365 Sign-In",
        [string]$WarningText = "Use your STANDARD domain account - not a Domain Admin account.",
        [string]$ExplanationText = "Domain Admin accounts are not licensed or eligible for the Microsoft Graph and Exchange Online permissions these tools request.",
        [scriptblock]$Launch
    )

    # Shared sign-in for everything MSToolkit runs as the signed-in Windows user: the
    # M365 tools, opening the Logs folder, and opening a log file. $Launch receives the
    # credential and must throw on failure (Start-Process -Credential does on a bad
    # password), which is how a wrong or outdated password is detected.
    # Local names are prefixed SI so they cannot shadow variables $Launch relies on.

    $SIUser = Get-MSToolkitSignedInWindowsUser
    # The Windows user is only used to decide whether a remembered sign-in belongs to
    # the person at this desk. The box itself starts empty unless a sign-in was saved.
    $SIName = ""
    $SIPassword = ""
    $SIRemember = $false
    $SIAuto = $false
    $SISaved = Get-MSToolkitRememberedM365Credential

    if ($SISaved -and ((-not $SIUser) -or (Test-MSToolkitSameAccount -First $SISaved.UserName -Second $SIUser))) {
        $SIName = $SISaved.UserName
        $SIPassword = $SISaved.GetNetworkCredential().Password
        $SIRemember = $true
        $SIAuto = Test-MSToolkitAutoSignIn
    }

    # Automatic sign-in, unless Shift is held to bring the sign-in box back.
    $SIShift = (([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Shift) -eq [System.Windows.Forms.Keys]::Shift)

    if ($SISaved -and $SIAuto -and -not $SIShift) {
        try {
            & $Launch $SISaved
            Write-OffboardOutput "$Purpose - signed in automatically as $($SISaved.UserName). Hold Shift when clicking to change this."
            return $true
        }
        catch {
            if (Test-MSToolkitCredentialFailure -ErrorRecord $_) {
                Write-OffboardOutput "Automatic sign-in as $($SISaved.UserName) failed: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                [void](Remove-MSToolkitRememberedM365Credential)
                Set-MSToolkitAutoSignIn -Enabled $false
                Write-OffboardOutput "The remembered password no longer works and has been cleared." ([System.Drawing.Color]::DarkOrange)
                $SIPassword = ""
                $SIRemember = $false
                $SIAuto = $false
            }
            else {
                # Signed in fine; the action itself failed. Keep the saved password.
                Write-OffboardOutput "$Purpose failed after signing in as $($SISaved.UserName): $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                return $false
            }
        }
    }

    $SIAttempt = 0
    $SIMessage = ""

    while ($true) {
        $SIAttempt++
        $SICred = Show-MSToolkitM365CredentialPrompt `
            -ToolName $Purpose `
            -Attempts $SIMessage `
            -UserName $SIName `
            -Password $SIPassword `
            -Remember $SIRemember `
            -AutoSignIn $SIAuto `
            -HeaderText $HeaderText `
            -WarningText $WarningText `
            -ExplanationText $ExplanationText

        if ($null -eq $SICred) {
            Write-OffboardOutput "$Purpose cancelled."
            return $false
        }

        $SIRemember = [bool]$SICred.Remember
        $SIAuto = [bool]$SICred.AutoSignIn

        if ($SICred.PSObject.Properties.Name -contains "Incomplete") {
            if (-not [string]::IsNullOrWhiteSpace($SICred.UserName)) {
                $SIName = $SICred.UserName
            }
            $SIPassword = ""
            if ((-not [string]::IsNullOrWhiteSpace($SICred.UserName)) -and ($SICred.UserName -notmatch '^[^\\@]+\\[^\\@]+$')) {
                $SIMessage = "Enter the user name as DOMAIN\username."
            }
            else {
                $SIMessage = "Enter both a user name and a password."
            }
            continue
        }

        $SIName = $SICred.UserName

        try {
            & $Launch $SICred

            # Only remember a password once it has actually worked.
            if ($SIRemember) {
                if (Save-MSToolkitRememberedM365Credential -Credential $SICred) {
                    Set-MSToolkitAutoSignIn -Enabled $SIAuto
                    if ($SIAuto) {
                        Write-OffboardOutput "Password remembered for $($SICred.UserName); you will be signed in automatically next time."
                    }
                    else {
                        Write-OffboardOutput "Password remembered for $($SICred.UserName) on this computer."
                    }
                }
                else {
                    Write-OffboardOutput "Could not save the password; you will be asked next time." ([System.Drawing.Color]::DarkOrange)
                }
            }
            else {
                Set-MSToolkitAutoSignIn -Enabled $false
                if (Remove-MSToolkitRememberedM365Credential) {
                    Write-OffboardOutput "Remembered password cleared."
                }
            }

            return $true
        }
        catch {
            if (-not (Test-MSToolkitCredentialFailure -ErrorRecord $_)) {
                # The sign-in worked; the action itself failed. Report it as what it is,
                # keep any saved password, and do not ask for the password again.
                Write-OffboardOutput "$Purpose failed after signing in as $($SICred.UserName): $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                return $false
            }

            Write-OffboardOutput "Sign-in failed for $($SICred.UserName) (attempt $SIAttempt): $($_.Exception.Message)" ([System.Drawing.Color]::Red)

            if (Remove-MSToolkitRememberedM365Credential) {
                Set-MSToolkitAutoSignIn -Enabled $false
                Write-OffboardOutput "The remembered password no longer works and has been cleared." ([System.Drawing.Color]::DarkOrange)
            }

            $SIPassword = ""
            $SIMessage = "Sign-in failed (attempt $SIAttempt). Check the user name and password."
        }
    }
}

function Start-MSToolkitM365Tool {
    param(
        [string]$ToolName,
        [string]$ScriptFile,
        [string]$ManageUser
    )

    $Script = Join-Path $PSScriptRoot $ScriptFile

    if (-not (Test-Path $Script)) {
        Write-OffboardOutput "ERROR: $ToolName script not found: $Script" ([System.Drawing.Color]::Red)
        return
    }

    Write-OffboardOutput "Signing in to launch $ToolName."

    [void](Invoke-MSToolkitAsSignedInUser -Purpose $ToolName -Launch {
        param($Credential)

        Write-OffboardOutput "Launching $ToolName as: $($Credential.UserName)"

        $Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$Script`" -ThemeMode `"$script:ThemeMode`""

        if ($ManageUser) {
            $Arguments += " -AutoConnect -ManageUser `"$ManageUser`""
        }

        $Process = Start-Process powershell.exe `
            -Credential $Credential `
            -WorkingDirectory $PSScriptRoot `
            -ArgumentList $Arguments `
            -PassThru `
            -ErrorAction Stop

        [void](Show-MSToolkitWindowForProcess -Process $Process -ProcessName 'powershell')
    })
}

# ---------------------------------------------------------------------------
# Undo journal
#
# Every change this tool makes to Active Directory is recorded to a file before
# and after it happens: what the account looked like, and what was taken away.
# That record outlives the window, so an offboarding run can be reversed days
# later - which is the case that matters, because the wrong user is rarely
# noticed within five minutes.
#
# Journals live in the running admin account's own %LOCALAPPDATA%, so only that
# account can change them. They are never deleted by the tool.
# ---------------------------------------------------------------------------
function Get-OffboardUndoFolder {
    # In this admin account's own local profile, which standard users cannot write
    # to. Not the shared Logs folder: an undo record says which groups to add the
    # account back to, so a record a standard user could edit would let them pick
    # groups for an admin to add someone to.
    $Folder = Join-Path $env:LOCALAPPDATA "MSToolkit\Offboard-Undo"
    if (-not (Test-Path -LiteralPath $Folder)) {
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
    }

    return $Folder
}

function Start-OffboardUndoRecord {
    # Opened the first time something is actually changed, not when a user is
    # merely loaded - a journal with no actions in it is just noise.
    param($User)

    if ($script:UndoRecord -and $script:UndoRecord.User.DistinguishedName -eq $User.DistinguishedName) {
        return $script:UndoRecord
    }

    $script:UndoRecord = [ordered]@{
        SchemaVersion = 1
        StartedUtc    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
        RanBy         = "$env:USERDOMAIN\$env:USERNAME"
        Server        = $script:WorkingServer
        Undone        = $false
        UndoneUtc     = $null
        User          = [ordered]@{
            Name              = [string]$User.Name
            DisplayName       = [string]$User.DisplayName
            SamAccountName    = [string]$User.SamAccountName
            UserPrincipalName = [string]$User.UserPrincipalName
            DistinguishedName = [string]$User.DistinguishedName
            Sid               = [string]$User.SID
        }
        Actions       = @()
    }

    $Stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:UndoRecordPath = Join-Path (Get-OffboardUndoFolder) "$($User.SamAccountName)-$Stamp.json"

    return $script:UndoRecord
}

function Save-OffboardUndoRecord {
    if (-not $script:UndoRecord -or -not $script:UndoRecordPath) { return }

    try {
        $script:UndoRecord | ConvertTo-Json -Depth 6 |
            Set-Content -LiteralPath $script:UndoRecordPath -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-AppLog "Could not write the undo record: $($_.Exception.Message)" 'WARNING'
    }
}

function Add-OffboardUndoAction {
    # Recorded and flushed to disk immediately, so a crash mid-run still leaves a
    # usable record of what had already happened.
    param($User, [string]$Type, $Details)

    [void](Start-OffboardUndoRecord -User $User)

    $script:UndoRecord.Actions += ,([ordered]@{
        Type    = $Type
        TimeUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
        Details = $Details
    })

    Save-OffboardUndoRecord

    if ($script:UndoRecordPath) {
        Write-AppLog "Undo record: $($script:UndoRecordPath)"
    }
}

function Get-OffboardUndoRecords {
    $Folder = Get-OffboardUndoFolder
    $Records = New-Object System.Collections.Generic.List[object]

    try {
        $Files = @(Get-ChildItem -LiteralPath $Folder -Filter "*.json" -File -ErrorAction Stop |
            Sort-Object LastWriteTime -Descending)
    }
    catch {
        return $Records.ToArray()
    }

    foreach ($File in $Files) {
        try {
            $Data = Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            $Summary = @()
            foreach ($Action in @($Data.Actions)) {
                switch ($Action.Type) {
                    'DisableAccount' { $Summary += 'disabled' }
                    'ClearPhone'     { $Summary += 'phone cleared' }
                    'RemoveGroups'   { $Summary += "$(@($Action.Details.Groups).Count) group(s) removed" }
                    default          { $Summary += $Action.Type }
                }
            }

            $Records.Add([pscustomobject]@{
                Path    = $File.FullName
                When    = $File.LastWriteTime
                User    = "$($Data.User.DisplayName) ($($Data.User.SamAccountName))"
                RanBy   = $Data.RanBy
                Actions = ($Summary -join ', ')
                Undone  = [bool]$Data.Undone
                Data    = $Data
            })
        }
        catch {
            Write-AppLog "Could not read undo record $($File.Name): $($_.Exception.Message)" 'WARNING'
        }
    }

    return $Records.ToArray()
}

function Invoke-OffboardUndoRecord {
    # Reverses in the opposite order to the run: groups back first, then the
    # phone numbers, then the account itself - so the account is only re-enabled
    # once its access is back.
    param($Record, [string]$Path)

    $Dn = [string]$Record.User.DistinguishedName
    $Server = $script:WorkingServer

    try {
        $Account = Get-ADUser -Identity $Dn -Server $Server -Properties Enabled,telephoneNumber,mobile,MemberOf -ErrorAction Stop
    }
    catch {
        Show-ErrorMessage ("The account in this record could not be found on $($Server).`r`n`r`n$Dn`r`n`r`n" +
            "If it has been moved or renamed, it has to be put back by hand.`r`n`r`n$($_.Exception.Message)")
        Write-AppLog "Undo failed - account not found: $Dn" 'ERROR'
        return
    }

    # The record names the account by DN and SID. If the SID no longer matches,
    # this is a different account now living at that DN - refuse.
    if ($Record.User.Sid -and ([string]$Account.SID.Value -ne [string]$Record.User.Sid)) {
        Show-ErrorMessage ("This record does not match the account now at:`r`n$Dn`r`n`r`n" +
            "Record SID: $($Record.User.Sid)`r`nAccount SID: $($Account.SID.Value)`r`n`r`nNothing was restored.")
        Write-AppLog "Undo refused - SID mismatch for $Dn." 'ERROR'
        return
    }

    $Restored = 0
    $Failed = 0

    # The checklist on screen belongs to whoever is loaded. Only clear its ticks
    # if that is the same account being restored - undoing an old run for someone
    # else must not touch the checklist in front of you.
    $SameUser = ($script:CurrentUser -and $script:CurrentUser.DistinguishedName -eq $Dn)

    foreach ($Action in @($Record.Actions)) {
        switch ($Action.Type) {

            'RemoveGroups' {
                $GroupsRestored = $false

                foreach ($Group in @($Action.Details.Groups)) {
                    # Offboarding never removes protected groups, so a record naming
                    # one was not written by it. Never add anyone to them from here.
                    try {
                        $GroupObject = Get-ADGroup -Identity $Group.DistinguishedName -Server $Server -Properties isCriticalSystemObject -ErrorAction Stop
                        $Protected = Get-MSToolkitCriticalGroupReason -Group $GroupObject
                        if ($Protected) {
                            $Failed++
                            Write-AppLog "Skipped $($Group.Name) - $Protected. Protected groups are never restored by Undo; add the account by hand if that is really intended." 'WARNING'
                            continue
                        }
                    }
                    catch {
                        $Failed++
                        Write-AppLog "Could not check $($Group.Name) before restoring it: $($_.Exception.Message)" 'ERROR'
                        continue
                    }

                    try {
                        Set-BusyState -Busy $true -StatusText "Restoring $($Group.Name)..."

                        Add-ADGroupMember -Identity $Group.DistinguishedName -Members $Dn `
                            -Server $Server -Confirm:$false -ErrorAction Stop

                        $Restored++
                        $GroupsRestored = $true
                        Write-AppLog "Restored membership of $($Group.Name)." 'SUCCESS'
                    }
                    catch {
                        if ("$($_.Exception.Message)" -match 'already a member') {
                            Write-AppLog "Already a member of $($Group.Name) - nothing to do."
                        }
                        else {
                            $Failed++
                            Write-AppLog "Could not restore $($Group.Name): $($_.Exception.Message)" 'ERROR'
                        }
                    }

                    [System.Windows.Forms.Application]::DoEvents()
                }

                if ($GroupsRestored -and $SameUser -and $script:chkGroups) {
                    $script:chkGroups.Checked = $false
                    Write-AppLog 'Security group step unticked - the memberships are back.'
                }
            }

            'ClearPhone' {
                $Replace = @{}
                if ($Action.Details.telephoneNumber) { $Replace['telephoneNumber'] = [string]$Action.Details.telephoneNumber }
                if ($Action.Details.mobile)          { $Replace['mobile'] = [string]$Action.Details.mobile }

                if ($Replace.Count -gt 0) {
                    try {
                        Set-BusyState -Busy $true -StatusText 'Restoring the phone number...'
                        Set-ADUser -Identity $Dn -Server $Server -Replace $Replace -Confirm:$false -ErrorAction Stop
                        $Restored++
                        Write-AppLog "Restored $($Replace.Keys -join ' and ')." 'SUCCESS'

                        if ($SameUser -and $script:chkPhone) {
                            $script:chkPhone.Checked = $false
                            Write-AppLog 'Phone number step unticked - the phone number is back.'
                        }
                    }
                    catch {
                        $Failed++
                        Write-AppLog "Could not restore the phone number: $($_.Exception.Message)" 'ERROR'
                    }
                }
            }

            'DisableAccount' {
                if ($Action.Details.PreviousEnabled) {
                    try {
                        Set-BusyState -Busy $true -StatusText 'Re-enabling the account...'
                        Enable-ADAccount -Identity $Dn -Server $Server -Confirm:$false -ErrorAction Stop

                        $Check = Get-ADUser -Identity $Dn -Server $Server -Properties Enabled -ErrorAction Stop

                        if ($Check.Enabled) {
                            $Restored++
                            Write-AppLog 'Account re-enabled.' 'SUCCESS'

                            if ($SameUser -and $script:chkDisable) {
                                $script:chkDisable.Checked = $false
                                Write-AppLog 'Disable step unticked - the account is enabled again.'
                            }
                        }
                        else {
                            $Failed++
                            Write-AppLog 'The account still reports as disabled after being re-enabled.' 'ERROR'
                        }
                    }
                    catch {
                        $Failed++
                        Write-AppLog "Could not re-enable the account: $($_.Exception.Message)" 'ERROR'
                    }
                }
                else {
                    Write-AppLog 'The account was already disabled before this run, so it has been left disabled.'
                }
            }
        }
    }

    # Mark the record so it is obvious which runs have been reversed.
    try {
        $Record.Undone = $true
        $Record.UndoneUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
        $Record | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-AppLog "Could not mark the record as undone: $($_.Exception.Message)" 'WARNING'
    }

    # The panel would otherwise still show the account as it was before the undo.
    if ($SameUser) {
        try {
            $script:CurrentUser = Get-ADUser -Identity $Dn -Server $Server `
                -Properties SID,Enabled,DisplayName,Description,Title,Department,Manager,telephoneNumber,mobile,DistinguishedName,isCriticalSystemObject -ErrorAction Stop

            $Phone = @($script:CurrentUser.telephoneNumber, $script:CurrentUser.mobile | Where-Object { $_ }) -join " / "
            if (-not $Phone) { $Phone = "(none set)" }

            $State = "Enabled"
            if (-not $script:CurrentUser.Enabled) { $State = "DISABLED" }

            $script:lblUserDetail.Text = "LOADED: $($script:CurrentUser.DisplayName)  |  $($script:CurrentUser.SamAccountName)  |  $State`r`n" +
                                         "Title: $($script:CurrentUser.Title)   Department: $($script:CurrentUser.Department)   Manager: $($script:ManagerDisplayName)`r`n" +
                                         "Phone: $Phone`r`n" +
                                         "$($script:CurrentUser.DistinguishedName)"
        }
        catch { }
    }

    Set-BusyState -Busy $false -StatusText "Undo complete - $Restored restored, $Failed failed."

    $Message = "Restored $Restored item(s) for $($Record.User.DisplayName)."
    if ($Failed -gt 0) { $Message += "`r`n$Failed could not be restored - see the Activity log." }
    $Message += "`r`n`r`nThis only reverses the Active Directory changes. Anything done outside Active Directory - MFA, physical access, Microsoft 365, Intune or other systems - has to be put back by hand."

    Show-InfoMessage $Message
}

function Show-OffboardUndoDialog {
    $Records = @(Get-OffboardUndoRecords)

    if ($Records.Count -eq 0) {
        Show-InfoMessage "There are no offboarding records to undo.`r`n`r`nRecords are written to:`r`n$(Get-OffboardUndoFolder)"
        return
    }

    $Dialog = New-Object System.Windows.Forms.Form
    $Dialog.Text = "Undo an offboarding"
    $Dialog.Size = New-Object System.Drawing.Size(980,520)
    $Dialog.StartPosition = "CenterParent"
    $Dialog.MinimumSize = New-Object System.Drawing.Size(760,400)

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Select a run to reverse. This puts back the Active Directory changes only: group memberships, phone numbers, and the account's enabled state."
    $Label.Location = New-Object System.Drawing.Point(12,12)
    $Label.Size = New-Object System.Drawing.Size(940,36)
    $Label.Anchor = "Top,Left,Right"
    $Dialog.Controls.Add($Label)

    $Grid = New-Object System.Windows.Forms.DataGridView
    $Grid.Location = New-Object System.Drawing.Point(12,54)
    $Grid.Size = New-Object System.Drawing.Size(940,360)
    $Grid.Anchor = "Top,Bottom,Left,Right"
    $Grid.ReadOnly = $true
    $Grid.AllowUserToAddRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.SelectionMode = "FullRowSelect"
    $Grid.MultiSelect = $false
    $Grid.AutoSizeColumnsMode = "Fill"
    [void]$Grid.Columns.Add("When","When")
    [void]$Grid.Columns.Add("User","User")
    [void]$Grid.Columns.Add("Actions","What was done")
    [void]$Grid.Columns.Add("RanBy","Run by")
    [void]$Grid.Columns.Add("Undone","Already undone")
    $Grid.Columns["Actions"].FillWeight = 160
    $Dialog.Controls.Add($Grid)

    foreach ($Record in $Records) {
        $Undone = "No"
        if ($Record.Undone) { $Undone = "Yes" }

        $Index = $Grid.Rows.Add($Record.When, $Record.User, $Record.Actions, $Record.RanBy, $Undone)
        $Grid.Rows[$Index].Tag = $Record

        if ($Record.Undone) {
            $Grid.Rows[$Index].DefaultCellStyle.ForeColor = (Get-MSToolkitThemePalette).MutedText
        }
    }

    $BtnUndo = New-Object System.Windows.Forms.Button
    $BtnUndo.Text = "Undo Selected"
    $BtnUndo.Location = New-Object System.Drawing.Point(12,428)
    $BtnUndo.Size = New-Object System.Drawing.Size(140,30)
    $BtnUndo.Anchor = "Bottom,Left"
    $BtnUndo.ForeColor = [System.Drawing.Color]::Red
    $BtnUndo.Add_Click({
        $Row = $Grid.SelectedRows | Select-Object -First 1
        if (-not $Row -or -not $Row.Tag) {
            Show-InfoMessage 'Select a run first.'
            return
        }

        $Record = $Row.Tag
        $Data = $Record.Data

        $What = @()
        foreach ($Action in @($Data.Actions)) {
            switch ($Action.Type) {
                'DisableAccount' { if ($Action.Details.PreviousEnabled) { $What += "  Re-enable the account" } }
                'ClearPhone'     { $What += "  Restore telephone '$($Action.Details.telephoneNumber)' and mobile '$($Action.Details.mobile)'" }
                'RemoveGroups'   { $What += "  Add back $(@($Action.Details.Groups).Count) group membership(s)" }
            }
        }

        $Note = ""
        if ($Record.Undone) { $Note = "`r`n`r`nThis run has already been undone once. Running it again is harmless but will do nothing new." }

        $Answer = [System.Windows.Forms.MessageBox]::Show(
            ("Reverse this offboarding?`r`n`r`n" +
             "User: $($Data.User.DisplayName) ($($Data.User.SamAccountName))`r`n" +
             "$($Data.User.DistinguishedName)`r`n" +
             "Run by $($Data.RanBy) at $($Data.StartedUtc) UTC`r`n`r`n" +
             "This will:`r`n$($What -join "`r`n")$Note`r`n`r`n" +
             "Active Directory only. MFA, physical access, Microsoft 365, Intune and everything else stay as they are."),
            "Confirm Undo",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        Write-AppLog "Undoing the offboarding of $($Data.User.SamAccountName) from $($Data.StartedUtc) UTC." 'WARNING'
        Invoke-OffboardUndoRecord -Record $Data -Path $Record.Path
        $Dialog.Close()
    })
    $Dialog.Controls.Add($BtnUndo)

    $BtnFolder = New-Object System.Windows.Forms.Button
    $BtnFolder.Text = "Copy Folder Path"
    $BtnFolder.Location = New-Object System.Drawing.Point(160,428)
    $BtnFolder.Size = New-Object System.Drawing.Size(140,30)
    $BtnFolder.Anchor = "Bottom,Left"
    $BtnFolder.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText((Get-OffboardUndoFolder))
        Write-AppLog "Undo folder path copied: $(Get-OffboardUndoFolder)"
    })
    $Dialog.Controls.Add($BtnFolder)

    $BtnClose = New-Object System.Windows.Forms.Button
    $BtnClose.Text = "Close"
    $BtnClose.Location = New-Object System.Drawing.Point(840,428)
    $BtnClose.Size = New-Object System.Drawing.Size(110,30)
    $BtnClose.Anchor = "Bottom,Right"
    $BtnClose.Add_Click({ $Dialog.Close() })
    $Dialog.Controls.Add($BtnClose)

    Apply-MSToolkitSharedTheme -Root $Dialog
    [void]$Dialog.ShowDialog($script:MainForm)
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ThemeMode = Resolve-MSToolkitThemeMode -ToolKey "OffboardUser" -RequestedMode $ThemeMode

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "Offboard User"
$MainForm.Size = New-Object System.Drawing.Size(1320,940)
$MainForm.MinimumSize = New-Object System.Drawing.Size(900,700)
$MainForm.StartPosition = "CenterScreen"
$script:MainForm = $MainForm

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = "Top"
$pnlHeader.Height = 72
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$pnlHeader.Tag = "TopBar"
$MainForm.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Offboard User"
$lblTitle.AutoSize = $true
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",15)
$lblTitle.Location = New-Object System.Drawing.Point(18,10)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Work top to bottom. The Active Directory steps tick themselves off when they succeed."
$lblSubtitle.AutoSize = $true
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(200,214,232)
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI",9)
$lblSubtitle.Location = New-Object System.Drawing.Point(20,42)
$pnlHeader.Controls.Add($lblSubtitle)

# --- status strip and activity log ---------------------------------------
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$script:lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblStatus.Text = "Select the user who is leaving."
[void]$statusStrip.Items.Add($script:lblStatus)
$MainForm.Controls.Add($statusStrip)

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Dock = "Bottom"
$pnlLog.Height = 150
$MainForm.Controls.Add($pnlLog)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Activity"
$lblLog.Location = New-Object System.Drawing.Point(12,4)
$lblLog.Size = New-Object System.Drawing.Size(120,18)
$pnlLog.Controls.Add($lblLog)

$script:txtLog = New-Object System.Windows.Forms.RichTextBox
$script:txtLog.Location = New-Object System.Drawing.Point(12,24)
$script:txtLog.Size = New-Object System.Drawing.Size(($MainForm.ClientSize.Width - 24),115)
$script:txtLog.Anchor = "Top,Bottom,Left,Right"
$script:txtLog.ReadOnly = $true
$script:txtLog.Font = New-Object System.Drawing.Font("Consolas",9)
$pnlLog.Controls.Add($script:txtLog)

# --- user selection -------------------------------------------------------
$pnlUser = New-Object System.Windows.Forms.Panel
$pnlUser.Dock = "Top"
$pnlUser.Height = 118
$MainForm.Controls.Add($pnlUser)

$lblUserCaption = New-Object System.Windows.Forms.Label
$lblUserCaption.Text = "User leaving:"
$lblUserCaption.Location = New-Object System.Drawing.Point(14,16)
$lblUserCaption.Size = New-Object System.Drawing.Size(90,20)
$pnlUser.Controls.Add($lblUserCaption)

$script:cboUser = New-Object System.Windows.Forms.ComboBox
$script:cboUser.Location = New-Object System.Drawing.Point(108,13)
$script:cboUser.Size = New-Object System.Drawing.Size(420,24)
$script:cboUser.DropDownStyle = "DropDown"

# Windows autocomplete is deliberately off. With it on there are two popups at
# once - the suggestion box over the dropdown list - and a click goes to the
# entry behind the suggestion rather than the one being read. The list is
# filtered here instead, so whatever is on screen is the only thing clickable.
$script:cboUser.AutoCompleteMode = "None"
$script:cboUser.AutoCompleteSource = "None"
$script:cboUser.MaxDropDownItems = 20
$script:cboUser.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { $_.SuppressKeyPress = $true } })

$pnlUser.Controls.Add($script:cboUser)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "Load User"
$btnLoad.Location = New-Object System.Drawing.Point(540,12)
$btnLoad.Size = New-Object System.Drawing.Size(110,26)
$btnLoad.Add_Click({ Invoke-LoadOffboardUser })
$pnlUser.Controls.Add($btnLoad)

# Everything the tool can do on its own, in one go. Red, because by the time it
# finishes the account is disabled and its group memberships are gone.
$btnAllLocal = New-Object System.Windows.Forms.Button
$btnAllLocal.Text = "Do All Local Actions"
$btnAllLocal.Location = New-Object System.Drawing.Point(660,12)
$btnAllLocal.Size = New-Object System.Drawing.Size(160,26)
$btnAllLocal.ForeColor = [System.Drawing.Color]::Red
$btnAllLocal.Font = New-Object System.Drawing.Font("Segoe UI Semibold",9)
$script:AllLocalTipText = "Exports the group memberships to $($script:OffboardExportFolder) first, then disables the account, removes the groups, clears the phone and replicates" + $(if ($script:OffboardSyncServer) { ", then requests the delta sync." } else { "." })
$btnAllLocal.Add_Click({
    param($ClickSender, $ClickArgs)

    try {
        Invoke-OffboardAllLocal
    }
    catch {
        Write-AppLog "Do All Local Actions failed: $($_.Exception.GetType().FullName)" 'ERROR'
        Write-AppLog "  $($_.Exception.Message)" 'ERROR'
        Set-BusyState -Busy $false -StatusText 'Do All Local Actions failed - see the Activity log.'
    }
})
$pnlUser.Controls.Add($btnAllLocal)
$script:StepToolTip.SetToolTip($btnAllLocal, $script:AllLocalTipText)
$script:StepButtons += $btnAllLocal

# Reverses a previous run - including one from days ago, since the records are
# kept on disk rather than in this window.
$btnUndo = New-Object System.Windows.Forms.Button
$btnUndo.Text = "Undo..."
$btnUndo.Location = New-Object System.Drawing.Point(830,12)
$btnUndo.Size = New-Object System.Drawing.Size(90,26)
$btnUndo.Add_Click({
    param($ClickSender, $ClickArgs)

    try {
        Show-OffboardUndoDialog
    }
    catch {
        Write-AppLog "Undo failed: $($_.Exception.GetType().FullName)" 'ERROR'
        Write-AppLog "  $($_.Exception.Message)" 'ERROR'
        Set-BusyState -Busy $false -StatusText 'Undo failed - see the Activity log.'
    }
})
$pnlUser.Controls.Add($btnUndo)
$script:StepToolTip.SetToolTip($btnUndo, "Reverse a previous offboarding: puts back group memberships, phone numbers and the account's enabled state. Records are kept for this admin account in $(Get-OffboardUndoFolder).")

$btnCopySummary = New-Object System.Windows.Forms.Button
$btnCopySummary.Text = "Copy Checklist"
$btnCopySummary.Location = New-Object System.Drawing.Point(928,12)
$btnCopySummary.Size = New-Object System.Drawing.Size(120,26)
$btnCopySummary.Add_Click({ Copy-OffboardSummary })
$pnlUser.Controls.Add($btnCopySummary)

$script:lblProgress = New-Object System.Windows.Forms.Label
$script:lblProgress.Text = ""
$script:lblProgress.Location = New-Object System.Drawing.Point(1060,17)
$script:lblProgress.Size = New-Object System.Drawing.Size(210,20)
$script:lblProgress.Anchor = "Top,Left,Right"
$script:lblProgress.Font = New-Object System.Drawing.Font("Segoe UI Semibold",9)
$pnlUser.Controls.Add($script:lblProgress)

$script:lblUserDetail = New-Object System.Windows.Forms.Label
$script:lblUserDetail.Text = ""
$script:lblUserDetail.Location = New-Object System.Drawing.Point(108,46)
$script:lblUserDetail.Size = New-Object System.Drawing.Size(1000,64)
$script:lblUserDetail.Tag = "LoadedUser"
$script:lblUserDetail.Anchor = "Top,Left,Right"
$script:lblUserDetail.Font = New-Object System.Drawing.Font("Segoe UI Semibold",8.5)
$pnlUser.Controls.Add($script:lblUserDetail)

# --- the checklist itself -------------------------------------------------
$pnlSteps = New-Object System.Windows.Forms.Panel
$pnlSteps.Dock = "Fill"
$pnlSteps.AutoScroll = $true
$pnlSteps.Padding = New-Object System.Windows.Forms.Padding(10)
$MainForm.Controls.Add($pnlSteps)
$pnlSteps.BringToFront()

$script:StepY = 10

function New-OffboardStep {
    # One step: numbered heading with its own checkbox, the instructions, any
    # sub-steps, and any buttons. Returns the checkbox so an action can tick it,
    # and so the progress count and the copied checklist can read it.
    param(
        [int]$Number,
        [string]$Title,
        [string]$Instructions,
        [string]$Warning,
        [string[]]$SubSteps = @(),
        [scriptblock[]]$ButtonActions = @(),
        [string[]]$ButtonLabels = @(),
        [string[]]$ButtonTips = @()
    )

    $Group = New-Object System.Windows.Forms.GroupBox
    $Group.Location = New-Object System.Drawing.Point(10,$script:StepY)
    $Group.Size = New-Object System.Drawing.Size(1100,10)
    $Group.Anchor = "Top,Left,Right"
    $Group.Text = ""
    $pnlSteps.Controls.Add($Group)

    $Check = New-Object System.Windows.Forms.CheckBox
    $Check.Text = "$Number. $Title"
    $Check.Location = New-Object System.Drawing.Point(12,14)
    $Check.Size = New-Object System.Drawing.Size(1060,22)
    $Check.Anchor = "Top,Left,Right"
    $Check.Font = New-Object System.Drawing.Font("Segoe UI Semibold",9.5)
    $Check.Add_CheckedChanged({ Update-StepProgress })

    # Ticking the step itself ticks all of its sub-steps, and clearing it clears
    # them. Click only fires on a real click, so the automatic ticks that come
    # from a finished action do not loop back through here.
    $Check.Add_Click({
        param($ClickSender, $ClickArgs)

        try {
            $Subs = $ClickSender.Tag
            if ($null -eq $Subs) { return }

            for ($SubIndex = 0; $SubIndex -lt $Subs.Count; $SubIndex++) {
                $Subs[$SubIndex].Checked = $ClickSender.Checked
            }
        }
        catch {
            Write-AppLog "Could not update the sub-steps: $($_.Exception.Message)" 'ERROR'
        }
    })

    $Group.Controls.Add($Check)

    $Body = New-Object System.Windows.Forms.Label
    $Body.Text = $Instructions
    $Body.Location = New-Object System.Drawing.Point(32,40)
    $Body.Size = New-Object System.Drawing.Size(1040,18)
    $Body.Anchor = "Top,Left,Right"
    $Body.AutoSize = $false
    $Group.Controls.Add($Body)

    # Measure the wrapped text so the box is tall enough for it. Measured against
    # a slightly narrower width than the label, and rounded up generously: the
    # estimate is optimistic at the right-hand edge and was clipping the last line.
    $Graphics = $Group.CreateGraphics()
    $Measured = $Graphics.MeasureString($Instructions, $Body.Font, 1010)
    $Body.Height = [int][math]::Ceiling($Measured.Height) + 12

    $Bottom = $Body.Bottom + 8

    if ($Warning) {
        # Stated in the step itself, not only in the tooltip - the order matters
        # and a tooltip is only seen by somebody who already hovered.
        $WarnLabel = New-Object System.Windows.Forms.Label
        $WarnLabel.Text = $Warning
        $WarnLabel.Location = New-Object System.Drawing.Point(32,$Bottom)
        $WarnLabel.Size = New-Object System.Drawing.Size(1040,18)
        $WarnLabel.Anchor = "Top,Left,Right"
        $WarnLabel.AutoSize = $false
        $WarnLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold",8.5)
        $WarnLabel.ForeColor = [System.Drawing.Color]::FromArgb(200,40,40)
        $WarnLabel.Tag = "StepWarning"

        $WarnMeasured = $Graphics.MeasureString($Warning, $WarnLabel.Font, 1010)
        $WarnLabel.Height = [int][math]::Ceiling($WarnMeasured.Height) + 10

        $Group.Controls.Add($WarnLabel)
        $Bottom = $WarnLabel.Bottom + 8
    }

    if ($SubSteps.Count -gt 0) {
        # Sub-steps are indented under the step, with round markers so they read
        # as parts of one step rather than steps in their own right. AutoCheck is
        # off so they toggle independently - a radio group would let only one be
        # ticked, which is not what a checklist needs.
        $SubFont = New-Object System.Drawing.Font("Segoe UI",8.5)
        $SubTop = $Bottom
        $SubList = @()

        foreach ($SubText in $SubSteps) {
            $Radio = New-Object System.Windows.Forms.RadioButton
            $Radio.Text = $SubText
            $Radio.Location = New-Object System.Drawing.Point(58,$Bottom)
            $Radio.Size = New-Object System.Drawing.Size(1010,18)
            $Radio.Anchor = "Top,Left,Right"
            $Radio.AutoSize = $false
            $Radio.AutoCheck = $false
            $Radio.Font = $SubFont

                $SubMeasured = $Graphics.MeasureString($SubText, $SubFont, 960)
            $Radio.Height = [int][math]::Ceiling($SubMeasured.Height) + 12

            $Radio.Tag = $Check
            $Radio.Add_Click({
                param($ClickSender, $ClickArgs)

                try {
                    $ClickSender.Checked = -not $ClickSender.Checked
                    Update-ParentStep -Parent $ClickSender.Tag
                }
                catch {
                    Write-AppLog "Sub-step toggle failed: $($_.Exception.Message)" 'ERROR'
                }
            })

            $Group.Controls.Add($Radio)
            $SubList += $Radio

            $Bottom = $Radio.Bottom + 2
        }

        # A thin rail down the left of the sub-steps, so the grouping is obvious
        # at a glance without relying on indentation alone.
        $Rail = New-Object System.Windows.Forms.Panel
        $Rail.Location = New-Object System.Drawing.Point(44,$SubTop)
        $Rail.Size = New-Object System.Drawing.Size(2,($Bottom - $SubTop - 2))
        $Rail.BackColor = [System.Drawing.Color]::FromArgb(142,170,219)
        $Rail.Tag = "SubStepRail"
        $Group.Controls.Add($Rail)

        $Check.Tag = $SubList
        $Bottom += 6
    }

    $Graphics.Dispose()

    if ($ButtonActions.Count -gt 0) {
        $X = 32
        for ($i = 0; $i -lt $ButtonActions.Count; $i++) {
            $Button = New-Object System.Windows.Forms.Button
            $Button.Text = $ButtonLabels[$i]
            $Button.Location = New-Object System.Drawing.Point($X,$Bottom)
            $Button.Size = New-Object System.Drawing.Size(190,28)

            $Action = $ButtonActions[$i]
            $Button.Add_Click({
                param($ClickSender, $ClickArgs)

                try {
                    & $ClickSender.Tag
                }
                catch {
                    # WinForms discards an exception that escapes a handler, so
                    # the button appears to do nothing. Always surface it.
                    Write-AppLog "Step failed: $($_.Exception.GetType().FullName)" 'ERROR'
                    Write-AppLog "  $($_.Exception.Message)" 'ERROR'
                    if ($_.InvocationInfo) {
                        Write-AppLog "  at line $($_.InvocationInfo.ScriptLineNumber): $("$($_.InvocationInfo.Line)".Trim())" 'ERROR'
                    }
                    Set-BusyState -Busy $false -StatusText 'The last step failed - see the Activity log.'
                }
            })
            $Button.Tag = $Action

            if ($i -lt $ButtonTips.Count -and $ButtonTips[$i]) {
                $script:StepToolTip.SetToolTip($Button, $ButtonTips[$i])
            }

            $Group.Controls.Add($Button)
            $script:StepButtons += $Button
            $X += 200
        }

        $Bottom += 36
    }

    $Group.Height = $Bottom + 8
    $script:StepY += $Group.Height + 10

    $script:StepChecks += $Check
    return $Check
}

function Update-ParentStep {
    # A step with sub-steps ticks itself once every sub-step is ticked, and
    # unticks the moment one is cleared.
    param($Parent)

    if (-not $Parent) { return }

    $Subs = $Parent.Tag
    if ($null -eq $Subs) { return }

    $AllDone = $true

    for ($Index = 0; $Index -lt $Subs.Count; $Index++) {
        if (-not $Subs[$Index].Checked) { $AllDone = $false; break }
    }

    $Parent.Checked = $AllDone
}

# ---------------------------------------------------------------------------
# The steps. Numbered as they are created, so a step left out because its
# setting is blank does not leave a gap in the numbering. Everything that names a
# system, server, product or document comes from MSToolkit Settings (Offboarding).
# ---------------------------------------------------------------------------
$script:NextStepNumber = 0
function Get-NextStepNumber {
    $script:NextStepNumber++
    return $script:NextStepNumber
}

function ConvertTo-OffboardLiteral {
    # Single-quoted PowerShell literal, for building a button action from a setting.
    param([string]$Text)
    return "'" + ($Text -replace "'", "''") + "'"
}

# MFA - only when an MFA system is set
if ($script:OffboardMfaName -or $script:OffboardMfaUrl) {
    $MfaName = if ($script:OffboardMfaName) { $script:OffboardMfaName } else { 'the MFA system' }
    $MfaLabels = @(); $MfaActions = @()
    if ($script:OffboardMfaUrl) {
        $MfaLabels += "Open $MfaName"
        $MfaActions += [scriptblock]::Create("Open-OffboardLink -Url $(ConvertTo-OffboardLiteral $script:OffboardMfaUrl) -Label $(ConvertTo-OffboardLiteral $MfaName)")
    }
    $null = New-OffboardStep -Number (Get-NextStepNumber) -Title "MFA - disable the user and remove their enrolled devices" `
        -Instructions "Sign in to $MfaName and work through both parts." `
        -SubSteps @(
            "Disable the user in $MfaName.",
            "Remove their enrolled phone or authenticator devices."
        ) `
        -ButtonLabels $MfaLabels `
        -ButtonActions $MfaActions
}

# Physical access - only when an access system is set
if ($script:OffboardAccessName -or $script:OffboardAccessHost) {
    $AccessName = if ($script:OffboardAccessName) { $script:OffboardAccessName } else { 'the access control system' }
    $AccessLabels = @(); $AccessActions = @()
    if ($script:OffboardAccessHost) {
        $AccessLabels += $(if ($script:OffboardAccessHost -match '^https?://') { "Open $AccessName" } else { "Remote Desktop to $($script:OffboardAccessHost)" })
        $AccessActions += { Open-OffboardAccessSystem }
    }
    $null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Physical access - disable the badge or access card" `
        -Instructions "Disable the user's badge or access card in $AccessName." `
        -ButtonLabels $AccessLabels `
        -ButtonActions $AccessActions
}

# Disable the AD account
$script:chkDisable = New-OffboardStep -Number (Get-NextStepNumber) -Title "Active Directory - disable the account" `
    -Instructions "Load the user above, then disable the account. This box ticks itself once the account is actually disabled." `
    -ButtonLabels @("Disable User") `
    -ButtonActions @({ Invoke-DisableOffboardUser })

# Security groups
$script:chkGroups = New-OffboardStep -Number (Get-NextStepNumber) -Title "Active Directory - remove security group memberships" `
    -Instructions "Protected groups such as Domain Admins are identified by SID and are never removed here." `
    -Warning "EXPORT THE LIST FIRST - once the memberships are removed, the CSV is the only record of what this user had. Do All Local Actions exports to $($script:OffboardExportFolder) for you before it removes anything." `
    -ButtonLabels @("Export Groups to CSV","Remove All Groups") `
    -ButtonTips @(
        "Do this FIRST. Saves every group the user belongs to, with the DN of each, so the memberships can be restored or audited later.",
        "Removes every security group membership except protected groups. Export first - once these are gone the only record is the CSV."
    ) `
    -ButtonActions @({ Export-OffboardGroups }, { Remove-OffboardGroups })

# Distribution groups
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Remove distribution group memberships" `
    -Instructions ("Opens M365 Distro Compare, signs in, and loads the departing user in Manage User Groups. Remove them from every distribution group they belong to.") `
    -Warning "EXPORT THE LIST FIRST - use Export Distribution Groups in the Manage User Groups window before removing anything." `
    -ButtonLabels @("Open M365 Distro Compare") `
    -ButtonTips @("Opens, signs in, and loads this user in Manage User Groups automatically. Click Export Distribution Groups BEFORE removing anything - the list is gone once the memberships are.") `
    -ButtonActions @({ Start-OffboardM365Tool -ScriptName "M365-Distribution-Group-Compare.ps1" -Label "M365 Distro Compare" })

# Microsoft 365 groups
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Remove Microsoft 365 group memberships" `
    -Instructions ("Opens M365 Group Compare, signs in, and loads the departing user in Manage User Groups. Remove them from their Entra security groups. " +
        "Groups synced from on-premises AD are handled in the Active Directory step above, not here.") `
    -Warning "EXPORT THE LIST FIRST - use Export Security Groups in the Manage User Groups window before removing anything." `
    -ButtonLabels @("Open M365 Group Compare") `
    -ButtonTips @("Opens, signs in, and loads this user in Manage User Groups automatically. Click Export Security Groups BEFORE removing anything - the list is gone once the memberships are.") `
    -ButtonActions @({ Start-OffboardM365Tool -ScriptName "M365-Group-Compare.ps1" -Label "M365 Group Compare" })

# Phone numbers
$script:chkPhone = New-OffboardStep -Number (Get-NextStepNumber) -Title "Active Directory - clear the phone number, if applicable" `
    -Instructions "Removes the telephone and mobile numbers from the account. If the account has no numbers set, the step is ticked with nothing to do." `
    -ButtonLabels @("Clear Phone Number") `
    -ButtonActions @({ Invoke-ClearOffboardPhone })

# Replication
$script:chkReplicate = New-OffboardStep -Number (Get-NextStepNumber) -Title "Replicate the selected domain controller" `
    -Instructions ("Pushes the change out so directory sync reads the disabled account rather than a stale copy. " +
        "repadmin /syncall /AdeP is enterprise-wide, so syncing the selected DC still reaches every DC in the forest.") `
    -ButtonLabels @("Replicate Selected DC") `
    -ButtonActions @({ Invoke-OffboardReplicate })

# Delta sync - only when an Entra Connect server is set
$script:chkDeltaSync = $null
if ($script:OffboardSyncServer) {
    $script:chkDeltaSync = New-OffboardStep -Number (Get-NextStepNumber) -Title "Entra Connect delta sync" `
        -Instructions "Requests a delta sync on $($script:OffboardSyncServer) so the disabled account reaches Microsoft 365. The cycle runs for a few minutes after the request returns." `
        -ButtonLabels @("Run Delta Sync") `
        -ButtonActions @({ Invoke-OffboardDeltaSync })
}

# The ticket, so other teams can start
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Update the offboarding ticket" `
    -Instructions ("Note that the account is disabled and synced, so other teams know to begin removing access to their own systems. " +
        "Everything above is done at this point; everything below can run alongside them.")

# Devices in Intune
$DeviceSubSteps = @()
if ($script:OffboardRmmName -and $script:OffboardRmmReport) {
    $DeviceSubSteps += "Run the $($script:OffboardRmmReport) report from $($script:OffboardRmmName) against the device and WAIT for it to arrive - once the PC is reset that information is gone."
}
$DeviceSubSteps += @(
    "Personal device: do the App Selective Wipe first (Apps > App Selective Wipe).",
    "Personal device: then go to the device and Retire it.",
    "Where applicable: Users > the user > Devices > expand the device > Remove company data.",
    "Company PC being kept: reset it from Intune using Autopilot Reset."
)
if ($script:OffboardRmmName -and $script:OffboardRmmReport) {
    $DeviceSubSteps += "After a new user is assigned to that PC, run the $($script:OffboardRmmReport) report from $($script:OffboardRmmName) again. If nobody is being assigned, this one is already done."
}
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Intune - personal and company devices" `
    -Instructions $(if ($script:OffboardRmmName -and $script:OffboardRmmReport) { "Take the report before anything is wiped, then deal with the devices." } else { "Deal with each of the user's devices." }) `
    -SubSteps $DeviceSubSteps `
    -ButtonLabels @("Open Intune Devices") `
    -ButtonActions @({ Open-OffboardLink -Url "https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/overview" -Label "Intune devices" })

# Device decommission
$DecomLabels = @(); $DecomActions = @(); $DecomTips = @()
if ($script:OffboardDeviceDocument) {
    $DecomLabels += "Open Decommission Checklist"
    $DecomActions += { Open-OffboardDocument -Value $script:OffboardDeviceDocument -Label "The device decommission checklist" }
    $DecomTips += "Opens $($script:OffboardDeviceDocument) under your own account."
}
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "PC / laptop decommission" `
    -Instructions ("If a company PC or laptop is being returned, retained or disposed of, work through your decommission process for that device - " +
        "keeping or reusing it, or removing it completely, along with its Intune, Entra, Active Directory and inventory records.") `
    -Warning "A device being DISPOSED of usually does not need resetting first if it will be securely destroyed and certified." `
    -ButtonLabels $DecomLabels `
    -ButtonTips $DecomTips `
    -ButtonActions $DecomActions

# Mailbox, delegation and licences
$MailSubSteps = @(
    "Discuss with the manager how long they need delegated access to the mailbox.",
    "Convert the mailbox to a shared mailbox in Exchange Online if the manager wants it for longer than 30/60/90 days.",
    "Delegate the shared mailbox to the manager - or delegate the full mailbox if it is only being kept for 30/60/90 days.",
    "Set a calendar reminder to remove the manager's delegated access at the agreed time."
)
if ($script:OffboardArchiveName) {
    $MailSubSteps += @(
        "If the manager wants the mailbox kept indefinitely, use $($script:OffboardArchiveName) to preserve the required email data.",
        "Find the email and mailbox data that needs to be retained in $($script:OffboardArchiveName).",
        "Restore it to an alternate location using an unlicensed account as the destination - create that unlicensed account in Microsoft 365 first."
    )
}
$MailSubSteps += @(
    "Forward the Teams phone number, if that was requested.",
    "Set a calendar reminder to remove the Microsoft 365 licences after the agreed period."
)
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Microsoft 365 - mailbox, delegation and licences" `
    -Instructions "Agree the retention period with the manager first - everything below follows from that answer." `
    -SubSteps $MailSubSteps `
    -ButtonLabels @("Open Microsoft 365 Admin") `
    -ButtonActions @({ Open-OffboardLink -Url "https://admin.cloud.microsoft/?#/homepage" -Label "the Microsoft 365 admin center" })

# Reminder to move the account
$DisabledOUText = if ($script:OffboardDisabledOU) { "the Disabled OU ($($script:OffboardDisabledOU))" } else { "your disabled-accounts OU" }
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Calendar reminder - move the account to the disabled OU" `
    -Instructions ("Set a reminder to move the Active Directory account to $DisabledOUText in 30/60/90 days or longer, depending on the user. " +
        "If the mailbox data has been preserved and the Teams phone does not need forwarding, you can usually move it now.") `
    -Warning "Check directory sync scope first: if the disabled OU is not synced to Microsoft Entra ID, moving the account stops it syncing and breaks the cloud side - mailbox, licences and delegation."

# Line-of-business systems - only when some are listed
$LobList = Get-OffboardLobList
if ($LobList.Count -gt 0) {
    $LobSubSteps = @(); $LobLabels = @(); $LobActions = @()
    foreach ($Lob in $LobList) {
        $Line = "Remove the user's $($Lob.Name) account - or tell its owner it can be removed."
        if ($Lob.Note) { $Line += " $($Lob.Note)" }
        $LobSubSteps += $Line

        # Room for five buttons on one row.
        if ($Lob.Url -and $LobLabels.Count -lt 5) {
            $LobLabels += "Open $($Lob.Name)"
            $LobActions += [scriptblock]::Create("Open-OffboardLink -Url $(ConvertTo-OffboardLiteral $Lob.Url) -Label $(ConvertTo-OffboardLiteral $Lob.Name)")
        }
    }
    $null = New-OffboardStep -Number (Get-NextStepNumber) -Title "Line-of-business application accounts" `
        -Instructions "Remove the user from each system below, or report to whoever administers it that the account can now be removed." `
        -SubSteps $LobSubSteps `
        -ButtonLabels $LobLabels `
        -ButtonActions $LobActions
}

# The record
$FileLabels = @(); $FileActions = @(); $FileTips = @()
if ($script:OffboardChecklistDocument) {
    $FileLabels += "Open Offboarding Checklist"
    $FileActions += { Open-OffboardDocument -Value $script:OffboardChecklistDocument -Label "The offboarding checklist" }
    $FileTips += "Opens $($script:OffboardChecklistDocument) under your own account."
}
$FileText = "Fill in and keep a record of this offboarding."
if ($script:OffboardChecklistLocation) { $FileText = "Fill in the checklist and file a copy in: $($script:OffboardChecklistLocation)." }
$null = New-OffboardStep -Number (Get-NextStepNumber) -Title "File the completed checklist" `
    -Instructions ("$FileText Finish, below, copies this window's checklist and opens it as a text file for the ticket.") `
    -ButtonLabels $FileLabels `
    -ButtonTips $FileTips `
    -ButtonActions $FileActions

# Finish: copy the checklist, open it as a text file, and close the tool. Deliberately
# outside the numbered steps - it is the way out, not another step.
$pnlFinish = New-Object System.Windows.Forms.Panel
$pnlFinish.Location = New-Object System.Drawing.Point(10,($script:StepY + 6))
$pnlFinish.Size = New-Object System.Drawing.Size(1100,64)
$pnlFinish.Anchor = "Top,Left,Right"
$pnlSteps.Controls.Add($pnlFinish)

$btnFinish = New-Object System.Windows.Forms.Button
$btnFinish.Text = "Finish"
$btnFinish.Location = New-Object System.Drawing.Point(2,14)
$btnFinish.Size = New-Object System.Drawing.Size(200,34)
$btnFinish.Font = New-Object System.Drawing.Font("Segoe UI Semibold",10)
$btnFinish.Add_Click({
    param($ClickSender, $ClickArgs)

    try {
        Invoke-OffboardFinish
    }
    catch {
        Write-AppLog "Finish failed: $($_.Exception.GetType().FullName)" 'ERROR'
        Write-AppLog "  $($_.Exception.Message)" 'ERROR'
    }
})
$pnlFinish.Controls.Add($btnFinish)

$lblFinish = New-Object System.Windows.Forms.Label
$lblFinish.Text = "Copies the checklist, opens it as a text file to save or paste into the ticket, then closes this window."
$lblFinish.Location = New-Object System.Drawing.Point(212,22)
$lblFinish.Size = New-Object System.Drawing.Size(880,20)
$lblFinish.Anchor = "Top,Left,Right"
$pnlFinish.Controls.Add($lblFinish)

$script:StepY += 76

Update-StepProgress

# --- theme, startup, show -------------------------------------------------
New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "OffboardUser"
Apply-MSToolkitSharedTheme -Root $MainForm

# Docking follows z-order, and Controls.Add alone put the user panel above the
# header. Set the order explicitly: highest index docks first, so the header is
# outermost at the top and the status strip outermost at the bottom.
$MainForm.Controls.SetChildIndex($pnlSteps, 0)
$MainForm.Controls.SetChildIndex($pnlLog, 1)
$MainForm.Controls.SetChildIndex($statusStrip, 2)
$MainForm.Controls.SetChildIndex($pnlUser, 3)
$MainForm.Controls.SetChildIndex($pnlHeader, 4)

$MainForm.Add_Shown({
    $MainForm.Activate()

    Write-AppLog 'Offboard User ready.'

    # The shared startup block above already picked a DC that answers on ADWS,
    # preferring the one MSToolkit passed in.
    $script:WorkingServer = $script:CurrentServer

    if (-not $script:WorkingServer) {
        Write-AppLog 'No domain controller answered on port 9389. The Active Directory steps will not work.' 'ERROR'
        $script:lblStatus.Text = 'No domain controller available.'
        return
    }

    Write-AppLog "Using domain controller: $($script:WorkingServer)"

    if ($script:ServerNote) {
        Write-AppLog $script:ServerNote 'WARNING'
    }

    $Loaded = Initialize-MSToolkitUserPicker -Combos @($script:cboUser)
    if ($Loaded -ge 0) {
        Write-AppLog "Loaded $Loaded user(s) into the list."
    }

    $script:lblStatus.Text = 'Select the user who is leaving, then work down the list.'
})

# Type-ahead filtering on every editable dropdown. Its own handler, not folded
# into another, so it runs on every launch rather than only when the tool is
# started with parameters - and so a failure here cannot stop the rest of
# start-up. Multiple Shown handlers chain.
$MainForm.Add_Shown({ Register-MSToolkitComboFiltersOn -Root $MainForm })

[void]$MainForm.ShowDialog()
