[CmdletBinding()]
param(
    [ValidateSet("Light","Dark")]
    [string]$ThemeMode,

    [string]$DefaultServer
)

$ErrorActionPreference = "Stop"

# Do not create the AD: PowerShell drive. This tool never uses it, and building it
# means connecting to a domain controller before anything else can run.
$env:ADPS_LoadDefaultDrive = 0
Import-Module ActiveDirectory

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
            $ToolProperty = "Theme_NewADUser"
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



# Hide PowerShell console window
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WindowHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$ConsolePtr = [WindowHelper]::GetConsoleWindow()

if ($ConsolePtr -ne [IntPtr]::Zero -and $host.Name -notlike "*ISE*") {
    [WindowHelper]::ShowWindow($ConsolePtr, 0) | Out-Null
}

function Get-CurrentADDomain {
    # Read locally first; no DC is contacted, so this cannot stall on an unreachable one.
    $LocalDomain = Get-MSToolkitLocalDomainName
    if ($LocalDomain) {
        return $LocalDomain
    }

    try {
        $DomainInfo = Get-ADDomain -Current LocalComputer -ErrorAction Stop

        if (-not $DomainInfo.DNSRoot) {
            throw "Active Directory did not return a DNS domain name."
        }

        return [string]$DomainInfo.DNSRoot
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Unable to determine the Active Directory domain for this computer.`r`n`r`n$($_.Exception.Message)",
            "Domain Discovery Error",
            "OK",
            "Error"
        ) | Out-Null

        exit 1
    }
}

$Domain = Get-CurrentADDomain

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

# New account naming - all from MSToolkit Settings (New account naming).
# UPN domain and Primary SMTP domain are required before an account can be created;
# the two onmicrosoft domains and Company are optional and simply skipped when blank.
$UPNDomain = Get-MSToolkitSetting -Name "UpnDomain"
$MailDomain = Get-MSToolkitSetting -Name "MailDomain"
$OnMicrosoftAliasDomain = Get-MSToolkitSetting -Name "OnMicrosoftDomain"
$OnMicrosoftMailAliasDomain = Get-MSToolkitSetting -Name "OnMicrosoftMailDomain"
$CompanyName = Get-MSToolkitSetting -Name "CompanyName"

# Derived from the DNS name (contoso.local -> DC=contoso,DC=local), so no DC is contacted.
# The Domain setting, when filled in, is used instead of the detected domain.
$DomainDN = 'DC=' + (((Get-MSToolkitSetting -Name "Domain" -Default $Domain) -split '\.') -join ',DC=')

# Default OU for new accounts: the Users OU from MSToolkit Settings, or the domain root.
# With no Users OU set, the picker starts at the domain root, but an account is never
# created there - an OU has to be picked or typed (checked in New-MSToolkitADUser).
$UsersOUConfigured = [bool](Get-MSToolkitSetting -Name "OUUsers")
$BaseOU = Get-MSToolkitSetting -Name "OUUsers" -Default $DomainDN

function Get-MSToolkitDomainControllers {
    # DC names over plain LDAP (port 389), which works against every DC and never
    # touches ADWS. Falls back to the original ADWS query if that fails.
    $LdapHostNames = @(Get-MSToolkitDomainControllerHostNames -DomainName $Domain)
    if ($LdapHostNames.Count -gt 0) {
        return @($LdapHostNames | ForEach-Object { ($_ -split '\.')[0].ToUpper() } | Sort-Object -Unique)
    }

    try {
        $Controllers = @(
            Get-ADDomainController `
                -Filter * `
                -Server $Domain `
                -ErrorAction Stop |
                Sort-Object Name |
                Select-Object -ExpandProperty Name
        )

        if ($Controllers.Count -eq 0) {
            throw "No domain controllers were returned for $Domain."
        }

        return $Controllers
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Unable to discover domain controllers for $Domain.`r`n`r`n$($_.Exception.Message)",
            "Domain Controller Discovery Error",
            "OK",
            "Error"
        ) | Out-Null
        exit 1
    }
}

$DCs = @(Get-MSToolkitDomainControllers)

function Show-Message {
    param(
        [string]$Message,
        [string]$Title = "New AD User"
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        "OK",
        "Information"
    ) | Out-Null
}

function Show-Error {
    param([string]$Message)

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "Error",
        "OK",
        "Error"
    ) | Out-Null
}

function Get-SelectedServer {
    if ($DCDropdown.SelectedItem) {
        return "$($DCDropdown.SelectedItem).$Domain"
    }

    return $Domain
}

function Write-OutputBox {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black
    )

    $OutputBox.SelectionStart = $OutputBox.TextLength
    $OutputBox.SelectionLength = 0
    $OutputBox.SelectionColor = (Convert-MSToolkitThemeColor $color)
    $OutputBox.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Text`r`n")
    $OutputBox.SelectionColor = $OutputBox.ForeColor
    $OutputBox.ScrollToCaret()
}


function Write-StartupSection {
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )

    $Palette = Get-MSToolkitThemePalette
    $NormalFont = $OutputBox.Font
    $BoldFont = New-Object System.Drawing.Font(
        $NormalFont.FontFamily,
        $NormalFont.Size,
        [System.Drawing.FontStyle]::Bold
    )

    try {
        $OutputBox.AppendText("`r`n")
        $OutputBox.SelectionStart = $OutputBox.TextLength
        $OutputBox.SelectionLength = 0
        $OutputBox.SelectionFont = $BoldFont
        $OutputBox.SelectionColor = $Palette.Section
        $OutputBox.AppendText("============================================================`r`n")
        $OutputBox.AppendText("  $Title`r`n")
        $OutputBox.AppendText("============================================================`r`n")
    }
    finally {
        $OutputBox.SelectionFont = $NormalFont
        $OutputBox.SelectionColor = $OutputBox.ForeColor
        $BoldFont.Dispose()
        $OutputBox.ScrollToCaret()
    }
}

function Write-StartupField {
    param(
        [Parameter(Mandatory)]
        [string]$Label,
        [AllowEmptyString()]
        [string]$Value,
        [System.Drawing.Color]$ValueColor
    )

    $Palette = Get-MSToolkitThemePalette
    if (-not $PSBoundParameters.ContainsKey('ValueColor')) {
        $ValueColor = $Palette.Text
    }

    $NormalFont = $OutputBox.Font
    $BoldFont = New-Object System.Drawing.Font(
        $NormalFont.FontFamily,
        $NormalFont.Size,
        [System.Drawing.FontStyle]::Bold
    )

    try {
        $OutputBox.SelectionStart = $OutputBox.TextLength
        $OutputBox.SelectionLength = 0
        $OutputBox.SelectionFont = $NormalFont
        $OutputBox.SelectionColor = $Palette.MutedText
        $OutputBox.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ")

        $OutputBox.SelectionFont = $BoldFont
        $OutputBox.SelectionColor = $Palette.Section
        $OutputBox.AppendText("${Label}: ")

        $OutputBox.SelectionFont = $NormalFont
        $OutputBox.SelectionColor = $ValueColor
        $OutputBox.AppendText("$Value`r`n")
    }
    finally {
        $OutputBox.SelectionFont = $NormalFont
        $OutputBox.SelectionColor = $OutputBox.ForeColor
        $BoldFont.Dispose()
        $OutputBox.ScrollToCaret()
    }
}

function Write-ResultSeparator {
    $OutputBox.AppendText("`r`n")
    $OutputBox.SelectionStart = $OutputBox.TextLength
    $OutputBox.SelectionLength = 0
    $OutputBox.SelectionColor = Convert-MSToolkitThemeColor ([System.Drawing.Color]::FromArgb(140,140,140))
    $OutputBox.AppendText("============================================================`r`n")
    $OutputBox.SelectionColor = $OutputBox.ForeColor
    $OutputBox.ScrollToCaret()
}

function Write-VerificationResult {
    param(
        [string]$Item,
        [bool]$Passed,
        [string]$Detail = ''
    )

    $Color = if ($Passed) { (Get-MSToolkitThemePalette).Success } else { (Get-MSToolkitThemePalette).Danger }
    $Status = if ($Passed) { 'Verified' } else { 'FAILED' }
    $Message = "$Item - $Status"
    if ($Detail) { $Message += ": $Detail" }
    Write-OutputBox $Message $Color
}

# ---------------------------------------------------------------------------
# Cascading OU picker. Level 0 lists the child OUs of the department base OU;
# choosing one adds another dropdown for its children, and so on. The current
# selection is tracked in $script:OUPickerCurrentDN.
# ---------------------------------------------------------------------------
function Get-MSToolkitOUNameFromDN {
    param([string]$DistinguishedName)

    if ($DistinguishedName -match '^OU=([^,]+)') {
        return $Matches[1]
    }

    return $DistinguishedName
}

function New-MSToolkitOUPickerItem {
    param(
        [string]$Name,
        [string]$DN,
        [bool]$IsStay = $false
    )

    $Item = [pscustomobject]@{
        Name   = $Name
        DN     = $DN
        IsStay = $IsStay
    }

    $Item | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Name } -Force
    return $Item
}

function Remove-MSToolkitOUPickerLevelsFrom {
    param([int]$Level)

    $ToRemove = @(
        $script:OUPickerFlow.Controls |
        Where-Object { $_ -is [System.Windows.Forms.ComboBox] -and $_.Tag -is [hashtable] -and $_.Tag.Level -ge $Level }
    )

    foreach ($Control in $ToRemove) {
        $script:OUPickerFlow.Controls.Remove($Control)
        $Control.Dispose()
    }
}

function Add-MSToolkitOUPickerLevel {
    param(
        [string]$ParentDN,
        [int]$Level
    )

    $Children = @()
    try {
        $Children = @(
            Get-ADOrganizationalUnit -SearchBase $ParentDN -SearchScope OneLevel -Filter * -Server $script:OUPickerServer -ErrorAction Stop |
            Sort-Object Name
        )
    }
    catch {
        return
    }

    if ($Children.Count -eq 0) {
        return
    }

    $ParentName = if ($ParentDN -match '^OU=') { Get-MSToolkitOUNameFromDN -DistinguishedName $ParentDN } else { "domain root" }

    $Combo = New-Object System.Windows.Forms.ComboBox
    $Combo.DropDownStyle = "DropDownList"
    $Combo.Width = 200
    $Combo.Margin = New-Object System.Windows.Forms.Padding(0,0,6,6)
    $Combo.Tag = @{ Level = $Level; ParentDN = $ParentDN }

    $null = $Combo.Items.Add((New-MSToolkitOUPickerItem -Name "(stay in $ParentName)" -DN $ParentDN -IsStay $true))

    foreach ($Child in $Children) {
        $null = $Combo.Items.Add((New-MSToolkitOUPickerItem -Name $Child.Name -DN $Child.DistinguishedName))
    }

    $Combo.SelectedIndex = 0

    $Combo.Add_SelectedIndexChanged({
        $ThisCombo = $this
        $ThisLevel = [int]$ThisCombo.Tag.Level
        $Selected = $ThisCombo.SelectedItem

        Remove-MSToolkitOUPickerLevelsFrom -Level ($ThisLevel + 1)

        if ($Selected -and -not $Selected.IsStay) {
            $script:OUPickerCurrentDN = $Selected.DN
            Add-MSToolkitOUPickerLevel -ParentDN $Selected.DN -Level ($ThisLevel + 1)
        }
        else {
            $script:OUPickerCurrentDN = [string]$ThisCombo.Tag.ParentDN
        }

        Update-MSToolkitOUPathLabel

        Apply-MSToolkitSharedTheme -Root $script:OUPickerFlow
    })

    $script:OUPickerFlow.Controls.Add($Combo)
}

function Initialize-MSToolkitOUPicker {
    param(
        [System.Windows.Forms.FlowLayoutPanel]$Flow,
        [System.Windows.Forms.Label]$PathLabel,
        [string]$Server,
        [string]$RootDN
    )

    $script:OUPickerFlow = $Flow
    $script:OUPickerPathLabel = $PathLabel
    $script:OUPickerServer = $Server
    $script:OUPickerCurrentDN = $RootDN

    $Flow.Controls.Clear()
    Add-MSToolkitOUPickerLevel -ParentDN $RootDN -Level 0

    Update-MSToolkitOUPathLabel
}

function Update-MSToolkitOUPathLabel {
    # The line under the Department OU picker. While the selection is the domain root
    # and no Users OU is set in MSToolkit Settings, it says so in the warning colour,
    # because an account cannot be created there. Otherwise it shows the selected DN.
    if (-not $script:OUPickerPathLabel) { return }

    $Palette = Get-MSToolkitThemePalette
    $SelectedDN = [string]$script:OUPickerCurrentDN

    if ((-not $UsersOUConfigured) -and ($SelectedDN -ieq $DomainDN)) {
        $script:OUPickerPathLabel.Text = "Selected OU: domain root - no Users OU is set in MSToolkit Settings, so pick an OU."
        $script:OUPickerPathLabel.ForeColor = $Palette.Warning
    }
    else {
        $script:OUPickerPathLabel.Text = "Selected OU: $SelectedDN"
        $script:OUPickerPathLabel.ForeColor = $Palette.MutedText
    }
}

function Reset-MSToolkitOUPicker {
    param([string]$Server)

    if ($script:OUPickerFlow) {
        Initialize-MSToolkitOUPicker -Flow $script:OUPickerFlow -PathLabel $script:OUPickerPathLabel -Server $Server -RootDN $BaseOU
    }
}

function Resolve-MSToolkitOUPathUnder {
    param(
        [string]$BaseDN,
        [string]$RelativePath,
        [string]$Server
    )

    $Segments = @(
        [regex]::Split($RelativePath, '\s*[\\/|;:>]\s*') |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($Segments.Count -eq 0) {
        return $BaseDN
    }

    $Current = $BaseDN

    foreach ($Segment in $Segments) {
        $Next = $null

        try {
            $Next = Get-ADOrganizationalUnit -SearchBase $Current -SearchScope OneLevel -Filter { Name -eq $Segment } -Server $Server -ErrorAction Stop |
                Select-Object -First 1 -ExpandProperty DistinguishedName
        }
        catch {
            return $null
        }

        if ([string]::IsNullOrWhiteSpace($Next)) {
            return $null
        }

        $Current = $Next
    }

    return $Current
}

function Get-OUPath {
    param(
        [string]$ChildOU,
        [string]$SelectedOU,
        [string]$Server
    )

    $Base = if (-not [string]::IsNullOrWhiteSpace($SelectedOU)) { $SelectedOU } else { $BaseOU }
    $Typed = "$ChildOU".Trim()

    # Nothing typed - use whatever the dropdowns are pointing at.
    if ([string]::IsNullOrWhiteSpace($Typed)) {
        return $Base
    }

    # A full DN always overrides the dropdown selection.
    if ($Typed -like "OU=*") {
        return $Typed
    }

    # Otherwise treat it as a path underneath the current selection.
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $Relative = Resolve-MSToolkitOUPathUnder -BaseDN $Base -RelativePath $Typed -Server $Server

        if (-not [string]::IsNullOrWhiteSpace($Relative)) {
            return $Relative
        }
    }

    # Fall back to the original behaviour so a single child name still works.
    return "OU=$Typed,$Base"
}


function ConvertTo-EmailSafeName {
    param([string]$Value)

    return ($Value.Trim().ToLowerInvariant() -replace "[^a-z0-9]", "")
}

function Get-GeneratedUsername {
    param(
        [string]$FirstName,
        [string]$LastName
    )

    $FirstNamePart = ConvertTo-EmailSafeName -Value $FirstName
    $LastNamePart = ConvertTo-EmailSafeName -Value $LastName

    if (-not $FirstNamePart -or -not $LastNamePart) {
        return ""
    }

    return "$FirstNamePart$($LastNamePart.Substring(0,1))"
}

function Update-GeneratedUsername {
    if ($null -eq $SamBox -or $null -eq $FirstNameBox -or $null -eq $LastNameBox) {
        return
    }

    $SamBox.Text = Get-GeneratedUsername -FirstName $FirstNameBox.Text -LastName $LastNameBox.Text
}

function Clear-FormFields {
    $FirstNameBox.Clear()
    $LastNameBox.Clear()
    $SamBox.Clear()
    $PasswordBox.Clear()
    $OUBox.Clear()
    Reset-MSToolkitOUPicker -Server (Get-SelectedServer)
    $TitleBox.Clear()
    $DepartmentBox.Clear()
    $DescriptionBox.Clear()
    $GroupsBox.Clear()
    $AddressBox.Clear()
    $ManagerBox.Clear()

    $ChangePasswordCheckbox.Checked = $true

    $FirstNameBox.Focus()
}


function Show-NewUserConfirmation {
    param(
        [string]$DisplayName,
        [string]$Sam,
        [string]$UPN,
        [string]$MailAddress,
        [string]$OnMicrosoftAliasAddress,
        [string]$OnMicrosoftMailAliasAddress,
        [string]$OUPath,
        [string]$MirrorSummary,
        [string]$AddressSummary,
        [string]$ManagerSummary
    )

    $ConfirmForm = New-Object System.Windows.Forms.Form
    $ConfirmForm.Text = "Confirm Create User"
    $ConfirmForm.Size = New-Object System.Drawing.Size(760,650)
    $ConfirmForm.MinimumSize = New-Object System.Drawing.Size(700,560)
    $ConfirmForm.StartPosition = "CenterScreen"
    $ConfirmForm.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
    $ConfirmForm.Font = New-Object System.Drawing.Font("Segoe UI",9)
    $ConfirmForm.MaximizeBox = $true
    $ConfirmForm.MinimizeBox = $false

    $Header = New-Object System.Windows.Forms.Panel
    $Header.Dock = "Top"
    $Header.Height = 70
    $Header.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $ConfirmForm.Controls.Add($Header)

    $HeaderTitle = New-Object System.Windows.Forms.Label
    $HeaderTitle.Text = "Confirm New Active Directory User"
    $HeaderTitle.AutoSize = $true
    $HeaderTitle.Location = New-Object System.Drawing.Point(18,14)
    $HeaderTitle.ForeColor = [System.Drawing.Color]::White
    $HeaderTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",15)
    $Header.Controls.Add($HeaderTitle)

    $HeaderSub = New-Object System.Windows.Forms.Label
    $HeaderSub.Text = "Review the account, email, location, and cloned settings before creation."
    $HeaderSub.AutoSize = $true
    $HeaderSub.Location = New-Object System.Drawing.Point(20,43)
    $HeaderSub.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
    $Header.Controls.Add($HeaderSub)

    $Details = New-Object System.Windows.Forms.RichTextBox
    $Details.Location = New-Object System.Drawing.Point(18,86)
    $Details.Size = New-Object System.Drawing.Size(708,475)
    $Details.Anchor = "Top,Bottom,Left,Right"
    $Details.ReadOnly = $true
    $Details.BackColor = [System.Drawing.Color]::White
    $Details.BorderStyle = "FixedSingle"
    $Details.Font = New-Object System.Drawing.Font("Segoe UI",9.5)
    $ConfirmForm.Controls.Add($Details)

    function Add-ConfirmationSection {
        param([string]$Title)

        $NormalFont = $Details.Font
        $BoldFont = New-Object System.Drawing.Font(
            $NormalFont.FontFamily,
            $NormalFont.Size,
            [System.Drawing.FontStyle]::Bold
        )

        $Details.SelectionStart = $Details.TextLength
        $Details.SelectionFont = $BoldFont
        $Details.SelectionColor = (Get-MSToolkitThemePalette).Section
        $Details.AppendText("`r`n============================================================`r`n")
        $Details.AppendText("$Title`r`n")
        $Details.AppendText("============================================================`r`n")
        $Details.SelectionFont = $NormalFont
        $Details.SelectionColor = $Details.ForeColor

        $BoldFont.Dispose()
    }

    function Add-ConfirmationField {
        param(
            [string]$Label,
            [string]$Value,
            [System.Drawing.Color]$LabelColor = (Get-MSToolkitThemePalette).Section,
            [System.Drawing.Color]$ValueColor = (Get-MSToolkitThemePalette).Text
        )

        $NormalFont = $Details.Font
        $BoldFont = New-Object System.Drawing.Font(
            $NormalFont.FontFamily,
            $NormalFont.Size,
            [System.Drawing.FontStyle]::Bold
        )

        $Details.SelectionStart = $Details.TextLength
        $Details.SelectionFont = $BoldFont
        $Details.SelectionColor = $LabelColor
        $Details.AppendText("${Label}: ")

        $Details.SelectionFont = $NormalFont
        $Details.SelectionColor = $ValueColor
        $Details.AppendText("$Value`r`n")

        $Details.SelectionFont = $NormalFont
        $Details.SelectionColor = $Details.ForeColor
        $BoldFont.Dispose()
    }

    Add-ConfirmationSection "ACCOUNT"
    Add-ConfirmationField "Display Name" $DisplayName
    Add-ConfirmationField "Username" $Sam
    Add-ConfirmationField "Microsoft 365 UPN / Sign-in" $UPN
    Add-ConfirmationField "Company" $(if ($CompanyName) { $CompanyName } else { "(not set in MSToolkit Settings - left blank)" })
    Add-ConfirmationField `
        "Password Must Be Changed at Next Logon" `
        "Yes" `
        ((Get-MSToolkitThemePalette).Danger) `
        ((Get-MSToolkitThemePalette).Danger)

    Add-ConfirmationSection "EMAIL"
    Add-ConfirmationField `
        "Primary SMTP / Email" `
        $MailAddress `
        ((Get-MSToolkitThemePalette).Success) `
        ((Get-MSToolkitThemePalette).Success)
    Add-ConfirmationField "Secondary SMTP Alias" $UPN
    Add-ConfirmationField "Microsoft 365 Tenant Alias" $(if ($OnMicrosoftAliasAddress) { $OnMicrosoftAliasAddress } else { "(not set in MSToolkit Settings - skipped)" })
    Add-ConfirmationField "Microsoft 365 Routing Alias" $(if ($OnMicrosoftMailAliasAddress) { $OnMicrosoftMailAliasAddress } else { "(not set in MSToolkit Settings - skipped)" })

    Add-ConfirmationSection "ACTIVE DIRECTORY LOCATION"
    Add-ConfirmationField "OU" $OUPath

    Add-ConfirmationSection "CLONED / MIRRORED SETTINGS"
    Add-ConfirmationField "Mirror Direct Groups From" $MirrorSummary
    Add-ConfirmationField "Clone Address From" $AddressSummary
    Add-ConfirmationField "Clone Manager From" $ManagerSummary

    $CreateButton = New-Object System.Windows.Forms.Button
    $CreateButton.Text = "Create User"
    $CreateButton.Size = New-Object System.Drawing.Size(130,36)
    $CreateButton.Location = New-Object System.Drawing.Point(450,575)
    $CreateButton.Anchor = "Bottom,Right"
    $CreateButton.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $CreateButton.ForeColor = [System.Drawing.Color]::White
    $CreateButton.FlatStyle = "Flat"
    $CreateButton.FlatAppearance.BorderSize = 0
    $CreateButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $ConfirmForm.Controls.Add($CreateButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Size = New-Object System.Drawing.Size(130,36)
    $CancelButton.Location = New-Object System.Drawing.Point(596,575)
    $CancelButton.Anchor = "Bottom,Right"
    $CancelButton.BackColor = [System.Drawing.Color]::White
    $CancelButton.FlatStyle = "Flat"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $ConfirmForm.Controls.Add($CancelButton)

    $ConfirmForm.AcceptButton = $CreateButton
    $ConfirmForm.CancelButton = $CancelButton

    Apply-MSToolkitSharedTheme -Root $ConfirmForm
    return $ConfirmForm.ShowDialog()
}

function New-MSToolkitADUser {
    $UserCreated = $false
    $CreatedSam = $null
    $CurrentStep = 'Validating input'

    try {
        $Server = Get-SelectedServer

        $FirstName = $FirstNameBox.Text.Trim()
        $LastName = $LastNameBox.Text.Trim()
        $ChildOU = $OUBox.Text.Trim()
        $Title = $TitleBox.Text.Trim()
        $Department = $DepartmentBox.Text.Trim()
        $Description = $DescriptionBox.Text.Trim()
        $MirrorUserSam = $GroupsBox.Text.Trim()
        $AddressUserSam = $AddressBox.Text.Trim()
        $ManagerUserSam = $ManagerBox.Text.Trim()
        $Password = $PasswordBox.Text

        # Refuse rather than create a half-formed account.
        $MissingNaming = New-Object System.Collections.Generic.List[string]
        if (-not $UPNDomain) { $MissingNaming.Add("UPN domain") }
        if (-not $MailDomain) { $MissingNaming.Add("Primary SMTP domain") }
        if ($MissingNaming.Count -gt 0) {
            Show-Error ("No account was created. These MSToolkit Settings are required and are blank: " + ($MissingNaming.ToArray() -join ", ") + ".`r`n`r`nEnter them in MSToolkit Settings under New account naming, save, then reopen Create New AD User.")
            return
        }

        if (-not $FirstName) {
            Show-Error "First name is required."
            return
        }

        if (-not $LastName) {
            Show-Error "Last name is required."
            return
        }

        $FirstNamePart = ConvertTo-EmailSafeName -Value $FirstName
        $LastNamePart = ConvertTo-EmailSafeName -Value $LastName

        if (-not $FirstNamePart) {
            Show-Error "First name must contain at least one letter or number for the email address."
            return
        }

        if (-not $LastNamePart) {
            Show-Error "Last name must contain at least one letter or number for the email address."
            return
        }

        $Sam = Get-GeneratedUsername -FirstName $FirstName -LastName $LastName
        $SamBox.Text = $Sam

        if ($Sam.Length -gt 20) {
            Show-Error "Generated username $Sam is longer than the 20-character sAMAccountName limit. Shorten the first name before creating the account."
            return
        }

        if (-not $Password) {
            Show-Error "Temporary password is required."
            return
        }

        $DisplayName = "$FirstName $LastName"
        $UPN = "$Sam@$UPNDomain"
        $MailAddress = "$FirstNamePart.$LastNamePart@$MailDomain"
        $PrimarySmtpAddress = $MailAddress
        $AliasSmtpAddress = $UPN
        # The onmicrosoft aliases are optional: blank domains in MSToolkit Settings skip them.
        $OnMicrosoftAliasAddress = if ($OnMicrosoftAliasDomain) { "$Sam@$OnMicrosoftAliasDomain" } else { "" }
        $OnMicrosoftMailAliasAddress = if ($OnMicrosoftMailAliasDomain) { "$Sam@$OnMicrosoftMailAliasDomain" } else { "" }
        $SecondaryAliases = @(@($UPN, $OnMicrosoftAliasAddress, $OnMicrosoftMailAliasAddress) | Where-Object { $_ })
        $ProxyAddresses = @("SMTP:$MailAddress") + @($SecondaryAliases | ForEach-Object { "smtp:$_" })
        $OUPath = Get-OUPath -ChildOU $ChildOU -SelectedOU $script:OUPickerCurrentDN -Server $Server

        # Never create an account directly in the domain root.
        if ([string]::IsNullOrWhiteSpace($OUPath) -or ($OUPath.Trim() -ieq $DomainDN)) {
            Show-Error ("No OU was chosen, so the account would be created in the domain root ($DomainDN). No account was created.`r`n`r`n" +
                "Pick an OU in the OU picker or type a child OU name. To have one chosen by default, set the Users OU in MSToolkit Settings (Organizational units).")
            return
        }
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

        $MirrorUser = $null
        $MirrorGroups = @()

        if ($MirrorUserSam) {
            try {
                $MirrorUser = Get-ADUser `
                    -Identity $MirrorUserSam `
                    -Server $Server `
                    -Properties MemberOf `
                    -ErrorAction Stop

                if ($MirrorUser.MemberOf) {
                    $ResolvedMirrorGroups = New-Object System.Collections.Generic.List[object]

                    foreach ($GroupDN in $MirrorUser.MemberOf) {
                        try {
                            $ResolvedGroup = Get-ADGroup `
                                -Identity $GroupDN `
                                -Server $Server `
                                -Properties Name,SamAccountName,GroupCategory,GroupScope `
                                -ErrorAction Stop

                            $ResolvedMirrorGroups.Add($ResolvedGroup)
                        }
                        catch {
                            Write-OutputBox "WARNING: Could not read mirrored group $GroupDN. $($_.Exception.Message)" ([System.Drawing.Color]::DarkOrange)
                        }
                    }

                    $MirrorGroups = @($ResolvedMirrorGroups.ToArray() | Sort-Object Name)
                }

                Write-OutputBox "Mirror user resolved: $($MirrorUser.SamAccountName) - $($MirrorGroups.Count) direct group(s) found."
            }
            catch {
                Show-Error "Could not find mirror user '$MirrorUserSam' using $Server.`r`n`r`n$($_.Exception.Message)"
                return
            }
        }


        $AddressUser = $null
        $AddressAttributes = @{}

        if ($AddressUserSam) {
            try {
                # Read the exact LDAP attributes used by the ADUC Address tab.
                $AddressUser = Get-ADUser `
                    -Identity $AddressUserSam `
                    -Server $Server `
                    -Properties streetAddress,l,st,postalCode,c,co,countryCode `
                    -ErrorAction Stop

                if ($AddressUser.streetAddress) {
                    $AddressAttributes['streetAddress'] = [string]$AddressUser.streetAddress
                }

                if ($AddressUser.l) {
                    $AddressAttributes['l'] = [string]$AddressUser.l
                }

                if ($AddressUser.st) {
                    $AddressAttributes['st'] = [string]$AddressUser.st
                }

                if ($AddressUser.postalCode) {
                    $AddressAttributes['postalCode'] = [string]$AddressUser.postalCode
                }

                if ($AddressUser.c) {
                    $AddressAttributes['c'] = [string]$AddressUser.c
                }

                if ($AddressUser.co) {
                    $AddressAttributes['co'] = [string]$AddressUser.co
                }

                if ($null -ne $AddressUser.countryCode -and [int]$AddressUser.countryCode -ne 0) {
                    $AddressAttributes['countryCode'] = [int]$AddressUser.countryCode
                }

                Write-OutputBox "Address clone user resolved: $($AddressUser.SamAccountName)"
                Write-OutputBox "Street: $($AddressUser.streetAddress)"
                Write-OutputBox "City: $($AddressUser.l)"
                Write-OutputBox "State: $($AddressUser.st)"
                Write-OutputBox "ZIP: $($AddressUser.postalCode)"
                Write-OutputBox "Country: $($AddressUser.co) ($($AddressUser.c))"
            }
            catch {
                Show-Error "Could not find address clone user '$AddressUserSam' using $Server.`r`n`r`n$($_.Exception.Message)"
                return
            }
        }


        $ManagerSourceUser = $null
        $ManagerUser = $null

        if ($ManagerUserSam) {
            try {
                $ManagerSourceUser = Get-ADUser `
                    -Identity $ManagerUserSam `
                    -Server $Server `
                    -Properties Manager `
                    -ErrorAction Stop

                if (-not $ManagerSourceUser.Manager) {
                    Show-Error "The selected manager clone user '$ManagerUserSam' does not have a Manager set in Active Directory."
                    return
                }

                $ManagerUser = Get-ADUser `
                    -Identity $ManagerSourceUser.Manager `
                    -Server $Server `
                    -Properties DisplayName,SamAccountName,DistinguishedName `
                    -ErrorAction Stop

                Write-OutputBox "Manager clone user resolved: $($ManagerSourceUser.SamAccountName)"
                Write-OutputBox "Manager to clone: $($ManagerUser.DisplayName) ($($ManagerUser.SamAccountName))"
            }
            catch {
                Show-Error "Could not resolve the manager from '$ManagerUserSam' using $Server.`r`n`r`n$($_.Exception.Message)"
                return
            }
        }

        $ExistingUser = Get-ADUser `
            -Filter "SamAccountName -eq '$Sam'" `
            -Server $Server `
            -ErrorAction SilentlyContinue

        if ($ExistingUser) {
            Show-Error "A user with username $Sam already exists."
            return
        }

        $ExistingUPN = Get-ADUser `
            -Filter "UserPrincipalName -eq '$UPN'" `
            -Server $Server `
            -ErrorAction SilentlyContinue

        if ($ExistingUPN) {
            Show-Error "A user with UPN $UPN already exists."
            return
        }

        $EmailLookupClauses = @(@($MailAddress) + $SecondaryAliases | ForEach-Object { "(proxyAddresses=SMTP:$_)(proxyAddresses=smtp:$_)" }) -join ""
        $EmailLookupFilter = "(|(mail=$MailAddress)$EmailLookupClauses)"
        $ExistingAddressUser = Get-ADUser `
            -LDAPFilter $EmailLookupFilter `
            -Server $Server `
            -Properties mail,proxyAddresses `
            -ErrorAction SilentlyContinue | Select-Object -First 1

        if ($ExistingAddressUser) {
            Show-Error "The UPN, mail address, or proxy alias is already assigned to $($ExistingAddressUser.SamAccountName)."
            return
        }

        $MirrorSummary = if ($MirrorUser) {
            "$($MirrorUser.SamAccountName) ($($MirrorGroups.Count) groups)"
        }
        else {
            "None"
        }

        $AddressSummary = if ($AddressUser) {
            $AddressUser.SamAccountName
        }
        else {
            "None"
        }

        $ManagerSummary = if ($ManagerSourceUser) {
            "$($ManagerSourceUser.SamAccountName) -> $($ManagerUser.DisplayName)"
        }
        else {
            "None"
        }

        $Confirm = Show-NewUserConfirmation `
            -DisplayName $DisplayName `
            -Sam $Sam `
            -UPN $UPN `
            -MailAddress $MailAddress `
            -OnMicrosoftAliasAddress $OnMicrosoftAliasAddress `
            -OnMicrosoftMailAliasAddress $OnMicrosoftMailAliasAddress `
            -OUPath $OUPath `
            -MirrorSummary $MirrorSummary `
            -AddressSummary $AddressSummary `
            -ManagerSummary $ManagerSummary

        if ($Confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-OutputBox "User creation cancelled."
            return
        }

        Write-ResultSeparator
        Write-OutputBox "Beginning user creation for $DisplayName ($Sam)..."
        Write-OutputBox "Using AD server: $Server"
        Write-OutputBox "Using OU path: $OUPath"

        $Params = @{
            GivenName             = $FirstName
            Surname               = $LastName
            Name                  = $DisplayName
            DisplayName           = $DisplayName
            SamAccountName        = $Sam
            UserPrincipalName     = $UPN
            EmailAddress          = $MailAddress
            AccountPassword       = $SecurePassword
            Enabled               = $true
            ChangePasswordAtLogon = $true
            Path                  = $OUPath
            Server                = $Server
        }

        if ($CompanyName) {
            $Params.Company = $CompanyName
        }

        if ($Title) {
            $Params.Title = $Title
        }

        if ($Department) {
            $Params.Department = $Department
        }

        if ($Description) {
            $Params.Description = $Description
        }

        $CurrentStep = 'Creating AD user'
        New-ADUser @Params
        $UserCreated = $true
        $CreatedSam = $Sam
        Write-OutputBox "AD user object created successfully: $DisplayName / $Sam" ([System.Drawing.Color]::Green)

        if ($AddressUser -and $AddressAttributes.Count -gt 0) {
            $CurrentStep = 'Applying cloned address attributes' 
            Set-ADUser `
                -Identity $Sam `
                -Server $Server `
                -Replace $AddressAttributes `
                -ErrorAction Stop

            Write-OutputBox "Applied $($AddressAttributes.Count) cloned address attribute(s) from $($AddressUser.SamAccountName)." ([System.Drawing.Color]::Green)
        }

        if ($ManagerUser) {
            $CurrentStep = 'Setting manager'
            Set-ADUser `
                -Identity $Sam `
                -Server $Server `
                -Manager $ManagerUser.DistinguishedName `
                -ErrorAction Stop

            Write-OutputBox "Set Manager: $($ManagerUser.DisplayName) ($($ManagerUser.SamAccountName))" ([System.Drawing.Color]::Green)
        }

        $CurrentStep = 'Setting primary mail attribute'
        Set-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Replace @{ mail = $MailAddress } `
            -ErrorAction Stop

        $CurrentStep = 'Setting proxyAddresses'
        Set-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Add @{ proxyAddresses = $ProxyAddresses } `
            -ErrorAction Stop

        Write-OutputBox "Created user: $DisplayName / $Sam using $Server" ([System.Drawing.Color]::Green)
        Write-OutputBox "Set Microsoft 365 UPN / sign-in: $UPN" ([System.Drawing.Color]::Green)
        Write-OutputBox "Set Primary SMTP / Email address: $MailAddress" ([System.Drawing.Color]::Green)
        if ($CompanyName) {
            Write-OutputBox "Set Company: $CompanyName" ([System.Drawing.Color]::Green)
        }
        else {
            Write-OutputBox "Company not set - no Company in MSToolkit Settings." ([System.Drawing.Color]::DimGray)
        }
        Write-OutputBox "Added secondary SMTP alias: $UPN" ([System.Drawing.Color]::Green)
        if ($OnMicrosoftAliasAddress) { Write-OutputBox "Added Microsoft 365 tenant alias: $OnMicrosoftAliasAddress" ([System.Drawing.Color]::Green) }
        if ($OnMicrosoftMailAliasAddress) { Write-OutputBox "Added Microsoft 365 routing alias: $OnMicrosoftMailAliasAddress" ([System.Drawing.Color]::Green) }
        Write-OutputBox "Set user must change password at next logon: Yes" ([System.Drawing.Color]::Green)

        if ($AddressUser) {
            Write-OutputBox "Cloned address from $($AddressUser.SamAccountName):" ([System.Drawing.Color]::Green)
            Write-OutputBox "Street: $($AddressUser.streetAddress)" ([System.Drawing.Color]::Green)
            Write-OutputBox "City: $($AddressUser.l)" ([System.Drawing.Color]::Green)
            Write-OutputBox "State: $($AddressUser.st)" ([System.Drawing.Color]::Green)
            Write-OutputBox "ZIP: $($AddressUser.postalCode)" ([System.Drawing.Color]::Green)
            Write-OutputBox "Country: $($AddressUser.co) ($($AddressUser.c))" ([System.Drawing.Color]::Green)
        }

        if ($MirrorUser) {
            $CurrentStep = 'Mirroring group memberships'
            Write-OutputBox "Mirroring $($MirrorGroups.Count) direct group membership(s) from $($MirrorUser.SamAccountName)..."

            $GroupAddSuccess = 0
            $GroupAddFailed = 0

            foreach ($Group in $MirrorGroups) {
                try {
                    Add-ADGroupMember `
                        -Identity $Group.DistinguishedName `
                        -Members $Sam `
                        -Server $Server `
                        -Confirm:$false `
                        -ErrorAction Stop

                    Write-OutputBox "Added $Sam to mirrored group: $($Group.Name)" ([System.Drawing.Color]::Green)
                    $GroupAddSuccess++
                }
                catch {
                    Write-OutputBox "WARNING: Could not add $Sam to mirrored group $($Group.Name). $($_.Exception.Message)" ([System.Drawing.Color]::DarkOrange)
                    $GroupAddFailed++
                }
            }

            Write-OutputBox "Group mirror complete. Successful: $GroupAddSuccess | Failed: $GroupAddFailed"
        }

        $CurrentStep = 'Final verification'
        $CreatedUser = Get-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Properties DisplayName,Enabled,Title,Department,Company,Manager,UserPrincipalName,EmailAddress,proxyAddresses,streetAddress,l,st,postalCode,c,co,countryCode,WhenCreated,DistinguishedName,MemberOf `
            -ErrorAction Stop

        Write-ResultSeparator
        Write-OutputBox "FINAL VERIFICATION" ([System.Drawing.Color]::FromArgb(31,58,93))

        $VerificationFailed = $false

        $UpnOk = ($CreatedUser.UserPrincipalName -eq $UPN)
        Write-VerificationResult -Item 'Microsoft 365 UPN / sign-in' -Passed $UpnOk -Detail $CreatedUser.UserPrincipalName
        if (-not $UpnOk) { $VerificationFailed = $true }

        $MailOk = ($CreatedUser.EmailAddress -eq $MailAddress)
        Write-VerificationResult -Item 'Primary SMTP / mail attribute' -Passed $MailOk -Detail $CreatedUser.EmailAddress
        if (-not $MailOk) { $VerificationFailed = $true }

        $ActualProxies = @($CreatedUser.proxyAddresses)
        $PrimaryProxyOk = ($ActualProxies -ccontains "SMTP:$MailAddress")
        Write-VerificationResult -Item 'Primary proxyAddress' -Passed $PrimaryProxyOk -Detail "SMTP:$MailAddress"
        if (-not $PrimaryProxyOk) { $VerificationFailed = $true }

        foreach ($ExpectedAlias in @($SecondaryAliases | ForEach-Object { "smtp:$_" })) {
            $AliasOk = ($ActualProxies -contains $ExpectedAlias)
            Write-VerificationResult -Item 'Proxy alias' -Passed $AliasOk -Detail $ExpectedAlias
            if (-not $AliasOk) { $VerificationFailed = $true }
        }

        if ($ManagerUser) {
            $ManagerOk = ($CreatedUser.Manager -eq $ManagerUser.DistinguishedName)
            Write-VerificationResult -Item 'Manager' -Passed $ManagerOk -Detail "$($ManagerUser.DisplayName) ($($ManagerUser.SamAccountName))"
            if (-not $ManagerOk) { $VerificationFailed = $true }
        }
        else {
            Write-OutputBox 'Manager - Not requested.' ([System.Drawing.Color]::DimGray)
        }

        if ($AddressUser -and $AddressAttributes.Count -gt 0) {
            $AddressMismatches = New-Object System.Collections.Generic.List[string]
            foreach ($Key in $AddressAttributes.Keys) {
                $ActualValue = [string]$CreatedUser.$Key
                $ExpectedValue = [string]$AddressAttributes[$Key]
                if ($ActualValue -ne $ExpectedValue) {
                    $AddressMismatches.Add("$Key expected '$ExpectedValue' but found '$ActualValue'")
                }
            }

            $AddressOk = ($AddressMismatches.Count -eq 0)
            $AddressDetail = if ($AddressOk) { "Matches $($AddressUser.SamAccountName)" } else { $AddressMismatches.ToArray() -join '; ' }
            Write-VerificationResult -Item 'Cloned address' -Passed $AddressOk -Detail $AddressDetail
            if (-not $AddressOk) { $VerificationFailed = $true }
        }
        else {
            Write-OutputBox 'Cloned address - Not requested.' ([System.Drawing.Color]::DimGray)
        }

        if ($MirrorUser) {
            $ActualGroupDns = @($CreatedUser.MemberOf)
            $MissingMirroredGroups = New-Object System.Collections.Generic.List[string]
            foreach ($Group in $MirrorGroups) {
                if ($ActualGroupDns -notcontains $Group.DistinguishedName) {
                    $MissingMirroredGroups.Add($Group.Name)
                }
            }

            $GroupsOk = ($MissingMirroredGroups.Count -eq 0)
            $GroupsDetail = if ($GroupsOk) { "$($MirrorGroups.Count) of $($MirrorGroups.Count) direct groups present" } else { "Missing: $($MissingMirroredGroups.ToArray() -join ', ')" }
            Write-VerificationResult -Item 'Mirrored direct group memberships' -Passed $GroupsOk -Detail $GroupsDetail
            if (-not $GroupsOk) { $VerificationFailed = $true }
        }
        else {
            Write-OutputBox 'Mirrored direct group memberships - Not requested.' ([System.Drawing.Color]::DimGray)
        }

        Write-ResultSeparator
        Write-OutputBox "Created User Summary:"
        Write-OutputBox "Name: $($CreatedUser.DisplayName)"
        Write-OutputBox "Username: $($CreatedUser.SamAccountName)"
        Write-OutputBox "UPN: $($CreatedUser.UserPrincipalName)"
        Write-OutputBox "Primary SMTP / Email address: $($CreatedUser.EmailAddress)"
        Write-OutputBox "ProxyAddresses: $($CreatedUser.proxyAddresses -join ', ')"
        Write-OutputBox "Enabled: $($CreatedUser.Enabled)"
        Write-OutputBox "Title: $($CreatedUser.Title)"
        Write-OutputBox "Department: $($CreatedUser.Department)"
        Write-OutputBox "Company: $($CreatedUser.Company)"
        if ($CreatedUser.Manager) {
            try {
                $CreatedManager = Get-ADUser -Identity $CreatedUser.Manager -Server $Server -Properties DisplayName,SamAccountName -ErrorAction Stop
                Write-OutputBox "Manager: $($CreatedManager.DisplayName) ($($CreatedManager.SamAccountName))"
            }
            catch {
                Write-OutputBox "Manager DN: $($CreatedUser.Manager)"
            }
        }
        else {
            Write-OutputBox "Manager: "
        }
        Write-OutputBox "Street: $($CreatedUser.streetAddress)"
        Write-OutputBox "City: $($CreatedUser.l)"
        Write-OutputBox "State: $($CreatedUser.st)"
        Write-OutputBox "ZIP: $($CreatedUser.postalCode)"
        Write-OutputBox "Country: $($CreatedUser.co) ($($CreatedUser.c))"
        Write-OutputBox "Created: $($CreatedUser.WhenCreated)"
        Write-OutputBox "DistinguishedName: $($CreatedUser.DistinguishedName)"

        if ($VerificationFailed) {
            Write-OutputBox "USER WAS CREATED, BUT FINAL VERIFICATION FAILED FOR ONE OR MORE ITEMS." ([System.Drawing.Color]::Red)
            Show-Error "User $DisplayName was created, but one or more final verification checks failed. Review the Activity Log before making additional changes."
            return
        }

        Write-OutputBox "FINAL VERIFICATION COMPLETE: All requested settings verified." ([System.Drawing.Color]::Green)
        Show-Message "User $DisplayName was created successfully and all requested settings were verified."

        Clear-FormFields
    }
    catch {
        Write-ResultSeparator

        if ($UserCreated) {
            Write-OutputBox "USER WAS CREATED, BUT POST-CREATION CONFIGURATION FAILED." ([System.Drawing.Color]::Red)
            Write-OutputBox "Username: $CreatedSam" ([System.Drawing.Color]::Red)
            Write-OutputBox "Failed step: $CurrentStep" ([System.Drawing.Color]::Red)
            Write-OutputBox "Error: $($_.Exception.Message)" ([System.Drawing.Color]::Red)

            Show-Error "The AD user '$CreatedSam' WAS created, but post-creation configuration failed.`r`n`r`nFailed step: $CurrentStep`r`n`r`n$($_.Exception.Message)`r`n`r`nDo not create the user again. Review the existing account and the Activity Log."
        }
        else {
            Write-OutputBox "USER CREATION FAILED BEFORE THE AD ACCOUNT WAS CREATED." ([System.Drawing.Color]::Red)
            Write-OutputBox "Failed step: $CurrentStep" ([System.Drawing.Color]::Red)
            Write-OutputBox "Error: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
            Show-Error $_.Exception.Message
        }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Create New AD User - $Domain"
# Open large by default, but never larger than the screen will allow.
$PreferredWidth = 1500
$PreferredHeight = 980
$WorkingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$FormWidth = [Math]::Min($PreferredWidth, [int]($WorkingArea.Width * 0.95))
$FormHeight = [Math]::Min($PreferredHeight, [int]($WorkingArea.Height * 0.95))

$form.Size = New-Object System.Drawing.Size($FormWidth,$FormHeight)
$form.MinimumSize = New-Object System.Drawing.Size(800,650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.AutoScroll = $true
# Set after the controls exist, from the real content height (see below).
$form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)

# Compare-tool style header
$HeaderPanel = New-Object System.Windows.Forms.Panel
# Docked rather than anchored so it always spans the window, whatever its width.
$HeaderPanel.Dock = "Top"
$HeaderPanel.Height = 78
$HeaderPanel.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$form.Controls.Add($HeaderPanel)

$HeaderTitle = New-Object System.Windows.Forms.Label
$HeaderTitle.Text = "Create New AD User"
$HeaderTitle.AutoSize = $true
$HeaderTitle.ForeColor = [System.Drawing.Color]::White
$HeaderTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",20)
$HeaderTitle.Location = New-Object System.Drawing.Point(20,12)
$HeaderPanel.Controls.Add($HeaderTitle)

$HeaderSubtitle = New-Object System.Windows.Forms.Label
$HeaderSubtitle.Text = "Create a new Active Directory account and optionally mirror groups, address, and manager"
$HeaderSubtitle.AutoSize = $true
$HeaderSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
$HeaderSubtitle.Font = New-Object System.Drawing.Font("Segoe UI",9.5)
$HeaderSubtitle.Location = New-Object System.Drawing.Point(23,49)
$HeaderPanel.Controls.Add($HeaderSubtitle)
$script:MSToolkitThemeToggleButton = New-MSToolkitThemeToggleButton -HeaderPanel $HeaderPanel -Form $form -ToolKey "NewADUser"

$LabelWidth = 150
$InputWidth = 520
$Left = 20
$Top = 98
$RowHeight = 34

function Add-Label {
    param(
        [string]$Text,
        [int]$Y
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($Left,$Y)
    $label.Size = New-Object System.Drawing.Size($LabelWidth,22)
    $label.Font = New-Object System.Drawing.Font("Segoe UI Semibold",9)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(45,55,65)
    $form.Controls.Add($label)
}

function Add-TextBox {
    param(
        [int]$Y,
        [bool]$Password = $false
    )

    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),$Y)
    $box.Size = New-Object System.Drawing.Size($InputWidth,24)
    $box.Anchor = "Top,Left,Right"
    $box.Font = New-Object System.Drawing.Font("Segoe UI",9.5)
    $box.BackColor = [System.Drawing.Color]::White
    $box.BorderStyle = "FixedSingle"

    if ($Password) {
        $box.UseSystemPasswordChar = $true
    }

    $form.Controls.Add($box)
    return $box
}

Add-Label "AD Server:" $Top

$DCDropdown = New-Object System.Windows.Forms.ComboBox
$DCDropdown.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),($Top - 2))
$DCDropdown.Size = New-Object System.Drawing.Size(180,24)
$DCDropdown.DropDownStyle = "DropDownList"
$DCDropdown.Font = New-Object System.Drawing.Font("Segoe UI",9.5)
$DCDropdown.FlatStyle = "Flat"

foreach ($DC in $DCs) {
    [void]$DCDropdown.Items.Add($DC)
}

# When launched from MSToolkit, prefer the DC selected in the main window.
# Accept either a short DC name (DC01) or FQDN (DC01.contoso.local).
$DefaultServerIndex = -1
if (-not [string]::IsNullOrWhiteSpace($DefaultServer)) {
    $DefaultServerShortName = ($DefaultServer -split '\.')[0]

    for ($Index = 0; $Index -lt $DCDropdown.Items.Count; $Index++) {
        if ([string]$DCDropdown.Items[$Index] -ieq $DefaultServerShortName) {
            $DefaultServerIndex = $Index
            break
        }
    }
}

if ($DefaultServerIndex -ge 0) {
    $DCDropdown.SelectedIndex = $DefaultServerIndex
}
elseif ($DCDropdown.Items.Count -gt 0) {
    $DCDropdown.SelectedIndex = 0
}
$form.Controls.Add($DCDropdown)

$Top += $RowHeight

Add-Label "First Name:" $Top
$FirstNameBox = Add-TextBox $Top

$Top += $RowHeight

Add-Label "Last Name:" $Top
$LastNameBox = Add-TextBox $Top

$Top += $RowHeight

Add-Label "Username:" $Top
$SamBox = Add-TextBox $Top
$SamBox.ReadOnly = $true
$SamBox.BackColor = [System.Drawing.Color]::FromArgb(238,241,245)
$SamBox.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)

$FirstNameBox.Add_TextChanged({ Update-GeneratedUsername })
$LastNameBox.Add_TextChanged({ Update-GeneratedUsername })

$Top += $RowHeight

Add-Label "Temp Password:" $Top
$PasswordBox = Add-TextBox $Top $true

$Top += $RowHeight

Add-Label "Department OU:" $Top

$OUPickerFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$OUPickerFlow.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),$Top)
$OUPickerFlow.Size = New-Object System.Drawing.Size(650,62)
$OUPickerFlow.AutoScroll = $true
$OUPickerFlow.FlowDirection = "LeftToRight"
$OUPickerFlow.WrapContents = $true
$OUPickerFlow.Anchor = "Top,Left,Right"
$form.Controls.Add($OUPickerFlow)

$Top += 66

$OUPathLabel = New-Object System.Windows.Forms.Label
$OUPathLabel.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),$Top)
$OUPathLabel.Size = New-Object System.Drawing.Size(650,20)
$OUPathLabel.Anchor = "Top,Left,Right"
$OUPathLabel.ForeColor = [System.Drawing.Color]::DimGray
$OUPathLabel.Font = New-Object System.Drawing.Font("Segoe UI",8.5)
$form.Controls.Add($OUPathLabel)

$Top += 26

Add-Label "Or type OU:" $Top
$OUBox = Add-TextBox $Top

$OUHelp = New-Object System.Windows.Forms.Label
$OUHelp.Text = "Optional. Type a child OU name or a path such as IT\Test under the selection above, or paste a full DN to override it."
$OUHelp.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),($Top + 24))
$OUHelp.Size = New-Object System.Drawing.Size(650,35)
$OUHelp.Anchor = "Top,Left,Right"
$OUHelp.ForeColor = [System.Drawing.Color]::DimGray
$OUHelp.Font = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Italic)
$form.Controls.Add($OUHelp)

$Top += 60

Add-Label "Title:" $Top
$TitleBox = Add-TextBox $Top

$Top += $RowHeight

Add-Label "Department:" $Top
$DepartmentBox = Add-TextBox $Top

$Top += $RowHeight

Add-Label "Description:" $Top
$DescriptionBox = Add-TextBox $Top

$Top += $RowHeight

Add-Label "Mirror Groups From:" $Top
$GroupsBox = Add-TextBox $Top

$GroupsHelp = New-Object System.Windows.Forms.Label
$GroupsHelp.Text = "Optional. Enter an existing user's sAMAccountName to copy their direct AD group memberships."
$GroupsHelp.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),($Top + 24))
$GroupsHelp.Size = New-Object System.Drawing.Size(650,22)
$GroupsHelp.ForeColor = [System.Drawing.Color]::DimGray
$GroupsHelp.Font = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Italic)
$form.Controls.Add($GroupsHelp)

$Top += 52

Add-Label "Clone Address From:" $Top
$AddressBox = Add-TextBox $Top

$AddressHelp = New-Object System.Windows.Forms.Label
$AddressHelp.Text = "Optional. Enter an existing user's sAMAccountName to copy street, city, state, ZIP, and country."
$AddressHelp.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),($Top + 24))
$AddressHelp.Size = New-Object System.Drawing.Size(650,22)
$AddressHelp.ForeColor = [System.Drawing.Color]::DimGray
$AddressHelp.Font = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Italic)
$form.Controls.Add($AddressHelp)

$Top += 52

Add-Label "Clone Manager From:" $Top
$ManagerBox = Add-TextBox $Top

$ManagerHelp = New-Object System.Windows.Forms.Label
$ManagerHelp.Text = "Optional. Enter an existing user's sAMAccountName to copy that user's Manager."
$ManagerHelp.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),($Top + 24))
$ManagerHelp.Size = New-Object System.Drawing.Size(650,22)
$ManagerHelp.ForeColor = [System.Drawing.Color]::DimGray
$ManagerHelp.Font = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Italic)
$form.Controls.Add($ManagerHelp)

$Top += 52

$ChangePasswordCheckbox = New-Object System.Windows.Forms.CheckBox
$ChangePasswordCheckbox.Text = "User must change password at next logon"
$ChangePasswordCheckbox.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),$Top)
$ChangePasswordCheckbox.Size = New-Object System.Drawing.Size(350,24)
$ChangePasswordCheckbox.Checked = $true
$ChangePasswordCheckbox.Enabled = $false
$ChangePasswordCheckbox.ForeColor = [System.Drawing.Color]::FromArgb(65,65,65)
$ChangePasswordCheckbox.Font = New-Object System.Drawing.Font("Segoe UI",9)
$form.Controls.Add($ChangePasswordCheckbox)

$Top += 42

$CreateButton = New-Object System.Windows.Forms.Button
$CreateButton.Text = "Create User"
$CreateButton.Location = New-Object System.Drawing.Point(($Left + $LabelWidth),$Top)
$CreateButton.Size = New-Object System.Drawing.Size(130,34)
$CreateButton.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$CreateButton.ForeColor = [System.Drawing.Color]::White
$CreateButton.FlatStyle = "Flat"
$CreateButton.FlatAppearance.BorderSize = 0
$CreateButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold",9.5)
$CreateButton.Add_Click({ New-MSToolkitADUser })
$form.Controls.Add($CreateButton)

$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.Text = "Close"
$CloseButton.Location = New-Object System.Drawing.Point(($Left + $LabelWidth + 145),$Top)
$CloseButton.Size = New-Object System.Drawing.Size(100,34)
$CloseButton.BackColor = [System.Drawing.Color]::White
$CloseButton.ForeColor = [System.Drawing.Color]::FromArgb(45,55,65)
$CloseButton.FlatStyle = "Flat"
$CloseButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(190,198,208)
$CloseButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold",9)
$CloseButton.Add_Click({ $form.Close() })
$form.Controls.Add($CloseButton)

$Top += 50

$OutputBox = New-Object System.Windows.Forms.RichTextBox
$OutputBox.Location = New-Object System.Drawing.Point(20,$Top)

# Size to the current window rather than a fixed width, because anchoring only
# takes effect on a later resize and the form is created at its final size.
# No Bottom anchor: on an AutoScroll form that fights the scrollable region and
# leaves the last lines unreachable.
$OutputWidth = [Math]::Max(840, ($form.ClientSize.Width - 40))
$OutputHeight = [Math]::Max(180, ($form.ClientSize.Height - $Top - 24))
$OutputBox.Size = New-Object System.Drawing.Size($OutputWidth,$OutputHeight)
$OutputBox.Anchor = "Top,Left,Right"
$OutputBox.ScrollBars = "Vertical"
$OutputBox.ReadOnly = $true
$OutputBox.BackColor = [System.Drawing.Color]::White
$OutputBox.ForeColor = [System.Drawing.Color]::FromArgb(35,40,45)
$OutputBox.Font = New-Object System.Drawing.Font("Consolas",9)
$OutputBox.BorderStyle = "FixedSingle"
$form.Controls.Add($OutputBox)

# Now that every control exists, tell the form how tall its content really is so
# a scrollbar appears when the window is smaller than the layout.
$form.AutoScrollMinSize = New-Object System.Drawing.Size(0,($OutputBox.Bottom + 20))

$form.Add_Resize({
    if ($OutputBox -and $form.ClientSize.Height -gt ($OutputBox.Top + 200)) {
        $OutputBox.Height = $form.ClientSize.Height - $OutputBox.Top - 24
        $form.AutoScrollMinSize = New-Object System.Drawing.Size(0,($OutputBox.Bottom + 20))
    }
})

$form.AcceptButton = $CreateButton

$DCDropdown.Add_SelectedIndexChanged({
    Reset-MSToolkitOUPicker -Server (Get-SelectedServer)
    Apply-MSToolkitSharedTheme -Root $OUPickerFlow
})

Initialize-MSToolkitOUPicker -Flow $OUPickerFlow -PathLabel $OUPathLabel -Server (Get-SelectedServer) -RootDN $BaseOU

Apply-MSToolkitSharedTheme -Root $form

# The shared theme pass resets the picker's path line to the muted colour; put the
# domain-root warning colour back, now and after every theme toggle.
Update-MSToolkitOUPathLabel
$script:MSToolkitThemeRefreshHook = { Update-MSToolkitOUPathLabel }

$StartupPalette = Get-MSToolkitThemePalette

Write-StartupSection "DIRECTORY CONNECTION"
Write-StartupField `
    -Label "Domain" `
    -Value $Domain `
    -ValueColor $StartupPalette.Success
Write-StartupField `
    -Label "Selected AD Server" `
    -Value (Get-SelectedServer) `
    -ValueColor $StartupPalette.Success
Write-StartupField `
    -Label "Default OU" `
    -Value $(if ($UsersOUConfigured) { $BaseOU } else { "Domain root ($BaseOU)" }) `
    -ValueColor $StartupPalette.Info

$NotSetText = "(not set in MSToolkit Settings)"

Write-StartupSection "ACCOUNT & EMAIL DEFAULTS"
Write-StartupField `
    -Label "Username / Microsoft 365 UPN" `
    -Value $(if ($UPNDomain) { "@$UPNDomain" } else { "$NotSetText - required" }) `
    -ValueColor $(if ($UPNDomain) { $StartupPalette.Info } else { $StartupPalette.Warning })
Write-StartupField `
    -Label "Primary SMTP / Email" `
    -Value $(if ($MailDomain) { "@$MailDomain" } else { "$NotSetText - required" }) `
    -ValueColor $(if ($MailDomain) { $StartupPalette.Info } else { $StartupPalette.Warning })
$StartupAliasDomains = @(@($UPNDomain, $OnMicrosoftAliasDomain, $OnMicrosoftMailAliasDomain) | Where-Object { $_ } | ForEach-Object { "@$_" })
Write-StartupField `
    -Label "Secondary Aliases" `
    -Value $(if ($StartupAliasDomains.Count -gt 0) { $StartupAliasDomains -join ", " } else { $NotSetText }) `
    -ValueColor $StartupPalette.Info
Write-StartupField `
    -Label "Company" `
    -Value $(if ($CompanyName) { $CompanyName } else { "$NotSetText - left blank on new accounts" }) `
    -ValueColor $StartupPalette.Text
if ((-not $UPNDomain) -or (-not $MailDomain)) {
    Write-StartupField `
        -Label "Before Creating Accounts" `
        -Value "Set UPN domain and Primary SMTP domain in MSToolkit Settings (New account naming), then reopen this tool." `
        -ValueColor $StartupPalette.Warning
}
Write-StartupField `
    -Label "Password at Next Logon" `
    -Value "User must change password is always enabled" `
    -ValueColor $StartupPalette.Warning

Write-StartupSection "CLONE / MIRROR OPTIONS"
Write-StartupField `
    -Label "Mirror Groups From" `
    -Value "Copies the selected user's direct AD group memberships to the new account." `
    -ValueColor $StartupPalette.Text
Write-StartupField `
    -Label "Clone Address From" `
    -Value "Copies street, city, state, ZIP, and country from the selected user." `
    -ValueColor $StartupPalette.Text
Write-StartupField `
    -Label "Clone Manager From" `
    -Value "Copies the selected user's Manager to the new account." `
    -ValueColor $StartupPalette.Text

Write-StartupSection "OU ENTRY"
Write-StartupField `
    -Label "Department OU" `
    -Value "Type only a child OU name such as IT or HR." `
    -ValueColor $StartupPalette.Text
Write-StartupField `
    -Label "Blank Department OU" `
    -Value "Creates the user directly in the OU selected in the picker. The domain root is never used." `
    -ValueColor $StartupPalette.MutedText
Write-StartupField `
    -Label "Full OU Path" `
    -Value "You may paste a full distinguishedName beginning with OU= when needed." `
    -ValueColor $StartupPalette.MutedText

try {
    $form.ShowDialog() | Out-Null
}
catch {
    Show-Error "SCRIPT ERROR:`r`n$($_.Exception.Message)"
}
