[CmdletBinding()]
param(
    [ValidateSet("Light","Dark")]
    [string]$ThemeMode
)

$ErrorActionPreference = 'Stop'

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

$script:WorkingServer = $null

function Get-MSToolkitServerArgs {
    # Splat into AD cmdlets so lookups go to a DC that answers on ADWS.
    $Result = @{}
    if ($script:WorkingServer) {
        $Result['Server'] = $script:WorkingServer
    }
    return $Result
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
            $ToolProperty = "Theme_InvestigateAccountLockout"
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
if (-not ("MSToolkitLockoutConsole.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace MSToolkitLockoutConsole
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
    $consoleHandle = [MSToolkitLockoutConsole.NativeMethods]::GetConsoleWindow()
    if ($consoleHandle -ne [IntPtr]::Zero) {
        [MSToolkitLockoutConsole.NativeMethods]::ShowWindow($consoleHandle, 0) | Out-Null
    }
}
catch {
    # Console hiding must never prevent the GUI from loading.
}

$script:CurrentUser = $null
$script:CurrentDcs = @()
$script:LockoutEvents = @()
$script:FailedEvents = @()
$script:BadPwdByDc = @()
$script:PdcHostName = $null
$script:InvestigationSummary = ''

function Show-InfoMessage {
    param(
        [string]$Message,
        [string]$Title = 'Account Lockout Investigation'
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Show-ErrorMessage {
    param([string]$Message)

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        'Account Lockout Investigation',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Write-AppLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level = 'INFO'
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

function Write-AppSeparator {
    if (-not $script:txtLog) { return }

    $script:txtLog.SelectionStart = $script:txtLog.TextLength
    $script:txtLog.SelectionLength = 0
    $script:txtLog.SelectionColor = Convert-MSToolkitThemeColor ([System.Drawing.Color]::FromArgb(120,120,120))
    $script:txtLog.AppendText("`r`n============================================================`r`n")
    $script:txtLog.SelectionColor = $script:txtLog.ForeColor
    $script:txtLog.ScrollToCaret()
}

function Get-LogonTypeDescription {
    param([object]$LogonType)

    $value = [string]$LogonType
    $description = switch ($value) {
        '2'  { 'Interactive' }
        '3'  { 'Network' }
        '4'  { 'Batch' }
        '5'  { 'Service' }
        '7'  { 'Unlock' }
        '8'  { 'NetworkCleartext' }
        '9'  { 'NewCredentials' }
        '10' { 'RemoteInteractive / RDP' }
        '11' { 'CachedInteractive' }
        default { 'Unknown' }
    }

    return "$value - $description"
}

function Get-AuthStatusDescription {
    param([object]$Code)

    $raw = [string]$Code
    if ([string]::IsNullOrWhiteSpace($raw) -or $raw -eq '0x0') {
        return $raw
    }

    $description = switch ($raw.ToUpperInvariant()) {
        '0XC0000064' { 'User name does not exist' }
        '0XC000006A' { 'Incorrect password' }
        '0XC000006D' { 'Bad user name or authentication information' }
        '0XC000006E' { 'Account restriction' }
        '0XC000006F' { 'Invalid logon hours' }
        '0XC0000070' { 'Unauthorized workstation' }
        '0XC0000071' { 'Expired password' }
        '0XC0000072' { 'Disabled account' }
        '0XC0000193' { 'Expired account' }
        '0XC0000224' { 'Password must be changed' }
        '0XC0000234' { 'Account locked out' }
        '0XC000015B' { 'Logon type not granted' }
        default      { 'Unknown / other status' }
    }

    return "$description ($raw)"
}

function Get-LikelyCauseText {
    $latestLockout = $script:LockoutEvents | Select-Object -First 1
    $recentFailed = @($script:FailedEvents | Select-Object -First 50)

    if ($recentFailed.Count -gt 0) {
        $sourceGroups = @(
            $recentFailed |
            ForEach-Object {
                $source = if (-not [string]::IsNullOrWhiteSpace([string]$_.SourceWorkstation) -and [string]$_.SourceWorkstation -ne '-') {
                    [string]$_.SourceWorkstation
                }
                elseif (-not [string]::IsNullOrWhiteSpace([string]$_.SourceIpAddress) -and [string]$_.SourceIpAddress -ne '-') {
                    [string]$_.SourceIpAddress
                }
                else {
                    '(source not recorded)'
                }
                [pscustomobject]@{ Source = $source }
            } |
            Group-Object Source |
            Sort-Object Count -Descending
        )

        $topSource = $sourceGroups | Select-Object -First 1
        $latestFailed = $recentFailed | Select-Object -First 1
        $reason = Get-AuthStatusDescription $latestFailed.SubStatus
        if ([string]::IsNullOrWhiteSpace($reason) -or $reason -eq '0x0') {
            $reason = Get-AuthStatusDescription $latestFailed.Status
        }

        return @"
Likely Cause / Evidence Summary:

Most frequent recent source: $($topSource.Name) ($($topSource.Count) failed logon(s) in the displayed recent sample)
Latest failed logon: $($latestFailed.TimeCreated)
Latest reporting DC: $($latestFailed.DomainController)
Latest logon type: $(Get-LogonTypeDescription $latestFailed.LogonType)
Latest failure code: $reason
Latest lockout source: $(if ($latestLockout) { $latestLockout.CallerComputerName } else { 'No 4740 source recorded' })

This is an evidence summary, not a guaranteed root-cause determination. Check the identified source for saved credentials, scheduled tasks, services, mapped drives, VPN, RDP sessions, or applications using an old password.
"@
    }

    if ($latestLockout) {
        return @"
Likely Cause / Evidence Summary:

A 4740 lockout event was found from: $($latestLockout.CallerComputerName)
Time: $($latestLockout.TimeCreated)
Reporting DC: $($latestLockout.DomainController)

No matching 4625 failed-logon details were collected. Investigate the caller computer for cached or saved credentials and re-run with 'Include failed logons (4625)' enabled if needed.
"@
    }

    return @"
Likely Cause / Evidence Summary:

No clear source was identified from the available domain-controller security logs in this lookback window.
"@
}

function Update-InvestigationSummary {
    if (-not $script:CurrentUser) {
        $script:InvestigationSummary = ''
        return
    }

    $latestLockout = $script:LockoutEvents | Select-Object -First 1
    $latestFailed = $script:FailedEvents | Select-Object -First 1

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Account Lockout Investigation Summary")
    $lines.Add("User: $($script:CurrentUser.SamAccountName) | $($script:CurrentUser.Name)")
    $lines.Add("Enabled: $($script:CurrentUser.Enabled)")
    $lines.Add("Locked Out: $($script:CurrentUser.LockedOut)")
    $lines.Add("Password Must Be Changed: $($script:PasswordMustBeChanged)")
    $lines.Add("Password Never Expires: $($script:PasswordNeverExpires)")
    $lines.Add("Password Last Set: $($script:CurrentUser.PasswordLastSet)")
    $lines.Add("Account Lockout Time: $($script:CurrentUser.AccountLockoutTime)")
    $lines.Add("PDC Emulator: $($script:PdcHostName)")
    $lines.Add("Lockout Events: $($script:LockoutEvents.Count)")
    $lines.Add("Failed Logons: $($script:FailedEvents.Count)")

    if ($latestLockout) {
        $lines.Add("Latest Lockout: $($latestLockout.TimeCreated)")
        $lines.Add("Latest Lockout Source: $($latestLockout.CallerComputerName)")
        $lines.Add("Latest Reporting DC: $($latestLockout.DomainController)")
    }

    if ($latestFailed) {
        $lines.Add("Latest Failed Logon: $($latestFailed.TimeCreated)")
        $lines.Add("Failed Logon Source Workstation: $($latestFailed.SourceWorkstation)")
        $lines.Add("Failed Logon Source IP: $($latestFailed.SourceIpAddress)")
        $lines.Add("Logon Type: $(Get-LogonTypeDescription $latestFailed.LogonType)")
        $lines.Add("Failure Status: $(Get-AuthStatusDescription $latestFailed.Status)")
        $lines.Add("Failure SubStatus: $(Get-AuthStatusDescription $latestFailed.SubStatus)")
    }

    $evidenceSummary = (Get-LikelyCauseText).Trim()
    $script:InvestigationSummary = (($lines -join "`r`n") + "`r`n`r`n" + $evidenceSummary)
}

function Copy-InvestigationSummary {
    if ([string]::IsNullOrWhiteSpace($script:InvestigationSummary)) {
        Show-InfoMessage 'Run an investigation first.'
        return
    }

    [System.Windows.Forms.Clipboard]::SetText($script:InvestigationSummary)
    Write-AppLog 'Investigation summary copied to the clipboard.' 'SUCCESS'
}

function Set-BusyState {
    param(
        [bool]$Busy,
        [string]$StatusText = ''
    )

    $script:btnInvestigate.Enabled = -not $Busy
    $script:txtUser.Enabled = -not $Busy
    $script:numLookback.Enabled = -not $Busy
    $script:chkFailed.Enabled = -not $Busy
    if ($script:btnCopySummary) { $script:btnCopySummary.Enabled = -not $Busy }

    if ($StatusText) {
        $script:lblStatus.Text = $StatusText
    }

    $script:MainForm.UseWaitCursor = $Busy

    # A ComboBox owns its own window handle, so clearing UseWaitCursor on the form
    # does not always restore the pointer over it. Reset the cursor explicitly.
    if (-not $Busy) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Clear-Results {
    $script:lblUserValue.Text = '-'
    $script:lblEnabledValue.Text = '-'
    $script:lblLockedValue.Text = '-'
    $script:lblPasswordLastSetValue.Text = '-'
    $script:lblLastBadValue.Text = '-'
    $script:lblBadPwdCountValue.Text = '-'
    $script:lblLatestSourceValue.Text = '-'
    $script:lblLatestLockoutValue.Text = '-'
    $script:lblLatestDcValue.Text = '-'
    $script:lblPwdMustChangeValue.Text = '-'
    $script:lblPwdNeverExpiresValue.Text = '-'
    $script:lblAccountLockoutTimeValue.Text = '-'
    $script:lblPdcValue.Text = '-'

    $script:dgvLockouts.Rows.Clear()
    $script:dgvBadPwd.Rows.Clear()
    $script:dgvFailed.Rows.Clear()

    $script:lblLockoutCount.Text = 'Lockout events: 0'
    $script:lblFailedCount.Text = 'Failed logons: 0'
    $script:InvestigationSummary = ''
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

function Get-WinEventNamedData {
    param(
        [Parameter(Mandatory)]
        $Event
    )

    $Data = @{}

    try {
        [xml]$EventXml = $Event.ToXml()

        foreach ($Node in @($EventXml.Event.EventData.Data)) {
            $Name = [string]$Node.Name
            if (-not [string]::IsNullOrWhiteSpace($Name)) {
                $Data[$Name] = [string]$Node.'#text'
            }
        }
    }
    catch {
        Write-AppLog "WARNING: Could not parse named event data for event $($Event.Id): $($_.Exception.Message)" 'WARNING'
    }

    return $Data
}

function Add-GridRow {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [object[]]$Values,
        [object]$Tag = $null
    )

    $index = $Grid.Rows.Add($Values)
    if ($null -ne $Tag) {
        $Grid.Rows[$index].Tag = $Tag
    }
    return $index
}

function Invoke-LockoutInvestigation {
    $userName = ConvertFrom-MSToolkitUserPickerLabel -Value $script:txtUser.Text
    $lookbackHours = [int]$script:numLookback.Value
    $includeFailed = $script:chkFailed.Checked

    if ([string]::IsNullOrWhiteSpace($userName)) {
        Show-InfoMessage 'Enter a username / sAMAccountName first.'
        return
    }

    try {
        Set-BusyState -Busy $true -StatusText "Investigating $userName..."
        Clear-Results
        Write-AppSeparator
        Write-AppLog "Starting investigation for: $userName"
        Write-AppLog "Resolving user: $userName"

        $start = (Get-Date).AddHours(-$lookbackHours)

        $InvServerArgs = Get-MSToolkitServerArgs

        $script:CurrentUser = Get-ADUser @InvServerArgs `
            -Identity $userName `
            -Properties LockedOut,badPwdCount,badPasswordTime,LastBadPasswordAttempt,PasswordLastSet,pwdLastSet,userAccountControl,Enabled,AccountLockoutTime,LastLogonDate,MemberOf,msDS-UserPasswordExpiryTimeComputed `
            -ErrorAction Stop

        $script:PasswordMustBeChanged = ([int64]$script:CurrentUser.pwdLastSet -eq 0)
        $script:PasswordNeverExpires = (([int64]$script:CurrentUser.userAccountControl -band 0x10000) -ne 0)

        $script:CurrentDcs = @(Get-ADDomainController @InvServerArgs -Filter * | Sort-Object HostName)
        $domainInfo = Get-ADDomain @InvServerArgs -ErrorAction Stop
        $script:PdcHostName = [string]$domainInfo.PDCEmulator

        Write-AppLog "Resolved $($script:CurrentUser.SamAccountName). Checking $($script:CurrentDcs.Count) domain controller(s)." 'SUCCESS'

        $script:lblUserValue.Text = "$($script:CurrentUser.SamAccountName)  |  $($script:CurrentUser.Name)"
        $script:lblEnabledValue.Text = [string]$script:CurrentUser.Enabled
        $script:lblLockedValue.Text = [string]$script:CurrentUser.LockedOut
        $script:lblPasswordLastSetValue.Text = if ($script:CurrentUser.PasswordLastSet) { [string]$script:CurrentUser.PasswordLastSet } else { '-' }
        $script:lblLastBadValue.Text = if ($script:CurrentUser.LastBadPasswordAttempt) { [string]$script:CurrentUser.LastBadPasswordAttempt } else { '-' }
        $script:lblBadPwdCountValue.Text = [string]$script:CurrentUser.badPwdCount
        $script:lblPwdMustChangeValue.Text = if ($script:PasswordMustBeChanged) { 'True' } else { 'False' }
        $script:lblPwdNeverExpiresValue.Text = if ($script:PasswordNeverExpires) { 'True' } else { 'False' }
        $script:lblAccountLockoutTimeValue.Text = if ($script:CurrentUser.AccountLockoutTime) { [string]$script:CurrentUser.AccountLockoutTime } else { '-' }
        $script:lblPdcValue.Text = if ($script:PdcHostName) { $script:PdcHostName } else { '-' }

        $script:lblEnabledValue.ForeColor = if ($script:CurrentUser.Enabled) { (Get-MSToolkitThemePalette).Success } else { (Get-MSToolkitThemePalette).Danger }
        $script:lblPwdMustChangeValue.ForeColor = if ($script:PasswordMustBeChanged) { (Get-MSToolkitThemePalette).Danger } else { (Get-MSToolkitThemePalette).Success }
        $script:lblPwdNeverExpiresValue.ForeColor = if ($script:PasswordNeverExpires) { (Get-MSToolkitThemePalette).Danger } else { (Get-MSToolkitThemePalette).Success }

        if ($script:CurrentUser.LockedOut -eq $true) {
            $script:lblLockedValue.ForeColor = (Get-MSToolkitThemePalette).Danger
        }
        else {
            $script:lblLockedValue.ForeColor = (Get-MSToolkitThemePalette).Success
        }

        # 4740 lockout events
        $lockoutList = New-Object System.Collections.Generic.List[object]

        foreach ($dc in $script:CurrentDcs) {
            try {
                Write-AppLog "Checking 4740 lockout events on $($dc.HostName)..."

                $events = Get-WinEvent `
                    -ComputerName $dc.HostName `
                    -FilterHashtable @{ LogName='Security'; Id=4740; StartTime=$start } `
                    -ErrorAction Stop

                foreach ($event in $events) {
                    $eventData = Get-WinEventNamedData -Event $event
                    $targetUserName = [string]$eventData['TargetUserName']
                    $callerComputerName = [string]$eventData['CallerComputerName']

                    if ($targetUserName -ieq $script:CurrentUser.SamAccountName) {
                        $lockoutList.Add([pscustomobject]@{
                            EventType          = 'Account Lockout'
                            TimeCreated        = $event.TimeCreated
                            DomainController   = $dc.HostName
                            TargetUser         = $targetUserName
                            CallerComputerName = $callerComputerName
                            EventId            = $event.Id
                        })
                    }
                }
            }
            catch {
                Write-AppLog "Could not read lockout events on $($dc.HostName): $($_.Exception.Message)" 'WARNING'
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        $script:LockoutEvents = @($lockoutList.ToArray() | Sort-Object TimeCreated -Descending)

        foreach ($item in $script:LockoutEvents) {
            [void](Add-GridRow -Grid $script:dgvLockouts -Values @(
                $item.TimeCreated,
                $item.DomainController,
                $item.TargetUser,
                $item.CallerComputerName
            ) -Tag $item)
        }

        $script:lblLockoutCount.Text = "Lockout events: $($script:LockoutEvents.Count)"

        $latest = $script:LockoutEvents | Select-Object -First 1
        if ($latest) {
            $script:lblLatestSourceValue.Text = if ([string]::IsNullOrWhiteSpace([string]$latest.CallerComputerName)) { '(not recorded)' } else { [string]$latest.CallerComputerName }
            $script:lblLatestLockoutValue.Text = [string]$latest.TimeCreated
            $script:lblLatestDcValue.Text = [string]$latest.DomainController
        }
        else {
            $script:lblLatestSourceValue.Text = 'No lockout events found'
            $script:lblLatestLockoutValue.Text = '-'
            $script:lblLatestDcValue.Text = '-'
        }

        # Per-DC bad password state
        $badPwdList = New-Object System.Collections.Generic.List[object]

        foreach ($dc in $script:CurrentDcs) {
            try {
                Write-AppLog "Checking bad-password counters on $($dc.HostName)..."

                # Without this a DC whose ADWS port does not answer costs ~42 seconds.
                if (-not (Test-MSToolkitAdwsPort -HostName $dc.HostName)) {
                    throw "ADWS (port 9389) is not reachable on $($dc.HostName) from this computer, so its counters could not be read."
                }

                $dcUser = Get-ADUser `
                    -Server $dc.HostName `
                    -Identity $script:CurrentUser.SamAccountName `
                    -Properties badPwdCount,badPasswordTime,LastBadPasswordAttempt,LockedOut,AccountLockoutTime,LastLogonDate `
                    -ErrorAction Stop

                $badPwdList.Add([pscustomobject]@{
                    DomainController       = $dc.HostName
                    LockedOut              = $dcUser.LockedOut
                    BadPwdCount            = $dcUser.badPwdCount
                    LastBadPasswordAttempt = $dcUser.LastBadPasswordAttempt
                    AccountLockoutTime     = $dcUser.AccountLockoutTime
                    LastLogonDate          = $dcUser.LastLogonDate
                    Status                 = 'OK'
                    Error                  = $null
                })
            }
            catch {
                Write-AppLog "Could not query $($dc.HostName): $($_.Exception.Message)" 'WARNING'

                $badPwdList.Add([pscustomobject]@{
                    DomainController       = $dc.HostName
                    LockedOut              = $null
                    BadPwdCount            = $null
                    LastBadPasswordAttempt = $null
                    AccountLockoutTime     = $null
                    LastLogonDate          = $null
                    Status                 = 'Error'
                    Error                  = $_.Exception.Message
                })
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        $script:BadPwdByDc = @(
            $badPwdList.ToArray() |
            Sort-Object -Property @{Expression='BadPwdCount';Descending=$true}, @{Expression='LastBadPasswordAttempt';Descending=$true}
        )

        $maxBadPwdCount = ($script:BadPwdByDc | Where-Object { $_.Status -eq 'OK' } | Measure-Object -Property BadPwdCount -Maximum).Maximum

        foreach ($item in $script:BadPwdByDc) {
            $rowIndex = Add-GridRow -Grid $script:dgvBadPwd -Values @(
                $(if ($script:PdcHostName -and $item.DomainController -ieq $script:PdcHostName) { "$($item.DomainController) (PDC)" } else { $item.DomainController }),
                $item.LockedOut,
                $item.BadPwdCount,
                $item.LastBadPasswordAttempt,
                $item.AccountLockoutTime,
                $item.LastLogonDate,
                $item.Status
            ) -Tag $item

            if ($item.Status -eq 'Error') {
                $script:dgvBadPwd.Rows[$rowIndex].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).WarningBackground
            }
            elseif ($item.LockedOut -eq $true) {
                $script:dgvBadPwd.Rows[$rowIndex].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).DangerBackground
            }
            elseif ($null -ne $maxBadPwdCount -and [int]$item.BadPwdCount -eq [int]$maxBadPwdCount -and [int]$maxBadPwdCount -gt 0) {
                $script:dgvBadPwd.Rows[$rowIndex].DefaultCellStyle.BackColor = (Get-MSToolkitThemePalette).InfoBackground
            }
        }

        # Optional 4625 failed logons
        $failedList = New-Object System.Collections.Generic.List[object]

        if ($includeFailed) {
            foreach ($dc in $script:CurrentDcs) {
                try {
                    Write-AppLog "Checking 4625 failed logons on $($dc.HostName)..."

                    $events = Get-WinEvent `
                        -ComputerName $dc.HostName `
                        -FilterHashtable @{ LogName='Security'; Id=4625; StartTime=$start } `
                        -ErrorAction Stop

                    foreach ($event in $events) {
                        if ($event.Properties[5].Value -ieq $script:CurrentUser.SamAccountName -or
                            $event.Properties[5].Value -ieq $userName) {

                            $failedList.Add([pscustomobject]@{
                                EventType         = 'Failed Logon'
                                TimeCreated       = $event.TimeCreated
                                DomainController  = $dc.HostName
                                TargetUser        = $event.Properties[5].Value
                                SourceWorkstation = $event.Properties[13].Value
                                SourceIpAddress   = $event.Properties[19].Value
                                LogonType         = $event.Properties[10].Value
                                FailureReason     = $event.Properties[8].Value
                                Status            = $event.Properties[7].Value
                                SubStatus         = $event.Properties[9].Value
                                EventId           = $event.Id
                            })
                        }
                    }
                }
                catch {
                    Write-AppLog "Could not read failed logons on $($dc.HostName): $($_.Exception.Message)" 'WARNING'
                }

                [System.Windows.Forms.Application]::DoEvents()
            }
        }

        $script:FailedEvents = @($failedList.ToArray() | Sort-Object TimeCreated -Descending)

        foreach ($item in ($script:FailedEvents | Select-Object -First 200)) {
            [void](Add-GridRow -Grid $script:dgvFailed -Values @(
                $item.TimeCreated,
                $item.DomainController,
                $item.TargetUser,
                $item.SourceWorkstation,
                $item.SourceIpAddress,
                (Get-LogonTypeDescription $item.LogonType),
                $item.FailureReason,
                (Get-AuthStatusDescription $item.Status),
                (Get-AuthStatusDescription $item.SubStatus)
            ) -Tag $item)
        }

        $script:lblFailedCount.Text = "Failed logons: $($script:FailedEvents.Count)"

        $script:txtTroubleshoot.Text = (Get-LikelyCauseText) + "`r`n" + $script:BaseTroubleshootingText
        Update-InvestigationSummary

        $script:lblStatus.Text = "Investigation complete. Lockouts: $($script:LockoutEvents.Count) | Failed logons: $($script:FailedEvents.Count)"
        Write-AppLog "Investigation complete. Lockouts=$($script:LockoutEvents.Count), failed logons=$($script:FailedEvents.Count)." 'SUCCESS'

        if ($script:LockoutEvents.Count -eq 0) {
            Show-InfoMessage "No 4740 account-lockout events were found for $($script:CurrentUser.SamAccountName) in the last $lookbackHours hour(s)." 'No Lockout Events Found'
        }
    }
    catch {
        $script:lblStatus.Text = 'Investigation failed.'
        Write-AppLog "Investigation failed: $($_.Exception.Message)" 'ERROR'
        Show-ErrorMessage "Unable to complete the investigation.`r`n`r`n$($_.Exception.Message)"
    }
    finally {
        Set-BusyState -Busy $false
    }
}

# ---------------- GUI ----------------

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'Account Lockout Investigation'
$MainForm.StartPosition = 'CenterScreen'
$MainForm.Size = New-Object System.Drawing.Size(1220,820)
$MainForm.MinimumSize = New-Object System.Drawing.Size(1080,720)
$MainForm.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
$MainForm.Font = New-Object System.Drawing.Font('Segoe UI',9)
$MainForm.FormBorderStyle = 'Sizable'
$MainForm.MaximizeBox = $true
$script:MainForm = $MainForm

# Header
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = 'Top'
$pnlHeader.Height = 78
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$MainForm.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Account Lockout Investigation'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',20)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20,12)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Investigate AD account lockouts, bad-password counters, source systems, and failed logons across domain controllers'
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI',9.5)
$lblSubtitle.AutoSize = $true
$lblSubtitle.Location = New-Object System.Drawing.Point(23,49)
$pnlHeader.Controls.Add($lblSubtitle)
$script:MSToolkitThemeToggleButton = New-MSToolkitThemeToggleButton -HeaderPanel $pnlHeader -Form $MainForm -ToolKey "InvestigateAccountLockout"

# Search panel
$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = 'Investigation'
$grpSearch.Location = New-Object System.Drawing.Point(20,95)
$grpSearch.Size = New-Object System.Drawing.Size(1165,105)
$grpSearch.Anchor = 'Top,Left,Right'
$grpSearch.BackColor = [System.Drawing.Color]::White
$grpSearch.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9.5)
$MainForm.Controls.Add($grpSearch)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = 'Username / sAMAccountName'
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(18,25)
$grpSearch.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.ComboBox
$txtUser.Location = New-Object System.Drawing.Point(20,49)
$txtUser.Size = New-Object System.Drawing.Size(360,24)
$txtUser.Font = New-Object System.Drawing.Font('Segoe UI',10)
$txtUser.DropDownStyle = 'DropDown'
$txtUser.AutoCompleteMode = 'None'
$txtUser.AutoCompleteSource = 'None'
$txtUser.MaxDropDownItems = 20
$grpSearch.Controls.Add($txtUser)
$script:txtUser = $txtUser

# What the user list covers, shown only when an OU setting leaves a gap:
#   no Users OU            -> orange: the list covers the whole domain
#   Users OU, no Admin OU  -> orange: admin accounts outside the Users OU are not listed
# Its colour is applied after the theme pass (see $script:MSToolkitThemeRefreshHook).
$UsersOUSet  = [bool](Get-MSToolkitSetting -Name "OUUsers")
$AdminsOUSet = [bool](Get-MSToolkitSetting -Name "OUAdmins")
$script:UserListNoteIsWarning = $true

$lblUsersOUWarning = New-Object System.Windows.Forms.Label
$lblUsersOUWarning.Text = if (-not $UsersOUSet) { 'No Users OU is set in MSToolkit Settings, so the user list covers the whole domain.' } elseif (-not $AdminsOUSet) { "No Admin accounts OU is set in MSToolkit Settings: admin accounts outside the Users OU aren't listed (typing a name still works)." } else { '' }
$lblUsersOUWarning.AutoSize = $true
$lblUsersOUWarning.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$lblUsersOUWarning.Location = New-Object System.Drawing.Point(20,80)
$lblUsersOUWarning.Visible = (-not $UsersOUSet) -or (-not $AdminsOUSet)
$grpSearch.Controls.Add($lblUsersOUWarning)
$script:lblUsersOUWarning = $lblUsersOUWarning

$lblLookback = New-Object System.Windows.Forms.Label
$lblLookback.Text = 'Lookback Hours'
$lblLookback.AutoSize = $true
$lblLookback.Location = New-Object System.Drawing.Point(410,25)
$grpSearch.Controls.Add($lblLookback)

$numLookback = New-Object System.Windows.Forms.NumericUpDown
$numLookback.Location = New-Object System.Drawing.Point(412,49)
$numLookback.Size = New-Object System.Drawing.Size(110,24)
$numLookback.Minimum = 1
$numLookback.Maximum = 720
$numLookback.Value = 24
$grpSearch.Controls.Add($numLookback)
$script:numLookback = $numLookback

$chkFailed = New-Object System.Windows.Forms.CheckBox
$chkFailed.Text = 'Include failed logons (4625)'
$chkFailed.Location = New-Object System.Drawing.Point(560,48)
$chkFailed.Size = New-Object System.Drawing.Size(210,25)
$grpSearch.Controls.Add($chkFailed)
$script:chkFailed = $chkFailed

$btnInvestigate = New-Object System.Windows.Forms.Button
$btnInvestigate.Text = 'Investigate'
$btnInvestigate.Location = New-Object System.Drawing.Point(975,40)
$btnInvestigate.Size = New-Object System.Drawing.Size(160,38)
$btnInvestigate.Anchor = 'Top,Right'
$btnInvestigate.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$btnInvestigate.ForeColor = [System.Drawing.Color]::White
$btnInvestigate.FlatStyle = 'Flat'
$btnInvestigate.FlatAppearance.BorderSize = 0
$btnInvestigate.Font = New-Object System.Drawing.Font('Segoe UI Semibold',10)
$btnInvestigate.Add_Click({ Invoke-LockoutInvestigation })
$grpSearch.Controls.Add($btnInvestigate)
$script:btnInvestigate = $btnInvestigate

$btnCopySummary = New-Object System.Windows.Forms.Button
$btnCopySummary.Text = 'Copy Summary'
$btnCopySummary.Location = New-Object System.Drawing.Point(790,40)
$btnCopySummary.Size = New-Object System.Drawing.Size(160,38)
$btnCopySummary.Anchor = 'Top,Right'
$btnCopySummary.BackColor = [System.Drawing.Color]::FromArgb(70,90,110)
$btnCopySummary.ForeColor = [System.Drawing.Color]::White
$btnCopySummary.FlatStyle = 'Flat'
$btnCopySummary.FlatAppearance.BorderSize = 0
$btnCopySummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',10)
$btnCopySummary.Add_Click({ Copy-InvestigationSummary })
$grpSearch.Controls.Add($btnCopySummary)
$script:btnCopySummary = $btnCopySummary

# Highlight summary panel
$grpSummary = New-Object System.Windows.Forms.GroupBox
$grpSummary.Text = 'Current Account / Latest Lockout'
$grpSummary.Location = New-Object System.Drawing.Point(20,212)
$grpSummary.Size = New-Object System.Drawing.Size(1165,200)
$grpSummary.Anchor = 'Top,Left,Right'
$grpSummary.BackColor = [System.Drawing.Color]::White
$grpSummary.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9.5)
$MainForm.Controls.Add($grpSummary)

function Add-SummaryItem {
    param(
        [string]$LabelText,
        [int]$X,
        [int]$Y,
        [int]$Width = 330
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Location = New-Object System.Drawing.Point($X,$Y)
    $label.Size = New-Object System.Drawing.Size($Width,18)
    $label.ForeColor = [System.Drawing.Color]::DimGray
    $label.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
    $grpSummary.Controls.Add($label)

    $value = New-Object System.Windows.Forms.Label
    $value.Text = '-'
    $value.Location = New-Object System.Drawing.Point($X,($Y + 20))
    $value.Size = New-Object System.Drawing.Size($Width,22)
    $value.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9.5)
    $value.AutoEllipsis = $true
    $grpSummary.Controls.Add($value)

    return $value
}

$script:lblUserValue = Add-SummaryItem 'User' 20 24 300
$script:lblEnabledValue = Add-SummaryItem 'Enabled' 340 24 90
$script:lblLockedValue = Add-SummaryItem 'Locked Out' 450 24 100
$script:lblPwdMustChangeValue = Add-SummaryItem 'Password Must Be Changed' 570 24 170
$script:lblPwdNeverExpiresValue = Add-SummaryItem 'Password Never Expires' 760 24 160
$script:lblBadPwdCountValue = Add-SummaryItem 'Bad Password Count' 940 24 180

$script:lblLatestSourceValue = Add-SummaryItem 'Latest Lockout Source' 20 79 280
$script:lblLatestLockoutValue = Add-SummaryItem 'Latest Lockout Time' 320 79 220
$script:lblLatestDcValue = Add-SummaryItem 'Reporting DC' 560 79 180
$script:lblLastBadValue = Add-SummaryItem 'Last Bad Password Attempt' 760 79 360

$script:lblPasswordLastSetValue = Add-SummaryItem 'Password Last Set' 20 134 260
$script:lblAccountLockoutTimeValue = Add-SummaryItem 'Account Lockout Time' 300 134 260
$script:lblPdcValue = Add-SummaryItem 'PDC Emulator' 580 134 260

# Tabs
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(20,425)
$tabs.Size = New-Object System.Drawing.Size(1165,205)
$tabs.Anchor = 'Top,Bottom,Left,Right'
$MainForm.Controls.Add($tabs)

# Lockouts tab
$tabLockouts = New-Object System.Windows.Forms.TabPage
$tabLockouts.Text = 'Lockout Events'
$tabLockouts.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabLockouts)

$lblLockoutCount = New-Object System.Windows.Forms.Label
$lblLockoutCount.Text = 'Lockout events: 0'
$lblLockoutCount.Location = New-Object System.Drawing.Point(10,10)
$lblLockoutCount.Size = New-Object System.Drawing.Size(300,20)
$lblLockoutCount.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabLockouts.Controls.Add($lblLockoutCount)
$script:lblLockoutCount = $lblLockoutCount

$dgvLockouts = New-Object System.Windows.Forms.DataGridView
$dgvLockouts.Location = New-Object System.Drawing.Point(10,35)
$dgvLockouts.Size = New-Object System.Drawing.Size(1135,170)
$dgvLockouts.Anchor = 'Top,Bottom,Left,Right'
$dgvLockouts.AllowUserToAddRows = $false
$dgvLockouts.AllowUserToDeleteRows = $false
$dgvLockouts.ReadOnly = $true
$dgvLockouts.RowHeadersVisible = $false
$dgvLockouts.SelectionMode = 'FullRowSelect'
$dgvLockouts.MultiSelect = $false
$dgvLockouts.AutoSizeColumnsMode = 'Fill'
$dgvLockouts.BackgroundColor = [System.Drawing.Color]::White
[void]$dgvLockouts.Columns.Add('TimeCreated','Time')
[void]$dgvLockouts.Columns.Add('DomainController','Domain Controller')
[void]$dgvLockouts.Columns.Add('TargetUser','User')
[void]$dgvLockouts.Columns.Add('CallerComputerName','Caller Computer')
$tabLockouts.Controls.Add($dgvLockouts)
$script:dgvLockouts = $dgvLockouts

# Bad password tab
$tabBadPwd = New-Object System.Windows.Forms.TabPage
$tabBadPwd.Text = 'Bad Password by DC'
$tabBadPwd.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabBadPwd)

$dgvBadPwd = New-Object System.Windows.Forms.DataGridView
$dgvBadPwd.Dock = 'Fill'
$dgvBadPwd.AllowUserToAddRows = $false
$dgvBadPwd.AllowUserToDeleteRows = $false
$dgvBadPwd.ReadOnly = $true
$dgvBadPwd.RowHeadersVisible = $false
$dgvBadPwd.SelectionMode = 'FullRowSelect'
$dgvBadPwd.MultiSelect = $false
$dgvBadPwd.AutoSizeColumnsMode = 'Fill'
$dgvBadPwd.BackgroundColor = [System.Drawing.Color]::White
[void]$dgvBadPwd.Columns.Add('DomainController','Domain Controller')
[void]$dgvBadPwd.Columns.Add('LockedOut','Locked Out')
[void]$dgvBadPwd.Columns.Add('BadPwdCount','BadPwd Count')
[void]$dgvBadPwd.Columns.Add('LastBadPasswordAttempt','Last Bad Password')
[void]$dgvBadPwd.Columns.Add('AccountLockoutTime','Lockout Time')
[void]$dgvBadPwd.Columns.Add('LastLogonDate','Last Logon')
[void]$dgvBadPwd.Columns.Add('Status','Status')
$tabBadPwd.Controls.Add($dgvBadPwd)
$script:dgvBadPwd = $dgvBadPwd

# Failed logons tab
$tabFailed = New-Object System.Windows.Forms.TabPage
$tabFailed.Text = 'Failed Logons'
$tabFailed.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabFailed)

$lblFailedCount = New-Object System.Windows.Forms.Label
$lblFailedCount.Text = 'Failed logons: 0'
$lblFailedCount.Location = New-Object System.Drawing.Point(10,10)
$lblFailedCount.Size = New-Object System.Drawing.Size(300,20)
$lblFailedCount.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$tabFailed.Controls.Add($lblFailedCount)
$script:lblFailedCount = $lblFailedCount

$dgvFailed = New-Object System.Windows.Forms.DataGridView
$dgvFailed.Location = New-Object System.Drawing.Point(10,35)
$dgvFailed.Size = New-Object System.Drawing.Size(1135,170)
$dgvFailed.Anchor = 'Top,Bottom,Left,Right'
$dgvFailed.AllowUserToAddRows = $false
$dgvFailed.AllowUserToDeleteRows = $false
$dgvFailed.ReadOnly = $true
$dgvFailed.RowHeadersVisible = $false
$dgvFailed.SelectionMode = 'FullRowSelect'
$dgvFailed.MultiSelect = $false
$dgvFailed.AutoSizeColumnsMode = 'Fill'
$dgvFailed.BackgroundColor = [System.Drawing.Color]::White
[void]$dgvFailed.Columns.Add('TimeCreated','Time')
[void]$dgvFailed.Columns.Add('DomainController','DC')
[void]$dgvFailed.Columns.Add('TargetUser','User')
[void]$dgvFailed.Columns.Add('SourceWorkstation','Workstation')
[void]$dgvFailed.Columns.Add('SourceIpAddress','Source IP')
[void]$dgvFailed.Columns.Add('LogonType','Logon Type')
[void]$dgvFailed.Columns.Add('FailureReason','Failure Reason')
[void]$dgvFailed.Columns.Add('Status','Status')
[void]$dgvFailed.Columns.Add('SubStatus','SubStatus')
$tabFailed.Controls.Add($dgvFailed)
$script:dgvFailed = $dgvFailed

# Troubleshooting tab
$tabTroubleshoot = New-Object System.Windows.Forms.TabPage
$tabTroubleshoot.Text = 'Troubleshooting'
$tabTroubleshoot.BackColor = [System.Drawing.Color]::White
$tabs.TabPages.Add($tabTroubleshoot)

$txtTroubleshoot = New-Object System.Windows.Forms.RichTextBox
$txtTroubleshoot.Dock = 'Fill'
$txtTroubleshoot.ReadOnly = $true
$txtTroubleshoot.BackColor = [System.Drawing.Color]::White
$txtTroubleshoot.Font = New-Object System.Drawing.Font('Segoe UI',9.5)
$script:BaseTroubleshootingText = @"
General causes to check:

• Mapped drives or saved Windows credentials using an old password
• Mobile mail client or Outlook profile with cached credentials
• VPN client saved credentials or failed reconnect attempts
• Scheduled task running as the user
• Windows service configured to run as the user
• Disconnected RDP session with old credentials
• Application pool, script, printer scan-to-folder, or third-party app using the user account
• Repeated Wi-Fi authentication attempts on a domain-joined laptop

Helpful commands for a suspected source workstation:

cmdkey /list

Get-ScheduledTask | Where-Object { `$_.Principal.UserId -match "username" }

Get-CimInstance Win32_Service | Where-Object { `$_.StartName -match "username" }

net use
"@
$txtTroubleshoot.Text = $script:BaseTroubleshootingText
$script:txtTroubleshoot = $txtTroubleshoot
$tabTroubleshoot.Controls.Add($txtTroubleshoot)

# Activity log
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Activity Log'
$grpLog.Location = New-Object System.Drawing.Point(20,640)
$grpLog.Size = New-Object System.Drawing.Size(1165,130)
$grpLog.Anchor = 'Bottom,Left,Right'
$grpLog.BackColor = [System.Drawing.Color]::White
$grpLog.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$MainForm.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(10,20)
$txtLog.Size = New-Object System.Drawing.Size(1145,98)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.ReadOnly = $true
$txtLog.BorderStyle = 'FixedSingle'
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = New-Object System.Drawing.Font('Consolas',8.5)
$txtLog.ScrollBars = 'Vertical'
$txtLog.WordWrap = $false
$grpLog.Controls.Add($txtLog)
$script:txtLog = $txtLog

# Status bar
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.SizingGrip = $false
$MainForm.Controls.Add($statusStrip)

$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = 'Ready.'
$lblStatus.Spring = $true
$lblStatus.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($lblStatus)
$script:lblStatus = $lblStatus

$txtUser.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Invoke-LockoutInvestigation
    }
})

$MainForm.Add_Shown({
    $MainForm.Activate()
    $txtUser.Focus()
    Write-AppLog 'Account Lockout Investigation GUI loaded.'

    $script:WorkingServer = Resolve-MSToolkitWorkingServer -DomainName (Get-MSToolkitLocalDomainName)
    if ($script:WorkingServer) {
        Write-AppLog "Using AD server: $($script:WorkingServer)"
    }
    else {
        Write-AppLog 'No domain controller answered on ADWS (port 9389). Lookups may be slow or fail.' 'WARNING'
    }

    $LoadedCount = Initialize-MSToolkitUserPicker -Combos @($txtUser) -Server ([string]$script:WorkingServer)

    if ($LoadedCount -ge 0) {
        Write-AppLog "Loaded $LoadedCount employee account(s) into the user list."
        if (-not (Get-MSToolkitSetting -Name "OUUsers")) {
            Write-AppLog 'No Users OU is set in MSToolkit Settings, so the user list covers the whole domain. Set it under Settings > Organizational units to narrow it.' 'WARNING'
        }
    }
    else {
        Write-AppLog 'Could not load the employee user list. Type a username instead.' 'WARNING'
    }

    try {
        $ShownServerArgs = Get-MSToolkitServerArgs
        $domainInfo = Get-ADDomain @ShownServerArgs -ErrorAction Stop
        Write-AppLog "PDC Emulator: $($domainInfo.PDCEmulator)"
    } catch { }
})

Apply-MSToolkitSharedTheme -Root $MainForm

# The shared theme pass cannot infer the note's colour; set it now and after every
# theme toggle.
$script:MSToolkitThemeRefreshHook = {
    if ($script:lblUsersOUWarning) {
        $NotePalette = Get-MSToolkitThemePalette
        $script:lblUsersOUWarning.ForeColor = if ($script:UserListNoteIsWarning) { $NotePalette.Warning } else { $NotePalette.MutedText }
    }
}
& $script:MSToolkitThemeRefreshHook
# Type-ahead filtering on every editable dropdown. Its own handler, not folded
# into another, so it runs on every launch rather than only when the tool is
# started with parameters - and so a failure here cannot stop the rest of
# start-up. Multiple Shown handlers chain.
$MainForm.Add_Shown({ Register-MSToolkitComboFiltersOn -Root $MainForm })

[void]$MainForm.ShowDialog()
