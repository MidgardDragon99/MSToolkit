$ErrorActionPreference = "Stop"

# Do not create the AD: PowerShell drive. MSToolkit never uses it, and building it means
# connecting to a domain controller before anything else can run. If that DC does not
# answer on ADWS (port 9389), startup waits out two ~21 second connection timeouts.
$env:ADPS_LoadDefaultDrive = 0
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WindowHelper {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags
    );

    public static IntPtr FindLargestVisibleWindowForProcess(int processId) {
        IntPtr bestHandle = IntPtr.Zero;
        long bestArea = 0;

        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            if (!IsWindowVisible(hWnd)) {
                return true;
            }

            uint windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);
            if (windowProcessId != (uint)processId) {
                return true;
            }

            RECT rect;
            if (!GetWindowRect(hWnd, out rect)) {
                return true;
            }

            long width = Math.Max(0, rect.Right - rect.Left);
            long height = Math.Max(0, rect.Bottom - rect.Top);
            long area = width * height;

            if (area > bestArea) {
                bestArea = area;
                bestHandle = hWnd;
            }

            return true;
        }, IntPtr.Zero);

        return bestHandle;
    }

    public static bool CenterWindowInWorkingArea(
        IntPtr hWnd,
        int areaLeft,
        int areaTop,
        int areaWidth,
        int areaHeight
    ) {
        RECT rect;
        if (!GetWindowRect(hWnd, out rect)) {
            return false;
        }

        int windowWidth = Math.Max(1, rect.Right - rect.Left);
        int windowHeight = Math.Max(1, rect.Bottom - rect.Top);

        int windowCenterX = rect.Left + (windowWidth / 2);
        int windowCenterY = rect.Top + (windowHeight / 2);

        // If Windows already opened the window on the MSToolkit monitor, leave its
        // normal placement alone. Only reposition windows that landed elsewhere.
        if (
            windowCenterX >= areaLeft &&
            windowCenterX < areaLeft + areaWidth &&
            windowCenterY >= areaTop &&
            windowCenterY < areaTop + areaHeight
        ) {
            return true;
        }

        int x = areaLeft + Math.Max(0, (areaWidth - windowWidth) / 2);
        int y = areaTop + Math.Max(0, (areaHeight - windowHeight) / 2);

        const uint SWP_NOSIZE = 0x0001;
        const uint SWP_NOZORDER = 0x0004;
        const uint SWP_NOACTIVATE = 0x0010;

        return SetWindowPos(
            hWnd,
            IntPtr.Zero,
            x,
            y,
            0,
            0,
            SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE
        );
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    // Reserves space on the right of an edit control so an overlaid button does not
    // sit on top of the text. Used by the password reveal in the M365 sign-in prompt.
    public static void SetRightMargin(IntPtr handle, int margin) {
        const int EM_SETMARGINS = 0xD3;
        const int EC_RIGHTMARGIN = 0x2;
        SendMessage(handle, EM_SETMARGINS, (IntPtr)EC_RIGHTMARGIN, (IntPtr)(margin << 16));
    }
}
"@

$ConsolePtr = [WindowHelper]::GetConsoleWindow()

if ($ConsolePtr -ne [IntPtr]::Zero -and $host.Name -notlike "*ISE*") {
    [WindowHelper]::ShowWindow($ConsolePtr, 0) | Out-Null
}

function Get-CurrentADDomain {
    # Windows already knows which domain this computer is joined to, so read it
    # locally first. This needs no network call and cannot land on an unreachable DC.
    try {
        $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop

        if ($ComputerSystem.PartOfDomain -and -not [string]::IsNullOrWhiteSpace($ComputerSystem.Domain)) {
            return [string]$ComputerSystem.Domain
        }
    }
    catch {
        # Fall through to asking Active Directory directly.
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


function Test-MSToolkitAdwsPort {
    param(
        [string]$HostName,
        [int]$TimeoutMilliseconds = 2000
    )

    # Quick check that a DC answers on ADWS (TCP 9389) before handing it to the AD
    # module, which would otherwise wait out two ~21 second connection timeouts.
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

function Get-MSToolkitDomainControllers {
    # Builds the DC list for the dropdown. Only the names are needed, so get them over
    # plain LDAP (port 389) through System.DirectoryServices. That works against every
    # DC and never touches ADWS (port 9389), so a DC whose ADWS port is unreachable
    # cannot stall or break startup.
    try {
        $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $Domain)
        $DomainObject = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)

        $LdapControllers = @(
            $DomainObject.DomainControllers |
                ForEach-Object { ([string]$_.Name -split '\.')[0].ToUpper() } |
                Sort-Object -Unique
        )

        if ($LdapControllers.Count -gt 0) {
            return $LdapControllers
        }
    }
    catch {
        # Fall through to the ADWS-based methods below.
    }

    # Fallbacks, only used if the LDAP lookup fails:
    #   1. The DC saved in settings.json
    #   2. A DC the locator reports as offering ADWS
    #   3. The domain name itself - the original behaviour
    # Each named DC is checked on port 9389 first, so one that will not answer is
    # skipped in about 2 seconds instead of stalling startup for 40+.
    $Candidates = New-Object System.Collections.Generic.List[string]

    try {
        $SettingsFile = Join-Path (Join-Path $env:APPDATA "MSToolkit") "settings.json"

        if (Test-Path -LiteralPath $SettingsFile) {
            $SavedDC = [string](Get-Content -LiteralPath $SettingsFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop).SelectedDC

            if (-not [string]::IsNullOrWhiteSpace($SavedDC)) {
                $Candidates.Add("$SavedDC.$Domain")
            }
        }
    }
    catch {
        # No usable saved preference; carry on with discovery.
    }

    try {
        $SeedDC = Get-ADDomainController -Discover -DomainName $Domain -Service ADWS -ErrorAction Stop
        $SeedHost = [string]($SeedDC.HostName | Select-Object -First 1)

        if (-not [string]::IsNullOrWhiteSpace($SeedHost)) {
            $Candidates.Add($SeedHost)
        }
    }
    catch {
        # Discovery failed; fall back to the domain name below.
    }

    $LastError = $null
    $Tried = New-Object System.Collections.Generic.List[string]

    foreach ($QueryServer in $Candidates) {
        if ($Tried.Contains($QueryServer.ToLower())) { continue }
        $Tried.Add($QueryServer.ToLower())

        if (-not (Test-MSToolkitAdwsPort -HostName $QueryServer)) {
            continue
        }

        try {
            $Controllers = @(
                Get-ADDomainController `
                    -Filter * `
                    -Server $QueryServer `
                    -ErrorAction Stop |
                    Sort-Object Name |
                    Select-Object -ExpandProperty Name
            )

            if ($Controllers.Count -gt 0) {
                return $Controllers
            }
        }
        catch {
            $LastError = $_
        }
    }

    # Last resort: the original behaviour, letting the AD module pick a DC itself.
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

$LogPath = Join-Path $PSScriptRoot "Logs"
New-Item -ItemType Directory -Path $LogPath -Force | Out-Null

# Per-user UI preference. If this file does not exist or cannot be read,
# MSToolkit starts in Light Mode.
$ThemeSettingsRoot = $env:APPDATA
if ([string]::IsNullOrWhiteSpace($ThemeSettingsRoot)) {
    $ThemeSettingsRoot = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
}
$ThemeSettingsDirectory = Join-Path $ThemeSettingsRoot "MSToolkit"
$ThemeSettingsPath = Join-Path $ThemeSettingsDirectory "settings.json"
$script:ThemeMode = "Light"
$script:PreferredDC = $null
$script:OutputEntries = New-Object System.Collections.Generic.List[object]

function Load-ThemePreference {
    $script:ThemeMode = "Light"
    $script:PreferredDC = $null

    try {
        if (Test-Path -LiteralPath $ThemeSettingsPath) {
            $Settings = Get-Content -LiteralPath $ThemeSettingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($Settings.Theme -in @("Light","Dark")) {
                $script:ThemeMode = [string]$Settings.Theme
            }
            if ($Settings.SelectedDC) {
                $script:PreferredDC = [string]$Settings.SelectedDC
            }
        }
    }
    catch {
        # Missing/corrupt settings must never prevent MSToolkit from opening.
        $script:ThemeMode = "Light"
        $script:PreferredDC = $null
    }
}

function Save-ThemePreference {
    try {
        New-Item -ItemType Directory -Path $ThemeSettingsDirectory -Force -ErrorAction Stop | Out-Null

        # Read-modify-write: settings.json also holds the site values from the Settings
        # window and each child tool's Theme_* key. Only Theme and SelectedDC change here.
        $Existing = $null
        if (Test-Path -LiteralPath $ThemeSettingsPath) {
            # If the file is there but cannot be read, leave it alone rather than
            # overwrite everything in it with just these two values.
            $Raw = Get-Content -LiteralPath $ThemeSettingsPath -Raw -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($Raw)) {
                $Existing = $Raw | ConvertFrom-Json -ErrorAction Stop
            }
        }
        if ($null -eq $Existing) { $Existing = [pscustomobject]@{} }

        $Values = [ordered]@{
            Theme      = $script:ThemeMode
            SelectedDC = $script:PreferredDC
        }

        foreach ($Key in @($Values.Keys)) {
            if ($Existing.PSObject.Properties.Name -contains $Key) {
                $Existing.$Key = $Values[$Key]
            }
            else {
                $Existing | Add-Member -MemberType NoteProperty -Name $Key -Value $Values[$Key]
            }
        }

        $Existing |
            ConvertTo-Json -Depth 5 |
            Set-Content -LiteralPath $ThemeSettingsPath -Encoding UTF8 -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Get-ThemePalette {
    if ($script:ThemeMode -eq "Dark") {
        return [pscustomobject]@{
            MainBackground      = [System.Drawing.Color]::FromArgb(30,32,36)
            PanelBackground     = [System.Drawing.Color]::FromArgb(38,41,46)
            InputBackground     = [System.Drawing.Color]::FromArgb(45,48,54)
            OutputBackground    = [System.Drawing.Color]::FromArgb(24,26,29)
            Text                = [System.Drawing.Color]::FromArgb(232,234,237)
            MutedText           = [System.Drawing.Color]::FromArgb(174,180,187)
            Border              = [System.Drawing.Color]::FromArgb(78,84,92)
            Section             = [System.Drawing.Color]::FromArgb(122,181,238)
            TopBar              = [System.Drawing.Color]::FromArgb(24,47,74)
            Accent              = [System.Drawing.Color]::FromArgb(0,120,215)
            ButtonBackground    = [System.Drawing.Color]::FromArgb(49,53,59)
            ButtonHover         = [System.Drawing.Color]::FromArgb(60,65,72)
            SelectionBackground = [System.Drawing.Color]::FromArgb(55,105,155)
            Success             = [System.Drawing.Color]::FromArgb(118,210,142)
            Danger              = [System.Drawing.Color]::FromArgb(255,125,125)
            Warning             = [System.Drawing.Color]::FromArgb(255,184,92)
            Info                = [System.Drawing.Color]::FromArgb(125,190,245)
            Separator           = [System.Drawing.Color]::FromArgb(125,132,140)
            DisabledBackground  = [System.Drawing.Color]::FromArgb(58,61,67)
            DisabledText        = [System.Drawing.Color]::FromArgb(165,170,177)
            DisabledBorder      = [System.Drawing.Color]::FromArgb(92,97,105)
        }
    }

    return [pscustomobject]@{
        MainBackground      = [System.Drawing.Color]::FromArgb(245,247,250)
        PanelBackground     = [System.Drawing.Color]::White
        InputBackground     = [System.Drawing.Color]::White
        OutputBackground    = [System.Drawing.Color]::White
        Text                = [System.Drawing.Color]::FromArgb(35,35,35)
        MutedText           = [System.Drawing.Color]::DimGray
        Border              = [System.Drawing.Color]::FromArgb(210,215,220)
        Section             = [System.Drawing.Color]::FromArgb(31,58,93)
        TopBar              = [System.Drawing.Color]::FromArgb(31,58,93)
        Accent              = [System.Drawing.Color]::FromArgb(0,120,215)
        ButtonBackground    = [System.Drawing.Color]::White
        ButtonHover         = [System.Drawing.Color]::FromArgb(242,246,250)
        SelectionBackground = [System.Drawing.Color]::FromArgb(0,120,215)
        Success             = [System.Drawing.Color]::ForestGreen
        Danger              = [System.Drawing.Color]::Red
        Warning             = [System.Drawing.Color]::DarkOrange
        Info                = [System.Drawing.Color]::FromArgb(35,90,145)
        Separator           = [System.Drawing.Color]::FromArgb(140,140,140)
        DisabledBackground  = [System.Drawing.Color]::FromArgb(238,240,243)
        DisabledText        = [System.Drawing.Color]::FromArgb(125,130,136)
        DisabledBorder      = [System.Drawing.Color]::FromArgb(195,200,206)
    }
}

# ---------------------------------------------------------------------------
# Toolkit settings
#
# Everything site-specific lives in one file, %APPDATA%\MSToolkit\settings.json,
# shared by this console and every child tool. Nothing here is hardcoded to an
# organisation: blank values fall back to what can be read from the machine and
# the directory, so a fresh deployment runs before anything is configured.
# ---------------------------------------------------------------------------
$script:MSToolkitAppName      = "MSToolkit"
$script:MSToolkitSettingsDir  = Join-Path $env:APPDATA $script:MSToolkitAppName
$script:MSToolkitSettingsPath = Join-Path $script:MSToolkitSettingsDir "settings.json"
$script:MSToolkitSettings     = $null

function Get-MSToolkitSettingDefaults {
    # Ordered so the Settings window can be built straight from this list.
    return [ordered]@{
        Domain                = ""   # DNS domain such as contoso.local; blank = read from this computer
        NetBiosName           = ""   # short NetBIOS name such as CONTOSO; blank = %USERDOMAIN%
        EntraSyncServer       = ""   # server running Entra Connect, for Delta Sync
        OUUsers               = ""   # user accounts
        OUComputers           = ""   # workstations
        OUServers             = ""   # servers
        OUDisabled            = ""   # disabled accounts
        OUSecurityGroups      = ""   # security groups
        OUDistributionGroups  = ""   # distribution groups (on-premises or synced from AD)
        OUServiceAccounts     = ""   # service / special function accounts
        OUAdmins              = ""   # admin accounts, listed alongside users in the pickers
        UpnDomain             = ""   # user principal name suffix
        MailDomain            = ""   # primary SMTP domain
        OnMicrosoftDomain     = ""   # tenant.onmicrosoft.com
        OnMicrosoftMailDomain = ""   # tenant.mail.onmicrosoft.com
        CompanyName           = ""   # written to the Company attribute on new accounts
        TenantId              = ""   # Entra tenant, used by the cloud tools
        ClientId              = ""   # app registration, used by the cloud tools
        ExpectedTenantDomain  = ""   # verified domain a cloud tool must see before it will change anything
        SharePointAdminUrl    = ""   # SharePoint admin center address, used by OneDrive / SharePoint
    }
}

function Get-MSToolkitSettings {
    param([switch]$Reload)

    if ($script:MSToolkitSettings -and -not $Reload) {
        return $script:MSToolkitSettings
    }

    $Settings = Get-MSToolkitSettingDefaults

    try {
        if (Test-Path -LiteralPath $script:MSToolkitSettingsPath) {
            $Saved = Get-Content -LiteralPath $script:MSToolkitSettingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            foreach ($Key in @($Settings.Keys)) {
                if ($Saved.PSObject.Properties.Name -contains $Key) {
                    $Settings[$Key] = [string]$Saved.$Key
                }
            }
        }
    }
    catch {
        # A damaged settings file must not stop the toolkit from starting.
    }

    $script:MSToolkitSettings = $Settings
    return $Settings
}

function Save-MSToolkitSettings {
    param([hashtable]$Settings)

    # Read-modify-write: Theme, SelectedDC and the per-tool Theme_* keys live in the
    # same file and are not part of this window.
    try {
        if (-not (Test-Path -LiteralPath $script:MSToolkitSettingsDir)) {
            New-Item -ItemType Directory -Path $script:MSToolkitSettingsDir -Force | Out-Null
        }

        $Existing = $null
        if (Test-Path -LiteralPath $script:MSToolkitSettingsPath) {
            try {
                $Existing = Get-Content -LiteralPath $script:MSToolkitSettingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            }
            catch { $Existing = $null }
        }

        if ($null -eq $Existing) { $Existing = [pscustomobject]@{} }

        foreach ($Key in @($Settings.Keys)) {
            $Value = [string]$Settings[$Key]

            if ($Existing.PSObject.Properties.Name -contains $Key) {
                $Existing.$Key = $Value
            }
            else {
                $Existing | Add-Member -MemberType NoteProperty -Name $Key -Value $Value
            }
        }

        $Existing | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:MSToolkitSettingsPath -Encoding UTF8
        $script:MSToolkitSettings = $Settings
        return $true
    }
    catch {
        return $false
    }
}

function Get-MSToolkitSetting {
    param(
        [string]$Name,
        [string]$Default = ""
    )

    $Value = [string](Get-MSToolkitSettings)[$Name]
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    return $Value.Trim()
}

function Test-MSToolkitConfigured {
    # True once at least one site value has been filled in and saved. The file alone
    # is not enough: the theme toggle, the DC dropdown and the child tools' Theme_*
    # keys also create settings.json, and Clear Settings saves every site key blank.
    # So Settings opens by itself whenever nothing is filled in.
    try {
        if (-not (Test-Path -LiteralPath $script:MSToolkitSettingsPath)) { return $false }

        $Raw = Get-Content -LiteralPath $script:MSToolkitSettingsPath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($Raw)) { return $false }

        $Saved = $Raw | ConvertFrom-Json -ErrorAction Stop
        $Names = @($Saved.PSObject.Properties.Name)

        foreach ($Key in @((Get-MSToolkitSettingDefaults).Keys)) {
            if (($Names -contains $Key) -and -not [string]::IsNullOrWhiteSpace([string]$Saved.$Key)) {
                return $true
            }
        }
    }
    catch {
        # Unreadable file: nothing usable is configured, so open Settings. Saving from
        # there rewrites the file.
        return $false
    }

    return $false
}

function Get-MSToolkitNetBiosName {
    # Configured value first, then the domain this session is authenticating against.
    $Configured = Get-MSToolkitSetting -Name "NetBiosName"
    if ($Configured) { return $Configured }

    if (-not [string]::IsNullOrWhiteSpace($env:USERDOMAIN)) { return $env:USERDOMAIN }
    return ""
}

function Get-MSToolkitDomainPrefix {
    # "CONTOSO\" style prefix for user name boxes; empty when nothing is known.
    $Name = Get-MSToolkitNetBiosName
    if ($Name) { return "$Name\" }
    return ""
}

function Get-MSToolkitDomainRootDN {
    # DC=contoso,DC=local derived from the DNS domain name - no directory call needed.
    $DnsName = Get-MSToolkitSetting -Name "Domain" -Default $Domain
    if ([string]::IsNullOrWhiteSpace($DnsName)) { return "" }
    return ('DC=' + (($DnsName -split '\.') -join ',DC='))
}

function Get-MSToolkitOU {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [switch]$RootFallback
    )

    # A configured OU, or - with -RootFallback - the domain root, so a button still
    # works before anyone has filled the setting in.
    $Configured = Get-MSToolkitSetting -Name $Name
    if ($Configured) { return $Configured }

    if ($RootFallback) { return (Get-MSToolkitDomainRootDN) }
    return ""
}

function Convert-OutputColorForTheme {
    param([System.Drawing.Color]$Color)

    $Palette = Get-ThemePalette
    if ($script:ThemeMode -ne "Dark") {
        return $Color
    }

    $Argb = $Color.ToArgb()

    if ($Argb -in @(
        ([System.Drawing.Color]::Black).ToArgb(),
        ([System.Drawing.Color]::FromArgb(35,35,35)).ToArgb()
    )) { return $Palette.Text }

    if ($Argb -eq ([System.Drawing.Color]::DimGray).ToArgb()) {
        return $Palette.MutedText
    }

    if ($Argb -in @(
        ([System.Drawing.Color]::Red).ToArgb(),
        ([System.Drawing.Color]::Firebrick).ToArgb()
    )) { return $Palette.Danger }

    if ($Argb -in @(
        ([System.Drawing.Color]::ForestGreen).ToArgb(),
        ([System.Drawing.Color]::DarkGreen).ToArgb()
    )) { return $Palette.Success }

    if ($Argb -eq ([System.Drawing.Color]::DarkOrange).ToArgb()) {
        return $Palette.Warning
    }

    if ($Argb -in @(
        ([System.Drawing.Color]::FromArgb(31,58,93)).ToArgb(),
        ([System.Drawing.Color]::FromArgb(35,90,145)).ToArgb()
    )) { return $Palette.Info }

    if ($Argb -eq ([System.Drawing.Color]::FromArgb(140,140,140)).ToArgb()) {
        return $Palette.Separator
    }

    if ($Argb -eq ([System.Drawing.Color]::FromArgb(70,70,70)).ToArgb()) {
        return $Palette.MutedText
    }

    # Catch-all: anything still too dark would be unreadable on the dark output
    # background, so lighten it while keeping its hue. This stops any colour that
    # was picked for Light Mode from disappearing here.
    if ($Color.GetBrightness() -lt 0.45) {
        $Factor = 0.7
        $R = [int][Math]::Round($Color.R + ((255 - $Color.R) * $Factor))
        $G = [int][Math]::Round($Color.G + ((255 - $Color.G) * $Factor))
        $B = [int][Math]::Round($Color.B + ((255 - $Color.B) * $Factor))

        return [System.Drawing.Color]::FromArgb($R,$G,$B)
    }

    return $Color
}

function Add-OutputEntry {
    param([Parameter(Mandatory)]$Entry)

    $script:OutputEntries.Add($Entry)
}

function Render-OutputEntry {
    param([Parameter(Mandatory)]$Entry)

    if (-not $OutputBox) { return }

    $NormalFont = $OutputBox.Font
    $BoldFont = $null

    try {
        switch ($Entry.Type) {
            "Line" {
                $OutputBox.SelectionStart = $OutputBox.TextLength
                $OutputBox.SelectionLength = 0
                $OutputBox.SelectionFont = $NormalFont
                $OutputBox.SelectionColor = Convert-OutputColorForTheme $Entry.Color
                $OutputBox.AppendText("$($Entry.Timestamp) - $($Entry.Text)`r`n")
            }

            "Field" {
                $BoldFont = New-Object System.Drawing.Font(
                    $NormalFont.FontFamily,
                    $NormalFont.Size,
                    [System.Drawing.FontStyle]::Bold
                )

                $OutputBox.SelectionStart = $OutputBox.TextLength
                $OutputBox.SelectionLength = 0
                $OutputBox.SelectionFont = $NormalFont
                $OutputBox.SelectionColor = $OutputBox.ForeColor
                $OutputBox.AppendText("$($Entry.Timestamp) - ")

                $OutputBox.SelectionFont = $BoldFont
                $OutputBox.SelectionColor = Convert-OutputColorForTheme $Entry.LabelColor
                $OutputBox.AppendText("$($Entry.Label): ")

                $OutputBox.SelectionFont = $NormalFont
                $OutputBox.SelectionColor = Convert-OutputColorForTheme $Entry.ValueColor
                $OutputBox.AppendText("$($Entry.Value)`r`n")
            }

            "Separator" {
                $BoldFont = New-Object System.Drawing.Font(
                    $NormalFont.FontFamily,
                    $NormalFont.Size,
                    [System.Drawing.FontStyle]::Bold
                )

                $OutputBox.AppendText("`r`n")
                $OutputBox.SelectionStart = $OutputBox.TextLength
                $OutputBox.SelectionLength = 0
                $OutputBox.SelectionFont = $BoldFont
                $OutputBox.SelectionColor = (Get-ThemePalette).Separator
                $OutputBox.AppendText("============================================================`r`n")
            }

            "StartupSection" {
                $BoldFont = New-Object System.Drawing.Font(
                    $NormalFont.FontFamily,
                    $NormalFont.Size,
                    [System.Drawing.FontStyle]::Bold
                )

                $OutputBox.AppendText("`r`n")
                $OutputBox.SelectionStart = $OutputBox.TextLength
                $OutputBox.SelectionLength = 0
                $OutputBox.SelectionFont = $BoldFont
                $OutputBox.SelectionColor = Convert-OutputColorForTheme $Entry.Color
                $OutputBox.AppendText("============================================================`r`n")
                $OutputBox.AppendText("  $($Entry.Title)`r`n")
                $OutputBox.AppendText("============================================================`r`n")
            }

            "Blank" {
                $OutputBox.AppendText("`r`n")
            }
        }
    }
    finally {
        if ($BoldFont) { $BoldFont.Dispose() }
        $OutputBox.SelectionFont = $NormalFont
        $OutputBox.SelectionColor = $OutputBox.ForeColor
    }
}

function Render-AllOutputEntries {
    if (-not $OutputBox) { return }

    $OutputBox.SuspendLayout()
    try {
        $OutputBox.Clear()
        foreach ($Entry in $script:OutputEntries) {
            Render-OutputEntry -Entry $Entry
        }
        $OutputBox.SelectionStart = $OutputBox.TextLength
        $OutputBox.ScrollToCaret()
    }
    finally {
        $OutputBox.ResumeLayout()
    }
}

function Apply-ThemeToControl {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)

    $Palette = Get-ThemePalette
    $Role = [string]$Control.Tag

    if (
        ($Control -is [System.Windows.Forms.Button] -or
         $Control -is [System.Windows.Forms.CheckBox] -or
         $Control -is [System.Windows.Forms.RadioButton]) -and
        -not $Control.PSObject.Properties['MSToolkitThemeEnabledHooked']
    ) {
        $Control | Add-Member -NotePropertyName MSToolkitThemeEnabledHooked -NotePropertyValue $true
        $Control.Add_EnabledChanged({
            param($sender,$eventArgs)
            Apply-ThemeToControl -Control $sender
        })
    }

    if ($Control -is [System.Windows.Forms.Form]) {
        $Control.BackColor = $Palette.MainBackground
        $Control.ForeColor = $Palette.Text
    }
    elseif ($Control -is [System.Windows.Forms.Panel]) {
        if ($Role -eq "TopBar") {
            $Control.BackColor = $Palette.TopBar
        }
        elseif ($Role -eq "TopDivider") {
            $Control.BackColor = $Palette.Separator
        }
        else {
            $Control.BackColor = $Palette.PanelBackground
            $Control.ForeColor = $Palette.Text
        }
    }
    elseif ($Control -is [System.Windows.Forms.RichTextBox]) {
        $Control.BackColor = $Palette.OutputBackground
        $Control.ForeColor = $Palette.Text
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
    elseif ($Control -is [System.Windows.Forms.DataGridView]) {
        $Control.BackgroundColor = $Palette.PanelBackground
        $Control.GridColor = $Palette.Border
        $Control.DefaultCellStyle.BackColor = $Palette.PanelBackground
        $Control.DefaultCellStyle.ForeColor = $Palette.Text
        $Control.DefaultCellStyle.SelectionBackColor = $Palette.SelectionBackground
        $Control.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
        $Control.AlternatingRowsDefaultCellStyle.BackColor = if ($script:ThemeMode -eq "Dark") {
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
        $Control.EnableHeadersVisualStyles = $false
    }
    elseif ($Control -is [System.Windows.Forms.Button]) {
        $Control.UseVisualStyleBackColor = $false
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $Control.FlatAppearance.BorderSize = 1
        $Control.FlatAppearance.BorderColor = $Palette.Border
        $Control.FlatAppearance.MouseOverBackColor = $Palette.ButtonHover

        switch ($Role) {
            "TopButton" {
                $Control.BackColor = $Palette.Accent
                $Control.ForeColor = [System.Drawing.Color]::White
                $Control.FlatAppearance.BorderSize = 0
                $Control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(20,135,225)
            }
            "ThemeButton" {
                $Control.BackColor = $Palette.Accent
                $Control.ForeColor = [System.Drawing.Color]::White
                $Control.FlatAppearance.BorderSize = 0
                $Control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(20,135,225)
            }
            "SidebarDanger" {
                $Control.BackColor = $Palette.ButtonBackground
                $Control.ForeColor = $Palette.Danger
            }
            "SidebarSuccess" {
                $Control.BackColor = $Palette.ButtonBackground
                $Control.ForeColor = $Palette.Success
            }
            default {
                $Control.BackColor = $Palette.ButtonBackground
                $Control.ForeColor = $Palette.Text
            }
        }
    }
    elseif ($Control -is [System.Windows.Forms.Label]) {
        if ($Role -eq "SectionLabel") {
            $Control.ForeColor = $Palette.Section
        }
        elseif ($Control.Parent -and ([string]$Control.Parent.Tag -eq "TopBar")) {
            $Control.ForeColor = [System.Drawing.Color]::White
        }
        elseif ($Role -eq "MutedLabel") {
            $Control.ForeColor = $Palette.MutedText
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

    # Disabled controls remain clearly visible in both themes while looking
    # intentionally unavailable rather than simply fading into the background.
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
        Apply-ThemeToControl -Control $Child
    }
}

function Apply-CurrentTheme {
    if (-not $form) { return }

    Apply-ThemeToControl -Control $form

    if ($ThemeToggleButton) {
        # The header stays navy in both themes, so the toggle does too.
        $ThemeToggleButton.BackColor = [System.Drawing.Color]::FromArgb(24,47,74)
        $ThemeToggleButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(45,74,110)

        if ($script:ThemeMode -eq "Dark") {
            # Sun: click to go back to Light Mode.
            $ThemeToggleButton.Text = [string][char]0x263C
            $ThemeToggleButton.ForeColor = [System.Drawing.Color]::FromArgb(255,214,102)
        }
        else {
            # Moon: click to go to Dark Mode.
            $ThemeToggleButton.Text = [string][char]0x263E
            $ThemeToggleButton.ForeColor = [System.Drawing.Color]::FromArgb(226,234,245)
        }
    }

    Render-AllOutputEntries
    if ($StatusPanel) {
        Update-StatusStrip
    }
    $form.Invalidate($true)
    $form.Refresh()
}

Load-ThemePreference

function Write-OutputBox {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black
    )

    if ($OutputBox) {
        $Entry = [pscustomobject]@{
            Type      = "Line"
            Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Text      = $Text
            Color     = $Color
        }
        Add-OutputEntry -Entry $Entry
        Render-OutputEntry -Entry $Entry
        $OutputBox.ScrollToCaret()
    }
}

function Write-ResultSeparator {
    if ($script:SuppressResultSeparator) {
        return
    }

    if ($OutputBox) {
        $Entry = [pscustomobject]@{ Type = "Separator" }
        Add-OutputEntry -Entry $Entry
        Render-OutputEntry -Entry $Entry
        $OutputBox.ScrollToCaret()
    }
}

function Write-OutputField {
    param(
        [string]$Label,
        [object]$Value,
        [System.Drawing.Color]$ValueColor = [System.Drawing.Color]::Black,
        [System.Drawing.Color]$LabelColor = [System.Drawing.Color]::Black
    )

    if ($OutputBox) {
        $Entry = [pscustomobject]@{
            Type       = "Field"
            Timestamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Label      = $Label
            Value      = $Value
            ValueColor = $ValueColor
            LabelColor = $LabelColor
        }
        Add-OutputEntry -Entry $Entry
        Render-OutputEntry -Entry $Entry
        $OutputBox.ScrollToCaret()
    }
}

function Write-StartupSection {
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(31,58,93)
    )

    if ($OutputBox) {
        $Entry = [pscustomobject]@{
            Type  = "StartupSection"
            Title = $Title
            Color = $Color
        }
        Add-OutputEntry -Entry $Entry
        Render-OutputEntry -Entry $Entry
        $OutputBox.ScrollToCaret()
    }
}

function Get-SelectedServer {
    if ($DCDropdown -and $DCDropdown.SelectedItem) {
        return "$($DCDropdown.SelectedItem).$Domain"
    }

    return $Domain
}

function Get-InputBox {
    param(
        [string]$Title,
        [string]$Prompt,
        [switch]$Password
    )

    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = $Title
    $inputForm.Size = New-Object System.Drawing.Size(520,210)
    $inputForm.StartPosition = "CenterScreen"
    $inputForm.FormBorderStyle = "FixedDialog"
    $inputForm.MaximizeBox = $false
    $inputForm.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(10,15)
    $label.Size = New-Object System.Drawing.Size(480,55)
    $inputForm.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(10,75)
    $textbox.Size = New-Object System.Drawing.Size(480,22)
    if ($Password) {
        $textbox.UseSystemPasswordChar = $true
    }
    $inputForm.Controls.Add($textbox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "OK"
    $button.Location = New-Object System.Drawing.Point(390,115)
    $button.Size = New-Object System.Drawing.Size(100,30)
    $button.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $button.Add_Click({
        $inputForm.Tag = $textbox.Text
        $inputForm.Close()
    })

    $inputForm.Controls.Add($button)
    $inputForm.AcceptButton = $button

    $inputForm.Add_Shown({
        $textbox.Focus()
    })

    Apply-ThemeToControl -Control $inputForm
    $inputForm.ShowDialog() | Out-Null

    return $inputForm.Tag
}

# ---------------------------------------------------------------------------
# MSToolkit safety protections for critical directory objects.
# These are intentionally hard blocks inside MSToolkit. Administrative work that
# legitimately needs to change these objects should be performed with an
# approved native/admin process rather than bypassed from this utility.
# ---------------------------------------------------------------------------
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

function Get-MSToolkitCriticalUserReason {
    param($User)

    $Sid = Get-MSToolkitSidString -DirectoryObject $User
    $Rid = Get-MSToolkitRidFromSid -Sid $Sid

    switch ($Rid) {
        500 { return "Built-in Administrator account (RID 500)" }
        501 { return "Built-in Guest account (RID 501)" }
        502 { return "KRBTGT account (RID 502)" }
    }

    if ($User.isCriticalSystemObject -eq $true) {
        return "Active Directory marks this user as a critical system object"
    }

    return $null
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

function Get-MSToolkitCriticalComputerReason {
    param($Computer)

    $UserAccountControl = 0
    $PrimaryGroupID = 0

    try { $UserAccountControl = [int64]$Computer.userAccountControl } catch { }
    try { $PrimaryGroupID = [int]$Computer.PrimaryGroupID } catch { }

    # SERVER_TRUST_ACCOUNT (0x2000 / 8192) identifies a Domain Controller account.
    if ((($UserAccountControl -band 8192) -ne 0) -or ($PrimaryGroupID -eq 516)) {
        return 'Domain Controller computer account'
    }

    if ($Computer.isCriticalSystemObject -eq $true) {
        return "Active Directory marks this computer as a critical system object"
    }

    return $null
}

function Stop-MSToolkitCriticalOperation {
    param(
        [string]$ObjectType,
        [string]$DisplayName,
        [string]$Operation,
        [string]$Reason
    )

    if ([string]::IsNullOrWhiteSpace($Reason)) {
        return $false
    }

    $Message = @"
OPERATION BLOCKED

$DisplayName is protected by MSToolkit.

Object type: $ObjectType
Requested operation: $Operation
Reason: $Reason

MSToolkit does not permit this operation on critical Active Directory objects. Use ADUC or another approved administrative process if the change is intentionally required.
"@

    Write-OutputBox "BLOCKED: $Operation on $DisplayName. $Reason" ([System.Drawing.Color]::Red)

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "MSToolkit Critical Object Protection",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Stop
    ) | Out-Null

    return $true
}

function Request-MSToolkitAccidentalDeletionOverride {
    param(
        [string]$ObjectType,
        [string]$DisplayName,
        [string]$DistinguishedName,
        [string]$Server
    )

    $Message = @"
$ObjectType '$DisplayName' is currently protected from accidental deletion.

Yes = remove the Protect object from accidental deletion setting and continue with the deletion.
No = cancel the deletion and leave the protection unchanged.

This override applies only to this requested deletion. Critical objects blocked by MSToolkit cannot be overridden here.
"@

    $Choice = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "Protected from Accidental Deletion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($Choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-OutputBox "Deletion cancelled. Accidental-deletion protection was left enabled for $DisplayName."
        return $false
    }

    try {
        Set-ADObject `
            -Identity $DistinguishedName `
            -ProtectedFromAccidentalDeletion $false `
            -Server $Server `
            -ErrorAction Stop

        Write-OutputBox "Removed accidental-deletion protection from $DisplayName so the requested deletion can continue." ([System.Drawing.Color]::DarkOrange)
        return $true
    }
    catch {
        Write-OutputBox "ERROR removing accidental-deletion protection from ${DisplayName}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
        [System.Windows.Forms.MessageBox]::Show(
            "MSToolkit could not remove accidental-deletion protection from '$DisplayName'.`r`n`r`n$($_.Exception.Message)`r`n`r`nThe object was not deleted.",
            "Unable to Remove Protection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}

function Restore-MSToolkitAccidentalDeletionProtection {
    param(
        [string]$DisplayName,
        [string]$DistinguishedName,
        [string]$Server
    )

    try {
        Set-ADObject `
            -Identity $DistinguishedName `
            -ProtectedFromAccidentalDeletion $true `
            -Server $Server `
            -ErrorAction Stop

        Write-OutputBox "Deletion failed; accidental-deletion protection was restored for $DisplayName." ([System.Drawing.Color]::DarkOrange)
    }
    catch {
        Write-OutputBox "WARNING: Deletion failed and MSToolkit could not restore accidental-deletion protection for ${DisplayName}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

# Keep external windows launched by MSToolkit on the same monitor as the main form.
# A short-lived timer watches the launched process because many child tools hide their
# PowerShell console before creating their actual WinForms window.
$script:WindowPlacementTimers = New-Object System.Collections.Generic.List[object]

function Move-ProcessWindowsToMSToolkitMonitor {
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Process]$Process
    )

    if (-not $Process -or -not $form -or $form.IsDisposed) {
        return
    }

    try {
        $TargetArea = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
    }
    catch {
        return
    }

    $State = [pscustomobject]@{
        ProcessId       = $Process.Id
        TargetLeft      = $TargetArea.Left
        TargetTop       = $TargetArea.Top
        TargetWidth     = $TargetArea.Width
        TargetHeight    = $TargetArea.Height
        LastMovedHandle = [IntPtr]::Zero
        Deadline         = (Get-Date).AddSeconds(15)
        Timer            = $null
        OwnerList        = $script:WindowPlacementTimers
    }

    $Timer = New-Object System.Windows.Forms.Timer
    $Timer.Interval = 200
    $State.Timer = $Timer
    $script:WindowPlacementTimers.Add($Timer)

    $TickHandler = {
        $StopTimer = $false

        try {
            if ((Get-Date) -ge $State.Deadline) {
                $StopTimer = $true
            }
            else {
                $RunningProcess = Get-Process -Id $State.ProcessId -ErrorAction Stop
                $WindowHandle = [WindowHelper]::FindLargestVisibleWindowForProcess($State.ProcessId)

                if (
                    $WindowHandle -ne [IntPtr]::Zero -and
                    $WindowHandle -ne $State.LastMovedHandle
                ) {
                    [void][WindowHelper]::CenterWindowInWorkingArea(
                        $WindowHandle,
                        $State.TargetLeft,
                        $State.TargetTop,
                        $State.TargetWidth,
                        $State.TargetHeight
                    )

                    $State.LastMovedHandle = $WindowHandle
                }
            }
        }
        catch {
            # If the process exited, there is nothing left to position.
            $StopTimer = $true
        }

        if ($StopTimer) {
            $State.Timer.Stop()
            $State.Timer.Dispose()
            [void]$State.OwnerList.Remove($State.Timer)
        }
    }.GetNewClosure()

    $Timer.Add_Tick($TickHandler)
    $Timer.Start()
}

function Start-MSToolkitADACWatchdog {
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Process]$Process
    )

    # ADAC sometimes stalls on its splash screen while it waits on a domain controller
    # whose ADWS port (9389) never answers, then crashes 40+ seconds later with no way
    # to close it in the meantime. A healthy launch connects in milliseconds, so a
    # connection attempt to 9389 still pending after several seconds means it is stuck.
    # This watches for exactly that, closes ADAC, and says which DC it was waiting on.

    if (-not $Process) { return }

    $State = [pscustomobject]@{
        ProcessId   = $Process.Id
        Started     = Get-Date
        FirstSeen   = @{}
        Timer       = $null
        OwnerList   = $script:WindowPlacementTimers
        # Passed in explicitly: GetNewClosure() runs the tick handler in its own scope,
        # where script-level functions are not guaranteed to resolve. A function's
        # scriptblock keeps its original scope, so everything it calls still works.
        WriteOutput = ${function:Write-OutputBox}
        # Set once the watchdog decides to act. The UAC prompt keeps the message loop
        # running, so without this each timer tick would launch another close attempt.
        Handling    = $false
    }

    $Timer = New-Object System.Windows.Forms.Timer
    $Timer.Interval = 1000
    $State.Timer = $Timer
    $script:WindowPlacementTimers.Add($Timer)

    $TickHandler = {
        if ($State.Handling) { return }

        $StopTimer = $false

        try {
            $Elapsed = ((Get-Date) - $State.Started).TotalSeconds

            # Stop watching once ADAC has exited, or once it has clearly loaded.
            $null = Get-Process -Id $State.ProcessId -ErrorAction Stop

            # ADAC has loaded once its main window is showing. The splash screen is a
            # small window; the main console is much larger. A loaded ADAC is never
            # touched, even if it opens another connection attempt later.
            $Loaded = $false
            $Handle = [WindowHelper]::FindLargestVisibleWindowForProcess($State.ProcessId)
            if ($Handle -ne [IntPtr]::Zero) {
                $Rect = New-Object 'WindowHelper+RECT'
                if ([WindowHelper]::GetWindowRect($Handle, [ref]$Rect)) {
                    $Width = $Rect.Right - $Rect.Left
                    $Height = $Rect.Bottom - $Rect.Top
                    if (($Width -ge 700) -and ($Height -ge 450)) {
                        $Loaded = $true
                    }
                }
            }

            if ($Loaded -or ($Elapsed -ge 60)) {
                $StopTimer = $true
            }
            else {
                $Pending = @(
                    Get-NetTCPConnection -OwningProcess $State.ProcessId -State SynSent -RemotePort 9389 -ErrorAction SilentlyContinue
                )

                $Now = Get-Date
                $Stuck = $null

                foreach ($Attempt in $Pending) {
                    $Key = "$($Attempt.RemoteAddress):$($Attempt.LocalPort)"

                    if (-not $State.FirstSeen.ContainsKey($Key)) {
                        $State.FirstSeen[$Key] = $Now
                    }
                    elseif ((($Now - $State.FirstSeen[$Key]).TotalSeconds -ge 5) -and ($Elapsed -ge 20)) {
                        $Stuck = $Attempt
                    }
                }

                if ($Stuck) {
                    # Act exactly once: stop the timer before anything that can block.
                    $State.Handling = $true
                    $State.Timer.Stop()

                    $Address = [string]$Stuck.RemoteAddress
                    $Name = $Address

                    try {
                        $Name = [System.Net.Dns]::GetHostEntry($Address).HostName
                    }
                    catch { }

                    # Try a normal close first, then confirm it actually worked.
                    $Closed = $false
                    try {
                        Stop-Process -Id $State.ProcessId -Force -ErrorAction Stop
                    }
                    catch { }

                    Start-Sleep -Milliseconds 500
                    if (-not (Get-Process -Id $State.ProcessId -ErrorAction SilentlyContinue)) {
                        $Closed = $true
                    }

                    # ADAC can run elevated while MSToolkit does not, and a non-elevated
                    # process is not allowed to end an elevated one. Fall back to an
                    # elevated taskkill, which may show a UAC prompt.
                    if (-not $Closed) {
                        try {
                            Start-Process -FilePath "taskkill.exe" `
                                -ArgumentList "/PID $($State.ProcessId) /F" `
                                -Verb RunAs `
                                -WindowStyle Hidden `
                                -Wait `
                                -ErrorAction Stop
                        }
                        catch { }

                        Start-Sleep -Milliseconds 500
                        if (-not (Get-Process -Id $State.ProcessId -ErrorAction SilentlyContinue)) {
                            $Closed = $true
                        }
                    }

                    if ($Closed) {
                        & $State.WriteOutput "ADAC was stuck waiting on $Name ($Address) - its ADWS port 9389 never answered - so it was closed. Launch ADAC again; it usually picks a reachable DC." ([System.Drawing.Color]::DarkOrange)
                    }
                    else {
                        & $State.WriteOutput "ADAC is stuck waiting on $Name ($Address) - its ADWS port 9389 never answered - but MSToolkit could not close it. End dsac.exe in Task Manager (Details tab), then launch ADAC again." ([System.Drawing.Color]::Red)
                    }

                    $StopTimer = $true
                }
            }
        }
        catch {
            # ADAC has already exited; nothing left to watch.
            $StopTimer = $true
        }

        if ($StopTimer) {
            $State.Timer.Stop()
            $State.Timer.Dispose()
            [void]$State.OwnerList.Remove($State.Timer)
        }
    }.GetNewClosure()

    $Timer.Add_Tick($TickHandler)
    $Timer.Start()
}

function Launch-ADAC {
    try {
        $DsacExe = Join-Path $env:SystemRoot "System32\dsac.exe"

        if (-not (Test-Path $DsacExe)) {
            Write-OutputBox "ERROR: Active Directory Administrative Center (dsac.exe) is not installed on this computer." ([System.Drawing.Color]::Red)
            return
        }

        # dsac.exe has no server parameter. It finds its own DC through ADWS discovery,
        # so it does not follow the AD Server dropdown.
        Write-OutputBox "Launching Active Directory Administrative Center. ADAC picks its own domain controller and does not follow the AD Server selection."
        $Process = Start-Process $DsacExe -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
        Start-MSToolkitADACWatchdog -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Active Directory Administrative Center: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Launch-ADUC {
    try {
        $Server = Get-SelectedServer

        Write-OutputBox "Opening Active Directory Users and Computers using $Server..."
        $Process = Start-Process "mmc.exe" `
            -ArgumentList "`"$env:SystemRoot\System32\dsa.msc`" /server=$Server" `
            -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Active Directory Users and Computers: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Launch-DNSManager {
    try {
        $DnsMsc = Join-Path $env:SystemRoot "System32\dnsmgmt.msc"

        if (-not (Test-Path $DnsMsc)) {
            Write-OutputBox "ERROR: DNS Manager (dnsmgmt.msc) is not installed on this computer. It comes with the Rsat.Dns.Tools capability." ([System.Drawing.Color]::Red)
            return
        }

        Write-OutputBox "Launching DNS Manager..."
        $Process = Start-Process "mmc.exe" -ArgumentList "`"$DnsMsc`"" -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching DNS Manager: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Launch-ComputerManagement {
    try {
        $Server = Get-SelectedServer

        if ([string]::IsNullOrWhiteSpace($Server)) {
            Write-OutputBox "ERROR: No DC selected." ([System.Drawing.Color]::Red)
            return
        }

        Write-OutputBox "Opening Computer Management on $Server..."
        $Process = Start-Process "mmc.exe" `
            -ArgumentList "`"$env:SystemRoot\System32\compmgmt.msc`" /computer=$Server" `
            -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Computer Management: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Launch-GPManagement {
    try {
        $GpmcMsc = Join-Path $env:SystemRoot "System32\gpmc.msc"

        if (-not (Test-Path $GpmcMsc)) {
            Write-OutputBox "ERROR: Group Policy Management (gpmc.msc) is not installed on this computer." ([System.Drawing.Color]::Red)
            return
        }

        Write-OutputBox "Launching Group Policy Management..."
        $Process = Start-Process "mmc.exe" -ArgumentList "`"$GpmcMsc`"" -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Group Policy Management: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Write-MSToolkitReportResult {
    param(
        [string]$ReportName,
        [string]$FilePath,
        [int]$RowCount,
        [string]$Server
    )

    Write-ResultSeparator
    Write-OutputBox "$ReportName complete." ([System.Drawing.Color]::Green)

    Write-OutputField -Label "Rows" -Value ([string]$RowCount)
    Write-OutputField -Label "AD server" -Value $Server
    Write-OutputField -Label "File" -Value (Split-Path -Leaf $FilePath)

    # Copy the folder, not the file, so it can be pasted straight into Explorer.
    $Folder = Split-Path -Parent $FilePath
    $Copied = $false

    try {
        Set-Clipboard -Value $Folder -ErrorAction Stop
        $Copied = $true
    }
    catch {
        try {
            [System.Windows.Forms.Clipboard]::SetText($Folder)
            $Copied = $true
        }
        catch { }
    }

    if ($Copied) {
        Write-OutputField -Label "Logs location copied to clipboard" -Value $Folder -LabelColor ([System.Drawing.Color]::Green) -ValueColor ([System.Drawing.Color]::Green)
        Write-OutputBox "Paste it into an Explorer address bar or the Run box. MSToolkit runs as a different account, so it cannot open an Explorer window in your session." ([System.Drawing.Color]::DimGray)
    }
    else {
        Write-OutputField -Label "Logs location" -Value $Folder
        Write-OutputBox "The path could not be copied to the clipboard - copy it from the line above." ([System.Drawing.Color]::DarkOrange)
    }

    Write-OutputField -Label "Full path" -Value $FilePath -ValueColor ([System.Drawing.Color]::DimGray)
    Write-OUBlankLine
}

function Open-MSToolkitFolderAsSignedInUser {
    param([string]$Path)

    # Explorer hands folder requests to the shell already running on the desktop,
    # which belongs to the signed-in Windows user and refuses requests from the admin
    # account MSToolkit runs as. Launching Explorer as that same user lets it through.
    return (Invoke-MSToolkitAsSignedInUser `
        -Purpose "Open Logs Folder" `
        -HeaderText "Open Folder" `
        -WarningText "Sign in as the account you are signed in to Windows with." `
        -ExplanationText "Explorer can only open windows on your desktop for the account that owns it, not the Domain Admin account MSToolkit runs as." `
        -Launch {
            param($Credential)

            Start-Process -FilePath (Join-Path $env:SystemRoot "explorer.exe") `
                -ArgumentList "`"$Path`"" `
                -Credential $Credential `
                -WorkingDirectory $env:SystemRoot `
                -ErrorAction Stop

            Write-OutputBox "Opened $Path in Explorer as $($Credential.UserName)."
        })
}

function Open-MSToolkitFileAsSignedInUser {
    param([string]$FilePath)

    # Opens a file in its default app (Excel for a CSV) as the signed-in Windows user,
    # where Office is licensed and signed in. "start" runs inside that user's own
    # process, so it does not depend on handing off to the desktop shell.
    $FileName = Split-Path -Leaf $FilePath

    return (Invoke-MSToolkitAsSignedInUser `
        -Purpose "Open $FileName" `
        -HeaderText "Open File" `
        -WarningText "Sign in as the account you are signed in to Windows with." `
        -ExplanationText "The file opens in its default app under your own account, where Excel and Office are licensed - not the Domain Admin account MSToolkit runs as." `
        -Launch {
            param($Credential)

            Start-Process -FilePath (Join-Path $env:SystemRoot "System32\cmd.exe") `
                -ArgumentList "/c start `"`" `"$FilePath`"" `
                -Credential $Credential `
                -WorkingDirectory $env:SystemRoot `
                -WindowStyle Hidden `
                -LoadUserProfile `
                -ErrorAction Stop

            Write-OutputBox "Opened $FilePath as $($Credential.UserName)."
        })
}

function Remove-MSToolkitFileAsSignedInUser {
    param([string]$FilePath)

    # Used when MSToolkit' own account is refused - typically when running from a copy
    # inside the signed-in user's profile or OneDrive folder. The launch step returns
    # normally once the sign-in works, so a delete that fails is never mistaken for a
    # wrong password; whether the file actually went is checked separately below.
    $FileName = Split-Path -Leaf $FilePath

    $SignedIn = Invoke-MSToolkitAsSignedInUser `
        -Purpose "Delete $FileName" `
        -HeaderText "Delete File" `
        -WarningText "Sign in as the account you are signed in to Windows with." `
        -ExplanationText "MSToolkit's own account is not allowed to delete this file, so it is deleted under your Windows account instead." `
        -Launch {
            param($Credential)

            # No -Wait: in Windows PowerShell 5.1, -Wait with -Credential can fail with
            # "Access is denied" because it tracks the new process through a job object.
            $null = Start-Process -FilePath (Join-Path $env:SystemRoot "System32\cmd.exe") `
                -ArgumentList "/c del /f /q `"$FilePath`"" `
                -Credential $Credential `
                -WorkingDirectory $env:SystemRoot `
                -WindowStyle Hidden `
                -ErrorAction Stop
        }

    if (-not $SignedIn) {
        return "NotDone"
    }

    # The delete runs in its own process; give it up to 10 seconds to finish.
    for ($i = 0; $i -lt 20; $i++) {
        if (-not (Test-Path -LiteralPath $FilePath)) {
            return "Deleted"
        }
        Start-Sleep -Milliseconds 500
    }

    return "Failed"
}

function Remove-MSToolkitLauncherUsername {
    # The runas launchers run as the person signed in to Windows and remember the admin
    # username in THAT account's %APPDATA%\MSToolkit\launcher-admin-username.txt. This
    # console runs as the admin account, which cannot reach that profile, so the file is
    # removed by a short hidden process started under the Windows account - the same
    # sign-in the Logs and file buttons use.
    # Returns Done, Requested (handed to the Windows account; cannot be checked from
    # here), or NotDone (sign-in cancelled).
    $FileName = "launcher-admin-username.txt"

    # Same account (MSToolkit not started through runas): remove it directly.
    $OwnFile = Join-Path $script:MSToolkitSettingsDir $FileName
    if (Test-Path -LiteralPath $OwnFile) {
        Remove-Item -LiteralPath $OwnFile -Force -ErrorAction SilentlyContinue
    }

    $CurrentAccount = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $WindowsAccount = Get-MSToolkitSignedInWindowsUser

    if ((-not $WindowsAccount) -or (Test-MSToolkitSameAccount -First $CurrentAccount -Second $WindowsAccount)) {
        if (Test-Path -LiteralPath $OwnFile) { return "NotDone" }
        return "Done"
    }

    # GetFolderPath resolves the Windows account's own AppData, including a redirected one.
    $RemoveCommand = "Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'MSToolkit\$FileName') -Force -ErrorAction SilentlyContinue"
    $EncodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($RemoveCommand))

    $SignedIn = Invoke-MSToolkitAsSignedInUser `
        -Purpose "Forget launcher username" `
        -HeaderText "Clear Settings" `
        -WarningText "Sign in as the account you are signed in to Windows with." `
        -ExplanationText "The launchers remember the admin username in your Windows account's profile, which MSToolkit's admin account cannot reach." `
        -Launch {
            param($Credential)

            # No -Wait: in Windows PowerShell 5.1, -Wait with -Credential can fail with
            # "Access is denied" because it tracks the new process through a job object.
            $null = Start-Process -FilePath (Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe") `
                -ArgumentList "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $EncodedCommand" `
                -Credential $Credential `
                -WorkingDirectory $env:SystemRoot `
                -WindowStyle Hidden `
                -ErrorAction Stop
        }

    if ($SignedIn) { return "Requested" }
    return "NotDone"
}

function Show-MSToolkitOUBrowser {
    param(
        [string]$CurrentDN,
        [System.Windows.Forms.Form]$Owner
    )

    # Reuses the cascading picker from Create OU / Move Object to OU. Returns the
    # chosen DN, or $null if cancelled.
    try {
        $Server = Get-SelectedServer
        $DomainRoot = (Get-ADDomain -Server $Server -ErrorAction Stop).DistinguishedName
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not read the domain root: $($_.Exception.Message)",
            "Select OU", "OK", "Error") | Out-Null
        return $null
    }

    $Palette = Get-ThemePalette

    $PickForm = New-Object System.Windows.Forms.Form
    $PickForm.Text = "Select an organizational unit"
    $PickForm.Size = New-Object System.Drawing.Size(720,330)
    $PickForm.StartPosition = "CenterParent"
    $PickForm.FormBorderStyle = "FixedDialog"
    $PickForm.MaximizeBox = $false
    $PickForm.MinimizeBox = $false

    $Flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $Flow.Location = New-Object System.Drawing.Point(16,16)
    $Flow.Size = New-Object System.Drawing.Size(676,150)
    $Flow.AutoScroll = $true
    $PickForm.Controls.Add($Flow)

    $PathLabel = New-Object System.Windows.Forms.Label
    $PathLabel.Location = New-Object System.Drawing.Point(16,176)
    $PathLabel.Size = New-Object System.Drawing.Size(676,40)
    $PickForm.Controls.Add($PathLabel)

    $HintLabel = New-Object System.Windows.Forms.Label
    $HintLabel.Text = "Leave at the domain root to search the whole domain."
    $HintLabel.Location = New-Object System.Drawing.Point(16,218)
    $HintLabel.Size = New-Object System.Drawing.Size(420,20)
    $HintLabel.ForeColor = $Palette.MutedText
    $PickForm.Controls.Add($HintLabel)

    $UseRoot = New-Object System.Windows.Forms.Button
    $UseRoot.Text = "Clear"
    $UseRoot.Location = New-Object System.Drawing.Point(16,250)
    $UseRoot.Size = New-Object System.Drawing.Size(90,30)
    $UseRoot.Add_Click({ $script:OUPickerCurrentDN = ""; $PickForm.DialogResult = [System.Windows.Forms.DialogResult]::OK })
    $PickForm.Controls.Add($UseRoot)

    $OkButton = New-Object System.Windows.Forms.Button
    $OkButton.Text = "Select"
    $OkButton.Location = New-Object System.Drawing.Point(490,250)
    $OkButton.Size = New-Object System.Drawing.Size(100,30)
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $PickForm.Controls.Add($OkButton)
    $PickForm.AcceptButton = $OkButton

    $CancelBtn = New-Object System.Windows.Forms.Button
    $CancelBtn.Text = "Cancel"
    $CancelBtn.Location = New-Object System.Drawing.Point(598,250)
    $CancelBtn.Size = New-Object System.Drawing.Size(94,30)
    $CancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $PickForm.Controls.Add($CancelBtn)
    $PickForm.CancelButton = $CancelBtn

    Initialize-MSToolkitOUPicker -Flow $Flow -PathLabel $PathLabel -Server $Server -DomainRoot $DomainRoot
    Apply-ThemeToControl -Control $PickForm

    if ($PickForm.ShowDialog($Owner) -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return [string]$script:OUPickerCurrentDN
}

function Show-MSToolkitSettings {
    $Settings = Get-MSToolkitSettings -Reload
    $Palette = Get-ThemePalette

    $SetForm = New-Object System.Windows.Forms.Form
    $SetForm.Text = "MSToolkit Settings"
    $SetForm.Size = New-Object System.Drawing.Size(860,760)
    $SetForm.StartPosition = "CenterScreen"
    $SetForm.FormBorderStyle = "Sizable"
    $SetForm.MinimumSize = New-Object System.Drawing.Size(820,640)

    $HeaderPanel = New-Object System.Windows.Forms.Panel
    $HeaderPanel.Dock = "Top"
    $HeaderPanel.Height = 62
    $HeaderPanel.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $HeaderPanel.Tag = "TopBar"
    $SetForm.Controls.Add($HeaderPanel)

    $HeaderTitle = New-Object System.Windows.Forms.Label
    $HeaderTitle.Text = "Settings"
    $HeaderTitle.AutoSize = $true
    $HeaderTitle.ForeColor = [System.Drawing.Color]::White
    $HeaderTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",14)
    $HeaderTitle.Location = New-Object System.Drawing.Point(18,10)
    $HeaderPanel.Controls.Add($HeaderTitle)

    $HeaderSub = New-Object System.Windows.Forms.Label
    $HeaderSub.Text = "Shared by this console and every tool it launches. Blank values fall back to the domain root or this computer."
    $HeaderSub.AutoSize = $true
    $HeaderSub.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
    $HeaderSub.Font = New-Object System.Drawing.Font("Segoe UI",9)
    $HeaderSub.Location = New-Object System.Drawing.Point(20,38)
    $HeaderPanel.Controls.Add($HeaderSub)

    $Scroll = New-Object System.Windows.Forms.Panel
    $Scroll.Location = New-Object System.Drawing.Point(0,62)
    $Scroll.Size = New-Object System.Drawing.Size($SetForm.ClientSize.Width,($SetForm.ClientSize.Height - 62 - 56))
    $Scroll.Anchor = "Top,Bottom,Left,Right"
    $Scroll.AutoScroll = $true
    $SetForm.Controls.Add($Scroll)

    $Boxes = @{}
    $Y = 14

    function Add-SettingsSection {
        param([string]$Text)

        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = $Text
        $Label.Location = New-Object System.Drawing.Point(18,$script:SettingsY)
        $Label.Size = New-Object System.Drawing.Size(400,22)
        $Label.Font = New-Object System.Drawing.Font("Segoe UI Semibold",10)
        $Label.ForeColor = [System.Drawing.Color]::FromArgb(0,120,215)
        $script:SettingsScroll.Controls.Add($Label)
        $script:SettingsY += 28
    }

    function Add-SettingsField {
        param(
            [string]$Key,
            [string]$Caption,
            [string]$Hint = "",
            [switch]$Browse
        )

        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = $Caption
        $Label.Location = New-Object System.Drawing.Point(28,($script:SettingsY + 4))
        $Label.Size = New-Object System.Drawing.Size(210,20)
        $script:SettingsScroll.Controls.Add($Label)

        $BoxWidth = 556
        if ($Browse) { $BoxWidth = 468 }

        $Box = New-Object System.Windows.Forms.TextBox
        $Box.Location = New-Object System.Drawing.Point(244,$script:SettingsY)
        $Box.Size = New-Object System.Drawing.Size($BoxWidth,24)
        $Box.Text = [string]$script:SettingsValues[$Key]
        $Box.Anchor = "Top,Left,Right"
        $script:SettingsScroll.Controls.Add($Box)
        $script:SettingsBoxes[$Key] = $Box

        if ($Browse) {
            $BrowseButton = New-Object System.Windows.Forms.Button
            $BrowseButton.Text = "Browse..."
            $BrowseButton.Location = New-Object System.Drawing.Point(720,($script:SettingsY - 1))
            $BrowseButton.Size = New-Object System.Drawing.Size(92,26)
            $BrowseButton.Anchor = "Top,Right"
            $BrowseButton.Tag = $Key
            $BrowseButton.Add_Click({
                $TargetKey = [string]$this.Tag
                $Chosen = Show-MSToolkitOUBrowser -CurrentDN $script:SettingsBoxes[$TargetKey].Text -Owner $script:SettingsForm
                if ($null -ne $Chosen) {
                    $script:SettingsBoxes[$TargetKey].Text = $Chosen
                }
            })
            $script:SettingsScroll.Controls.Add($BrowseButton)
        }

        $script:SettingsY += 28

        if ($Hint) {
            $HintLabel = New-Object System.Windows.Forms.Label
            $HintLabel.Text = $Hint
            $HintLabel.Location = New-Object System.Drawing.Point(246,$script:SettingsY)
            $HintLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
            # Wrap long hints at the text box width instead of cutting them off.
            $HintLabel.MaximumSize = New-Object System.Drawing.Size(560,0)
            $HintLabel.AutoSize = $true
            $HintLabel.ForeColor = $script:SettingsPalette.MutedText
            $HintLabel.Tag = "SettingsHint"
            $script:SettingsScroll.Controls.Add($HintLabel)
            $script:SettingsY += ($HintLabel.GetPreferredSize((New-Object System.Drawing.Size(560,0))).Height + 2)
        }

        $script:SettingsY += 6
    }

    $script:SettingsScroll  = $Scroll
    $script:SettingsBoxes   = $Boxes
    $script:SettingsValues  = $Settings
    $script:SettingsY       = $Y
    $script:SettingsForm    = $SetForm
    $script:SettingsPalette = $Palette

    Add-SettingsSection "Directory"
    Add-SettingsField -Key "Domain"          -Caption "DNS domain:"          -Hint "DNS name of the Active Directory domain, for example contoso.local. Builds the domain root (DC=contoso,DC=local) used when an OU below is blank, and the Managed Service Accounts container. Blank uses the domain this computer is joined to."
    Add-SettingsField -Key "NetBiosName"     -Caption "Short domain name:"   -Hint "Pre-Windows 2000 (NetBIOS) domain name, for example CONTOSO. Not currently used: sign-in boxes take the user name exactly as typed, with no domain added."
    Add-SettingsField -Key "EntraSyncServer" -Caption "Entra Connect server:" -Hint "Name of the server running Microsoft Entra Connect Sync. Delta Sync runs Start-ADSyncSyncCycle -PolicyType Delta there over PowerShell remoting. Blank disables Delta Sync."

    Add-SettingsSection "Organizational units"
    Add-SettingsField -Key "OUUsers"           -Caption "Users:"             -Browse -Hint "Distinguished name of the OU that holds user accounts (OU=...,DC=...). Get Employee OUs lists the OUs directly under it; Create New AD User, Compare User Groups and Investigate Account Lockout start from it. Blank searches from the domain root."
    Add-SettingsField -Key "OUAdmins"          -Caption "Admin accounts:"    -Browse -Hint "Distinguished name of the OU that holds admin accounts, when they are kept apart from users. Compare User Groups and Investigate Account Lockout list them alongside users. Leave blank if admin accounts sit inside the Users OU."
    Add-SettingsField -Key "OUComputers"       -Caption "Computers:"         -Browse -Hint "Distinguished name of the OU that holds workstations. Get Computer OUs lists every OU under it, at all levels. Blank lists only the top-level OUs of the domain root, like the other OU buttons."
    Add-SettingsField -Key "OUServers"         -Caption "Servers:"           -Browse -Hint "Distinguished name of the OU that holds member servers. Get Server OUs lists the OUs directly under it. Blank searches from the domain root."
    Add-SettingsField -Key "OUDisabled"        -Caption "Disabled accounts:" -Browse -Hint "Distinguished name of the OU where disabled accounts are moved. Get Disabled OUs lists the OUs directly under it. Blank searches from the domain root."
    Add-SettingsField -Key "OUSecurityGroups"  -Caption "Security groups:"   -Browse -Hint "Distinguished name of the OU that holds security groups. Get Security Groups lists every security group under it, including sub-OUs. Blank lists every security group in the domain, labelled as the domain root."
    Add-SettingsField -Key "OUDistributionGroups" -Caption "Distribution groups:" -Browse -Hint "Optional. Distinguished name of the OU that holds distribution groups created in AD (on-premises Exchange, or synced to Microsoft 365). Get Distribution Groups lists every distribution group under it. Blank lists every distribution group in the domain, labelled as the domain root. Groups that exist only in Exchange Online are not in AD - use M365 Distro Compare/Add for those."
    Add-SettingsField -Key "OUServiceAccounts" -Caption "Service accounts:"  -Browse -Hint "Distinguished name of the OU that holds service accounts. Get Special Function Accounts lists the user accounts directly in it, not in sub-OUs. Required for that button; it asks for this when blank."

    Add-SettingsSection "New account naming"
    Add-SettingsField -Key "UpnDomain"             -Caption "UPN domain:"             -Hint "Domain part of the userPrincipalName on new accounts (username@contoso.com); also added as a secondary smtp: proxy address. Must be a UPN suffix in Active Directory Domains and Trusts and a verified domain in Microsoft 365."
    Add-SettingsField -Key "MailDomain"            -Caption "Primary SMTP domain:"    -Hint "Domain part of the mail attribute and the primary SMTP: proxy address on new accounts, for example contoso.com. Must be an accepted domain in Exchange Online."
    Add-SettingsField -Key "OnMicrosoftDomain"     -Caption "onmicrosoft domain:"     -Hint "The tenant's initial domain, for example contoso.onmicrosoft.com. Added to new accounts as a secondary smtp: proxy address. Listed in the Microsoft 365 admin center under Settings > Domains."
    Add-SettingsField -Key "OnMicrosoftMailDomain" -Caption "mail.onmicrosoft domain:" -Hint "The Exchange Online routing domain, for example contoso.mail.onmicrosoft.com. Added to new accounts as a secondary smtp: proxy address. Listed by Get-AcceptedDomain in Exchange Online PowerShell."
    Add-SettingsField -Key "CompanyName"           -Caption "Company:"                -Hint "Written exactly as typed to the Company attribute (company; Organization tab in Active Directory Users and Computers) on new accounts. This is free text, not an OU name - copy it from the Company field of an existing user."

    Add-SettingsSection "Microsoft 365 and Intune"
    Add-SettingsField -Key "TenantId"             -Caption "Tenant ID:"              -Hint "Optional. Directory (tenant) ID, a GUID, from the app registration's Overview page in the Microsoft Entra admin center. IntuneTools' Set to Org Defaults copies it in, and Teams Block Number signs in to this tenant and checks it matches. Leave blank to enter it in IntuneTools later; Teams Block Number then skips the check."
    Add-SettingsField -Key "ClientId"             -Caption "Client ID:"              -Hint "Optional. Application (client) ID of the app registration IntuneTools signs in through. Only IntuneTools uses it, via Set to Org Defaults; leave blank to enter it in IntuneTools later. No client secret is stored."
    Add-SettingsField -Key "ExpectedTenantDomain" -Caption "Expected tenant domain:" -Hint "A verified domain of your tenant, for example contoso.com. Teams Block Number compares it with the verified domains of the tenant you sign in to and stays read-only unless it is listed. Blank keeps changes disabled."
    Add-SettingsField -Key "SharePointAdminUrl"   -Caption "SharePoint admin URL:"    -Hint "Optional. The SharePoint admin center address, for example https://contoso-admin.sharepoint.com. OneDrive / SharePoint fills its admin URL box with it. Blank uses https://<name>-admin.sharepoint.com built from the onmicrosoft domain above; if that is blank too, type it in the tool."

    $ButtonY = $SetForm.ClientSize.Height - 44

    $SaveButton = New-Object System.Windows.Forms.Button
    $SaveButton.Text = "Save"
    $SaveButton.Location = New-Object System.Drawing.Point(($SetForm.ClientSize.Width - 214),$ButtonY)
    $SaveButton.Size = New-Object System.Drawing.Size(100,30)
    $SaveButton.Anchor = "Bottom,Right"
    $SetForm.Controls.Add($SaveButton)
    $SetForm.AcceptButton = $SaveButton

    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Text = "Cancel"
    $CloseButton.Location = New-Object System.Drawing.Point(($SetForm.ClientSize.Width - 106),$ButtonY)
    $CloseButton.Size = New-Object System.Drawing.Size(94,30)
    $CloseButton.Anchor = "Bottom,Right"
    $CloseButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $SetForm.Controls.Add($CloseButton)
    $SetForm.CancelButton = $CloseButton

    $PathLabelBottom = New-Object System.Windows.Forms.Label
    $PathLabelBottom.Text = $script:MSToolkitSettingsPath
    $PathLabelBottom.Location = New-Object System.Drawing.Point(16,($ButtonY + 7))
    $PathLabelBottom.Size = New-Object System.Drawing.Size(($SetForm.ClientSize.Width - 350),18)
    $PathLabelBottom.Anchor = "Bottom,Left,Right"
    $PathLabelBottom.AutoEllipsis = $true
    $PathLabelBottom.Font = New-Object System.Drawing.Font("Segoe UI",8)
    $SetForm.Controls.Add($PathLabelBottom)

    $ClearButton = New-Object System.Windows.Forms.Button
    $ClearButton.Text = "Clear Settings"
    $ClearButton.Location = New-Object System.Drawing.Point(($SetForm.ClientSize.Width - 326),$ButtonY)
    $ClearButton.Size = New-Object System.Drawing.Size(104,30)
    $ClearButton.Anchor = "Bottom,Right"
    $SetForm.Controls.Add($ClearButton)

    $ClearButton.Add_Click({
        # Are-you-sure box with an opt-in to forget the saved password as well.
        $Palette = Get-ThemePalette
        $SavedCred = Get-MSToolkitRememberedM365Credential

        $ClearForm = New-Object System.Windows.Forms.Form
        $ClearForm.Text = "Clear Settings"
        $ClearForm.ClientSize = New-Object System.Drawing.Size(500,364)
        $ClearForm.StartPosition = "CenterParent"
        $ClearForm.FormBorderStyle = "FixedDialog"
        $ClearForm.MaximizeBox = $false
        $ClearForm.MinimizeBox = $false
        $ClearForm.ShowInTaskbar = $false

        $ClearText = New-Object System.Windows.Forms.Label
        $ClearText.Location = New-Object System.Drawing.Point(16,14)
        $ClearText.Size = New-Object System.Drawing.Size(468,168)
        $ClearText.Text = "Are you sure? Every value in the Settings window is cleared and the blank settings are saved straight away.`r`n`r`n" +
            "- OU buttons search from the domain root again`r`n" +
            "- Delta Sync is disabled until an Entra Connect server is set`r`n" +
            "- New account naming values are empty`r`n" +
            "- IntuneTools' Set to Org Defaults finds nothing`r`n`r`n" +
            "The theme and the selected domain controller are kept. Settings opens by itself at the next launch."
        $ClearForm.Controls.Add($ClearText)

        $ForgetBox = New-Object System.Windows.Forms.CheckBox
        $ForgetBox.Text = "Also forget all saved passwords"
        $ForgetBox.Location = New-Object System.Drawing.Point(16,188)
        $ForgetBox.Size = New-Object System.Drawing.Size(468,22)
        $ForgetBox.Checked = $false
        $ClearForm.Controls.Add($ForgetBox)

        $ForgetHint = New-Object System.Windows.Forms.Label
        $ForgetHint.Location = New-Object System.Drawing.Point(34,212)
        $ForgetHint.Size = New-Object System.Drawing.Size(450,34)
        $ForgetHint.Font = New-Object System.Drawing.Font("Segoe UI",8)
        if ($SavedCred) {
            $ForgetHint.Text = "Removes the remembered sign-in for $($SavedCred.UserName) used by the Microsoft 365 tools, Logs and file buttons, and turns off automatic sign-in."
        }
        else {
            $ForgetHint.Text = "No password is saved for this account on this computer."
            $ForgetBox.Enabled = $false
        }
        $ClearForm.Controls.Add($ForgetHint)

        $LauncherBox = New-Object System.Windows.Forms.CheckBox
        $LauncherBox.Text = "Also forget the admin username remembered by the launchers"
        $LauncherBox.Location = New-Object System.Drawing.Point(16,252)
        $LauncherBox.Size = New-Object System.Drawing.Size(468,22)
        $LauncherBox.Checked = $false
        $ClearForm.Controls.Add($LauncherBox)

        $LauncherHint = New-Object System.Windows.Forms.Label
        $LauncherHint.Location = New-Object System.Drawing.Point(34,276)
        $LauncherHint.Size = New-Object System.Drawing.Size(450,34)
        $LauncherHint.Font = New-Object System.Drawing.Font("Segoe UI",8)
        $LauncherHint.Text = "Launch-MSToolkit.bat and the other launchers ask for the username again. It is removed under your Windows account, so you may be asked for that password."
        $ClearForm.Controls.Add($LauncherHint)

        $YesButton = New-Object System.Windows.Forms.Button
        $YesButton.Text = "Yes"
        $YesButton.Location = New-Object System.Drawing.Point(290,320)
        $YesButton.Size = New-Object System.Drawing.Size(94,30)
        $YesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        $ClearForm.Controls.Add($YesButton)

        $NoButton = New-Object System.Windows.Forms.Button
        $NoButton.Text = "No"
        $NoButton.Location = New-Object System.Drawing.Point(390,320)
        $NoButton.Size = New-Object System.Drawing.Size(94,30)
        $NoButton.DialogResult = [System.Windows.Forms.DialogResult]::No
        $ClearForm.Controls.Add($NoButton)

        # No is the safe default: Enter and Esc both leave everything as it is.
        $ClearForm.AcceptButton = $NoButton
        $ClearForm.CancelButton = $NoButton
        $ClearForm.Add_Shown({ $NoButton.Focus() })

        Apply-ThemeToControl -Control $ClearForm
        $ForgetHint.ForeColor = $Palette.MutedText
        $LauncherHint.ForeColor = $Palette.MutedText

        $Answer = $ClearForm.ShowDialog($script:SettingsForm)
        $ForgetPasswords = [bool]$ForgetBox.Checked
        $ForgetLauncher = [bool]$LauncherBox.Checked
        $ClearForm.Dispose()

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        foreach ($Box in $script:SettingsBoxes.Values) { $Box.Text = "" }

        if (Save-MSToolkitSettings -Settings (Get-MSToolkitSettingDefaults)) {
            Write-OutputBox "Settings cleared and saved to $($script:MSToolkitSettingsPath)" ([System.Drawing.Color]::DarkOrange)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Could not write $($script:MSToolkitSettingsPath).",
                "Settings", "OK", "Error") | Out-Null
        }

        # Before the password is forgotten, so a saved Windows sign-in can still be used.
        if ($ForgetLauncher) {
            switch (Remove-MSToolkitLauncherUsername) {
                "Done"      { Write-OutputBox "Launcher admin username forgotten." ([System.Drawing.Color]::DarkOrange) }
                "Requested" { Write-OutputBox "Launcher admin username removal sent to your Windows account. The launchers will ask for the username next time." ([System.Drawing.Color]::DarkOrange) }
                default     { Write-OutputBox "Launcher admin username was not removed." ([System.Drawing.Color]::Red) }
            }
        }

        if ($ForgetPasswords) {
            Set-MSToolkitAutoSignIn -Enabled $false
            if (Remove-MSToolkitRememberedM365Credential) {
                if (Test-Path -LiteralPath $script:MSToolkitM365CredentialPath) {
                    Write-OutputBox "Could not remove the saved password at $($script:MSToolkitM365CredentialPath)." ([System.Drawing.Color]::Red)
                }
                else {
                    Write-OutputBox "Saved password forgotten; automatic sign-in is off." ([System.Drawing.Color]::DarkOrange)
                }
            }
        }
    })

    $SaveButton.Add_Click({
        $Updated = Get-MSToolkitSettingDefaults
        foreach ($Key in @($Updated.Keys)) {
            if ($script:SettingsBoxes.ContainsKey($Key)) {
                $Updated[$Key] = "$($script:SettingsBoxes[$Key].Text)".Trim()
            }
        }

        if (Save-MSToolkitSettings -Settings $Updated) {
            Write-OutputBox "Settings saved to $($script:MSToolkitSettingsPath)" ([System.Drawing.Color]::Green)
            Write-OutputBox "Tools launched from here will pick them up straight away." ([System.Drawing.Color]::DimGray)
            $script:SettingsForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Could not write $($script:MSToolkitSettingsPath).",
                "Settings", "OK", "Error") | Out-Null
        }
    })

    Apply-ThemeToControl -Control $SetForm

    # Keep the hint lines muted - the theme pass treats them as ordinary labels.
    foreach ($Control in $Scroll.Controls) {
        if ("$($Control.Tag)" -eq "SettingsHint") {
            $Control.ForeColor = $Palette.MutedText
        }
    }

    [void]$SetForm.ShowDialog()
}

function Show-LogsPath {
    try {
        $Target = $LogPath

        if (-not (Test-Path -LiteralPath $Target)) {
            New-Item -ItemType Directory -Path $Target -Force | Out-Null
        }

        $Palette = Get-ThemePalette

        $LogForm = New-Object System.Windows.Forms.Form
        $LogForm.Text = "MSToolkit Logs"
        $LogForm.Size = New-Object System.Drawing.Size(1040,680)
        $LogForm.MinimumSize = New-Object System.Drawing.Size(960,520)
        $LogForm.StartPosition = "CenterScreen"
        $LogForm.FormBorderStyle = "Sizable"

        $HeaderPanel = New-Object System.Windows.Forms.Panel
        $HeaderPanel.Dock = "Top"
        $HeaderPanel.Height = 66
        $HeaderPanel.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $LogForm.Controls.Add($HeaderPanel)

        $HeaderTitle = New-Object System.Windows.Forms.Label
        $HeaderTitle.Text = "Logs and Reports"
        $HeaderTitle.AutoSize = $true
        $HeaderTitle.ForeColor = [System.Drawing.Color]::White
        $HeaderTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",14)
        $HeaderTitle.Location = New-Object System.Drawing.Point(18,10)
        $HeaderPanel.Controls.Add($HeaderTitle)

        $HeaderPath = New-Object System.Windows.Forms.Label
        $HeaderPath.Text = $Target
        $HeaderPath.AutoSize = $true
        $HeaderPath.ForeColor = [System.Drawing.Color]::FromArgb(218,228,240)
        $HeaderPath.Font = New-Object System.Drawing.Font("Segoe UI",9)
        $HeaderPath.Location = New-Object System.Drawing.Point(20,40)
        $HeaderPanel.Controls.Add($HeaderPath)

        $FileList = New-Object System.Windows.Forms.ListView
        $FileList.Location = New-Object System.Drawing.Point(12,76)
        $FileList.Size = New-Object System.Drawing.Size(($LogForm.ClientSize.Width - 24),230)
        $FileList.Anchor = "Top,Left,Right"
        $FileList.View = "Details"
        $FileList.FullRowSelect = $true
        $FileList.MultiSelect = $false
        $FileList.HideSelection = $false
        $FileList.Font = New-Object System.Drawing.Font("Segoe UI",9)
        [void]$FileList.Columns.Add("Name",430)
        [void]$FileList.Columns.Add("Modified",170)
        [void]$FileList.Columns.Add("Size",90)
        [void]$FileList.Columns.Add("Type",90)
        $LogForm.Controls.Add($FileList)

        $PreviewLabel = New-Object System.Windows.Forms.Label
        $PreviewLabel.Text = "Preview"
        $PreviewLabel.Location = New-Object System.Drawing.Point(12,314)
        $PreviewLabel.Size = New-Object System.Drawing.Size(200,18)
        $PreviewLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold",9)
        $LogForm.Controls.Add($PreviewLabel)

        $Preview = New-Object System.Windows.Forms.RichTextBox
        $Preview.Location = New-Object System.Drawing.Point(12,336)
        $Preview.Size = New-Object System.Drawing.Size(($LogForm.ClientSize.Width - 24),($LogForm.ClientSize.Height - 336 - 56))
        $Preview.Anchor = "Top,Bottom,Left,Right"
        $Preview.ReadOnly = $true
        $Preview.WordWrap = $false
        $Preview.Font = New-Object System.Drawing.Font("Consolas",9)
        $Preview.BorderStyle = "FixedSingle"
        $LogForm.Controls.Add($Preview)

        $ButtonY = $LogForm.ClientSize.Height - 44

        $btnRefresh = New-Object System.Windows.Forms.Button
        $btnRefresh.Text = "Refresh"
        $btnRefresh.Location = New-Object System.Drawing.Point(12,$ButtonY)
        $btnRefresh.Size = New-Object System.Drawing.Size(90,30)
        $btnRefresh.Anchor = "Bottom,Left"
        $LogForm.Controls.Add($btnRefresh)

        $btnOpen = New-Object System.Windows.Forms.Button
        $btnOpen.Text = "Open"
        $btnOpen.Location = New-Object System.Drawing.Point(110,$ButtonY)
        $btnOpen.Size = New-Object System.Drawing.Size(80,30)
        $btnOpen.Anchor = "Bottom,Left"
        $LogForm.Controls.Add($btnOpen)

        $OpenTip = New-Object System.Windows.Forms.ToolTip
        $OpenTip.SetToolTip($btnOpen, "Opens the selected file in its default app (Excel for CSVs), signed in as your Windows account")

        $btnOpenFolder = New-Object System.Windows.Forms.Button
        $btnOpenFolder.Text = "Open Folder"
        $btnOpenFolder.Location = New-Object System.Drawing.Point(198,$ButtonY)
        $btnOpenFolder.Size = New-Object System.Drawing.Size(110,30)
        $btnOpenFolder.Anchor = "Bottom,Left"
        $LogForm.Controls.Add($btnOpenFolder)

        $OpenFolderTip = New-Object System.Windows.Forms.ToolTip
        $OpenFolderTip.SetToolTip($btnOpenFolder, "Opens this folder in Explorer, signed in as your Windows account")

        $btnDelete = New-Object System.Windows.Forms.Button
        $btnDelete.Text = "Delete"
        $btnDelete.Location = New-Object System.Drawing.Point(316,$ButtonY)
        $btnDelete.Size = New-Object System.Drawing.Size(80,30)
        $btnDelete.Anchor = "Bottom,Left"
        $btnDelete.ForeColor = [System.Drawing.Color]::Red
        $LogForm.Controls.Add($btnDelete)

        $DeleteTip = New-Object System.Windows.Forms.ToolTip
        $DeleteTip.SetToolTip($btnDelete, "Permanently deletes the selected file - it does not go to the Recycle Bin")

        $btnCopyPath = New-Object System.Windows.Forms.Button
        $btnCopyPath.Text = "Copy Folder Path"
        $btnCopyPath.Location = New-Object System.Drawing.Point(404,$ButtonY)
        $btnCopyPath.Size = New-Object System.Drawing.Size(140,30)
        $btnCopyPath.Anchor = "Bottom,Left"
        $LogForm.Controls.Add($btnCopyPath)

        $btnCopyFile = New-Object System.Windows.Forms.Button
        $btnCopyFile.Text = "Copy File Path"
        $btnCopyFile.Location = New-Object System.Drawing.Point(552,$ButtonY)
        $btnCopyFile.Size = New-Object System.Drawing.Size(130,30)
        $btnCopyFile.Anchor = "Bottom,Left"
        $LogForm.Controls.Add($btnCopyFile)

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = "Close"
        $btnClose.Location = New-Object System.Drawing.Point(($LogForm.ClientSize.Width - 102),$ButtonY)
        $btnClose.Size = New-Object System.Drawing.Size(90,30)
        $btnClose.Anchor = "Bottom,Right"
        $btnClose.Add_Click({ $LogForm.Close() })
        $LogForm.Controls.Add($btnClose)
        $LogForm.CancelButton = $btnClose

        $StatusLabel = New-Object System.Windows.Forms.Label
        $StatusLabel.Location = New-Object System.Drawing.Point(692,($ButtonY + 7))

        # Sized to the gap between the last button and the right-anchored Close,
        # so it cannot run underneath it when the window is narrowed.
        $StatusWidth = $LogForm.ClientSize.Width - 692 - 112
        if ($StatusWidth -lt 60) { $StatusWidth = 60 }

        $StatusLabel.Size = New-Object System.Drawing.Size($StatusWidth,20)
        $StatusLabel.Anchor = "Bottom,Left,Right"
        $StatusLabel.AutoEllipsis = $true
        $StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI",8.5)
        $LogForm.Controls.Add($StatusLabel)

        # ------- behaviour -------

        $LoadFiles = {
            $FileList.Items.Clear()
            $Preview.Clear()

            $Files = @(
                Get-ChildItem -LiteralPath $Target -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending
            )

            foreach ($LogFile in $Files) {
                $SizeText = if ($LogFile.Length -ge 1MB) {
                    "$([math]::Round($LogFile.Length / 1MB, 1)) MB"
                }
                else {
                    "$([math]::Round($LogFile.Length / 1KB, 1)) KB"
                }

                $Item = New-Object System.Windows.Forms.ListViewItem($LogFile.Name)
                [void]$Item.SubItems.Add($LogFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
                [void]$Item.SubItems.Add($SizeText)
                [void]$Item.SubItems.Add($LogFile.Extension.TrimStart('.').ToUpper())
                $Item.Tag = $LogFile.FullName
                [void]$FileList.Items.Add($Item)
            }

            $StatusLabel.Text = "$($Files.Count) file(s)"
        }

        $ShowPreview = {
            if ($FileList.SelectedItems.Count -eq 0) { return }

            $Path = [string]$FileList.SelectedItems[0].Tag
            $Preview.Clear()

            try {
                $Info = Get-Item -LiteralPath $Path -ErrorAction Stop

                if ($Info.Length -gt 5MB) {
                    $Preview.Text = "File is $([math]::Round($Info.Length / 1MB,1)) MB. Showing the last 500 lines." + [Environment]::NewLine + [Environment]::NewLine
                    $Preview.AppendText((Get-Content -LiteralPath $Path -Tail 500 -ErrorAction Stop) -join [Environment]::NewLine)
                }
                else {
                    $Preview.Text = (Get-Content -LiteralPath $Path -ErrorAction Stop) -join [Environment]::NewLine
                }

                $PreviewLabel.Text = "Preview - $($Info.Name)"
            }
            catch {
                $Preview.Text = "Could not read this file: $($_.Exception.Message)"
            }
        }

        $btnRefresh.Add_Click($LoadFiles)
        $FileList.Add_SelectedIndexChanged($ShowPreview)

        $btnOpen.Add_Click({
            if ($FileList.SelectedItems.Count -eq 0) {
                $StatusLabel.Text = "Select a file first."
                return
            }

            $SelectedFile = [string]$FileList.SelectedItems[0].Tag
            $StatusLabel.Text = "Signing in to open the file..."

            # Opens in the default app (Excel for a CSV) under the signed-in Windows
            # account, where Office is licensed - the Domain Admin session is not.
            if (Open-MSToolkitFileAsSignedInUser -FilePath $SelectedFile) {
                $StatusLabel.Text = "File opened."
            }
            else {
                $StatusLabel.Text = "Open cancelled."
            }
        })

        $btnDelete.Add_Click({
            if ($FileList.SelectedItems.Count -eq 0) {
                $StatusLabel.Text = "Select a file first."
                return
            }

            $DeletePath = [string]$FileList.SelectedItems[0].Tag
            $DeleteName = Split-Path -Leaf $DeletePath

            # MSToolkit created these files under its own account, so it can delete them
            # itself - no sign-in needed. Permanent: Remove-Item bypasses the Recycle Bin.
            $Confirm = [System.Windows.Forms.MessageBox]::Show(
                "Permanently delete this file?`r`n`r`n$DeleteName`r`n`r`nFolder: $Target`r`n`r`nIt will not go to the Recycle Bin.",
                "Delete Log File",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )

            if ($Confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $StatusLabel.Text = "Delete cancelled."
                return
            }

            $Deleted = $false
            try {
                Remove-Item -LiteralPath $DeletePath -Force -ErrorAction Stop
                $Deleted = $true
            }
            catch {
                # Refused under MSToolkit' account - retry as the signed-in Windows user.
                $StatusLabel.Text = "Not allowed as the admin account - signing in to delete..."
                Write-OutputBox "Could not delete $DeleteName as the admin account ($($_.Exception.Message)). Retrying as your Windows account."

                switch (Remove-MSToolkitFileAsSignedInUser -FilePath $DeletePath) {
                    "Deleted"   { $Deleted = $true }
                    "NotDone"   { $StatusLabel.Text = "Delete not completed - see Activity Output." }
                    default     { $StatusLabel.Text = "Could not delete $DeleteName under your Windows account either." }
                }
            }

            if ($Deleted) {
                Write-OutputBox "Deleted log file: $DeletePath"
                & $LoadFiles
                $StatusLabel.Text = "Deleted $DeleteName."
            }
        })

        $btnOpenFolder.Add_Click({
            $StatusLabel.Text = "Signing in to open the folder..."

            if (Open-MSToolkitFolderAsSignedInUser -Path $Target) {
                $StatusLabel.Text = "Folder opened in Explorer."
            }
            else {
                $StatusLabel.Text = "Open folder cancelled."
            }
        })

        $btnCopyPath.Add_Click({
            try {
                Set-Clipboard -Value $Target -ErrorAction Stop
                $StatusLabel.Text = "Folder path copied to the clipboard."
            }
            catch {
                try {
                    [System.Windows.Forms.Clipboard]::SetText($Target)
                    $StatusLabel.Text = "Folder path copied to the clipboard."
                }
                catch {
                    $StatusLabel.Text = "Could not copy the path."
                }
            }
        })

        $btnCopyFile.Add_Click({
            if ($FileList.SelectedItems.Count -eq 0) {
                $StatusLabel.Text = "Select a file first."
                return
            }

            $Path = [string]$FileList.SelectedItems[0].Tag

            try {
                Set-Clipboard -Value $Path -ErrorAction Stop
                $StatusLabel.Text = "File path copied to the clipboard."
            }
            catch {
                try {
                    [System.Windows.Forms.Clipboard]::SetText($Path)
                    $StatusLabel.Text = "File path copied to the clipboard."
                }
                catch {
                    $StatusLabel.Text = "Could not copy the path."
                }
            }
        })

        & $LoadFiles

        Apply-ThemeToControl -Control $LogForm
        $Preview.BackColor = $Palette.OutputBackground
        $Preview.ForeColor = $Palette.Text
        $FileList.BackColor = $Palette.InputBackground
        $FileList.ForeColor = $Palette.Text

        Write-OutputBox "Opened the logs list for: $Target"

        [void]$LogForm.ShowDialog()
    }
    catch {
        Write-OutputBox "ERROR opening the logs list: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Launch-PrintManagement {
    try {
        $PrintManagementMsc = Join-Path $env:SystemRoot "System32\printmanagement.msc"

        if (-not (Test-Path $PrintManagementMsc)) {
            Write-OutputBox "ERROR: Print Management (printmanagement.msc) is not installed on this computer." ([System.Drawing.Color]::Red)
            return
        }

        Write-OutputBox "Launching Print Management..."
        $Process = Start-Process "mmc.exe" -ArgumentList "`"$PrintManagementMsc`"" -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Print Management: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Launch-RemoteDesktop {
    try {
        $Server = Get-SelectedServer

        if ([string]::IsNullOrWhiteSpace($Server)) {
            Write-OutputBox "ERROR: No DC selected." ([System.Drawing.Color]::Red)
            return
        }

        Write-OutputBox "Opening Remote Desktop to $Server..."
        $Process = Start-Process "mstsc.exe" -ArgumentList "/v:$Server" -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Remote Desktop: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Launch-PowerShell {
    try {
        Write-OutputBox "Opening PowerShell in $PSScriptRoot..."
        $Process = Start-Process "powershell.exe" -WorkingDirectory $PSScriptRoot -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching PowerShell: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}


function Launch-ISE {
    try {
        $IseExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell_ise.exe"

        if (-not (Test-Path $IseExe)) {
            Write-OutputBox "ERROR: Windows PowerShell ISE (powershell_ise.exe) is not installed on this computer." ([System.Drawing.Color]::Red)
            return
        }

        Write-OutputBox "Opening Windows PowerShell ISE in $PSScriptRoot..."
        $Process = Start-Process $IseExe -WorkingDirectory $PSScriptRoot -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Windows PowerShell ISE: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}


function Clear-ActivityOutput {
    if ($script:OutputEntries) {
        $script:OutputEntries.Clear()
    }

    if ($OutputBox) {
        $OutputBox.Clear()
    }
}

function Copy-ActivityOutput {
    if (-not $OutputBox -or [string]::IsNullOrWhiteSpace($OutputBox.Text)) {
        Write-OutputBox "Nothing is currently available to copy from Activity Output." ([System.Drawing.Color]::DarkOrange)
        return
    }

    try {
        [System.Windows.Forms.Clipboard]::SetText($OutputBox.Text)
        Write-OutputBox "Activity Output copied to the clipboard." (Get-ThemePalette).Success
    }
    catch {
        Write-OutputBox "ERROR copying Activity Output: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Update-StatusStrip {
    if (-not $StatusPanel -or -not $StatusAccountLabel -or -not $StatusDomainLabel -or -not $StatusServerLabel) {
        return
    }

    $Palette = Get-ThemePalette
    $IdentityName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Server = Get-SelectedServer

    $StatusAccountLabel.Text = "Account: $IdentityName"
    $StatusAccountLabel.ForeColor = $Palette.Text
    $StatusDomainLabel.Text = "Domain: $Domain"
    $StatusDomainLabel.ForeColor = $Palette.Text

    try {
        Get-ADRootDSE -Server $Server -ErrorAction Stop | Out-Null
        $StatusServerLabel.Text = "AD Server: $Server  |  Reachable"
        $StatusServerLabel.ForeColor = $Palette.Success
    }
    catch {
        # The check uses Get-ADRootDSE, which only works over ADWS (TCP 9389). A DC can
        # still answer LDAP, RDP and everything else, so say exactly what failed.
        $StatusServerLabel.Text = "AD Server: $Server  |  ADWS unreachable (port 9389)"
        $StatusServerLabel.ForeColor = $Palette.Danger
    }
}

function Launch-NewADUserScript {
    try {
        $Script = Join-Path $PSScriptRoot "NewADUser.ps1"

        if (-not (Test-Path $Script)) {
            Write-OutputBox "ERROR: New user script not found: $Script" ([System.Drawing.Color]::Red)
            return
        }

        $SelectedServer = Get-SelectedServer
        Write-OutputBox "Launching New AD User script using selected AD server: $SelectedServer"

        $Process = Start-Process powershell.exe `
            -WorkingDirectory $PSScriptRoot `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$Script`" -ThemeMode `"$script:ThemeMode`" -DefaultServer `"$SelectedServer`"" `
            -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching New AD User script: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
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

    $Palette = Get-ThemePalette

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

    Apply-ThemeToControl -Control $CredForm

    $AutoHint.ForeColor = $Palette.MutedText
    $UserHint.ForeColor = $Palette.MutedText

    # Blend the reveal control into the field and keep typed text clear of it.
    $RevealButton.BackColor = $PassBox.BackColor
    $RevealButton.ForeColor = $Palette.MutedText

    $CredForm.Add_Shown({
        try {
            [WindowHelper]::SetRightMargin($PassBox.Handle, ($RevealButton.Width + 10))
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
            Write-OutputBox "$Purpose - signed in automatically as $($SISaved.UserName). Hold Shift when clicking to change this."
            return $true
        }
        catch {
            if (Test-MSToolkitCredentialFailure -ErrorRecord $_) {
                Write-OutputBox "Automatic sign-in as $($SISaved.UserName) failed: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                [void](Remove-MSToolkitRememberedM365Credential)
                Set-MSToolkitAutoSignIn -Enabled $false
                Write-OutputBox "The remembered password no longer works and has been cleared." ([System.Drawing.Color]::DarkOrange)
                $SIPassword = ""
                $SIRemember = $false
                $SIAuto = $false
            }
            else {
                # Signed in fine; the action itself failed. Keep the saved password.
                Write-OutputBox "$Purpose failed after signing in as $($SISaved.UserName): $($_.Exception.Message)" ([System.Drawing.Color]::Red)
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
            Write-OutputBox "$Purpose cancelled."
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
                        Write-OutputBox "Password remembered for $($SICred.UserName); you will be signed in automatically next time."
                    }
                    else {
                        Write-OutputBox "Password remembered for $($SICred.UserName) on this computer."
                    }
                }
                else {
                    Write-OutputBox "Could not save the password; you will be asked next time." ([System.Drawing.Color]::DarkOrange)
                }
            }
            else {
                Set-MSToolkitAutoSignIn -Enabled $false
                if (Remove-MSToolkitRememberedM365Credential) {
                    Write-OutputBox "Remembered password cleared."
                }
            }

            return $true
        }
        catch {
            if (-not (Test-MSToolkitCredentialFailure -ErrorRecord $_)) {
                # The sign-in worked; the action itself failed. Report it as what it is,
                # keep any saved password, and do not ask for the password again.
                Write-OutputBox "$Purpose failed after signing in as $($SICred.UserName): $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                return $false
            }

            Write-OutputBox "Sign-in failed for $($SICred.UserName) (attempt $SIAttempt): $($_.Exception.Message)" ([System.Drawing.Color]::Red)

            if (Remove-MSToolkitRememberedM365Credential) {
                Set-MSToolkitAutoSignIn -Enabled $false
                Write-OutputBox "The remembered password no longer works and has been cleared." ([System.Drawing.Color]::DarkOrange)
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
        [string]$ExtraArguments = ""
    )

    $Script = Join-Path $PSScriptRoot $ScriptFile

    if (-not (Test-Path $Script)) {
        Write-OutputBox "ERROR: $ToolName script not found: $Script" ([System.Drawing.Color]::Red)
        return
    }

    Write-OutputBox "Signing in to launch $ToolName."

    [void](Invoke-MSToolkitAsSignedInUser -Purpose $ToolName -Launch {
        param($Credential)

        Write-OutputBox "Launching $ToolName as: $($Credential.UserName)"

        $Process = Start-Process powershell.exe `
            -Credential $Credential `
            -WorkingDirectory $PSScriptRoot `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$Script`" -ThemeMode `"$script:ThemeMode`"$ExtraArguments" `
            -PassThru `
            -ErrorAction Stop

        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    })
}

function Launch-M365GroupCompareScript {
    Start-MSToolkitM365Tool -ToolName "M365 Group Compare/Add" -ScriptFile "M365-Group-Compare.ps1"
}

function Launch-M365DistributionGroupCompareScript {
    Start-MSToolkitM365Tool -ToolName "M365 Distro Compare/Add" -ScriptFile "M365-Distribution-Group-Compare.ps1"
}

function Get-MSToolkitSharePointAdminUrl {
    # The SharePoint admin URL setting, or one built from the onmicrosoft domain
    # (contoso.onmicrosoft.com -> https://contoso-admin.sharepoint.com), or "".
    $Configured = Get-MSToolkitSetting -Name "SharePointAdminUrl"
    if ($Configured) { return $Configured.TrimEnd('/') }

    $OnMicrosoft = Get-MSToolkitSetting -Name "OnMicrosoftDomain"
    if ($OnMicrosoft -match '^([A-Za-z0-9-]+)\.onmicrosoft\.com$') {
        return "https://$($Matches[1])-admin.sharepoint.com"
    }

    return ""
}

function Launch-M365OneDriveTools {
    # Runs as the signed-in user, whose profile cannot see this console's
    # settings.json, so the admin URL is handed over at launch. It only pre-fills
    # the tool's admin URL box.
    $OneDriveArguments = ""

    $AdminUrl = Get-MSToolkitSharePointAdminUrl
    if ($AdminUrl) { $OneDriveArguments += " -AdminUrl `"$AdminUrl`"" }

    Start-MSToolkitM365Tool -ToolName "OneDrive and SharePoint Tools" -ScriptFile "M365-OneDrive-SharePoint-Tools.ps1" -ExtraArguments $OneDriveArguments
}

function Launch-M365ExchangeOnlineTools {
    Start-MSToolkitM365Tool -ToolName "Exchange Online Tools" -ScriptFile "M365-Exchange-Online-Tools.ps1"
}

function Launch-IntuneTools {
    # IntuneTools runs as the signed-in user, whose profile cannot see this console's
    # settings.json, so the org Tenant ID and Client ID are handed over at launch. They
    # only feed its "Set to Org Defaults" button; nothing is filled in automatically.
    $IntuneArguments = ""

    $OrgTenantId = Get-MSToolkitSetting -Name "TenantId"
    if ($OrgTenantId) { $IntuneArguments += " -OrgTenantId `"$OrgTenantId`"" }

    $OrgClientId = Get-MSToolkitSetting -Name "ClientId"
    if ($OrgClientId) { $IntuneArguments += " -OrgClientId `"$OrgClientId`"" }

    Start-MSToolkitM365Tool -ToolName "Intune Tools" -ScriptFile "IntuneTools.ps1" -ExtraArguments $IntuneArguments
}

function Launch-M365TeamsBlockNumber {
    # Teams Block Number runs as the signed-in user, whose profile cannot see this
    # console's settings.json, so its tenant check values are handed over at launch.
    $TeamsArguments = ""

    $ExpectedDomain = Get-MSToolkitSetting -Name "ExpectedTenantDomain"
    if ($ExpectedDomain) { $TeamsArguments += " -ExpectedTenantDomain `"$ExpectedDomain`"" }

    $OrgTenantId = Get-MSToolkitSetting -Name "TenantId"
    if ($OrgTenantId) { $TeamsArguments += " -TenantId `"$OrgTenantId`"" }

    Start-MSToolkitM365Tool -ToolName "M365 Teams Block Number" -ScriptFile "M365-Teams-Block-Number.ps1" -ExtraArguments $TeamsArguments
}

function Launch-M365ConditionalAccessScript {
    Start-MSToolkitM365Tool -ToolName "M365 Conditional Access" -ScriptFile "M365-Conditional-Access-User-Manager.ps1"
}




function Invoke-CompareUserGroupsInTool {
    try {
        $Script = Join-Path $PSScriptRoot "Compare-UserGroups.ps1"

        if (-not (Test-Path $Script)) {
            Write-OutputBox "ERROR: Compare User Groups script not found: $Script" ([System.Drawing.Color]::Red)
            return
        }

        $Server = Get-SelectedServer
        Write-OutputBox "Launching Compare User Groups window using $Server..."

        $Process = Start-Process powershell.exe `
            -WorkingDirectory $PSScriptRoot `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$Script`" -Server `"$Server`" -ThemeMode `"$script:ThemeMode`"" `
            -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Compare User Groups window: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Invoke-InvestigateAccountLockoutInTool {
    try {
        $Script = Join-Path $PSScriptRoot "Investigate-AccountLockout.ps1"

        if (-not (Test-Path $Script)) {
            Write-OutputBox "ERROR: Investigate Account Lockout script not found: $Script" ([System.Drawing.Color]::Red)
            return
        }

        Write-OutputBox "Launching Investigate Account Lockout window..."

        $Process = Start-Process powershell.exe `
            -WorkingDirectory $PSScriptRoot `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$Script`" -ThemeMode `"$script:ThemeMode`"" `
            -PassThru
        Move-ProcessWindowsToMSToolkitMonitor -Process $Process
    }
    catch {
        Write-OutputBox "ERROR launching Investigate Account Lockout window: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Show-ADUserInfo {
    $Sam = Get-InputBox "Get AD User" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer

        $User = Get-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Properties DisplayName,Enabled,LockedOut,PasswordLastSet,pwdLastSet,PasswordNeverExpires,LastLogonDate,MemberOf,Department,Title,Company,EmailAddress,Manager,Office,TelephoneNumber,DistinguishedName

        $ManagerDisplay = ""

        if ($User.Manager) {
            try {
                $ManagerUser = Get-ADUser `
                    -Identity $User.Manager `
                    -Server $Server `
                    -Properties DisplayName,SamAccountName `
                    -ErrorAction Stop

                $ManagerDisplay = "$($ManagerUser.DisplayName) ($($ManagerUser.SamAccountName))"
            }
            catch {
                $ManagerDisplay = $User.Manager
            }
        }

        Write-OutputField "AD Server Used" $Server
        Write-OutputField "User found" $User.DisplayName
        Write-OutputField "Username" $User.SamAccountName

        if ($User.Enabled -eq $true) {
            Write-OutputField "Enabled" "True" ([System.Drawing.Color]::Green) ([System.Drawing.Color]::Green)
        }
        else {
            Write-OutputField "Enabled" "False" ([System.Drawing.Color]::Red) ([System.Drawing.Color]::Red)
        }

        if ($User.LockedOut -eq $true) {
            Write-OutputField "Locked Out" "True" ([System.Drawing.Color]::Red) ([System.Drawing.Color]::Red)
        }
        else {
            Write-OutputField "Locked Out" "False" ([System.Drawing.Color]::Green) ([System.Drawing.Color]::Green)
        }

        if ([int64]$User.pwdLastSet -eq 0) {
            Write-OutputField "Password Must Be Changed" "True" ([System.Drawing.Color]::Red) ([System.Drawing.Color]::Red)
        }
        else {
            Write-OutputField "Password Must Be Changed" "False" ([System.Drawing.Color]::Green) ([System.Drawing.Color]::Green)
        }

        if ($User.PasswordNeverExpires -eq $true) {
            Write-OutputField "Password Never Expires" "True" ([System.Drawing.Color]::Red) ([System.Drawing.Color]::Red)
        }
        else {
            Write-OutputField "Password Never Expires" "False" ([System.Drawing.Color]::Green) ([System.Drawing.Color]::Green)
        }

        Write-OutputField "Manager" $ManagerDisplay
        Write-OutputField "Office" $User.Office
        Write-OutputField "Telephone Number" $User.TelephoneNumber
        Write-OutputField "E-mail" $User.EmailAddress
        Write-OutputField "Department" $User.Department
        Write-OutputField "Title" $User.Title
        Write-OutputField "Company" $User.Company
        Write-OutputField "Password Last Set" $User.PasswordLastSet
        Write-OutputField "Last Logon Date" $User.LastLogonDate
        Write-OutputField "DistinguishedName" $User.DistinguishedName
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Unlock-ADUserAccount {
    $Sam = Get-InputBox "Unlock AD User" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    try {
        $Server = Get-SelectedServer
        $User = Get-ADUser -Identity $Sam -Server $Server -Properties SID,isCriticalSystemObject,DisplayName -ErrorAction Stop
        $CriticalReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Unlock account" -Reason $CriticalReason) { return }

        Unlock-ADAccount `
            -Identity $User.DistinguishedName `
            -Server $Server

        Write-OutputBox "Unlocked account: $($User.SamAccountName) using $Server" ([System.Drawing.Color]::Green)
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Enable-ADUserAccount {
    $Sam = Get-InputBox "Enable AD User" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    try {
        $Server = Get-SelectedServer
        $User = Get-ADUser -Identity $Sam -Server $Server -Properties SID,isCriticalSystemObject,DisplayName -ErrorAction Stop
        $CriticalReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Enable account" -Reason $CriticalReason) { return }

        Enable-ADAccount `
            -Identity $User.DistinguishedName `
            -Server $Server

        Write-OutputBox "Enabled account: $($User.SamAccountName) using $Server" ([System.Drawing.Color]::Green)
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Disable-ADUserAccount {
    $Sam = Get-InputBox "Disable AD User" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    try {
        $Server = Get-SelectedServer
        $User = Get-ADUser -Identity $Sam -Server $Server -Properties SID,isCriticalSystemObject,DisplayName,Enabled,DistinguishedName -ErrorAction Stop
        $CriticalReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Disable account" -Reason $CriticalReason) { return }

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to disable $($User.DisplayName) ($($User.SamAccountName))?",
            "Confirm Disable",
            "YesNo",
            "Warning"
        )

        if ($Confirm -eq "Yes") {
            Disable-ADAccount `
                -Identity $User.DistinguishedName `
                -Server $Server

            Write-OutputBox "Disabled account: $($User.SamAccountName) using $Server" ([System.Drawing.Color]::DarkOrange)
        }
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Delete-ADUserAccount {
    $Sam = Get-InputBox "Delete AD User" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    $ProtectionRemoved = $false
    $User = $null
    $Server = Get-SelectedServer

    try {
        $User = Get-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Properties DisplayName,SamAccountName,Enabled,Description,WhenCreated,DistinguishedName,ProtectedFromAccidentalDeletion,SID,isCriticalSystemObject `
            -ErrorAction Stop

        $CriticalReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Delete user" -Reason $CriticalReason) { return }

        $ConfirmText = @"
You are about to DELETE the following Active Directory user:

Display Name: $($User.DisplayName)
Username: $($User.SamAccountName)
Enabled: $($User.Enabled)
Protected from Accidental Deletion: $($User.ProtectedFromAccidentalDeletion)
Created: $($User.WhenCreated)
Description: $($User.Description)
Distinguished Name: $($User.DistinguishedName)

This is permanent unless the object is restored from AD Recycle Bin/backups.

Continue with deletion?
"@

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            $ConfirmText,
            "Confirm Delete AD User",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($Confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-OutputBox "Delete cancelled for $($User.SamAccountName)."
            return
        }

        if ($User.ProtectedFromAccidentalDeletion) {
            $OverrideApproved = Request-MSToolkitAccidentalDeletionOverride `
                -ObjectType "AD user" `
                -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" `
                -DistinguishedName $User.DistinguishedName `
                -Server $Server

            if (-not $OverrideApproved) { return }
            $ProtectionRemoved = $true
        }

        try {
            Remove-ADUser `
                -Identity $User.DistinguishedName `
                -Server $Server `
                -Confirm:$false `
                -ErrorAction Stop

            Write-OutputBox "Deleted AD user: $($User.DisplayName) / $($User.SamAccountName) using $Server" ([System.Drawing.Color]::Red)
        }
        catch {
            if ($ProtectionRemoved) {
                Restore-MSToolkitAccidentalDeletionProtection -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -DistinguishedName $User.DistinguishedName -Server $Server
            }
            throw
        }
    }
    catch {
        Write-OutputBox "ERROR deleting user: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Reset-ADUserPassword {
    $Sam = Get-InputBox "Reset Password" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    try {
        $Server = Get-SelectedServer
        $User = Get-ADUser -Identity $Sam -Server $Server -Properties SID,isCriticalSystemObject,DisplayName,DistinguishedName -ErrorAction Stop
        $CriticalReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Reset password" -Reason $CriticalReason) { return }

        $TempPassword = Get-InputBox "Temporary Password" "Enter temporary password:" -Password
        if (-not $TempPassword) { return }

        $SecurePassword = ConvertTo-SecureString $TempPassword -AsPlainText -Force

        Set-ADAccountPassword `
            -Identity $User.DistinguishedName `
            -NewPassword $SecurePassword `
            -Reset `
            -Server $Server

        Set-ADUser `
            -Identity $User.DistinguishedName `
            -ChangePasswordAtLogon $true `
            -Server $Server

        Write-OutputBox "Password reset for $($User.SamAccountName) using $Server. Change at next logon enabled." ([System.Drawing.Color]::Green)
        $TempPassword = $null
        $SecurePassword = $null
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Force-PasswordChange {
    $Sam = Get-InputBox "Force Password Change" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    try {
        $Server = Get-SelectedServer
        $User = Get-ADUser -Identity $Sam -Server $Server -Properties SID,isCriticalSystemObject,DisplayName,DistinguishedName -ErrorAction Stop
        $CriticalReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Force password change" -Reason $CriticalReason) { return }

        Set-ADUser `
            -Identity $User.DistinguishedName `
            -ChangePasswordAtLogon $true `
            -Server $Server

        Write-OutputBox "Set change password at next logon for $($User.SamAccountName) using $Server" ([System.Drawing.Color]::Green)
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ADUserOU {
    $Sam = Get-InputBox "Get User OU" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer

        $User = Get-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Properties DisplayName,DistinguishedName,Enabled,Description

        $PathColor = [System.Drawing.Color]::DimGray
        $DetailColor = [System.Drawing.Color]::FromArgb(70,70,70)

        # Strip the leading CN= component to get the OU the user actually sits in.
        $Parts = [regex]::Split($User.DistinguishedName, '(?<!\\),')
        $ParentDN = ($Parts[1..($Parts.Count - 1)] -join ',')

        Write-ReadableSectionHeader -Title "OU Location for $($User.DisplayName) ($Sam)" -Server $Server -BasePath ""

        Write-OutputField -Label "Name" -Value $User.Name -ValueColor $DetailColor
        Write-OutputField -Label "sAMAccountName" -Value $User.SamAccountName -ValueColor $DetailColor
        Write-OutputField -Label "Enabled" -Value ([string]$User.Enabled) -ValueColor $DetailColor
        Write-OutputField -Label "Parent OU" -Value $ParentDN -ValueColor $PathColor
        Write-OutputField -Label "Object DN" -Value $User.DistinguishedName -ValueColor $PathColor

        $OUNames = Get-MSToolkitOUNamesFromDN -DistinguishedName $User.DistinguishedName

        if ($OUNames.Count -gt 0) {
            $TopDown = @($OUNames)
            [array]::Reverse($TopDown)
            Write-OutputField -Label "OU Path" -Value ($TopDown -join " \ ") -ValueColor ([System.Drawing.Color]::DarkCyan)
        }

        Write-OUBlankLine
    }
    catch {
        Write-OutputBox "ERROR getting OU for ${Sam}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ADUserGroups {
    $Sam = Get-InputBox "Get User Groups" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer

        $User = Get-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Properties DisplayName,MemberOf

        $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $PathColor = [System.Drawing.Color]::DimGray
        $DetailColor = [System.Drawing.Color]::FromArgb(70,70,70)
        $NameColors = Get-ReadableNameColors

        Write-ReadableSectionHeader -Title "Direct Groups for $($User.DisplayName) ($Sam)" -Server $Server -BasePath ""

        if (-not $User.MemberOf -or $User.MemberOf.Count -eq 0) {
            Write-OutputBox "$Sam is not a direct member of any groups." ([System.Drawing.Color]::DarkOrange)
            Write-OUBlankLine
            return
        }

        $Groups = @(
            foreach ($GroupDN in $User.MemberOf) {
                try {
                    Get-ADGroup `
                        -Identity $GroupDN `
                        -Server $Server `
                        -Properties Name,GroupCategory,GroupScope,DistinguishedName
                }
                catch {
                    Write-OutputBox "WARNING: Could not read group: $GroupDN" ([System.Drawing.Color]::DarkOrange)
                }
            }
        ) | Sort-Object Name

        $ColorIndex = 0

        foreach ($Group in $Groups) {
            $NameColor = $NameColors[$ColorIndex % $NameColors.Count]

            Write-OutputBox $Group.Name $NameColor
            Write-OutputField "Scope / Category" "$($Group.GroupScope) / $($Group.GroupCategory)" $DetailColor
            Write-OutputBox "  $($Group.DistinguishedName)" $PathColor
            Write-OUBlankLine

            $ColorIndex++
        }

        Write-OutputBox "Total direct groups: $($Groups.Count)" $SectionColor
        Write-OUBlankLine
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Add-UserToGroup {
    $Sam = Get-InputBox "Add User to Group" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    $GroupName = Get-InputBox "Add User to Group" "Enter group name / sAMAccountName / distinguishedName:"
    if (-not $GroupName) { return }

    try {
        $Server = Get-SelectedServer
        $User = Get-ADUser -Identity $Sam -Server $Server -Properties SID,isCriticalSystemObject,DisplayName,DistinguishedName -ErrorAction Stop
        $Group = Get-ADGroup -Identity $GroupName -Server $Server -Properties SID,isCriticalSystemObject,Name,SamAccountName,DistinguishedName -ErrorAction Stop

        $CriticalUserReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Change group membership" -Reason $CriticalUserReason) { return }

        $CriticalGroupReason = Get-MSToolkitCriticalGroupReason -Group $Group
        if (Stop-MSToolkitCriticalOperation -ObjectType "Group" -DisplayName "$($Group.Name)" -Operation "Add group member" -Reason $CriticalGroupReason) { return }

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add user $($User.DisplayName) ($($User.SamAccountName)) to group $($Group.Name)?",
            "Confirm Add to Group",
            "YesNo",
            "Question"
        )

        if ($Confirm -eq "Yes") {
            Add-ADGroupMember `
                -Identity $Group.DistinguishedName `
                -Members $User.DistinguishedName `
                -Server $Server `
                -Confirm:$false

            Write-OutputBox "Added user $($User.SamAccountName) to group $($Group.Name) using $Server" ([System.Drawing.Color]::Green)
        }
        else {
            Write-OutputBox "Add user to group cancelled."
        }
    }
    catch {
        Write-OutputBox "ERROR adding user to group: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Remove-UserFromGroup {
    $Sam = Get-InputBox "Remove User from Group" "Enter username / sAMAccountName:"
    if (-not $Sam) { return }

    try {
        $Server = Get-SelectedServer

        Write-OutputBox "Getting direct groups for $Sam using $Server..."

        $User = Get-ADUser `
            -Identity $Sam `
            -Server $Server `
            -Properties MemberOf,SID,DisplayName `
            -ErrorAction Stop

        $CriticalUserReason = Get-MSToolkitCriticalUserReason -User $User
        if (Stop-MSToolkitCriticalOperation -ObjectType "User" -DisplayName "$($User.DisplayName) ($($User.SamAccountName))" -Operation "Change group membership" -Reason $CriticalUserReason) { return }

        if (-not $User.MemberOf -or $User.MemberOf.Count -eq 0) {
            Write-OutputBox "$Sam is not a direct member of any removable groups." ([System.Drawing.Color]::DarkOrange)
            return
        }

        $Groups = New-Object System.Collections.Generic.List[object]

        foreach ($GroupDN in $User.MemberOf) {
            try {
                $Group = Get-ADGroup `
                    -Identity $GroupDN `
                    -Server $Server `
                    -Properties Name,SamAccountName,GroupCategory,GroupScope,Description,SID,isCriticalSystemObject,DistinguishedName `
                    -ErrorAction Stop

                $Groups.Add($Group)
            }
            catch {
                Write-OutputBox "WARNING: Could not read group: $GroupDN" ([System.Drawing.Color]::DarkOrange)
            }
        }

        $SortedGroups = @($Groups.ToArray() | Sort-Object Name)

        if ($SortedGroups.Count -eq 0) {
            Write-OutputBox "No readable direct groups were found for $Sam." ([System.Drawing.Color]::DarkOrange)
            return
        }

        $Picker = New-Object System.Windows.Forms.Form
        $Picker.Text = "Remove User from Group - Select Group"
        $Picker.Size = New-Object System.Drawing.Size(760,500)
        $Picker.StartPosition = "CenterScreen"
        $Picker.FormBorderStyle = "FixedDialog"
        $Picker.MaximizeBox = $false
        $Picker.MinimizeBox = $false
        $Picker.Font = New-Object System.Drawing.Font("Segoe UI",9)

        $Label = New-Object System.Windows.Forms.Label
        $Label.Text = "Select one or more direct groups to remove $Sam from. Use Ctrl or Shift for multiple selections:"
        $Label.Location = New-Object System.Drawing.Point(15,15)
        $Label.Size = New-Object System.Drawing.Size(710,25)
        $Picker.Controls.Add($Label)

        $Grid = New-Object System.Windows.Forms.DataGridView
        $Grid.Location = New-Object System.Drawing.Point(15,45)
        $Grid.Size = New-Object System.Drawing.Size(715,350)
        $Grid.ReadOnly = $true
        $Grid.MultiSelect = $true
        $Grid.SelectionMode = "FullRowSelect"
        $Grid.AllowUserToAddRows = $false
        $Grid.AllowUserToDeleteRows = $false
        $Grid.RowHeadersVisible = $false
        $Grid.AutoSizeColumnsMode = "Fill"
        $Grid.BackgroundColor = [System.Drawing.Color]::White

        [void]$Grid.Columns.Add("GroupName","Group Name")
        [void]$Grid.Columns.Add("Scope","Scope")
        [void]$Grid.Columns.Add("Category","Category")
        [void]$Grid.Columns.Add("Description","Description")

        foreach ($Group in $SortedGroups) {
            $Index = $Grid.Rows.Add(
                $Group.Name,
                [string]$Group.GroupScope,
                [string]$Group.GroupCategory,
                $Group.Description
            )
            $Grid.Rows[$Index].Tag = $Group
        }

        $Picker.Controls.Add($Grid)

        $RemoveButton = New-Object System.Windows.Forms.Button
        $RemoveButton.Text = "Select Group(s)"
        $RemoveButton.Location = New-Object System.Drawing.Point(520,410)
        $RemoveButton.Size = New-Object System.Drawing.Size(100,32)
        $RemoveButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $Picker.Controls.Add($RemoveButton)
        $Picker.AcceptButton = $RemoveButton

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Text = "Cancel"
        $CancelButton.Location = New-Object System.Drawing.Point(630,410)
        $CancelButton.Size = New-Object System.Drawing.Size(100,32)
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $Picker.Controls.Add($CancelButton)
        $Picker.CancelButton = $CancelButton

        $Grid.Add_CellDoubleClick({
            if ($Grid.SelectedRows.Count -gt 0) {
                $Picker.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $Picker.Close()
            }
        })

        $Picker.Add_Shown({
            if ($Grid.Rows.Count -gt 0) {
                $Grid.Rows[0].Selected = $true
                $Grid.CurrentCell = $Grid.Rows[0].Cells[0]
            }
            $Picker.Activate()
        })

        Apply-ThemeToControl -Control $Picker
        $Result = $Picker.ShowDialog()

        if ($Result -ne [System.Windows.Forms.DialogResult]::OK -or $Grid.SelectedRows.Count -eq 0) {
            $Picker.Dispose()
            Write-OutputBox "Remove user from group cancelled."
            return
        }

        $SelectedGroups = New-Object System.Collections.Generic.List[object]

        foreach ($SelectedRow in $Grid.SelectedRows) {
            if ($SelectedRow.Tag) {
                $SelectedGroups.Add($SelectedRow.Tag)
            }
        }

        $Picker.Dispose()

        if ($SelectedGroups.Count -eq 0) {
            Write-OutputBox "Remove user from group cancelled."
            return
        }

        foreach ($SelectedGroup in $SelectedGroups.ToArray()) {
            $CriticalGroupReason = Get-MSToolkitCriticalGroupReason -Group $SelectedGroup
            if (Stop-MSToolkitCriticalOperation -ObjectType "Group" -DisplayName "$($SelectedGroup.Name)" -Operation "Remove group member" -Reason $CriticalGroupReason) {
                Write-OutputBox "No group memberships were changed because the selection included a critical protected group." ([System.Drawing.Color]::Red)
                return
            }
        }

        $GroupNames = @($SelectedGroups.ToArray() | Sort-Object Name | ForEach-Object { $_.Name })
        $GroupListText = ($GroupNames -join "`r`n - ")

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove user $Sam from the following $($SelectedGroups.Count) group(s)?`r`n`r`n - $GroupListText",
            "Confirm Remove from Groups",
            "YesNo",
            "Warning"
        )

        if ($Confirm -ne "Yes") {
            Write-OutputBox "Remove user from group cancelled."
            return
        }

        $SuccessCount = 0
        $FailureCount = 0

        foreach ($SelectedGroup in $SelectedGroups.ToArray()) {
            try {
                Remove-ADGroupMember `
                    -Identity $SelectedGroup.DistinguishedName `
                    -Members $User.DistinguishedName `
                    -Server $Server `
                    -Confirm:$false `
                    -ErrorAction Stop

                Write-OutputBox "Removed user $Sam from group $($SelectedGroup.Name) using $Server" ([System.Drawing.Color]::DarkOrange)
                $SuccessCount++
            }
            catch {
                Write-OutputBox "ERROR removing $Sam from $($SelectedGroup.Name): $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                $FailureCount++
            }
        }

        Write-OutputBox "Remove User from Group complete. Successful: $SuccessCount | Failed: $FailureCount"
    }
    catch {
        Write-OutputBox "ERROR removing user from group: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Create-ADGroup {
    $GroupName = Get-InputBox "Create AD Group" "Enter new group name:"
    if (-not $GroupName) { return }

    $OUPath = Get-InputBox "Create AD Group" "Enter OU distinguishedName for the group. Leave blank for the default Users container."
    $Description = Get-InputBox "Create AD Group" "Enter group description. Leave blank for no description."

    $Confirm = [System.Windows.Forms.MessageBox]::Show(
        "Create security group $GroupName?",
        "Confirm Create Group",
        "YesNo",
        "Question"
    )

    if ($Confirm -eq "Yes") {
        try {
            $Server = Get-SelectedServer

            $Params = @{
                Name           = $GroupName
                SamAccountName = $GroupName
                GroupScope     = "Global"
                GroupCategory  = "Security"
                Server         = $Server
            }

            if ($OUPath) {
                $Params.Path = $OUPath
            }

            if ($Description) {
                $Params.Description = $Description
            }

            New-ADGroup @Params

            Write-OutputBox "Created AD group: $GroupName using $Server" ([System.Drawing.Color]::Green)
        }
        catch {
            Write-OutputBox "ERROR creating group: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
        }
    }
    else {
        Write-OutputBox "Create group cancelled."
    }
}

function Delete-ADGroup {
    $GroupName = Get-InputBox "Delete AD Group" "Enter group name / sAMAccountName / distinguishedName:"
    if (-not $GroupName) { return }

    $ProtectionRemoved = $false
    $Group = $null
    $Server = Get-SelectedServer

    try {
        $Group = Get-ADGroup `
            -Identity $GroupName `
            -Server $Server `
            -Properties Name,SamAccountName,GroupCategory,GroupScope,Description,WhenCreated,DistinguishedName,ProtectedFromAccidentalDeletion,SID,isCriticalSystemObject `
            -ErrorAction Stop

        $CriticalReason = Get-MSToolkitCriticalGroupReason -Group $Group
        if (Stop-MSToolkitCriticalOperation -ObjectType "Group" -DisplayName "$($Group.Name)" -Operation "Delete group" -Reason $CriticalReason) { return }

        $ConfirmText = @"
You are about to DELETE the following Active Directory group:

Name: $($Group.Name)
sAMAccountName: $($Group.SamAccountName)
Scope: $($Group.GroupScope)
Category: $($Group.GroupCategory)
Protected from Accidental Deletion: $($Group.ProtectedFromAccidentalDeletion)
Created: $($Group.WhenCreated)
Description: $($Group.Description)
Distinguished Name: $($Group.DistinguishedName)

This is permanent unless the object is restored from AD Recycle Bin/backups.

Continue with deletion?
"@

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            $ConfirmText,
            "Confirm Delete AD Group",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($Confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-OutputBox "Delete group cancelled for $($Group.Name)."
            return
        }

        if ($Group.ProtectedFromAccidentalDeletion) {
            $OverrideApproved = Request-MSToolkitAccidentalDeletionOverride `
                -ObjectType "AD group" `
                -DisplayName $Group.Name `
                -DistinguishedName $Group.DistinguishedName `
                -Server $Server

            if (-not $OverrideApproved) { return }
            $ProtectionRemoved = $true
        }

        try {
            Remove-ADGroup `
                -Identity $Group.DistinguishedName `
                -Server $Server `
                -Confirm:$false `
                -ErrorAction Stop

            Write-OutputBox "Deleted AD group: $($Group.Name) using $Server" ([System.Drawing.Color]::Red)
        }
        catch {
            if ($ProtectionRemoved) {
                Restore-MSToolkitAccidentalDeletionProtection -DisplayName $Group.Name -DistinguishedName $Group.DistinguishedName -Server $Server
            }
            throw
        }
    }
    catch {
        Write-OutputBox "ERROR deleting group: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ADGroupInfo {
    $GroupName = Get-InputBox "Get AD Group" "Enter group name / sAMAccountName / distinguishedName:"
    if (-not $GroupName) { return }

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer

        $Group = Get-ADGroup `
            -Identity $GroupName `
            -Server $Server `
            -Properties Description,GroupCategory,GroupScope,ManagedBy,WhenCreated,WhenChanged,DistinguishedName

        $ManagedByDisplay = ""

        if ($Group.ManagedBy) {
            try {
                $ManagedByObject = Get-ADObject `
                    -Identity $Group.ManagedBy `
                    -Server $Server `
                    -Properties DisplayName,SamAccountName,ObjectClass `
                    -ErrorAction Stop

                $ManagedByName = if ($ManagedByObject.DisplayName) { $ManagedByObject.DisplayName } else { $ManagedByObject.Name }

                if ($ManagedByObject.SamAccountName) {
                    $ManagedByDisplay = "$ManagedByName ($($ManagedByObject.SamAccountName)) [$($ManagedByObject.ObjectClass)]"
                }
                else {
                    $ManagedByDisplay = "$ManagedByName [$($ManagedByObject.ObjectClass)]"
                }
            }
            catch {
                $ManagedByDisplay = $Group.ManagedBy
            }
        }

        $Members = @(Get-ADGroupMember -Identity $Group.DistinguishedName -Server $Server -ErrorAction Stop)

        Write-OutputBox "Group found using ${Server}:"
        Write-OutputField "Name" $Group.Name
        Write-OutputField "SamAccountName" $Group.SamAccountName
        Write-OutputField "Scope" $Group.GroupScope
        Write-OutputField "Category" $Group.GroupCategory
        Write-OutputField "Description" $Group.Description
        Write-OutputField "Managed By" $ManagedByDisplay
        Write-OutputField "Member Count" $Members.Count
        Write-OutputField "Created" $Group.WhenCreated
        Write-OutputField "Changed" $Group.WhenChanged
        Write-OutputField "DistinguishedName" $Group.DistinguishedName
    }
    catch {
        Write-OutputBox "ERROR getting group: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ADGroupMembers {
    $GroupName = Get-InputBox "Get Group Members" "Enter group name / sAMAccountName / distinguishedName:"
    if (-not $GroupName) { return }

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer

        $Group = Get-ADGroup `
            -Identity $GroupName `
            -Server $Server `
            -Properties DistinguishedName `
            -ErrorAction Stop

        $Members = @(
            Get-ADGroupMember `
                -Identity $Group.DistinguishedName `
                -Server $Server |
                Sort-Object objectClass,Name
        )

        $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $PathColor = [System.Drawing.Color]::DimGray

        Write-ReadableSectionHeader -Title "Members of $($Group.Name)" -Server $Server -BasePath $Group.DistinguishedName

        if (-not $Members -or $Members.Count -eq 0) {
            Write-OutputBox "Group $($Group.Name) has no members." ([System.Drawing.Color]::DarkOrange)
            Write-OUBlankLine
            return
        }

        foreach ($Member in $Members) {
            $MemberColor = switch ($Member.objectClass.ToString().ToLowerInvariant()) {
                "user"     { [System.Drawing.Color]::RoyalBlue }
                "group"    { [System.Drawing.Color]::DarkViolet }
                "computer" { [System.Drawing.Color]::ForestGreen }
                default    { [System.Drawing.Color]::DarkOrange }
            }

            Write-OutputBox "$($Member.Name) [$($Member.objectClass)]" $MemberColor
            Write-OutputBox "  $($Member.DistinguishedName)" $PathColor
            Write-OUBlankLine
        }

        Write-OutputBox "Total members: $($Members.Count)" $SectionColor
        Write-OUBlankLine
    }
    catch {
        Write-OutputBox "ERROR getting group members: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ADComputerInfo {
    $ComputerName = Get-InputBox "Get AD Computer" "Enter computer name:"
    if (-not $ComputerName) { return }

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer

        $Computer = Get-ADComputer `
            -Identity $ComputerName `
            -Server $Server `
            -Properties DNSHostName,Enabled,Description,LastLogonDate,OperatingSystem,WhenCreated,WhenChanged,DistinguishedName

        Write-OutputBox "Computer found using ${Server}:"
        Write-OutputField "Name" $Computer.Name
        Write-OutputField "DNS Hostname" $Computer.DNSHostName

        if ($Computer.Enabled -eq $true) {
            Write-OutputField "Enabled" "True" ([System.Drawing.Color]::Green) ([System.Drawing.Color]::Green)
        }
        else {
            Write-OutputField "Enabled" "False" ([System.Drawing.Color]::Red) ([System.Drawing.Color]::Red)
        }

        Write-OutputField "Description" $Computer.Description
        Write-OutputField "Last Logon Date" $Computer.LastLogonDate
        Write-OutputField "Operating System" $Computer.OperatingSystem
        Write-OutputField "Created" $Computer.WhenCreated
        Write-OutputField "Changed" $Computer.WhenChanged
        Write-OutputField "DistinguishedName" $Computer.DistinguishedName
    }
    catch {
        Write-OutputBox "ERROR getting computer: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ADComputerOU {
    $ComputerName = Get-InputBox "Get Computer OU" "Enter computer name:"
    if (-not $ComputerName) { return }

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer

        $Computer = Get-ADComputer `
            -Identity $ComputerName `
            -Server $Server `
            -Properties DNSHostName,DistinguishedName,Enabled,Description,OperatingSystem

        $PathColor = [System.Drawing.Color]::DimGray
        $DetailColor = [System.Drawing.Color]::FromArgb(70,70,70)

        # Strip the leading CN= component to get the OU the computer sits in.
        $Parts = [regex]::Split($Computer.DistinguishedName, '(?<!\\),')
        $ParentDN = ($Parts[1..($Parts.Count - 1)] -join ',')

        Write-ReadableSectionHeader -Title "OU Location for $($Computer.Name)" -Server $Server -BasePath ""

        Write-OutputField -Label "Name" -Value $Computer.Name -ValueColor $DetailColor
        Write-OutputField -Label "DNS Hostname" -Value $Computer.DNSHostName -ValueColor $DetailColor
        Write-OutputField -Label "Enabled" -Value ([string]$Computer.Enabled) -ValueColor $DetailColor
        Write-OutputField -Label "Operating System" -Value $Computer.OperatingSystem -ValueColor $DetailColor
        Write-OutputField -Label "Parent OU" -Value $ParentDN -ValueColor $PathColor
        Write-OutputField -Label "Object DN" -Value $Computer.DistinguishedName -ValueColor $PathColor

        $OUNames = Get-MSToolkitOUNamesFromDN -DistinguishedName $Computer.DistinguishedName

        if ($OUNames.Count -gt 0) {
            $TopDown = @($OUNames)
            [array]::Reverse($TopDown)
            Write-OutputField -Label "OU Path" -Value ($TopDown -join " \ ") -ValueColor ([System.Drawing.Color]::DarkCyan)
        }

        Write-OUBlankLine
    }
    catch {
        Write-OutputBox "ERROR getting OU for ${ComputerName}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Enable-ADComputerAccount {
    $ComputerName = Get-InputBox "Enable AD Computer" "Enter computer name:"
    if (-not $ComputerName) { return }

    try {
        $Server = Get-SelectedServer
        $Computer = Get-ADComputer `
            -Identity $ComputerName `
            -Server $Server `
            -Properties DNSHostName,Enabled,DistinguishedName,userAccountControl,PrimaryGroupID,SID,isCriticalSystemObject `
            -ErrorAction Stop

        $CriticalReason = Get-MSToolkitCriticalComputerReason -Computer $Computer
        if (Stop-MSToolkitCriticalOperation -ObjectType "Computer" -DisplayName "$($Computer.Name)" -Operation "Enable computer account" -Reason $CriticalReason) { return }

        if ($Computer.Enabled -eq $true) {
            Write-OutputBox "Computer account $($Computer.Name) is already enabled. No change made." ([System.Drawing.Color]::DarkOrange)
            return
        }

        Enable-ADAccount `
            -Identity $Computer.DistinguishedName `
            -Server $Server

        Write-OutputBox "Enabled computer account: $($Computer.Name) using $Server" ([System.Drawing.Color]::Green)
    }
    catch {
        Write-OutputBox "ERROR enabling computer: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Reset-ADComputerAccount {
    $ComputerName = Get-InputBox "Reset AD Computer Account" "Enter computer name:"
    if (-not $ComputerName) { return }

    try {
        $Server = Get-SelectedServer

        $Computer = Get-ADComputer `
            -Identity $ComputerName `
            -Server $Server `
            -Properties DNSHostName,Enabled,OperatingSystem,Description,WhenCreated,DistinguishedName,PasswordLastSet,userAccountControl,PrimaryGroupID,SID,isCriticalSystemObject `
            -ErrorAction Stop

        $CriticalReason = Get-MSToolkitCriticalComputerReason -Computer $Computer
        if (Stop-MSToolkitCriticalOperation -ObjectType "Computer" -DisplayName "$($Computer.Name)" -Operation "Reset computer account" -Reason $CriticalReason) { return }

        $ConfirmText = @"
You are about to RESET the machine account password for:

Computer Name: $($Computer.Name)
DNS Host Name: $($Computer.DNSHostName)
Enabled: $($Computer.Enabled)
Operating System: $($Computer.OperatingSystem)
Password Last Set: $($Computer.PasswordLastSet)
Created: $($Computer.WhenCreated)
Description: $($Computer.Description)
Distinguished Name: $($Computer.DistinguishedName)

This BREAKS the trust relationship immediately. Nobody will be able to log on
to this computer with domain credentials until it is rejoined to the domain.

Only do this when you are about to rejoin or re-image the machine. To repair a
broken trust without rejoining, run this ON the computer instead:

    Test-ComputerSecureChannel -Repair -Credential (Get-Credential)

Continue with the reset?
"@

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            $ConfirmText,
            "Confirm Reset Computer Account",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($Confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-OutputBox "Reset cancelled for computer $($Computer.Name)."
            return
        }

        # The default machine account password is the computer name in lower case,
        # truncated to 14 characters - the same value a pre-staged account gets.
        $DefaultPassword = $Computer.Name.ToLower()
        if ($DefaultPassword.Length -gt 14) {
            $DefaultPassword = $DefaultPassword.Substring(0,14)
        }

        $SecurePassword = ConvertTo-SecureString $DefaultPassword -AsPlainText -Force

        Set-ADAccountPassword `
            -Identity $Computer.DistinguishedName `
            -Server $Server `
            -Reset `
            -NewPassword $SecurePassword `
            -ErrorAction Stop

        Write-OutputBox "Reset machine account password for: $($Computer.Name) using $Server" ([System.Drawing.Color]::Red)
        Write-OutputBox "The computer must now be rejoined to the domain before anyone can log on to it." ([System.Drawing.Color]::DarkOrange)
    }
    catch {
        Write-OutputBox "ERROR resetting computer account: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Disable-ADComputerAccount {
    $ComputerName = Get-InputBox "Disable AD Computer" "Enter computer name:"
    if (-not $ComputerName) { return }

    try {
        $Server = Get-SelectedServer
        $Computer = Get-ADComputer `
            -Identity $ComputerName `
            -Server $Server `
            -Properties DNSHostName,Enabled,DistinguishedName,userAccountControl,PrimaryGroupID,SID,isCriticalSystemObject `
            -ErrorAction Stop

        $CriticalReason = Get-MSToolkitCriticalComputerReason -Computer $Computer
        if (Stop-MSToolkitCriticalOperation -ObjectType "Computer" -DisplayName "$($Computer.Name)" -Operation "Disable computer account" -Reason $CriticalReason) { return }

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to disable computer $($Computer.Name)?",
            "Confirm Disable Computer",
            "YesNo",
            "Warning"
        )

        if ($Confirm -eq "Yes") {
            Disable-ADAccount `
                -Identity $Computer.DistinguishedName `
                -Server $Server

            Write-OutputBox "Disabled computer account: $($Computer.Name) using $Server" ([System.Drawing.Color]::DarkOrange)
        }
    }
    catch {
        Write-OutputBox "ERROR disabling computer: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Delete-ADComputerAccount {
    $ComputerName = Get-InputBox "Delete AD Computer" "Enter computer name:"
    if (-not $ComputerName) { return }

    $ProtectionRemoved = $false
    $Computer = $null
    $Server = Get-SelectedServer

    try {
        $Computer = Get-ADComputer `
            -Identity $ComputerName `
            -Server $Server `
            -Properties DNSHostName,Enabled,OperatingSystem,Description,WhenCreated,DistinguishedName,ProtectedFromAccidentalDeletion,userAccountControl,PrimaryGroupID,SID,isCriticalSystemObject `
            -ErrorAction Stop

        $CriticalReason = Get-MSToolkitCriticalComputerReason -Computer $Computer
        if (Stop-MSToolkitCriticalOperation -ObjectType "Computer" -DisplayName "$($Computer.Name)" -Operation "Delete computer account" -Reason $CriticalReason) { return }

        $ConfirmText = @"
You are about to DELETE the following Active Directory computer:

Computer Name: $($Computer.Name)
DNS Host Name: $($Computer.DNSHostName)
Enabled: $($Computer.Enabled)
Operating System: $($Computer.OperatingSystem)
Protected from Accidental Deletion: $($Computer.ProtectedFromAccidentalDeletion)
Created: $($Computer.WhenCreated)
Description: $($Computer.Description)
Distinguished Name: $($Computer.DistinguishedName)

This is permanent unless the object is restored from AD Recycle Bin/backups.

Continue with deletion?
"@

        $Confirm = [System.Windows.Forms.MessageBox]::Show(
            $ConfirmText,
            "Confirm Delete AD Computer",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($Confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-OutputBox "Delete cancelled for computer $($Computer.Name)."
            return
        }

        if ($Computer.ProtectedFromAccidentalDeletion) {
            $OverrideApproved = Request-MSToolkitAccidentalDeletionOverride `
                -ObjectType "AD computer" `
                -DisplayName $Computer.Name `
                -DistinguishedName $Computer.DistinguishedName `
                -Server $Server

            if (-not $OverrideApproved) { return }
            $ProtectionRemoved = $true
        }

        try {
            Remove-ADComputer `
                -Identity $Computer.DistinguishedName `
                -Server $Server `
                -Confirm:$false `
                -ErrorAction Stop

            Write-OutputBox "Deleted AD computer: $($Computer.Name) using $Server" ([System.Drawing.Color]::Red)
        }
        catch {
            if ($ProtectionRemoved) {
                Restore-MSToolkitAccidentalDeletionProtection -DisplayName $Computer.Name -DistinguishedName $Computer.DistinguishedName -Server $Server
            }
            throw
        }
    }
    catch {
        Write-OutputBox "ERROR deleting computer: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Write-OUBlankLine {
    if ($OutputBox) {
        $Entry = [pscustomobject]@{ Type = "Blank" }
        Add-OutputEntry -Entry $Entry
        Render-OutputEntry -Entry $Entry
        $OutputBox.ScrollToCaret()
    }
}

function Get-OUNameFromDN {
    param([string]$DistinguishedName)

    if ($DistinguishedName -match '^OU=([^,]+)') {
        return $Matches[1]
    }

    return $DistinguishedName
}


function Get-ReadableNameColors {
    return @(
        [System.Drawing.Color]::RoyalBlue,
        [System.Drawing.Color]::ForestGreen,
        [System.Drawing.Color]::DarkViolet,
        [System.Drawing.Color]::DarkOrange,
        [System.Drawing.Color]::Teal,
        [System.Drawing.Color]::Brown
    )
}

function Write-ReadableSectionHeader {
    param(
        [string]$Title,
        [string]$Server,
        [string]$BasePath,
        # Settings caption of an OU that is not filled in (e.g. "Computers"). When given,
        # the listing is labelled as having fallen back to the domain root.
        [string]$FallbackSetting = ""
    )

    $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $PathColor = [System.Drawing.Color]::DimGray

    if ($FallbackSetting) { $Title = "$Title - domain root" }

    Write-OUBlankLine
    Write-OutputBox "============================================================" $SectionColor
    Write-OutputBox $Title $SectionColor
    Write-OutputBox "AD Server: $Server" $PathColor
    if ($BasePath) {
        if ($FallbackSetting) {
            Write-OutputBox "Base path: $BasePath (domain root)" $PathColor
        }
        else {
            Write-OutputBox "Base path: $BasePath" $PathColor
        }
    }
    if ($FallbackSetting) {
        Write-OutputBox "No $FallbackSetting OU is set in Settings, so this uses the domain root. Set it under Settings > Organizational units." ([System.Drawing.Color]::DarkOrange)
    }
    Write-OutputBox "============================================================" $SectionColor
    Write-OUBlankLine
}

function Get-MSToolkitOUFallbackCaption {
    # The Settings caption to report when an OU setting is blank, or "" when it is set.
    param(
        [string]$Name,
        [string]$Caption
    )

    if (Get-MSToolkitOU -Name $Name) { return "" }
    return $Caption
}

function Write-MSToolkitOUFallbackReminder {
    # Repeats the domain-root note after a long listing, where the output pane stops
    # scrolling, so it is not only at the top where it has scrolled out of view.
    param([string]$FallbackSetting)

    if (-not $FallbackSetting) { return }

    Write-OutputBox "Reminder: no $FallbackSetting OU is set in Settings, so the list above is from the domain root. Set it under Settings > Organizational units." ([System.Drawing.Color]::DarkOrange)
    Write-OUBlankLine
}

function Get-MSToolkitOUPathTopDown {
    param([string]$DistinguishedName)

    $Names = @(Get-MSToolkitOUNamesFromDN -DistinguishedName $DistinguishedName)

    if ($Names.Count -eq 0) {
        return $DistinguishedName
    }

    $TopDown = @($Names)
    [array]::Reverse($TopDown)
    return ($TopDown -join " \ ")
}

function Get-ChildOUs {
    param(
        [string]$BasePath,
        [string]$Label,
        [switch]$Recurse,
        [string]$FallbackSetting = ""
    )

    try {
        $Server = Get-SelectedServer

        $Scope = if ($Recurse) { "Subtree" } else { "OneLevel" }

        $OUs = Get-ADOrganizationalUnit `
            -SearchBase $BasePath `
            -SearchScope $Scope `
            -Filter * `
            -Server $Server `
            -Properties DistinguishedName |
            Sort-Object @{ Expression = { Get-MSToolkitOUPathTopDown -DistinguishedName $_.DistinguishedName } }

        $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $BaseColor = [System.Drawing.Color]::DarkCyan
        $PathColor = [System.Drawing.Color]::DimGray
        $NameColors = Get-ReadableNameColors

        Write-ReadableSectionHeader -Title $Label -Server $Server -BasePath $BasePath -FallbackSetting $FallbackSetting

        # Show the base OU first so the hierarchy is immediately obvious.
        $BaseName = if ($BasePath -match '^OU=') { Get-OUNameFromDN -DistinguishedName $BasePath } else { "Domain root" }
        Write-OutputBox "$BaseName" $BaseColor
        Write-OutputBox "  $BasePath" $PathColor
        Write-OUBlankLine

        if (-not $OUs) {
            Write-OutputBox "No child OUs found under $BaseName." ([System.Drawing.Color]::DarkOrange)
            Write-OUBlankLine
            return
        }

        $ColorIndex = 0
        $BaseDepth = @(Get-MSToolkitOUNamesFromDN -DistinguishedName $BasePath).Count

        foreach ($OU in $OUs) {
            $NameColor = $NameColors[$ColorIndex % $NameColors.Count]

            # Indent each level so the hierarchy is visible in a recursive listing.
            $Depth = @(Get-MSToolkitOUNamesFromDN -DistinguishedName $OU.DistinguishedName).Count - $BaseDepth
            if ($Depth -lt 1) { $Depth = 1 }
            $Indent = "    " * ($Depth - 1)

            Write-OutputBox "$Indent$($OU.Name)" $NameColor
            Write-OutputBox "$Indent  $($OU.DistinguishedName)" $PathColor
            Write-OUBlankLine

            $ColorIndex++
        }

        $TotalLabel = if ($Recurse) { "Total OUs (all levels)" } else { "Total child OUs" }
        Write-OutputBox "${TotalLabel}: $($OUs.Count)" $SectionColor
        Write-OUBlankLine
        Write-MSToolkitOUFallbackReminder -FallbackSetting $FallbackSetting
    }
    catch {
        Write-OutputBox "ERROR getting OUs from ${Label}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ComputerOUs {
    $FallbackCaption = Get-MSToolkitOUFallbackCaption -Name "OUComputers" -Caption "Computers"

    if ($FallbackCaption) {
        # No Computers OU set: list only the top level of the domain root, the same as
        # the other OU buttons, instead of every OU in the domain.
        Get-ChildOUs `
            -Label "Computer OUs" `
            -BasePath (Get-MSToolkitOU -Name "OUComputers" -RootFallback) `
            -FallbackSetting $FallbackCaption
        return
    }

    Get-ChildOUs `
        -Label "Computer OUs (all levels)" `
        -BasePath (Get-MSToolkitOU -Name "OUComputers") `
        -Recurse
}

function Get-EmployeeOUs {
    Get-ChildOUs `
        -Label "User OU" `
        -BasePath (Get-MSToolkitOU -Name "OUUsers" -RootFallback) `
        -FallbackSetting (Get-MSToolkitOUFallbackCaption -Name "OUUsers" -Caption "Users")
}

function Get-DisabledOUs {
    Get-ChildOUs `
        -Label "Disabled OU" `
        -BasePath (Get-MSToolkitOU -Name "OUDisabled" -RootFallback) `
        -FallbackSetting (Get-MSToolkitOUFallbackCaption -Name "OUDisabled" -Caption "Disabled accounts")
}

function Get-TopLevelOUs {
    try {
        $Server = Get-SelectedServer
        $DomainRoot = (Get-ADDomain -Server $Server -ErrorAction Stop).DistinguishedName

        Get-ChildOUs -Label "Top-Level OUs" -BasePath $DomainRoot
    }
    catch {
        Write-OutputBox "ERROR listing top-level OUs: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-ServerOUs {
    Get-ChildOUs `
        -Label "Server OU" `
        -BasePath (Get-MSToolkitOU -Name "OUServers" -RootFallback) `
        -FallbackSetting (Get-MSToolkitOUFallbackCaption -Name "OUServers" -Caption "Servers")
}

function Show-MSToolkitOUChooser {
    param(
        [string[]]$Candidates,
        [string]$Prompt
    )

    $ChooserForm = New-Object System.Windows.Forms.Form
    $ChooserForm.Text = "Select Organizational Unit"
    $ChooserForm.Size = New-Object System.Drawing.Size(760,360)
    $ChooserForm.StartPosition = "CenterScreen"
    $ChooserForm.FormBorderStyle = "FixedDialog"
    $ChooserForm.MaximizeBox = $false
    $ChooserForm.MinimizeBox = $false

    $PromptLabel = New-Object System.Windows.Forms.Label
    $PromptLabel.Text = $Prompt
    $PromptLabel.Location = New-Object System.Drawing.Point(12,12)
    $PromptLabel.Size = New-Object System.Drawing.Size(720,20)
    $ChooserForm.Controls.Add($PromptLabel)

    $ChooserList = New-Object System.Windows.Forms.ListBox
    $ChooserList.Location = New-Object System.Drawing.Point(12,38)
    $ChooserList.Size = New-Object System.Drawing.Size(720,215)
    foreach ($Candidate in $Candidates) {
        $null = $ChooserList.Items.Add($Candidate)
    }
    $ChooserList.SelectedIndex = 0
    $ChooserForm.Controls.Add($ChooserList)

    $OkButton = New-Object System.Windows.Forms.Button
    $OkButton.Text = "Use This OU"
    $OkButton.Location = New-Object System.Drawing.Point(512,268)
    $OkButton.Size = New-Object System.Drawing.Size(110,30)
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ChooserForm.Controls.Add($OkButton)
    $ChooserForm.AcceptButton = $OkButton

    $CancelBtn = New-Object System.Windows.Forms.Button
    $CancelBtn.Text = "Cancel"
    $CancelBtn.Location = New-Object System.Drawing.Point(632,268)
    $CancelBtn.Size = New-Object System.Drawing.Size(100,30)
    $CancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $ChooserForm.Controls.Add($CancelBtn)
    $ChooserForm.CancelButton = $CancelBtn

    Apply-ThemeToControl -Control $ChooserForm
    $Result = $ChooserForm.ShowDialog()

    if ($Result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return [string]$ChooserList.SelectedItem
}

function Get-MSToolkitOUNamesFromDN {
    param([string]$DistinguishedName)

    # Split on commas that are not backslash-escaped, so an OU name such as
    # An OU name containing a comma, such as "Example Company, Inc", survives intact.
    $Parts = [regex]::Split($DistinguishedName, '(?<!\\),')
    $Names = New-Object System.Collections.Generic.List[string]

    foreach ($Part in $Parts) {
        if ($Part -match '^\s*OU=(.+)$') {
            $Names.Add(($Matches[1] -replace '\\,', ','))
        }
    }

    # Index 0 is the OU itself, then each ancestor working outward.
    return $Names.ToArray()
}

function Test-MSToolkitOUAncestorMatch {
    param(
        [string[]]$AncestorsTopDown,
        [string[]]$RequiredTopDown
    )

    # Every required name must appear in order, but levels may be skipped so
    # "Employees\IT" still matches OU=IT,OU=Staff,OU=Employees.
    $Index = 0

    foreach ($Required in $RequiredTopDown) {
        $Matched = $false

        while ($Index -lt $AncestorsTopDown.Count) {
            if ($AncestorsTopDown[$Index] -ieq $Required) {
                $Matched = $true
                $Index++
                break
            }

            $Index++
        }

        if (-not $Matched) {
            return $false
        }
    }

    return $true
}

function Resolve-MSToolkitOUPath {
    param(
        [string]$NameOrDN,
        [string]$Server
    )

    $Value = "$NameOrDN".Trim()

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    # Already a distinguished name - use it as entered, preserving any escaping.
    if ($Value -match '^(OU|CN|DC)=') {
        try {
            $Exact = Get-ADObject -Identity $Value -Server $Server -ErrorAction Stop
            return $Exact.DistinguishedName
        }
        catch {
            Write-OutputBox "ERROR: '$Value' could not be resolved on ${Server}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
            return $null
        }
    }

    # Accept a friendly path such as Employees\IT, Employees/IT, Employees | IT,
    # Employees ; IT, Employees : IT or Employees > IT. Hyphens are never
    # separators because OU names such as "Servers - No Policy" contain them.
    $Segments = @(
        [regex]::Split($Value, '\s*[\\/|;:>]\s*') |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($Segments.Count -eq 0) {
        return $null
    }

    $LeafName = $Segments[$Segments.Count - 1]
    $RequiredAncestors = @()

    if ($Segments.Count -gt 1) {
        $RequiredAncestors = @($Segments[0..($Segments.Count - 2)])
    }

    $Found = @()
    try {
        $Found = @(
            Get-ADOrganizationalUnit -Filter { Name -eq $LeafName } -Server $Server -ErrorAction Stop |
            Select-Object -ExpandProperty DistinguishedName |
            Sort-Object
        )
    }
    catch {
        Write-OutputBox "ERROR searching for OU '$LeafName' on ${Server}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
        return $null
    }

    if ($Found.Count -eq 0) {
        Write-OutputBox "ERROR: No organizational unit named '$LeafName' was found on $Server." ([System.Drawing.Color]::Red)
        return $null
    }

    # Narrow by the parent names the admin supplied, if any.
    if ($RequiredAncestors.Count -gt 0) {
        $Narrowed = New-Object System.Collections.Generic.List[string]

        foreach ($Candidate in $Found) {
            $OUNames = Get-MSToolkitOUNamesFromDN -DistinguishedName $Candidate

            $AncestorsTopDown = @()
            if ($OUNames.Count -gt 1) {
                $AncestorsTopDown = @($OUNames[1..($OUNames.Count - 1)])
                [array]::Reverse($AncestorsTopDown)
            }

            if (Test-MSToolkitOUAncestorMatch -AncestorsTopDown $AncestorsTopDown -RequiredTopDown $RequiredAncestors) {
                $Narrowed.Add($Candidate)
            }
        }

        if ($Narrowed.Count -eq 0) {
            Write-OutputBox "ERROR: Found $($Found.Count) OU(s) named '$LeafName' on $Server, but none of them sit under '$($RequiredAncestors -join '\')'." ([System.Drawing.Color]::Red)
            return $null
        }

        $Found = @($Narrowed.ToArray())
    }

    if ($Found.Count -eq 1) {
        return $Found[0]
    }

    Write-OutputBox "'$Value' matches $($Found.Count) organizational units. Select the one you want." ([System.Drawing.Color]::DarkOrange)
    return (Show-MSToolkitOUChooser -Candidates $Found -Prompt "More than one OU matches '$Value'. Choose the intended location:")
}

# ---------------------------------------------------------------------------
# Cascading OU picker. Level 0 lists the domain root and its top-level OUs;
# choosing one adds another dropdown for its children, and so on. The current
# selection is tracked in $script:OUPickerCurrentDN.
# ---------------------------------------------------------------------------
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

    $ParentName = if ($Level -eq 0) { "domain root" } else { Get-OUNameFromDN -DistinguishedName $ParentDN }

    $Combo = New-Object System.Windows.Forms.ComboBox
    $Combo.DropDownStyle = "DropDownList"
    $Combo.Width = 215
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

        if ($script:OUPickerPathLabel) {
            $script:OUPickerPathLabel.Text = "Selected: $script:OUPickerCurrentDN"
        }

        Apply-ThemeToControl -Control $script:OUPickerFlow
    })

    $script:OUPickerFlow.Controls.Add($Combo)
}

function Initialize-MSToolkitOUPicker {
    param(
        [System.Windows.Forms.FlowLayoutPanel]$Flow,
        [System.Windows.Forms.Label]$PathLabel,
        [string]$Server,
        [string]$DomainRoot
    )

    $script:OUPickerFlow = $Flow
    $script:OUPickerPathLabel = $PathLabel
    $script:OUPickerServer = $Server
    $script:OUPickerCurrentDN = $DomainRoot

    $Flow.Controls.Clear()
    Add-MSToolkitOUPickerLevel -ParentDN $DomainRoot -Level 0

    if ($PathLabel) {
        $PathLabel.Text = "Selected: $DomainRoot"
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

function Resolve-MSToolkitOUSelection {
    param(
        [string]$BaseDN,
        [string]$TypedValue,
        [string]$Server
    )

    $Typed = "$TypedValue".Trim()

    # Nothing typed - use whatever the dropdowns are pointing at.
    if ([string]::IsNullOrWhiteSpace($Typed)) {
        return $BaseDN
    }

    # A full DN always overrides the dropdown selection.
    if ($Typed -match '^(OU|CN|DC)=') {
        return (Resolve-MSToolkitOUPath -NameOrDN $Typed -Server $Server)
    }

    # Otherwise treat it as a path underneath the current selection.
    $Relative = Resolve-MSToolkitOUPathUnder -BaseDN $BaseDN -RelativePath $Typed -Server $Server

    if (-not [string]::IsNullOrWhiteSpace($Relative)) {
        return $Relative
    }

    # Fall back to a domain-wide search so plain names still work.
    Write-OutputBox "'$Typed' was not found under the selected OU. Searching the whole domain." ([System.Drawing.Color]::DarkOrange)
    return (Resolve-MSToolkitOUPath -NameOrDN $Typed -Server $Server)
}

function Show-CreateOU {
    try {
        $Server = Get-SelectedServer

        try {
            $DomainRoot = (Get-ADDomain -Server $Server -ErrorAction Stop).DistinguishedName
        }
        catch {
            Write-OutputBox "ERROR resolving the domain root from ${Server}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
            return
        }

        $CreateForm = New-Object System.Windows.Forms.Form
        $CreateForm.Text = "Create Organizational Unit"
        $CreateForm.Size = New-Object System.Drawing.Size(720,420)
        $CreateForm.StartPosition = "CenterScreen"
        $CreateForm.FormBorderStyle = "FixedDialog"
        $CreateForm.MaximizeBox = $false
        $CreateForm.MinimizeBox = $false

        $NameLabel = New-Object System.Windows.Forms.Label
        $NameLabel.Text = "New OU name:"
        $NameLabel.Location = New-Object System.Drawing.Point(12,15)
        $NameLabel.Size = New-Object System.Drawing.Size(680,20)
        $CreateForm.Controls.Add($NameLabel)

        $NameBox = New-Object System.Windows.Forms.TextBox
        $NameBox.Location = New-Object System.Drawing.Point(12,38)
        $NameBox.Size = New-Object System.Drawing.Size(680,22)
        $CreateForm.Controls.Add($NameBox)

        $PickerLabel = New-Object System.Windows.Forms.Label
        $PickerLabel.Text = "Parent OU - start at the domain root and drill down. Each selection adds the next level."
        $PickerLabel.Location = New-Object System.Drawing.Point(12,75)
        $PickerLabel.Size = New-Object System.Drawing.Size(680,20)
        $CreateForm.Controls.Add($PickerLabel)

        $PickerFlow = New-Object System.Windows.Forms.FlowLayoutPanel
        $PickerFlow.Location = New-Object System.Drawing.Point(12,98)
        $PickerFlow.Size = New-Object System.Drawing.Size(680,100)
        $PickerFlow.AutoScroll = $true
        $PickerFlow.FlowDirection = "LeftToRight"
        $PickerFlow.WrapContents = $true
        $PickerFlow.Tag = "PickerFlow"
        $CreateForm.Controls.Add($PickerFlow)

        $PathLabel = New-Object System.Windows.Forms.Label
        $PathLabel.Location = New-Object System.Drawing.Point(12,202)
        $PathLabel.Size = New-Object System.Drawing.Size(680,20)
        $CreateForm.Controls.Add($PathLabel)

        $TypedLabel = New-Object System.Windows.Forms.Label
        $TypedLabel.Text = "Optional - type a path under the selection above (IT\Test), or a full DN to override it:"
        $TypedLabel.Location = New-Object System.Drawing.Point(12,230)
        $TypedLabel.Size = New-Object System.Drawing.Size(680,20)
        $CreateForm.Controls.Add($TypedLabel)

        $TypedBox = New-Object System.Windows.Forms.TextBox
        $TypedBox.Location = New-Object System.Drawing.Point(12,253)
        $TypedBox.Size = New-Object System.Drawing.Size(680,22)
        $CreateForm.Controls.Add($TypedBox)

        $OkButton = New-Object System.Windows.Forms.Button
        $OkButton.Text = "Continue"
        $OkButton.Location = New-Object System.Drawing.Point(482,320)
        $OkButton.Size = New-Object System.Drawing.Size(100,30)
        $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CreateForm.Controls.Add($OkButton)
        $CreateForm.AcceptButton = $OkButton

        $CancelBtn = New-Object System.Windows.Forms.Button
        $CancelBtn.Text = "Cancel"
        $CancelBtn.Location = New-Object System.Drawing.Point(592,320)
        $CancelBtn.Size = New-Object System.Drawing.Size(100,30)
        $CancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $CreateForm.Controls.Add($CancelBtn)
        $CreateForm.CancelButton = $CancelBtn

        Initialize-MSToolkitOUPicker -Flow $PickerFlow -PathLabel $PathLabel -Server $Server -DomainRoot $DomainRoot

        $CreateForm.Add_Shown({ $NameBox.Focus() })

        Apply-ThemeToControl -Control $CreateForm
        $DialogResult = $CreateForm.ShowDialog()

        if ($DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $OUName = "$($NameBox.Text)".Trim()
        $SelectedBase = $script:OUPickerCurrentDN

        if ([string]::IsNullOrWhiteSpace($OUName)) {
            Write-OutputBox "ERROR: No OU name entered." ([System.Drawing.Color]::Red)
            return
        }

        $ParentPath = Resolve-MSToolkitOUSelection -BaseDN $SelectedBase -TypedValue $TypedBox.Text -Server $Server

        if ([string]::IsNullOrWhiteSpace($ParentPath)) {
            return
        }

        # The new OU name becomes an RDN, so a comma in the name must stay escaped in the DN.
        $EscapedName = $OUName -replace ',', '\,'
        $NewDN = "OU=$EscapedName,$ParentPath"

        $Existing = $null
        try {
            $Existing = Get-ADOrganizationalUnit -Identity $NewDN -Server $Server -ErrorAction SilentlyContinue
        }
        catch { }

        if ($Existing) {
            Write-OutputBox "ERROR: An OU already exists at $NewDN" ([System.Drawing.Color]::Red)
            return
        }

        $ConfirmMessage = @"
Create this organizational unit?

Name: $OUName

Parent:
$ParentPath

Resulting DN:
$NewDN

The new OU will be created with Protect object from accidental deletion enabled. No Group Policy objects are linked to it by default.
"@

        $Answer = [System.Windows.Forms.MessageBox]::Show(
            $ConfirmMessage,
            "Confirm Create OU",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-OutputBox "OU creation cancelled." ([System.Drawing.Color]::DarkOrange)
            return
        }

        try {
            New-ADOrganizationalUnit `
                -Name $OUName `
                -Path $ParentPath `
                -Server $Server `
                -ProtectedFromAccidentalDeletion $true `
                -ErrorAction Stop

            $Created = Get-ADOrganizationalUnit -Identity $NewDN -Server $Server -Properties ProtectedFromAccidentalDeletion -ErrorAction SilentlyContinue

            Write-ResultSeparator
            Write-OutputBox "Organizational unit '$OUName' created successfully." ([System.Drawing.Color]::Green)
            Write-OutputField -Label "Distinguished Name" -Value $(if ($Created) { $Created.DistinguishedName } else { $NewDN })
            Write-OutputField -Label "Protected from accidental deletion" -Value $(if ($Created) { [string]$Created.ProtectedFromAccidentalDeletion } else { "True" })
            Write-OutputField -Label "Created using" -Value $Server
        }
        catch {
            Write-OutputBox "ERROR creating OU '$OUName': $($_.Exception.Message)" ([System.Drawing.Color]::Red)
        }
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Show-MoveObjectToOU {
    try {
        $Server = Get-SelectedServer

        try {
            $DomainRoot = (Get-ADDomain -Server $Server -ErrorAction Stop).DistinguishedName
        }
        catch {
            Write-OutputBox "ERROR resolving the domain root from ${Server}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
            return
        }

        $MoveForm = New-Object System.Windows.Forms.Form
        $MoveForm.Text = "Move Object to OU"
        $MoveForm.Size = New-Object System.Drawing.Size(720,440)
        $MoveForm.StartPosition = "CenterScreen"
        $MoveForm.FormBorderStyle = "FixedDialog"
        $MoveForm.MaximizeBox = $false
        $MoveForm.MinimizeBox = $false

        $NameLabel = New-Object System.Windows.Forms.Label
        $NameLabel.Text = "Object to move (sAMAccountName, user name, computer name, or group name):"
        $NameLabel.Location = New-Object System.Drawing.Point(12,15)
        $NameLabel.Size = New-Object System.Drawing.Size(680,20)
        $MoveForm.Controls.Add($NameLabel)

        $NameBox = New-Object System.Windows.Forms.TextBox
        $NameBox.Location = New-Object System.Drawing.Point(12,38)
        $NameBox.Size = New-Object System.Drawing.Size(680,22)
        $MoveForm.Controls.Add($NameBox)

        $PickerLabel = New-Object System.Windows.Forms.Label
        $PickerLabel.Text = "Target OU - start at the domain root and drill down. Each selection adds the next level."
        $PickerLabel.Location = New-Object System.Drawing.Point(12,75)
        $PickerLabel.Size = New-Object System.Drawing.Size(680,20)
        $MoveForm.Controls.Add($PickerLabel)

        $PickerFlow = New-Object System.Windows.Forms.FlowLayoutPanel
        $PickerFlow.Location = New-Object System.Drawing.Point(12,98)
        $PickerFlow.Size = New-Object System.Drawing.Size(680,100)
        $PickerFlow.AutoScroll = $true
        $PickerFlow.FlowDirection = "LeftToRight"
        $PickerFlow.WrapContents = $true
        $PickerFlow.Tag = "PickerFlow"
        $MoveForm.Controls.Add($PickerFlow)

        $PathLabel = New-Object System.Windows.Forms.Label
        $PathLabel.Location = New-Object System.Drawing.Point(12,202)
        $PathLabel.Size = New-Object System.Drawing.Size(680,20)
        $MoveForm.Controls.Add($PathLabel)

        $TypedLabel = New-Object System.Windows.Forms.Label
        $TypedLabel.Text = "Optional - type a path under the selection above (IT\Test), or a full DN to override it:"
        $TypedLabel.Location = New-Object System.Drawing.Point(12,230)
        $TypedLabel.Size = New-Object System.Drawing.Size(680,20)
        $MoveForm.Controls.Add($TypedLabel)

        $TypedBox = New-Object System.Windows.Forms.TextBox
        $TypedBox.Location = New-Object System.Drawing.Point(12,253)
        $TypedBox.Size = New-Object System.Drawing.Size(680,22)
        $MoveForm.Controls.Add($TypedBox)

        $NoteLabel = New-Object System.Windows.Forms.Label
        $NoteLabel.Text = "Moving an object changes which Group Policy objects apply to it. Critical directory objects are blocked."
        $NoteLabel.Location = New-Object System.Drawing.Point(12,285)
        $NoteLabel.Size = New-Object System.Drawing.Size(680,36)
        $MoveForm.Controls.Add($NoteLabel)

        $OkButton = New-Object System.Windows.Forms.Button
        $OkButton.Text = "Continue"
        $OkButton.Location = New-Object System.Drawing.Point(482,340)
        $OkButton.Size = New-Object System.Drawing.Size(100,30)
        $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $MoveForm.Controls.Add($OkButton)
        $MoveForm.AcceptButton = $OkButton

        $CancelBtn = New-Object System.Windows.Forms.Button
        $CancelBtn.Text = "Cancel"
        $CancelBtn.Location = New-Object System.Drawing.Point(592,340)
        $CancelBtn.Size = New-Object System.Drawing.Size(100,30)
        $CancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $MoveForm.Controls.Add($CancelBtn)
        $MoveForm.CancelButton = $CancelBtn

        Initialize-MSToolkitOUPicker -Flow $PickerFlow -PathLabel $PathLabel -Server $Server -DomainRoot $DomainRoot

        $MoveForm.Add_Shown({ $NameBox.Focus() })

        Apply-ThemeToControl -Control $MoveForm
        $DialogResult = $MoveForm.ShowDialog()

        if ($DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        $ObjectName = "$($NameBox.Text)".Trim()
        $SelectedBase = $script:OUPickerCurrentDN

        if ([string]::IsNullOrWhiteSpace($ObjectName)) {
            Write-OutputBox "ERROR: No object name entered." ([System.Drawing.Color]::Red)
            return
        }

        # Resolve the real directory object before doing anything with it.
        $Resolved = $null
        $ObjectType = $null

        try {
            $Resolved = Get-ADUser -Identity $ObjectName -Server $Server -Properties Name,SamAccountName,Description,DistinguishedName,ProtectedFromAccidentalDeletion,Enabled,SID,userAccountControl,isCriticalSystemObject -ErrorAction Stop
            $ObjectType = "User"
        }
        catch {
            try {
                $Resolved = Get-ADComputer -Identity $ObjectName -Server $Server -Properties Name,SamAccountName,Description,DistinguishedName,ProtectedFromAccidentalDeletion,Enabled,SID,userAccountControl,PrimaryGroupID,isCriticalSystemObject -ErrorAction Stop
                $ObjectType = "Computer"
            }
            catch {
                try {
                    $Resolved = Get-ADGroup -Identity $ObjectName -Server $Server -Properties Name,SamAccountName,Description,DistinguishedName,ProtectedFromAccidentalDeletion,SID,isCriticalSystemObject -ErrorAction Stop
                    $ObjectType = "Group"
                }
                catch {
                    Write-OutputBox "ERROR: Could not find a user, computer, or group named '$ObjectName' on $Server." ([System.Drawing.Color]::Red)
                    return
                }
            }
        }

        # Hard block on critical directory objects, before any confirmation.
        $Reason = switch ($ObjectType) {
            "User"     { Get-MSToolkitCriticalUserReason -User $Resolved }
            "Computer" { Get-MSToolkitCriticalComputerReason -Computer $Resolved }
            "Group"    { Get-MSToolkitCriticalGroupReason -Group $Resolved }
        }

        if (Stop-MSToolkitCriticalOperation -ObjectType $ObjectType -DisplayName $Resolved.Name -Operation "Move object to another OU" -Reason $Reason) {
            return
        }

        $TargetOU = Resolve-MSToolkitOUSelection -BaseDN $SelectedBase -TypedValue $TypedBox.Text -Server $Server

        if ([string]::IsNullOrWhiteSpace($TargetOU)) {
            return
        }

        # Compare the object's immediate parent, not a DN suffix - a suffix match
        # would treat anything nested deeper under the target as "already there".
        $CurrentParts = [regex]::Split($Resolved.DistinguishedName, '(?<!\\),')
        $CurrentParent = ($CurrentParts[1..($CurrentParts.Count - 1)] -join ',')

        if ($CurrentParent -ieq $TargetOU) {
            Write-OutputBox "$ObjectType '$($Resolved.Name)' is already located directly in $TargetOU. No move performed." ([System.Drawing.Color]::DarkOrange)
            return
        }

        $ProtectedText = if ($Resolved.ProtectedFromAccidentalDeletion) { "True" } else { "False" }

        $ConfirmMessage = @"
Move this $($ObjectType.ToLower()) to a different organizational unit?

Name: $($Resolved.Name)
sAMAccountName: $($Resolved.SamAccountName)
Description: $($Resolved.Description)
Protected from accidental deletion: $ProtectedText

Current location:
$($Resolved.DistinguishedName)

New location:
$TargetOU

Moving an object changes which Group Policy objects apply to it and can affect delegated permissions that are linked to its current OU.
"@

        $Answer = [System.Windows.Forms.MessageBox]::Show(
            $ConfirmMessage,
            "Confirm Move",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-OutputBox "Move cancelled for $($Resolved.Name)." ([System.Drawing.Color]::DarkOrange)
            return
        }

        try {
            Move-ADObject -Identity $Resolved.DistinguishedName -TargetPath $TargetOU -Server $Server -ErrorAction Stop

            $Moved = Get-ADObject -Identity ("CN=" + $Resolved.Name + "," + $TargetOU) -Server $Server -ErrorAction SilentlyContinue

            Write-ResultSeparator
            Write-OutputBox "$ObjectType '$($Resolved.Name)' moved successfully." ([System.Drawing.Color]::Green)
            Write-OutputField -Label "Previous DN" -Value $Resolved.DistinguishedName
            Write-OutputField -Label "New DN" -Value $(if ($Moved) { $Moved.DistinguishedName } else { "$TargetOU (verify in ADUC)" })
        }
        catch {
            if ($Resolved.ProtectedFromAccidentalDeletion) {
                Write-OutputBox "ERROR moving $($Resolved.Name): $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                Write-OutputBox "This object is protected from accidental deletion, which also blocks moves. Clear that setting in ADUC if the move is intended." ([System.Drawing.Color]::DarkOrange)
            }
            else {
                Write-OutputBox "ERROR moving $($Resolved.Name): $($_.Exception.Message)" ([System.Drawing.Color]::Red)
            }
        }
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-MSToolkitSecurityGroups {
    $BasePath = Get-MSToolkitOU -Name "OUSecurityGroups" -RootFallback

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer
        $Groups = @(
            Get-ADGroup `
                -SearchBase $BasePath `
                -SearchScope Subtree `
                -Filter * `
                -Server $Server `
                -Properties GroupCategory,GroupScope,Description,DistinguishedName |
                Where-Object { $_.GroupCategory -eq 'Security' } |
                Sort-Object Name
        )

        $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $PathColor = [System.Drawing.Color]::DimGray
        $DetailColor = [System.Drawing.Color]::FromArgb(70,70,70)
        $NameColors = Get-ReadableNameColors

        $FallbackCaption = Get-MSToolkitOUFallbackCaption -Name "OUSecurityGroups" -Caption "Security groups"
        Write-ReadableSectionHeader -Title "Security Groups" -Server $Server -BasePath $BasePath -FallbackSetting $FallbackCaption

        if (-not $Groups -or $Groups.Count -eq 0) {
            Write-OutputBox "No security groups found." ([System.Drawing.Color]::DarkOrange)
            Write-OUBlankLine
            return
        }

        $ColorIndex = 0
        foreach ($Group in $Groups) {
            $NameColor = $NameColors[$ColorIndex % $NameColors.Count]

            Write-OutputBox $Group.Name $NameColor
            Write-OutputBox "  sAMAccountName: $($Group.SamAccountName)" $DetailColor
            Write-OutputBox "  Scope: $($Group.GroupScope) | Category: $($Group.GroupCategory)" $DetailColor
            if ($Group.Description) {
                Write-OutputBox "  Description: $($Group.Description)" $DetailColor
            }
            Write-OutputBox "  $($Group.DistinguishedName)" $PathColor
            Write-OUBlankLine

            $ColorIndex++
        }

        Write-OutputBox "Total security groups: $($Groups.Count)" $SectionColor
        Write-OUBlankLine
        Write-MSToolkitOUFallbackReminder -FallbackSetting $FallbackCaption
    }
    catch {
        Write-OutputBox "ERROR getting Security Groups: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-MSToolkitDistributionGroups {
    # Distribution groups as AD sees them: created in AD for on-premises Exchange, or
    # synced to Microsoft 365. Read-only. Groups created in Exchange Online never exist
    # in AD, so an empty result points to the cloud tool instead.
    $BasePath = Get-MSToolkitOU -Name "OUDistributionGroups" -RootFallback

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer
        $Groups = @(
            Get-ADGroup `
                -SearchBase $BasePath `
                -SearchScope Subtree `
                -Filter 'GroupCategory -eq "Distribution"' `
                -Server $Server `
                -Properties GroupCategory,GroupScope,Description,mail,DistinguishedName |
                Sort-Object Name
        )

        $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $PathColor = [System.Drawing.Color]::DimGray
        $DetailColor = [System.Drawing.Color]::FromArgb(70,70,70)
        $NameColors = Get-ReadableNameColors

        $FallbackCaption = Get-MSToolkitOUFallbackCaption -Name "OUDistributionGroups" -Caption "Distribution groups"
        Write-ReadableSectionHeader -Title "Distribution Groups" -Server $Server -BasePath $BasePath -FallbackSetting $FallbackCaption

        if (-not $Groups -or $Groups.Count -eq 0) {
            Write-OutputBox "No distribution groups were found in AD here." ([System.Drawing.Color]::DarkOrange)
            Write-OutputBox "Distribution groups created in Exchange Online exist only in the cloud, not in AD. Use M365 Distro Compare/Add for those." ([System.Drawing.Color]::DimGray)
            Write-OUBlankLine
            return
        }

        $ColorIndex = 0
        foreach ($Group in $Groups) {
            $NameColor = $NameColors[$ColorIndex % $NameColors.Count]

            Write-OutputBox $Group.Name $NameColor
            Write-OutputBox "  sAMAccountName: $($Group.SamAccountName)" $DetailColor
            Write-OutputBox "  Scope: $($Group.GroupScope) | Category: $($Group.GroupCategory)" $DetailColor
            if ($Group.mail) {
                Write-OutputBox "  Email: $($Group.mail)" $DetailColor
            }
            else {
                Write-OutputBox "  Email: (no mail attribute - not mail-enabled from AD)" $DetailColor
            }
            if ($Group.Description) {
                Write-OutputBox "  Description: $($Group.Description)" $DetailColor
            }
            Write-OutputBox "  $($Group.DistinguishedName)" $PathColor
            Write-OUBlankLine

            $ColorIndex++
        }

        Write-OutputBox "Total distribution groups: $($Groups.Count)" $SectionColor
        Write-OUBlankLine
        Write-MSToolkitOUFallbackReminder -FallbackSetting $FallbackCaption
    }
    catch {
        Write-OutputBox "ERROR getting Distribution Groups: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-MSToolkitManagedServiceAccounts {
    $BasePath = "CN=Managed Service Accounts," + (Get-MSToolkitDomainRootDN)

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer
        $Accounts = @(
            Get-ADServiceAccount `
                -Filter * `
                -SearchBase $BasePath `
                -SearchScope OneLevel `
                -Server $Server `
                -Properties Description,DNSHostName,Enabled,DistinguishedName |
                Sort-Object Name
        )

        $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $PathColor = [System.Drawing.Color]::DimGray
        $DetailColor = [System.Drawing.Color]::FromArgb(70,70,70)
        $NameColors = Get-ReadableNameColors

        Write-ReadableSectionHeader -Title "Managed Service Accounts" -Server $Server -BasePath $BasePath

        if (-not $Accounts -or $Accounts.Count -eq 0) {
            Write-OutputBox "No managed service accounts found." ([System.Drawing.Color]::DarkOrange)
            Write-OUBlankLine
            return
        }

        $ColorIndex = 0
        foreach ($Account in $Accounts) {
            $NameColor = $NameColors[$ColorIndex % $NameColors.Count]

            Write-OutputBox $Account.Name $NameColor
            Write-OutputField "sAMAccountName" $Account.SamAccountName $DetailColor
            if ($Account.Enabled -eq $true) {
                Write-OutputField "Enabled" "True" ([System.Drawing.Color]::Green) ([System.Drawing.Color]::Green)
            }
            else {
                Write-OutputField "Enabled" "False" ([System.Drawing.Color]::Red) ([System.Drawing.Color]::Red)
            }
            if ($Account.DNSHostName) {
                Write-OutputField "DNS Host Name" $Account.DNSHostName $DetailColor
            }
            if ($Account.Description) {
                Write-OutputField "Description" $Account.Description $DetailColor
            }
            Write-OutputBox "  $($Account.DistinguishedName)" $PathColor
            Write-OUBlankLine

            $ColorIndex++
        }

        Write-OutputBox "Total managed service accounts: $($Accounts.Count)" $SectionColor
        Write-OUBlankLine
    }
    catch {
        Write-OutputBox "ERROR getting Managed Service Accounts: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-MSToolkitSpecialFunctionAccounts {
    # This lists user accounts directly inside one OU, so the domain root is no useful
    # stand-in: nothing normally sits there. Ask for the setting instead.
    if (-not (Get-MSToolkitOU -Name "OUServiceAccounts")) {
        Write-ResultSeparator
        Write-OutputBox "Get Special Function Accounts needs the Service accounts OU. It is not set in Settings." ([System.Drawing.Color]::DarkOrange)
        Write-OutputBox "Set it under Settings > Organizational units > Service accounts, then try again." ([System.Drawing.Color]::DimGray)
        return
    }

    $BasePath = Get-MSToolkitOU -Name "OUServiceAccounts" -RootFallback

    Write-ResultSeparator

    try {
        $Server = Get-SelectedServer
        $Accounts = @(
            Get-ADUser `
                -SearchBase $BasePath `
                -SearchScope OneLevel `
                -Filter * `
                -Server $Server `
                -Properties DisplayName,Enabled,Description,LastLogonDate,DistinguishedName |
                Sort-Object Name
        )

        $SectionColor = [System.Drawing.Color]::FromArgb(31,58,93)
        $PathColor = [System.Drawing.Color]::DimGray
        $DetailColor = [System.Drawing.Color]::FromArgb(70,70,70)
        $NameColors = Get-ReadableNameColors

        Write-ReadableSectionHeader -Title "Special Function Accounts" -Server $Server -BasePath $BasePath

        if (-not $Accounts -or $Accounts.Count -eq 0) {
            Write-OutputBox "No special function accounts found." ([System.Drawing.Color]::DarkOrange)
            Write-OUBlankLine
            return
        }

        $ColorIndex = 0
        foreach ($Account in $Accounts) {
            $NameColor = $NameColors[$ColorIndex % $NameColors.Count]
            $DisplayName = if ($Account.DisplayName) { $Account.DisplayName } else { $Account.Name }

            Write-OutputBox $DisplayName $NameColor
            Write-OutputField "sAMAccountName" $Account.SamAccountName $DetailColor
            if ($Account.Enabled -eq $true) {
                Write-OutputField "Enabled" "True" ([System.Drawing.Color]::Green) ([System.Drawing.Color]::Green)
            }
            else {
                Write-OutputField "Enabled" "False" ([System.Drawing.Color]::Red) ([System.Drawing.Color]::Red)
            }
            Write-OutputField "Last Logon Date" $Account.LastLogonDate $DetailColor
            if ($Account.Description) {
                Write-OutputField "Description" $Account.Description $DetailColor
            }
            Write-OutputBox "  $($Account.DistinguishedName)" $PathColor
            Write-OUBlankLine

            $ColorIndex++
        }

        Write-OutputBox "Total special function accounts: $($Accounts.Count)" $SectionColor
        Write-OUBlankLine
    }
    catch {
        Write-OutputBox "ERROR getting Special Function Accounts: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-AllRequestedOUs {
    Write-ResultSeparator

    $PreviousSuppressSeparator = $script:SuppressResultSeparator
    $script:SuppressResultSeparator = $true

    try {
        # Only the OUs that are filled in. Blank ones would each fall back to the domain
        # root and repeat the same listing, so they are named once instead.
        $Common = @(
            @{ Name = "OUUsers";     Caption = "Users";             List = { Get-EmployeeOUs } },
            @{ Name = "OUComputers"; Caption = "Computers";         List = { Get-ComputerOUs } },
            @{ Name = "OUServers";   Caption = "Servers";           List = { Get-ServerOUs } },
            @{ Name = "OUDisabled";  Caption = "Disabled accounts"; List = { Get-DisabledOUs } }
        )

        $NotSet = New-Object System.Collections.Generic.List[string]
        $Listed = 0

        foreach ($Item in $Common) {
            if (Get-MSToolkitOU -Name $Item.Name) {
                if ($Listed -gt 0) { Write-OUBlankLine }
                & $Item.List
                $Listed++
            }
            else {
                $NotSet.Add($Item.Caption)
            }
        }

        if ($NotSet.Count -gt 0) {
            if ($Listed -gt 0) { Write-OUBlankLine }
            Write-OutputBox "Not listed - no OU is set in Settings for: $($NotSet.ToArray() -join ', ')." ([System.Drawing.Color]::DarkOrange)
            Write-OutputBox "Set them under Settings > Organizational units, or use Get Top-Level OUs to see the domain root." ([System.Drawing.Color]::DimGray)
            Write-OUBlankLine
        }
    }
    finally {
        $script:SuppressResultSeparator = $PreviousSuppressSeparator
    }
}

function Get-PasswordExpiringSoon {
    try {
        $Server = Get-SelectedServer

        $DaysInput = Get-InputBox -Title "Password Expiring Soon" -Prompt "Report on enabled accounts whose password expires within how many days?"

        if ([string]::IsNullOrWhiteSpace($DaysInput)) {
            return
        }

        $Days = 0
        if (-not [int]::TryParse($DaysInput.Trim(), [ref]$Days) -or $Days -lt 1) {
            Write-OutputBox "ERROR: Enter a whole number of days greater than zero." ([System.Drawing.Color]::Red)
            return
        }

        $Now = Get-Date
        $Cutoff = $Now.AddDays($Days)
        $DateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $File = Join-Path $LogPath "PasswordExpiringSoon-${Days}Days-$DateStamp.csv"

        $Results = Get-ADUser `
            -Filter { Enabled -eq $true } `
            -Server $Server `
            -Properties PasswordLastSet,PasswordNeverExpires,Department,Title,EmailAddress,msDS-UserPasswordExpiryTimeComputed |
        Where-Object {
            -not $_.PasswordNeverExpires -and
            $_.'msDS-UserPasswordExpiryTimeComputed' -ne $null -and
            $_.'msDS-UserPasswordExpiryTimeComputed' -ne 0 -and
            $_.'msDS-UserPasswordExpiryTimeComputed' -lt 9223372036854775807
        } |
        ForEach-Object {
            $Expiry = [datetime]::FromFileTime($_.'msDS-UserPasswordExpiryTimeComputed')

            [pscustomobject]@{
                Name           = $_.Name
                SamAccountName = $_.SamAccountName
                PasswordExpires = $Expiry
                DaysRemaining  = [math]::Floor(($Expiry - $Now).TotalDays)
                PasswordLastSet = $_.PasswordLastSet
                Department     = $_.Department
                Title          = $_.Title
                EmailAddress   = $_.EmailAddress
            }
        } |
        Where-Object { $_.PasswordExpires -ge $Now -and $_.PasswordExpires -le $Cutoff } |
        Sort-Object PasswordExpires

        $Results | Export-Csv $File -NoTypeInformation

        Write-MSToolkitReportResult -ReportName "Password Expiring Soon report ($Days days)" -FilePath $File -RowCount (@($Results).Count) -Server $Server
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Invoke-DCHealthCheck {
    try {
        Write-ResultSeparator
        Write-OutputBox "Running domain controller health check..." ([System.Drawing.Color]::FromArgb(31,58,93))

        # Query through the selected DC rather than automatic discovery, which can
        # land on a DC whose ADWS port does not answer.
        $ControllerList = @(Get-ADDomainController -Filter * -Server (Get-SelectedServer) | Sort-Object HostName)

        foreach ($Controller in $ControllerList) {
            $HostName = $Controller.HostName

            Write-OutputField -Label "Domain Controller" -Value $HostName -LabelColor ([System.Drawing.Color]::FromArgb(31,58,93))
            Write-OutputField -Label "  Site" -Value $Controller.Site

            # ADWS reachability - this is what the AD PowerShell module and ADAC use.
            $AdwsReachable = $false
            try {
                $null = Get-ADRootDSE -Server $HostName -ErrorAction Stop
                $AdwsReachable = $true
                Write-OutputField -Label "  ADWS" -Value "Reachable" -ValueColor ([System.Drawing.Color]::Green)
            }
            catch {
                Write-OutputField -Label "  ADWS" -Value "UNREACHABLE - $($_.Exception.Message)" -ValueColor ([System.Drawing.Color]::Red)
            }

            [System.Windows.Forms.Application]::DoEvents()

            # ADWS advertising - whether it registered with the DC locator at startup.
            # A DC can answer ADWS yet not be advertising (event 1005), which drops it out of
            # automatic ADWS discovery silently. Read over RPC, so this works even when
            # ADWS itself is unreachable.
            try {
                $AdvertEvent = Get-WinEvent -ComputerName $HostName -FilterHashtable @{
                    LogName = 'Active Directory Web Services'
                    Id      = 1005,1006
                } -MaxEvents 1 -ErrorAction Stop

                if ($AdvertEvent.Id -eq 1006) {
                    Write-OutputField -Label "  ADWS advertising" -Value "Yes (since $($AdvertEvent.TimeCreated))" -ValueColor ([System.Drawing.Color]::Green)
                }
                else {
                    Write-OutputField -Label "  ADWS advertising" -Value "NO - failed at startup $($AdvertEvent.TimeCreated) (event 1005). Restart ADWS to re-register." -ValueColor ([System.Drawing.Color]::Red)
                }
            }
            catch {
                if ($_.Exception.Message -match 'No events were found') {
                    Write-OutputField -Label "  ADWS advertising" -Value "Unknown - no advertising events in the ADWS log" -ValueColor ([System.Drawing.Color]::DarkOrange)
                }
                else {
                    Write-OutputField -Label "  ADWS advertising" -Value "Could not query - $($_.Exception.Message)" -ValueColor ([System.Drawing.Color]::DarkOrange)
                }
            }

            [System.Windows.Forms.Application]::DoEvents()

            # Replication failures reported against this DC. This also goes over ADWS, so if
            # the check above failed it would only wait out the same timeout again.
            if (-not $AdwsReachable) {
                Write-OutputField -Label "  Replication" -Value "Skipped - ADWS unreachable" -ValueColor ([System.Drawing.Color]::DarkOrange)
            }
            else {
                try {
                    $Failures = @(Get-ADReplicationFailure -Target $HostName -ErrorAction Stop)

                    if ($Failures.Count -eq 0) {
                        Write-OutputField -Label "  Replication" -Value "No failures reported" -ValueColor ([System.Drawing.Color]::Green)
                    }
                    else {
                        foreach ($Failure in $Failures) {
                            Write-OutputField -Label "  Replication" -Value "FAILURE from $($Failure.Partner) - $($Failure.FailureCount) failure(s), last error $($Failure.LastError)" -ValueColor ([System.Drawing.Color]::Red)
                        }
                    }
                }
                catch {
                    Write-OutputField -Label "  Replication" -Value "Could not query - $($_.Exception.Message)" -ValueColor ([System.Drawing.Color]::DarkOrange)
                }
            }

            [System.Windows.Forms.Application]::DoEvents()

            # Security log capacity and how far back it actually reaches.
            try {
                $SecurityLog = Get-WinEvent -ListLog Security -ComputerName $HostName -ErrorAction Stop
                $MaxMB = [math]::Round($SecurityLog.MaximumSizeInBytes / 1MB, 1)

                $OldestText = "unknown"
                try {
                    $OldestRecord = Get-WinEvent -ComputerName $HostName -LogName Security -MaxEvents 1 -Oldest -ErrorAction Stop
                    $OldestText = "oldest event $($OldestRecord.TimeCreated)"
                }
                catch {
                    $OldestText = "oldest event unavailable"
                }

                $LogColor = if ($MaxMB -lt 192) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::Green }
                Write-OutputField -Label "  Security log" -Value "$MaxMB MB max, $OldestText" -ValueColor $LogColor
            }
            catch {
                Write-OutputField -Label "  Security log" -Value "Could not query - $($_.Exception.Message)" -ValueColor ([System.Drawing.Color]::DarkOrange)
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        Write-OutputBox "Domain controller health check complete. Checked $($ControllerList.Count) DC(s)." ([System.Drawing.Color]::Green)
    }
    catch {
        Write-OutputBox "ERROR running DC health check: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-InactiveUsers90Days {
    try {
        $Server = Get-SelectedServer
        $Cutoff = (Get-Date).AddDays(-90)
        $DateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $File = Join-Path $LogPath "InactiveUsers90Days-$DateStamp.csv"

        # Captured rather than piped straight to Export-Csv so the row count can be reported.
        $Results = @(
            Get-ADUser `
                -Filter {
                    Enabled -eq $true -and LastLogonDate -lt $Cutoff
                } `
                -Server $Server `
                -Properties LastLogonDate,Department,Title,EmailAddress |
            Select-Object Name,SamAccountName,Enabled,LastLogonDate,Department,Title,EmailAddress |
            Sort-Object LastLogonDate
        )

        $Results | Export-Csv $File -NoTypeInformation

        Write-MSToolkitReportResult -ReportName "90-Day Inactive Users report" -FilePath $File -RowCount (@($Results).Count) -Server $Server
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Get-InactiveComputers90Days {
    try {
        $Server = Get-SelectedServer
        $Cutoff = (Get-Date).AddDays(-90)
        $DateStamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $File = Join-Path $LogPath "InactiveComputers90Days-$DateStamp.csv"

        # Captured rather than piped straight to Export-Csv so the row count can be reported.
        $Results = @(
            Get-ADComputer `
                -Filter {
                    Enabled -eq $true -and LastLogonDate -lt $Cutoff
                } `
                -Server $Server `
                -Properties LastLogonDate,OperatingSystem |
            Select-Object Name,DNSHostName,Enabled,LastLogonDate,OperatingSystem |
            Sort-Object LastLogonDate
        )

        $Results | Export-Csv $File -NoTypeInformation

        Write-MSToolkitReportResult -ReportName "90-Day Inactive Computers report" -FilePath $File -RowCount (@($Results).Count) -Server $Server
    }
    catch {
        Write-OutputBox "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}


function Invoke-EntraDeltaSync {
    $EntraSyncServer = Get-MSToolkitSetting -Name "EntraSyncServer"

    if ([string]::IsNullOrWhiteSpace($EntraSyncServer)) {
        Write-OutputBox "No Entra Connect server is configured. Set one in Settings before running a delta sync." ([System.Drawing.Color]::DarkOrange)
        return
    }

    Write-ResultSeparator

    try {
        Write-OutputBox "Starting Microsoft Entra Connect delta sync on $EntraSyncServer..."

        $Result = Invoke-Command -ComputerName $EntraSyncServer -ScriptBlock {
            Import-Module ADSync -ErrorAction Stop
            Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        } -ErrorAction Stop

        if ($Result) {
            foreach ($Item in $Result) {
                if ($Item.PSObject.Properties.Name -contains "Result") {
                    Write-OutputBox "Entra delta sync result: $($Item.Result)"
                }
                else {
                    Write-OutputBox ($Item | Out-String).Trim()
                }
            }
        }

        Write-OutputBox "Microsoft Entra Connect delta sync command completed on $EntraSyncServer." ([System.Drawing.Color]::Green)
    }
    catch {
        Write-OutputBox "ERROR running Entra delta sync on ${EntraSyncServer}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Set-ReplicationButtonsEnabled {
    param([bool]$Enabled)

    foreach ($Button in @($SelectedDCButton, $ReplicateAllButton)) {
        if ($Button) {
            $Button.Enabled = $Enabled
        }
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Sync-DC {
    param([string]$DC)

    try {
        Write-OutputBox "Starting replication sync on $DC..."
        Write-OutputBox "Running repadmin /syncall $DC.$Domain /AdeP - this can take several minutes."

        $global:LASTEXITCODE = 0

        # Stream repadmin output line by line instead of collecting it, and pump
        # the WinForms message loop between lines. Collecting the output meant
        # nothing appeared until the command finished, and the window stopped
        # repainting for the whole run.
        repadmin /syncall "$DC.$Domain" /AdeP 2>&1 | ForEach-Object {
            $Line = ($_ | Out-String).Trim()

            if (-not [string]::IsNullOrWhiteSpace($Line)) {
                Write-OutputBox $Line
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($LASTEXITCODE -eq 0) {
            Write-OutputBox "Replication command completed for $DC." ([System.Drawing.Color]::Green)
        }
        else {
            Write-OutputBox "Replication command for $DC reported errors. repadmin exit code: $LASTEXITCODE" ([System.Drawing.Color]::Red)
        }
    }
    catch {
        Write-OutputBox "ERROR syncing ${DC}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
}

function Sync-SelectedDC {
    try {
        if (-not $DCDropdown.SelectedItem) {
            Write-OutputBox "ERROR: No DC selected." ([System.Drawing.Color]::Red)
            return
        }

        Set-ReplicationButtonsEnabled -Enabled $false

        Write-ResultSeparator
        Sync-DC -DC $DCDropdown.SelectedItem.ToString()
    }
    catch {
        Write-OutputBox "ERROR syncing selected DC: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
    finally {
        Set-ReplicationButtonsEnabled -Enabled $true
    }
}

function Sync-AllDCs {
    try {
        Set-ReplicationButtonsEnabled -Enabled $false

        Write-ResultSeparator

        $FirstDC = $true

        foreach ($DC in $DCs) {
            if (-not $FirstDC) {
                Write-ResultSeparator
            }

            Sync-DC -DC $DC
            $FirstDC = $false
        }
    }
    catch {
        Write-OutputBox "ERROR syncing all DCs: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
    }
    finally {
        Set-ReplicationButtonsEnabled -Enabled $true
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "MSToolkit - $Domain"
$form.Size = New-Object System.Drawing.Size(1400,800)
$form.MinimumSize = New-Object System.Drawing.Size(1400,600)
$form.StartPosition = "CenterScreen"
$form.WindowState = "Maximized"
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.FormBorderStyle = "Sizable"
$form.AutoSize = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Top utility bar - styled like the compare tools without changing control positions.
$TopBar = New-Object System.Windows.Forms.Panel
$TopBar.Location = New-Object System.Drawing.Point(0,0)
$TopBar.Dock = "Top"
$TopBar.Height = 38
$TopBar.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$TopBar.Tag = "TopBar"
$form.Controls.Add($TopBar)

# Second strip for the per-server consoles and shells, styled like the top bar.
$BottomBar = New-Object System.Windows.Forms.Panel
$BottomBar.Dock = "Bottom"
$BottomBar.Height = 38
$BottomBar.BackColor = [System.Drawing.Color]::FromArgb(31,58,93)
$BottomBar.Tag = "TopBar"
$form.Controls.Add($BottomBar)

$DCLabel = New-Object System.Windows.Forms.Label
$DCLabel.Text = "AD Server:"
$DCLabel.Location = New-Object System.Drawing.Point(10,12)
$DCLabel.Size = New-Object System.Drawing.Size(75,20)
$DCLabel.ForeColor = [System.Drawing.Color]::White
$DCLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$TopBar.Controls.Add($DCLabel)

$DCDropdown = New-Object System.Windows.Forms.ComboBox
$DCDropdown.Location = New-Object System.Drawing.Point(85,8)
$DCDropdown.Size = New-Object System.Drawing.Size(150,24)
$DCDropdown.DropDownStyle = "DropDownList"

foreach ($DC in $DCs) {
    [void]$DCDropdown.Items.Add($DC)
}

$PreferredDCIndex = -1
if ($script:PreferredDC) {
    for ($i = 0; $i -lt $DCDropdown.Items.Count; $i++) {
        if ([string]$DCDropdown.Items[$i] -ieq $script:PreferredDC) {
            $PreferredDCIndex = $i
            break
        }
    }
}

if ($PreferredDCIndex -ge 0) {
    $DCDropdown.SelectedIndex = $PreferredDCIndex
}
else {
    $DCDropdown.SelectedIndex = 0
    if ($DCDropdown.SelectedItem) {
        $script:PreferredDC = [string]$DCDropdown.SelectedItem
        [void](Save-ThemePreference)
    }
}

$DCDropdown.Add_SelectedIndexChanged({
    # Keep the bottom-bar copy showing the same server.
    if ($DCDropdownBottom -and ($DCDropdownBottom.SelectedIndex -ne $DCDropdown.SelectedIndex)) {
        $script:SyncingDCDropdowns = $true
        try { $DCDropdownBottom.SelectedIndex = $DCDropdown.SelectedIndex }
        finally { $script:SyncingDCDropdowns = $false }
    }

    if ($DCDropdown.SelectedItem) {
        $script:PreferredDC = [string]$DCDropdown.SelectedItem
        $Saved = Save-ThemePreference
        Update-StatusStrip
        Write-OutputBox "Selected AD server: $(Get-SelectedServer)" (Get-ThemePalette).Info

        if (-not $Saved) {
            Write-OutputBox "WARNING: The selected AD server changed for this session, but the preference could not be saved to $ThemeSettingsPath" ([System.Drawing.Color]::DarkOrange)
        }
    }
})

$DCDropdown.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$TopBar.Controls.Add($DCDropdown)

# Mirror of the AD Server selector on the bottom bar. The top dropdown stays the
# single source of truth - Get-SelectedServer and the replication buttons read it -
# so this one only forwards its selection upward and is kept in step from above.
$DCLabelBottom = New-Object System.Windows.Forms.Label
$DCLabelBottom.Text = "AD Server:"
$DCLabelBottom.Location = New-Object System.Drawing.Point(10,12)
$DCLabelBottom.Size = New-Object System.Drawing.Size(75,20)
$DCLabelBottom.ForeColor = [System.Drawing.Color]::White
$DCLabelBottom.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$BottomBar.Controls.Add($DCLabelBottom)

$DCDropdownBottom = New-Object System.Windows.Forms.ComboBox
$DCDropdownBottom.Location = New-Object System.Drawing.Point(85,8)
$DCDropdownBottom.Size = New-Object System.Drawing.Size(150,24)
$DCDropdownBottom.DropDownStyle = "DropDownList"
$DCDropdownBottom.Font = $DCDropdown.Font

foreach ($DCItem in $DCDropdown.Items) {
    [void]$DCDropdownBottom.Items.Add($DCItem)
}

$DCDropdownBottom.SelectedIndex = $DCDropdown.SelectedIndex

$DCDropdownBottom.Add_SelectedIndexChanged({
    if ($script:SyncingDCDropdowns) { return }

    if ($DCDropdown.SelectedIndex -ne $DCDropdownBottom.SelectedIndex) {
        $DCDropdown.SelectedIndex = $DCDropdownBottom.SelectedIndex
    }
})

$BottomBar.Controls.Add($DCDropdownBottom)

$SelectedDCButton = New-Object System.Windows.Forms.Button
$SelectedDCButton.Text = "REPL Selected DC"
$SelectedDCButton.Location = New-Object System.Drawing.Point(245,6)
$SelectedDCButton.Size = New-Object System.Drawing.Size(118,28)
$SelectedDCButton.Add_Click({ Sync-SelectedDC })
$SelectedDCButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$SelectedDCButton.ForeColor = [System.Drawing.Color]::White
$SelectedDCButton.FlatStyle = "Flat"
$SelectedDCButton.FlatAppearance.BorderSize = 0
$SelectedDCButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$SelectedDCButton.Tag = "TopButton"
$TopBar.Controls.Add($SelectedDCButton)

$ReplicateAllButton = New-Object System.Windows.Forms.Button
$ReplicateAllButton.Text = "REPL All DCs"
$ReplicateAllButton.Location = New-Object System.Drawing.Point(369,6)
$ReplicateAllButton.Size = New-Object System.Drawing.Size(96,28)
$ReplicateAllButton.Add_Click({ Sync-AllDCs })
$ReplicateAllButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$ReplicateAllButton.ForeColor = [System.Drawing.Color]::White
$ReplicateAllButton.FlatStyle = "Flat"
$ReplicateAllButton.FlatAppearance.BorderSize = 0
$ReplicateAllButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$ReplicateAllButton.Tag = "TopButton"
$TopBar.Controls.Add($ReplicateAllButton)

$DeltaSyncButton = New-Object System.Windows.Forms.Button
$DeltaSyncButton.Text = "Delta Sync"
$DeltaSyncButton.Location = New-Object System.Drawing.Point(471,6)
$DeltaSyncButton.Size = New-Object System.Drawing.Size(90,28)
$DeltaSyncButton.Add_Click({ Invoke-EntraDeltaSync })
$DeltaSyncButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$DeltaSyncButton.ForeColor = [System.Drawing.Color]::White
$DeltaSyncButton.FlatStyle = "Flat"
$DeltaSyncButton.FlatAppearance.BorderSize = 0
$DeltaSyncButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$DeltaSyncButton.Tag = "TopButton"
$TopBar.Controls.Add($DeltaSyncButton)

$DCHealthButton = New-Object System.Windows.Forms.Button
$DCHealthButton.Text = "DC Health"
$DCHealthButton.Location = New-Object System.Drawing.Point(715,6)
$DCHealthButton.Size = New-Object System.Drawing.Size(85,28)
$DCHealthButton.Add_Click({ Invoke-DCHealthCheck })
$DCHealthButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$DCHealthButton.ForeColor = [System.Drawing.Color]::White
$DCHealthButton.FlatStyle = "Flat"
$DCHealthButton.FlatAppearance.BorderSize = 0
$DCHealthButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$DCHealthButton.Tag = "TopButton"
$TopBar.Controls.Add($DCHealthButton)

$TopDividerOne = New-Object System.Windows.Forms.Panel
$TopDividerOne.Location = New-Object System.Drawing.Point(569,8)
$TopDividerOne.Size = New-Object System.Drawing.Size(2,22)
$TopDividerOne.BackColor = [System.Drawing.Color]::FromArgb(140,140,140)
$TopDividerOne.Tag = "TopDivider"
$TopBar.Controls.Add($TopDividerOne)

$ADACButton = New-Object System.Windows.Forms.Button
$ADACButton.Text = "ADAC"
$ADACButton.Location = New-Object System.Drawing.Point(579,6)
$ADACButton.Size = New-Object System.Drawing.Size(62,28)
$ADACButton.Add_Click({ Launch-ADAC })
$ADACButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$ADACButton.ForeColor = [System.Drawing.Color]::White
$ADACButton.FlatStyle = "Flat"
$ADACButton.FlatAppearance.BorderSize = 0
$ADACButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$ADACButton.Tag = "TopButton"
$TopBar.Controls.Add($ADACButton)

$ADUCButton = New-Object System.Windows.Forms.Button
$ADUCButton.Text = "ADUC"
$ADUCButton.Location = New-Object System.Drawing.Point(647,6)
$ADUCButton.Size = New-Object System.Drawing.Size(62,28)
$ADUCButton.Add_Click({ Launch-ADUC })
$ADUCButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$ADUCButton.ForeColor = [System.Drawing.Color]::White
$ADUCButton.FlatStyle = "Flat"
$ADUCButton.FlatAppearance.BorderSize = 0
$ADUCButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$ADUCButton.Tag = "TopButton"
$TopBar.Controls.Add($ADUCButton)


$DNSButton = New-Object System.Windows.Forms.Button
$DNSButton.Text = "DNS"
$DNSButton.Location = New-Object System.Drawing.Point(806,6)
$DNSButton.Size = New-Object System.Drawing.Size(52,28)
$DNSButton.Add_Click({ Launch-DNSManager })
$DNSButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$DNSButton.ForeColor = [System.Drawing.Color]::White
$DNSButton.FlatStyle = "Flat"
$DNSButton.FlatAppearance.BorderSize = 0
$DNSButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$DNSButton.Tag = "TopButton"
$TopBar.Controls.Add($DNSButton)



$GPManagementButton = New-Object System.Windows.Forms.Button
$GPManagementButton.Text = "GPO"
$GPManagementButton.Location = New-Object System.Drawing.Point(864,6)
$GPManagementButton.Size = New-Object System.Drawing.Size(58,28)
$GPManagementButton.Add_Click({ Launch-GPManagement })
$GPManagementButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$GPManagementButton.ForeColor = [System.Drawing.Color]::White
$GPManagementButton.FlatStyle = "Flat"
$GPManagementButton.FlatAppearance.BorderSize = 0
$GPManagementButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$GPManagementButton.Tag = "TopButton"
$TopBar.Controls.Add($GPManagementButton)


$PrintManagementButton = New-Object System.Windows.Forms.Button
$PrintManagementButton.Text = "Print MGMT"
$PrintManagementButton.Location = New-Object System.Drawing.Point(514,6)
$PrintManagementButton.Size = New-Object System.Drawing.Size(88,28)
$PrintManagementButton.Add_Click({ Launch-PrintManagement })
$PrintManagementButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$PrintManagementButton.ForeColor = [System.Drawing.Color]::White
$PrintManagementButton.FlatStyle = "Flat"
$PrintManagementButton.FlatAppearance.BorderSize = 0
$PrintManagementButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$PrintManagementButton.Tag = "TopButton"
$BottomBar.Controls.Add($PrintManagementButton)

$ComputerMgmtButton = New-Object System.Windows.Forms.Button
$ComputerMgmtButton.Text = "Computer MGMT"
$ComputerMgmtButton.Location = New-Object System.Drawing.Point(245,6)
$ComputerMgmtButton.Size = New-Object System.Drawing.Size(118,28)
$ComputerMgmtButton.Add_Click({ Launch-ComputerManagement })
$ComputerMgmtButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$ComputerMgmtButton.ForeColor = [System.Drawing.Color]::White
$ComputerMgmtButton.FlatStyle = "Flat"
$ComputerMgmtButton.FlatAppearance.BorderSize = 0
$ComputerMgmtButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$ComputerMgmtButton.Tag = "TopButton"
$BottomBar.Controls.Add($ComputerMgmtButton)


$PowerShellButton = New-Object System.Windows.Forms.Button
$PowerShellButton.Text = "PowerShell"
$PowerShellButton.Location = New-Object System.Drawing.Point(423,6)
$PowerShellButton.Size = New-Object System.Drawing.Size(85,28)
$PowerShellButton.Add_Click({ Launch-PowerShell })
$PowerShellButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$PowerShellButton.ForeColor = [System.Drawing.Color]::White
$PowerShellButton.FlatStyle = "Flat"
$PowerShellButton.FlatAppearance.BorderSize = 0
$PowerShellButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$PowerShellButton.Tag = "TopButton"
$BottomBar.Controls.Add($PowerShellButton)

$ISEButton = New-Object System.Windows.Forms.Button
$ISEButton.Text = "ISE"
$ISEButton.Location = New-Object System.Drawing.Point(369,6)
$ISEButton.Size = New-Object System.Drawing.Size(48,28)
$ISEButton.Add_Click({ Launch-ISE })
$ISEButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$ISEButton.ForeColor = [System.Drawing.Color]::White
$ISEButton.FlatStyle = "Flat"
$ISEButton.FlatAppearance.BorderSize = 0
$ISEButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$ISEButton.Tag = "TopButton"
$BottomBar.Controls.Add($ISEButton)

$RemoteDesktopButton = New-Object System.Windows.Forms.Button
$RemoteDesktopButton.Text = "RDP"
$RemoteDesktopButton.Location = New-Object System.Drawing.Point(608,6)
$RemoteDesktopButton.Size = New-Object System.Drawing.Size(52,28)
$RemoteDesktopButton.Add_Click({ Launch-RemoteDesktop })
$RemoteDesktopButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$RemoteDesktopButton.ForeColor = [System.Drawing.Color]::White
$RemoteDesktopButton.FlatStyle = "Flat"
$RemoteDesktopButton.FlatAppearance.BorderSize = 0
$RemoteDesktopButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$RemoteDesktopButton.Tag = "TopButton"
$BottomBar.Controls.Add($RemoteDesktopButton)

$TopDividerThree = New-Object System.Windows.Forms.Panel
$TopDividerThree.Location = New-Object System.Drawing.Point(668,8)
$TopDividerThree.Size = New-Object System.Drawing.Size(2,22)
$TopDividerThree.BackColor = [System.Drawing.Color]::FromArgb(140,140,140)
$TopDividerThree.Tag = "TopDivider"
$BottomBar.Controls.Add($TopDividerThree)

$LogsButton = New-Object System.Windows.Forms.Button
$LogsButton.Text = "Logs"
$LogsButton.Location = New-Object System.Drawing.Point(678,6)
$LogsButton.Size = New-Object System.Drawing.Size(62,28)
$LogsButton.Add_Click({ Show-LogsPath })
$LogsButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$LogsButton.ForeColor = [System.Drawing.Color]::White
$LogsButton.FlatStyle = "Flat"
$LogsButton.FlatAppearance.BorderSize = 0
$LogsButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$LogsButton.Tag = "TopButton"
$BottomBar.Controls.Add($LogsButton)

$SettingsButton = New-Object System.Windows.Forms.Button
$SettingsButton.Text = "Settings"
$SettingsButton.Size = New-Object System.Drawing.Size(84,28)
$SettingsButton.Location = New-Object System.Drawing.Point(($BottomBar.ClientSize.Width - 94),6)
$SettingsButton.Anchor = "Top,Right"
$SettingsButton.Add_Click({ Show-MSToolkitSettings })
$SettingsButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$SettingsButton.ForeColor = [System.Drawing.Color]::White
$SettingsButton.FlatStyle = "Flat"
$SettingsButton.FlatAppearance.BorderSize = 0
$SettingsButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$SettingsButton.Tag = "TopButton"
$BottomBar.Controls.Add($SettingsButton)

# Identical to the toggle on every child tool: same size, same navy that blends
# into the header, same glyphs, same hover highlight.
$ThemeToggleButton = New-Object System.Windows.Forms.Button
$ThemeToggleButton.Size = New-Object System.Drawing.Size(44,32)
$ThemeToggleButton.Anchor = "Top,Right"
$ThemeToggleButton.BackColor = [System.Drawing.Color]::FromArgb(24,47,74)
$ThemeToggleButton.ForeColor = [System.Drawing.Color]::White
$ThemeToggleButton.FlatStyle = "Flat"
$ThemeToggleButton.FlatAppearance.BorderSize = 0
$ThemeToggleButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(45,74,110)
$ThemeToggleButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(18,36,58)
$ThemeToggleButton.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 15)
$ThemeToggleButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$ThemeToggleButton.TabStop = $false
$ThemeToggleButton.Tag = "ThemeButton"

$ThemeToggleTip = New-Object System.Windows.Forms.ToolTip
$ThemeToggleTip.SetToolTip($ThemeToggleButton, "Switch between Light and Dark mode")
$ThemeToggleButton.Add_Click({
    if ($script:ThemeMode -eq "Dark") {
        $script:ThemeMode = "Light"
    }
    else {
        $script:ThemeMode = "Dark"
    }

    if ($DCDropdown -and $DCDropdown.SelectedItem) {
        $script:PreferredDC = [string]$DCDropdown.SelectedItem
    }
    $Saved = Save-ThemePreference
    Apply-CurrentTheme

    if (-not $Saved) {
        Write-OutputBox "WARNING: Theme changed for this session, but the preference could not be saved to $ThemeSettingsPath" ([System.Drawing.Color]::DarkOrange)
    }
})
$TopBar.Controls.Add($ThemeToggleButton)

# Keep only the Light/Dark mode toggle right-aligned with a visible
# dark-blue gutter at the right edge. PowerShell remains in the normal
# left-to-right utility-button row immediately after GPO.
$ThemeToggleRightPadding = 15
$PositionThemeToggleButton = {
    $ThemeToggleButton.Location = New-Object System.Drawing.Point(
        ($TopBar.ClientSize.Width - $ThemeToggleButton.Width - $ThemeToggleRightPadding),
        4
    )
}
$TopBar.Add_Resize($PositionThemeToggleButton)
& $PositionThemeToggleButton

if ($script:ThemeMode -eq "Dark") {
    $ThemeToggleButton.Text = [string][char]0x263C
    $ThemeToggleButton.ForeColor = [System.Drawing.Color]::FromArgb(255,214,102)
}
else {
    $ThemeToggleButton.Text = [string][char]0x263E
    $ThemeToggleButton.ForeColor = [System.Drawing.Color]::FromArgb(236,242,250)
}

$StatusPanel = New-Object System.Windows.Forms.Panel
$StatusPanel.Location = New-Object System.Drawing.Point(0,38)

# Width comes from the live client area. A fixed width here was laid out for the
# old 1200px form, which left a blank strip on the right once the window grew.
$StatusPanel.Size = New-Object System.Drawing.Size($form.ClientSize.Width,30)
$StatusPanel.Anchor = "Top,Left,Right"
$StatusPanel.BackColor = [System.Drawing.Color]::White
$StatusPanel.BorderStyle = "FixedSingle"
$StatusPanel.Tag = "StatusBar"
$form.Controls.Add($StatusPanel)

$StatusAccountLabel = New-Object System.Windows.Forms.Label
$StatusAccountLabel.Location = New-Object System.Drawing.Point(10,6)
$StatusAccountLabel.Size = New-Object System.Drawing.Size(300,20)
$StatusAccountLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$StatusPanel.Controls.Add($StatusAccountLabel)

$StatusDomainLabel = New-Object System.Windows.Forms.Label
$StatusDomainLabel.Location = New-Object System.Drawing.Point(320,6)
$StatusDomainLabel.Size = New-Object System.Drawing.Size(220,20)
$StatusDomainLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$StatusPanel.Controls.Add($StatusDomainLabel)

$StatusServerLabel = New-Object System.Windows.Forms.Label
$StatusServerLabel.Location = New-Object System.Drawing.Point(550,6)
$StatusServerLabel.Size = New-Object System.Drawing.Size(610,20)
$StatusServerLabel.Anchor = "Top,Left,Right"
$StatusServerLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8.5)
$StatusPanel.Controls.Add($StatusServerLabel)

$LeftButtonPanel = New-Object System.Windows.Forms.Panel
$LeftButtonPanel.Location = New-Object System.Drawing.Point(10,72)
$LeftButtonPanel.Size = New-Object System.Drawing.Size(250,($form.ClientSize.Height - 120))
$LeftButtonPanel.AutoScroll = $true
$LeftButtonPanel.Anchor = "Top,Bottom,Left"
$LeftButtonPanel.BackColor = [System.Drawing.Color]::White
$LeftButtonPanel.BorderStyle = "FixedSingle"
$LeftButtonPanel.Tag = "Sidebar"
$form.Controls.Add($LeftButtonPanel)

$RightButtonPanel = New-Object System.Windows.Forms.Panel
$RightButtonPanel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 260),72)
$RightButtonPanel.Size = New-Object System.Drawing.Size(250,($form.ClientSize.Height - 120))
$RightButtonPanel.AutoScroll = $true
$RightButtonPanel.Anchor = "Top,Bottom,Right"
$RightButtonPanel.BackColor = [System.Drawing.Color]::White
$RightButtonPanel.BorderStyle = "FixedSingle"
$RightButtonPanel.Tag = "Sidebar"
$form.Controls.Add($RightButtonPanel)

$OutputPanel = New-Object System.Windows.Forms.Panel
$OutputPanel.Location = New-Object System.Drawing.Point(270,72)
$OutputPanel.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 540),($form.ClientSize.Height - 120))
$OutputPanel.Anchor = "Top,Bottom,Left,Right"
$OutputPanel.BackColor = [System.Drawing.Color]::White
$OutputPanel.BorderStyle = "FixedSingle"
$OutputPanel.Tag = "OutputContainer"
$form.Controls.Add($OutputPanel)

$OutputToolbar = New-Object System.Windows.Forms.Panel
$OutputToolbar.Location = New-Object System.Drawing.Point(0,0)
# Sized from the panel rather than a fixed 648, which was laid out for the old
# narrower window and left dead space down the right-hand side.
$OutputToolbar.Size = New-Object System.Drawing.Size($OutputPanel.ClientSize.Width,34)
$OutputToolbar.Anchor = "Top,Left,Right"
$OutputToolbar.BackColor = [System.Drawing.Color]::White
$OutputToolbar.Tag = "OutputToolbar"
$OutputPanel.Controls.Add($OutputToolbar)

$OutputTitleLabel = New-Object System.Windows.Forms.Label
$OutputTitleLabel.Text = "Activity Output"
$OutputTitleLabel.Location = New-Object System.Drawing.Point(10,8)
$OutputTitleLabel.Size = New-Object System.Drawing.Size(140,20)
$OutputTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$OutputTitleLabel.Tag = "SectionLabel"
$OutputToolbar.Controls.Add($OutputTitleLabel)

$CopyOutputButton = New-Object System.Windows.Forms.Button
$CopyOutputButton.Text = "Copy Output"
$CopyOutputButton.Location = New-Object System.Drawing.Point(($OutputToolbar.ClientSize.Width - 196),3)
$CopyOutputButton.Size = New-Object System.Drawing.Size(92,27)
$CopyOutputButton.Anchor = "Top,Right"
$CopyOutputButton.Tag = "OutputUtility"
$CopyOutputButton.Add_Click({ Copy-ActivityOutput })
$OutputToolbar.Controls.Add($CopyOutputButton)

$ClearOutputButton = New-Object System.Windows.Forms.Button
$ClearOutputButton.Text = "Clear Output"
$ClearOutputButton.Location = New-Object System.Drawing.Point(($OutputToolbar.ClientSize.Width - 98),3)
$ClearOutputButton.Size = New-Object System.Drawing.Size(88,27)
$ClearOutputButton.Anchor = "Top,Right"
$ClearOutputButton.Tag = "OutputUtility"
$ClearOutputButton.Add_Click({ Clear-ActivityOutput })
$OutputToolbar.Controls.Add($ClearOutputButton)

$OutputBox = New-Object System.Windows.Forms.RichTextBox
$OutputBox.Location = New-Object System.Drawing.Point(0,34)
$OutputBox.Size = New-Object System.Drawing.Size($OutputPanel.ClientSize.Width,($OutputPanel.ClientSize.Height - 34))
$OutputBox.Multiline = $true
$OutputBox.ScrollBars = "Vertical"
$OutputBox.ReadOnly = $true
$OutputBox.Anchor = "Top,Bottom,Left,Right"
$OutputBox.BackColor = [System.Drawing.Color]::White
$OutputBox.ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
$OutputBox.BorderStyle = "FixedSingle"
$OutputBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$OutputBox.Tag = "Output"
$OutputPanel.Controls.Add($OutputBox)

$LeftY = 10
$RightY = 10

function Add-Section {
    param(
        [System.Windows.Forms.Panel]$Panel,
        [ref]$YPosition,
        [string]$Text
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point(10,$YPosition.Value)
    $label.Size = New-Object System.Drawing.Size((Get-MSToolkitSidebarItemWidth -Panel $Panel),22)
    $label.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(31,58,93)
    $label.Tag = "SectionLabel"

    $Panel.Controls.Add($label)
    $YPosition.Value += 26
}

function Get-MSToolkitSidebarItemWidth {
    param([System.Windows.Forms.Panel]$Panel)

    # Leave room for the vertical scrollbar AutoScroll adds once the section list
    # is taller than the panel, otherwise it sits on top of the button edge.
    $Width = $Panel.ClientSize.Width - 20 - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth

    if ($Width -lt 120) {
        $Width = 120
    }

    return $Width
}

function Add-Button {
    param(
        [System.Windows.Forms.Panel]$Panel,
        [ref]$YPosition,
        [string]$Text,
        [scriptblock]$Action,
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point(10,$YPosition.Value)
    $button.Size = New-Object System.Drawing.Size((Get-MSToolkitSidebarItemWidth -Panel $Panel),30)
    $button.BackColor = [System.Drawing.Color]::White
    $button.ForeColor = $ForeColor
    $button.FlatStyle = "Flat"
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(210,215,220)
    $button.FlatAppearance.BorderSize = 1
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $button.Padding = New-Object System.Windows.Forms.Padding(8,0,0,0)

    if ($ForeColor.ToArgb() -eq ([System.Drawing.Color]::Red).ToArgb()) {
        $button.Tag = "SidebarDanger"
    }
    elseif ($ForeColor.ToArgb() -eq ([System.Drawing.Color]::ForestGreen).ToArgb()) {
        $button.Tag = "SidebarSuccess"
    }
    else {
        $button.Tag = "SidebarDefault"
    }

    $button.Add_Click($Action)

    $Panel.Controls.Add($button)
    $YPosition.Value += 38
}

# Left sidebar: Users and Groups
Add-Section -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Users"
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Create New User" -Action { Launch-NewADUserScript } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get User" -Action { Show-ADUserInfo }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get User Groups" -Action { Get-ADUserGroups }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get User OU" -Action { Get-ADUserOU }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Unlock User" -Action { Unlock-ADUserAccount }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Reset Password" -Action { Reset-ADUserPassword }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Force Password Change" -Action { Force-PasswordChange }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Enable User" -Action { Enable-ADUserAccount }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Disable User" -Action { Disable-ADUserAccount }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Investigate Lockout" -Action { Invoke-InvestigateAccountLockoutInTool } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Delete User" -Action { Delete-ADUserAccount } -ForeColor ([System.Drawing.Color]::Red)

$LeftY += 8

Add-Section -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Service Accounts"
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get Managed Service Accounts" -Action { Get-MSToolkitManagedServiceAccounts }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get Special Function Accounts" -Action { Get-MSToolkitSpecialFunctionAccounts }

$LeftY += 8

Add-Section -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Groups"
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get Group" -Action { Get-ADGroupInfo }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get Group Members" -Action { Get-ADGroupMembers }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get Security Groups" -Action { Get-MSToolkitSecurityGroups }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Get Distribution Groups" -Action { Get-MSToolkitDistributionGroups }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Add User to Group" -Action { Add-UserToGroup }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Remove User from Group" -Action { Remove-UserFromGroup }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Create Group" -Action { Create-ADGroup }
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Compare/Manage User Groups" -Action { Invoke-CompareUserGroupsInTool } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $LeftButtonPanel -YPosition ([ref]$LeftY) -Text "Delete Group" -Action { Delete-ADGroup } -ForeColor ([System.Drawing.Color]::Red)

# Right sidebar: M365, Computers, OUs, and Reports
Add-Section -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "M365"
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Block Teams Numbers" -Action { Launch-M365TeamsBlockNumber } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Exchange Online" -Action { Launch-M365ExchangeOnlineTools } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Intune Tools" -Action { Launch-IntuneTools } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "M365 Group Compare/Add" -Action { Launch-M365GroupCompareScript } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "M365 Distro Compare/Add" -Action { Launch-M365DistributionGroupCompareScript } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "M365 Conditional Access" -Action { Launch-M365ConditionalAccessScript } -ForeColor ([System.Drawing.Color]::ForestGreen)
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "OneDrive / SharePoint" -Action { Launch-M365OneDriveTools } -ForeColor ([System.Drawing.Color]::ForestGreen)

$RightY += 8

Add-Section -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Computers"
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get Computer" -Action { Get-ADComputerInfo }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get Computer OU" -Action { Get-ADComputerOU }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Enable Computer" -Action { Enable-ADComputerAccount }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Disable Computer" -Action { Disable-ADComputerAccount }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Delete Computer" -Action { Delete-ADComputerAccount } -ForeColor ([System.Drawing.Color]::Red)
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Reset Computer Account" -Action { Reset-ADComputerAccount } -ForeColor ([System.Drawing.Color]::Red)

$RightY += 8

Add-Section -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "OUs"
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get Top-Level OUs" -Action { Get-TopLevelOUs }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get Employee OUs" -Action { Get-EmployeeOUs }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get Computer OUs" -Action { Get-ComputerOUs }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get Server OUs" -Action { Get-ServerOUs }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get Disabled OUs" -Action { Get-DisabledOUs }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Get All Common OUs" -Action { Get-AllRequestedOUs }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Create OU" -Action { Show-CreateOU }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Move Object to OU" -Action { Show-MoveObjectToOU } -ForeColor ([System.Drawing.Color]::Red)

$RightY += 8

Add-Section -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Reports"
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "90-Day Inactive Users" -Action { Get-InactiveUsers90Days }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "90-Day Inactive Computers" -Action { Get-InactiveComputers90Days }
Add-Button -Panel $RightButtonPanel -YPosition ([ref]$RightY) -Text "Password Expiring Soon" -Action { Get-PasswordExpiringSoon }

Apply-CurrentTheme

$StartupNavy   = [System.Drawing.Color]::FromArgb(31,58,93)
$StartupBlue   = [System.Drawing.Color]::FromArgb(35,90,145)
$StartupGreen  = [System.Drawing.Color]::DarkGreen
$StartupGray   = [System.Drawing.Color]::DimGray
$StartupOrange = [System.Drawing.Color]::DarkOrange

Write-StartupSection "DIRECTORY CONNECTION"
Write-OutputField `
    -Label "Domain" `
    -Value $Domain `
    -LabelColor $StartupNavy `
    -ValueColor $StartupGreen
Write-OutputField `
    -Label "Domain Controllers" `
    -Value ($DCs -join ', ') `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Selected AD Server" `
    -Value (Get-SelectedServer) `
    -LabelColor $StartupNavy `
    -ValueColor $StartupGreen
Write-OutputBox `
    "User, group, computer, OU, and report actions will use the selected AD server." `
    $StartupGray

Write-StartupSection "DOMAIN TOOLS"
Write-OutputField `
    -Label "ADAC" `
    -Value "Opens Active Directory Administrative Center in its own window. It picks its own domain controller rather than following the AD Server selection." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "ADUC" `
    -Value "Opens in its own window and connects to the currently selected AD server." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "DNS" `
    -Value "Opens DNS Manager in its own window." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Computer MGMT" `
    -Value "Opens Computer Management connected to the currently selected AD server." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "GPO" `
    -Value "Opens Group Policy Management in its own window using its normal domain controller selection behavior." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Print MGMT" `
    -Value "Opens Print Management in its own window for print server and printer administration." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "RDP" `
    -Value "Opens Remote Desktop to the currently selected AD server." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "ISE" `
    -Value "Opens Windows PowerShell ISE in its own window." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "DC Health" `
    -Value "Read-only check of every DC: ADWS reachability and advertising, replication failures, and Security log retention." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Delta Sync" `
    -Value "Remotely starts a Microsoft Entra Connect delta synchronization on the server set in Settings." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Replication" `
    -Value "Use REPL Selected DC for one DC or REPL All DCs for all DCs." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue

Write-StartupSection "FILES AND AD LAUNCHERS"
Write-OutputField `
    -Label "Settings" `
    -Value "Domain, organizational units, naming and cloud IDs. Shared with every tool launched from here; blank values fall back to the domain root." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Logs" `
    -Value "Opens a list of the log and report files this session writes, with a preview pane." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Logs" `
    -Value "Save by default to the Logs folder beside this MSToolkit script: $LogPath" `
    -LabelColor $StartupNavy `
    -ValueColor $StartupGray
Write-OutputField `
    -Label "Create New User" `
    -Value "Launches NewADUser.ps1 from the same folder in its own window." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Compare/Manage User Groups" `
    -Value "Launches Compare-UserGroups.ps1 in its own window using the selected AD server." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Investigate Lockout" `
    -Value "Launches Investigate-AccountLockout.ps1 in its own window." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue

Write-StartupSection "MICROSOFT 365"
Write-OutputBox `
    "Block Teams Numbers, Exchange Online, Intune Tools, M365 Group Compare/Add, M365 Distro Compare/Add, M365 Conditional Access, and OneDrive / SharePoint each launch their own windows." `
    $StartupBlue
Write-OutputField `
    -Label "Block Teams Numbers" `
    -Value "Blocks inbound Teams caller numbers. Launches in its own window as your standard domain account, like the other M365 tools." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Exchange Online" `
    -Value "Inbox rules, message trace, group delivery, mailbox report, quarantine, header analysis, mail flow rules and folder permissions." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Intune Tools" `
    -Value "Intune and Entra administration, including Win32 app packaging. Signs in to Microsoft Graph with device code flow under your standard account." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "OneDrive / SharePoint" `
    -Value "Grant and remove access to a user's OneDrive, audit who has it, offboarding handover, restore a deleted OneDrive, storage and sharing report." `
    -LabelColor $StartupNavy `
    -ValueColor $StartupBlue
Write-OutputField `
    -Label "Authentication" `
    -Value "M365 tools prompt for a STANDARD domain account and authenticate to Microsoft 365 separately from this Domain Admin session." `
    -LabelColor $StartupOrange `
    -ValueColor $StartupOrange

# First run: nothing is configured yet, so open Settings once the window is up.
if (-not (Test-MSToolkitConfigured)) {
    Write-OutputBox ""
    Write-OutputBox "No settings are filled in for $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) yet. Opening Settings - fill in what applies and Save." ([System.Drawing.Color]::DarkOrange)
    Write-OutputBox "Settings opens on each launch until at least one value is saved; after that, use the Settings button." ([System.Drawing.Color]::DimGray)
    Write-OutputBox "Settings are kept per account: $($script:MSToolkitSettingsPath)" ([System.Drawing.Color]::DimGray)
    Write-OutputBox "Anything left blank falls back to the domain root, so the tool works either way." ([System.Drawing.Color]::DimGray)

    $form.Add_Shown({ Show-MSToolkitSettings })
}
else {
    # Settings are saved, but some OU buttons may still be falling back. Name them once
    # here so a blank OU is not only noticed when its button is clicked.
    $FallbackOUs = New-Object System.Collections.Generic.List[string]
    foreach ($Item in @(
        @{ Name = "OUUsers";          Caption = "Users" },
        @{ Name = "OUComputers";      Caption = "Computers" },
        @{ Name = "OUServers";        Caption = "Servers" },
        @{ Name = "OUDisabled";       Caption = "Disabled accounts" },
        @{ Name = "OUSecurityGroups"; Caption = "Security groups" }
    )) {
        if (-not (Get-MSToolkitOU -Name $Item.Name)) { $FallbackOUs.Add($Item.Caption) }
    }
    $ServiceOUMissing = -not (Get-MSToolkitOU -Name "OUServiceAccounts")

    if ($FallbackOUs.Count -gt 0 -or $ServiceOUMissing) {
        Write-OutputBox ""
        if ($FallbackOUs.Count -gt 0) {
            Write-OutputBox "Not set in Settings - these OU buttons use the domain root: $($FallbackOUs.ToArray() -join ', ')." ([System.Drawing.Color]::DarkOrange)
        }
        if ($ServiceOUMissing) {
            Write-OutputBox "Not set in Settings - Service accounts, which Get Special Function Accounts requires." ([System.Drawing.Color]::DarkOrange)
        }
        Write-OutputBox "Set them under Settings > Organizational units." ([System.Drawing.Color]::DimGray)
    }
}

try {
    $form.ShowDialog() | Out-Null
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "SCRIPT ERROR:`r`n$($_.Exception.Message)`r`n`r`n$($_.ScriptStackTrace)",
        "MSToolkit Error",
        "OK",
        "Error"
    )
}
