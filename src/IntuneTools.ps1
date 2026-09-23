<#
    IntuneTools.ps1
    Intune / Entra administration utility (PowerShell 5.1 + WinForms)

    Cloud operations run against Microsoft Graph using delegated interactive
    authentication (device code flow). Your own Intune RBAC and Conditional
    Access apply, and every action is attributed to you in the audit log.

    Local operations are limited to Win32 app packaging (IntuneWinAppUtil.exe).

    First run: open the Settings tab, enter Tenant ID and Client ID, Save,
    then click Connect.
#>

# Optional. MSToolkit passes its current theme so this tool opens to match; on its own
# it uses the theme saved in config.json, as before.
# OrgTenantId / OrgClientId: MSToolkit passes the Tenant ID and Client ID from its own
# Settings. They are used only by the "Set to Org Defaults" button - never filled in
# automatically.
param(
    [ValidateSet('Light','Dark')]
    [string]$ThemeMode,
    [string]$OrgTenantId,
    [string]$OrgClientId
)

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')       | Out-Null
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Win32 helpers used to bring this window back to the front after the browser
# takes focus during sign-in. Optional - the tool works without them.
$script:NativeWindowHelper = $false
try {
    if (-not ([System.Management.Automation.PSTypeName]'IntuneTools.Native').Type) {
        Add-Type -Namespace 'IntuneTools' -Name 'Native' -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@ -ErrorAction Stop
    }
    $script:NativeWindowHelper = $true
} catch {
    $script:NativeWindowHelper = $false
}

# ---------------------------------------------------------------------------
# Organization defaults
#   Nothing organization-specific is stored in this script. Tenant ID and Client ID
#   start blank and are saved on the Settings tab like any other setting. The
#   "Set to Org Defaults" button copies them from MSToolkit's Settings - see
#   Get-MSToolkitOrgDefaults. Neither value is a secret: the tenant ID identifies
#   the directory and the client ID identifies the app registration. There is no
#   client secret in this tool by design.
# ---------------------------------------------------------------------------
$script:DefaultPackagingRoot = ''    # e.g. '\\fileserver\share\IntunePackaging'  (optional)

# ---------------------------------------------------------------------------
# Script state
# ---------------------------------------------------------------------------
$script:AppDir      = Join-Path (Join-Path $env:APPDATA 'MSToolkit') 'IntuneTools'
$script:ConfigPath  = Join-Path $script:AppDir 'config.json'
$script:TokenPath   = Join-Path $script:AppDir 'token.dat'
$script:ScriptRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path

$script:AccessToken = $null
$script:TokenExpiry = [datetime]::MinValue
$script:Account     = $null
$script:TenantName  = $null

$script:Config = [ordered]@{
    TenantId            = ''
    ClientId            = ''
    PackagingRoot       = ''
    IntuneWinAppUtil    = ''
    StaleDays           = 90
    SignInMethod        = 'Browser'
    Theme               = 'Light'
    PkgRootFollowsUtil  = $false
}

$script:Cache = @{}      # grid name -> hashtable of id -> raw object

if (-not (Test-Path $script:AppDir)) {
    New-Item -Path $script:AppDir -ItemType Directory -Force | Out-Null
}

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
function Load-Config {
    if (Test-Path $script:ConfigPath) {
        try {
            $raw = Get-Content -Path $script:ConfigPath -Raw | ConvertFrom-Json
            foreach ($k in @($script:Config.Keys)) {
                if ($raw.PSObject.Properties.Name -contains $k) {
                    $script:Config[$k] = $raw.$k
                }
            }
        } catch {
            # corrupt config - ignore and use defaults
        }
    }
    # the organization default for the packaging root fills it when the local config has none
    if ([string]::IsNullOrWhiteSpace($script:Config.PackagingRoot) -and $script:DefaultPackagingRoot) { $script:Config.PackagingRoot = $script:DefaultPackagingRoot }

    if ([string]::IsNullOrWhiteSpace($script:Config.IntuneWinAppUtil)) {
        $guess = Join-Path $script:ScriptRoot 'IntuneWinAppUtil.exe'
        if (Test-Path $guess) { $script:Config.IntuneWinAppUtil = $guess }
    }
}

function Get-MSToolkitOrgDefaults {
    # Tenant ID and Client ID from MSToolkit's Settings, for "Set to Org Defaults".
    #   1. The values MSToolkit passed on launch (-OrgTenantId / -OrgClientId). MSToolkit
    #      runs as an admin account, so its settings.json is in that account's profile,
    #      which this tool - running as you - cannot read.
    #   2. Otherwise %APPDATA%\MSToolkit\settings.json for the account running this tool,
    #      for when MSToolkit and this tool run as the same account.
    $result = [pscustomobject]@{ TenantId = ''; ClientId = ''; Source = '' }

    if (-not [string]::IsNullOrWhiteSpace($OrgTenantId) -or -not [string]::IsNullOrWhiteSpace($OrgClientId)) {
        $result.TenantId = "$OrgTenantId".Trim()
        $result.ClientId = "$OrgClientId".Trim()
        $result.Source   = 'MSToolkit Settings (passed in when MSToolkit opened this tool)'
        return $result
    }

    try {
        $path = Join-Path (Join-Path $env:APPDATA 'MSToolkit') 'settings.json'
        if (Test-Path -LiteralPath $path) {
            $saved = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($saved.PSObject.Properties.Name -contains 'TenantId') { $result.TenantId = "$($saved.TenantId)".Trim() }
            if ($saved.PSObject.Properties.Name -contains 'ClientId') { $result.ClientId = "$($saved.ClientId)".Trim() }
            if ($result.TenantId -or $result.ClientId) { $result.Source = $path }
        }
    } catch {
        # Unreadable MSToolkit settings: report nothing found.
    }

    return $result
}

function Save-Config {
    try {
        [pscustomobject]$script:Config | ConvertTo-Json -Depth 4 |
            Set-Content -Path $script:ConfigPath -Encoding UTF8
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Could not save config:`r`n$($_.Exception.Message)",
            'IntuneTools', 'OK', 'Error') | Out-Null
        return $false
    }
}

# ---------------------------------------------------------------------------
# Token cache (DPAPI - current user, current machine only)
# ---------------------------------------------------------------------------
function Save-RefreshToken {
    param([string]$RefreshToken)
    if ([string]::IsNullOrWhiteSpace($RefreshToken)) { return }
    try {
        $sec = ConvertTo-SecureString -String $RefreshToken -AsPlainText -Force
        ConvertFrom-SecureString -SecureString $sec |
            Set-Content -Path $script:TokenPath -Encoding UTF8
    } catch { }
}

function Get-RefreshToken {
    if (-not (Test-Path $script:TokenPath)) { return $null }
    try {
        $enc = Get-Content -Path $script:TokenPath -Raw
        $sec = ConvertTo-SecureString -String $enc
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
        try   { return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch { return $null }
}

function Clear-TokenCache {
    $script:AccessToken = $null
    $script:TokenExpiry = [datetime]::MinValue
    $script:Account     = $null
    $script:TenantName  = $null
    if (Test-Path $script:TokenPath) { Remove-Item $script:TokenPath -Force -ErrorAction SilentlyContinue }
}

# ---------------------------------------------------------------------------
# Authentication - device code flow, no external modules required
# ---------------------------------------------------------------------------
function Get-TokenEndpoint {
    "https://login.microsoftonline.com/$($script:Config.TenantId)/oauth2/v2.0/token"
}

function Read-UpnFromToken {
    param([string]$Jwt)
    try {
        $payload = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
        if ($json.upn)               { return $json.upn }
        if ($json.preferred_username){ return $json.preferred_username }
        return $json.name
    } catch { return 'signed in' }
}

function Set-TokenFromResponse {
    param($Response)
    $script:AccessToken = $Response.access_token
    $script:TokenExpiry = (Get-Date).AddSeconds([int]$Response.expires_in - 120)
    $script:Account     = Read-UpnFromToken -Jwt $Response.access_token
    if ($Response.refresh_token) { Save-RefreshToken -RefreshToken $Response.refresh_token }
}

function Invoke-TokenRefresh {
    $rt = Get-RefreshToken
    if (-not $rt) { return $false }
    $body = @{
        client_id     = $script:Config.ClientId
        grant_type    = 'refresh_token'
        refresh_token = $rt
        scope         = 'https://graph.microsoft.com/.default offline_access'
    }
    try {
        $resp = Invoke-RestMethod -Method POST -Uri (Get-TokenEndpoint) -Body $body `
                    -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        Set-TokenFromResponse -Response $resp
        return $true
    } catch {
        Clear-TokenCache
        return $false
    }
}

function Start-DeviceCodeSignIn {
    $url  = "https://login.microsoftonline.com/$($script:Config.TenantId)/oauth2/v2.0/devicecode"
    $body = @{
        client_id = $script:Config.ClientId
        scope     = 'https://graph.microsoft.com/.default offline_access'
    }

    try {
        $dc = Invoke-RestMethod -Method POST -Uri $url -Body $body `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not start sign-in.`r`n`r`n$(Get-GraphErrorText $_)",
            'Sign in failed', 'OK', 'Error') | Out-Null
        return $false
    }

    # --- sign-in dialog -----------------------------------------------------
    $dlg               = New-Object System.Windows.Forms.Form
    $dlg.Text          = 'Sign in to Microsoft Graph'
    $dlg.Size          = New-Object System.Drawing.Size(480, 260)
    $dlg.StartPosition = 'CenterScreen'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox   = $false
    $dlg.MinimizeBox   = $false

    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = "1. Click Open Browser (or go to https://microsoft.com/devicelogin)`r`n2. Enter the code below`r`n3. Sign in with your admin account"
    $lbl.Location  = New-Object System.Drawing.Point(15, 15)
    $lbl.Size      = New-Object System.Drawing.Size(440, 60)
    $dlg.Controls.Add($lbl)

    $txtCode           = New-Object System.Windows.Forms.TextBox
    $txtCode.Text      = $dc.user_code
    $txtCode.ReadOnly  = $true
    $txtCode.Font      = New-Object System.Drawing.Font('Consolas', 16, [System.Drawing.FontStyle]::Bold)
    $txtCode.TextAlign = 'Center'
    $txtCode.Location  = New-Object System.Drawing.Point(15, 80)
    $txtCode.Size      = New-Object System.Drawing.Size(280, 40)
    $dlg.Controls.Add($txtCode)

    $btnCopy          = New-Object System.Windows.Forms.Button
    $btnCopy.Text     = 'Copy Code'
    $btnCopy.Location = New-Object System.Drawing.Point(305, 80)
    $btnCopy.Size     = New-Object System.Drawing.Size(140, 32)
    $btnCopy.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($txtCode.Text) })
    $dlg.Controls.Add($btnCopy)

    $btnOpen          = New-Object System.Windows.Forms.Button
    $btnOpen.Text     = 'Open Browser'
    $btnOpen.Location = New-Object System.Drawing.Point(15, 135)
    $btnOpen.Size     = New-Object System.Drawing.Size(140, 32)
    $btnOpen.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($txtCode.Text)
        Start-Process 'https://microsoft.com/devicelogin'
    })
    $dlg.Controls.Add($btnOpen)

    $lblStatus          = New-Object System.Windows.Forms.Label
    $lblStatus.Text     = 'Waiting for sign-in...'
    $lblStatus.Location = New-Object System.Drawing.Point(15, 180)
    $lblStatus.Size     = New-Object System.Drawing.Size(440, 40)
    $dlg.Controls.Add($lblStatus)

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = 'Cancel'
    $btnCancel.Location     = New-Object System.Drawing.Point(320, 135)
    $btnCancel.Size         = New-Object System.Drawing.Size(125, 32)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    $script:SignInOk = $false
    $interval = 5
    if ($dc.interval) { $interval = [int]$dc.interval }
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)

    $timer          = New-Object System.Windows.Forms.Timer
    $timer.Interval = $interval * 1000
    $timer.Add_Tick({
        if ((Get-Date) -gt $deadline) {
            $timer.Stop()
            $lblStatus.Text = 'Code expired. Close and try again.'
            return
        }
        $pollBody = @{
            client_id   = $script:Config.ClientId
            grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
            device_code = $dc.device_code
        }
        try {
            $resp = Invoke-RestMethod -Method POST -Uri (Get-TokenEndpoint) -Body $pollBody `
                        -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            Set-TokenFromResponse -Response $resp
            $script:SignInOk = $true
            $timer.Stop()
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        } catch {
            $err = Get-GraphErrorObject $_
            switch ($err.error) {
                'authorization_pending' { $lblStatus.Text = 'Waiting for sign-in...' }
                'slow_down'             { $timer.Interval = $timer.Interval + 5000 }
                'authorization_declined'{ $timer.Stop(); $lblStatus.Text = 'Sign-in was declined.' }
                'expired_token'         { $timer.Stop(); $lblStatus.Text = 'Code expired. Close and try again.' }
                default                 {
                    $timer.Stop()
                    $lblStatus.Text = "Error: $($err.error) - $($err.error_description)"
                }
            }
        }
    })
    $timer.Start()

    Set-ControlTheme -Control $dlg
    [void]$dlg.ShowDialog()
    $timer.Stop()
    $timer.Dispose()
    $dlg.Dispose()
    Show-MainWindow

    return $script:SignInOk
}

# ---------------------------------------------------------------------------
# Authentication - interactive browser sign-in (authorization code + PKCE)
#   Opens the normal Microsoft sign-in page in the default browser and catches
#   the reply on a loopback listener. No code to copy, no client secret.
#   Returns 'ok', 'cancelled' or 'unavailable'.
# ---------------------------------------------------------------------------
function New-PkcePair {
    $bytes = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($bytes)
    $verifier = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier))
    $challenge = [Convert]::ToBase64String($hash).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $sha.Dispose()
    return @{ Verifier = $verifier; Challenge = $challenge }
}

function Get-FreeLoopbackPort {
    $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $l.Start()
    $port = ([System.Net.IPEndPoint]$l.LocalEndpoint).Port
    $l.Stop()
    return $port
}

function Start-BrowserSignIn {
    param([switch]$ForceAccountPicker)

    $listener = $null
    try {
        $port     = Get-FreeLoopbackPort
        $redirect = "http://localhost:$port/"

        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add($redirect)
        try {
            $listener.Start()
        } catch {
            # loopback listener blocked (URL ACL, local firewall, policy)
            return 'unavailable'
        }

        $pkce  = New-PkcePair
        $state = [Guid]::NewGuid().ToString('N')
        $scope = 'https://graph.microsoft.com/.default offline_access'

        $authUrl = "https://login.microsoftonline.com/$($script:Config.TenantId)/oauth2/v2.0/authorize" +
            "?client_id=$([Uri]::EscapeDataString($script:Config.ClientId))" +
            "&response_type=code" +
            "&redirect_uri=$([Uri]::EscapeDataString($redirect))" +
            "&response_mode=query" +
            "&scope=$([Uri]::EscapeDataString($scope))" +
            "&state=$state" +
            "&code_challenge=$($pkce.Challenge)" +
            "&code_challenge_method=S256"
        if ($ForceAccountPicker) { $authUrl += "&prompt=select_account" }

        # small wait window so the user can cancel if they close the browser
        $wait               = New-Object System.Windows.Forms.Form
        $wait.Text          = 'Signing in...'
        $wait.Size          = New-Object System.Drawing.Size(440, 170)
        $wait.StartPosition = 'CenterScreen'
        $wait.FormBorderStyle = 'FixedDialog'
        $wait.MaximizeBox   = $false
        $wait.MinimizeBox   = $false

        $wl = New-Label -Text "A sign-in window has opened in your browser.`r`nComplete the sign-in there - this window closes on its own." -X 15 -Y 15 -W 400 -H 50
        $wait.Controls.Add($wl)

        $wb = New-Button -Text 'Open browser again' -X 15 -Y 80 -W 160 -H 30
        $wb.Add_Click({ Start-Process $authUrl })
        $wait.Controls.Add($wb)

        $script:SignInCancelled = $false
        $wc = New-Button -Text 'Cancel' -X 300 -Y 80 -W 110 -H 30
        $wc.Add_Click({ $script:SignInCancelled = $true })
        $wait.Controls.Add($wc)
        $wait.Add_FormClosing({ $script:SignInCancelled = $true })

        Set-ControlTheme -Control $wait
        $wait.Show()
        [System.Windows.Forms.Application]::DoEvents()
        Start-Process $authUrl

        # wait for the browser to come back, keeping the UI responsive
        $task     = $listener.GetContextAsync()
        $deadline = (Get-Date).AddSeconds(180)
        while (-not $task.AsyncWaitHandle.WaitOne(200)) {
            [System.Windows.Forms.Application]::DoEvents()
            if ($script:SignInCancelled) { break }
            if ((Get-Date) -gt $deadline) { break }
        }

        if (-not $task.IsCompleted) {
            $wait.Close(); $wait.Dispose()
            Show-MainWindow
            return 'cancelled'
        }

        $ctx    = $task.Result
        $code   = $ctx.Request.QueryString['code']
        $rstate = $ctx.Request.QueryString['state']
        $err    = $ctx.Request.QueryString['error']
        $errDsc = $ctx.Request.QueryString['error_description']

        $bodyHtml = if ($code) {
            @'
<html><head><title>Signed in</title></head>
<body style="font-family:Segoe UI;padding:40px;color:#202020">
<h2 id="h">Signed in</h2>
<p id="p">IntuneTools has the sign-in. Returning to the app...</p>
<script>
// Browsers only allow a script to close a window that a script opened. This
// tab was opened by the OS, so these attempts are usually refused - the page
// tidies itself up instead.
function bye() {
  try { window.close(); } catch (e) {}
  try { window.open('', '_self'); window.close(); } catch (e) {}
}
bye();
setTimeout(bye, 400);
setTimeout(function () {
  document.getElementById('h').textContent = 'Signed in';
  document.getElementById('p').textContent = 'You can close this tab.';
}, 1500);
</script>
</body></html>
'@
        } else {
            "<html><body style='font-family:Segoe UI;padding:40px'><h2>Sign-in failed</h2><p>$err</p><p>$errDsc</p></body></html>"
        }
        $buf = [Text.Encoding]::UTF8.GetBytes($bodyHtml)
        $ctx.Response.ContentType     = 'text/html'
        $ctx.Response.ContentLength64 = $buf.Length
        $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
        $ctx.Response.OutputStream.Close()

        $wait.Close(); $wait.Dispose()
        Show-MainWindow

        if ($err) {
            Show-ErrorBox "Sign-in failed:`r`n$err`r`n$errDsc"
            return 'cancelled'
        }
        if (-not $code)            { return 'cancelled' }
        if ($rstate -ne $state)    { Show-ErrorBox 'Sign-in state mismatch - the reply did not match this request. Try again.'; return 'cancelled' }

        $body = @{
            client_id     = $script:Config.ClientId
            grant_type    = 'authorization_code'
            code          = $code
            redirect_uri  = $redirect
            code_verifier = $pkce.Verifier
            scope         = $scope
        }
        try {
            $resp = Invoke-RestMethod -Method POST -Uri (Get-TokenEndpoint) -Body $body `
                        -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            Set-TokenFromResponse -Response $resp
            return 'ok'
        } catch {
            Show-ErrorBox "Token request failed:`r`n$(Get-GraphErrorText $_)"
            return 'cancelled'
        }
    } catch {
        Show-ErrorBox "Browser sign-in error:`r`n$($_.Exception.Message)"
        return 'unavailable'
    } finally {
        if ($listener) {
            try { $listener.Stop() } catch { }
            try { $listener.Close() } catch { }
        }
    }
}

function Connect-IntuneGraph {
    param([switch]$ForceAccountPicker)

    if ([string]::IsNullOrWhiteSpace($script:Config.TenantId) -or
        [string]::IsNullOrWhiteSpace($script:Config.ClientId)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Enter your Tenant ID and Client ID on the Settings tab first, then Save.',
            'IntuneTools', 'OK', 'Warning') | Out-Null
        return $false
    }

    # silent first - only prompt when there is no usable cached refresh token
    if (-not $ForceAccountPicker) {
        if (Invoke-TokenRefresh) { return $true }
    }

    if ($script:Config.SignInMethod -eq 'Device code') {
        return (Start-DeviceCodeSignIn)
    }

    $result = Start-BrowserSignIn -ForceAccountPicker:$ForceAccountPicker
    switch ($result) {
        'ok'        { return $true }
        'cancelled' { return $false }
        default {
            if (Confirm-Action -Message "The browser sign-in listener could not start on this machine.`r`n`r`nFall back to device code sign-in?" -Title 'Sign in') {
                return (Start-DeviceCodeSignIn)
            }
            return $false
        }
    }
}

function Test-GraphConnected {
    if ($script:AccessToken -and (Get-Date) -lt $script:TokenExpiry) { return $true }
    if ($script:AccessToken -or (Test-Path $script:TokenPath)) {
        if (Invoke-TokenRefresh) { return $true }
    }
    return $false
}

function Assert-Connected {
    if (Test-GraphConnected) { return $true }
    [System.Windows.Forms.MessageBox]::Show('Not connected. Click Connect first.',
        'IntuneTools', 'OK', 'Warning') | Out-Null
    return $false
}

# ---------------------------------------------------------------------------
# Graph error handling
# ---------------------------------------------------------------------------
function Get-GraphErrorObject {
    param($ErrorRecord)
    try {
        $resp = $ErrorRecord.Exception.Response
        if ($resp) {
            $stream = $resp.GetResponseStream()
            $stream.Position = 0
            $reader = New-Object IO.StreamReader($stream)
            $text = $reader.ReadToEnd()
            $reader.Close()
            if ($text) { return ($text | ConvertFrom-Json) }
        }
    } catch { }
    return [pscustomobject]@{ error = 'unknown'; error_description = $ErrorRecord.Exception.Message }
}

function Get-GraphErrorText {
    param($ErrorRecord)
    $o = Get-GraphErrorObject $ErrorRecord
    if ($o.error -and $o.error.message) { return "$($o.error.code): $($o.error.message)" }   # Graph style
    if ($o.error_description)           { return "$($o.error): $($o.error_description)" }    # OAuth style
    return $ErrorRecord.Exception.Message
}

# ---------------------------------------------------------------------------
# Graph request wrapper
#   -Uri may be a relative path ('/deviceManagement/managedDevices') or a full URL
#   -All follows @odata.nextLink and returns every page
# ---------------------------------------------------------------------------
function Invoke-Graph {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        $Body,
        [switch]$Beta,
        [switch]$All,
        [switch]$Raw
    )

    if ($Uri -notmatch '^https?://') {
        $ver = if ($Beta) { 'beta' } else { 'v1.0' }
        $Uri = "https://graph.microsoft.com/$ver" + $Uri
    }

    $results = New-Object System.Collections.Generic.List[object]
    $next    = $Uri
    $retries = 0

    while ($next) {
        $headers = @{ Authorization = "Bearer $script:AccessToken" }
        if ($next -match '/(groups|users|directory)') { $headers['ConsistencyLevel'] = 'eventual' }
        $params = @{
            Method      = $Method
            Uri         = $next
            Headers     = $headers
            ErrorAction = 'Stop'
        }
        if ($null -ne $Body) {
            $params.Body        = if ($Body -is [string]) { $Body } else { ($Body | ConvertTo-Json -Depth 10) }
            $params.ContentType = 'application/json'
        }

        try {
            $resp = Invoke-RestMethod @params
        } catch {
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }

            if ($status -eq 401 -and $retries -lt 1) {
                $retries++
                if (Invoke-TokenRefresh) { continue }
                throw "Session expired. Click Connect again."
            }
            if ($status -eq 429 -and $retries -lt 5) {
                $retries++
                $wait = 10
                try { if ($_.Exception.Response.Headers['Retry-After']) { $wait = [int]$_.Exception.Response.Headers['Retry-After'] } } catch { }
                Start-Sleep -Seconds $wait
                continue
            }
            throw (Get-GraphErrorText $_)
        }

        if ($Raw) { return $resp }

        if ($null -ne $resp -and $resp.PSObject.Properties.Name -contains 'value') {
            foreach ($v in $resp.value) { [void]$results.Add($v) }
        } elseif ($null -ne $resp) {
            [void]$results.Add($resp)
        }

        if ($All -and $resp.PSObject.Properties.Name -contains '@odata.nextLink') {
            $next = $resp.'@odata.nextLink'
        } else {
            $next = $null
        }
    }

    return $results.ToArray()
}

# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Theming - light / dark
# ---------------------------------------------------------------------------
$script:Theme            = $null
$script:IsConnected      = $false
$script:LastWindowState  = [System.Windows.Forms.FormWindowState]::Maximized

function Get-ThemePalette {
    # Colour set matched to MSToolkit so both utilities look like one family.
    param([string]$Name)

    function C { param([int]$R, [int]$G, [int]$B) [System.Drawing.Color]::FromArgb($R, $G, $B) }

    if ($Name -eq 'Dark') {
        return @{
            Name           = 'Dark'
            MainBack       = (C 30 32 36)
            PanelBack      = (C 38 41 46)
            InputBack      = (C 45 48 54)
            OutputBack     = (C 24 26 29)
            Text           = (C 232 234 237)
            MutedText      = (C 174 180 187)
            Border         = (C 78 84 92)
            Section        = (C 122 181 238)
            TopBar         = (C 24 47 74)
            Accent         = (C 0 120 215)
            AccentHover    = (C 20 135 225)
            ButtonBack     = (C 49 53 59)
            ButtonHover    = (C 60 65 72)
            SelectionBack  = (C 55 105 155)
            GridAltBack    = (C 42 45 50)
            Success        = (C 118 210 142)
            Danger         = (C 255 125 125)
            Warning        = (C 255 184 92)
            Info           = (C 125 190 245)
            Separator      = (C 125 132 140)
            DisabledBack   = (C 58 61 67)
            DisabledText   = (C 165 170 177)
            DisabledBorder = (C 92 97 105)
        }
    }

    return @{
        Name           = 'Light'
        MainBack       = (C 245 247 250)
        PanelBack      = [System.Drawing.Color]::White
        InputBack      = [System.Drawing.Color]::White
        OutputBack     = [System.Drawing.Color]::White
        Text           = (C 35 35 35)
        MutedText      = [System.Drawing.Color]::DimGray
        Border         = (C 210 215 220)
        Section        = (C 31 58 93)
        TopBar         = (C 31 58 93)
        Accent         = (C 0 120 215)
        AccentHover    = (C 20 135 225)
        ButtonBack     = [System.Drawing.Color]::White
        ButtonHover    = (C 242 246 250)
        SelectionBack  = (C 0 120 215)
        GridAltBack    = (C 250 251 252)
        Success        = [System.Drawing.Color]::ForestGreen
        Danger         = [System.Drawing.Color]::Red
        Warning        = [System.Drawing.Color]::DarkOrange
        Info           = (C 35 90 145)
        Separator      = (C 140 140 140)
        DisabledBack   = (C 238 240 243)
        DisabledText   = (C 125 130 136)
        DisabledBorder = (C 195 200 206)
    }
}

function Set-ControlTheme {
    param([System.Windows.Forms.Control]$Control)

    $T = $script:Theme
    if (-not $T -or -not $Control) { return }

    $role = [string]$Control.Tag

    if ($Control -is [System.Windows.Forms.DataGridView]) {
        $Control.EnableHeadersVisualStyles = $false
        $Control.BackgroundColor = $T.PanelBack
        $Control.GridColor       = $T.Border
        $Control.ForeColor       = $T.Text
        $Control.BorderStyle     = 'FixedSingle'
        $Control.DefaultCellStyle.BackColor          = $T.PanelBack
        $Control.DefaultCellStyle.ForeColor          = $T.Text
        $Control.DefaultCellStyle.SelectionBackColor = $T.SelectionBack
        $Control.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
        $Control.AlternatingRowsDefaultCellStyle.BackColor = $T.GridAltBack
        $Control.AlternatingRowsDefaultCellStyle.ForeColor = $T.Text
        $Control.ColumnHeadersDefaultCellStyle.BackColor          = $T.ButtonBack
        $Control.ColumnHeadersDefaultCellStyle.ForeColor          = $T.Text
        $Control.ColumnHeadersDefaultCellStyle.SelectionBackColor = $T.ButtonBack
        $Control.ColumnHeadersDefaultCellStyle.SelectionForeColor = $T.Text
        $Control.RowHeadersDefaultCellStyle.BackColor = $T.ButtonBack
    }
    elseif ($Control -is [System.Windows.Forms.Form]) {
        $Control.BackColor = $T.MainBack
        $Control.ForeColor = $T.Text
    }
    elseif ($Control -is [System.Windows.Forms.Button]) {
        $Control.UseVisualStyleBackColor = $false
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $Control.FlatAppearance.BorderSize  = 1
        $Control.FlatAppearance.BorderColor = $T.Border
        $Control.FlatAppearance.MouseOverBackColor = $T.ButtonHover

        $onTopBar = ($Control.Parent -and ([string]$Control.Parent.Tag -eq 'TopBar'))
        if ($onTopBar) {
            $Control.BackColor = $T.Accent
            $Control.ForeColor = [System.Drawing.Color]::White
            $Control.FlatAppearance.BorderSize = 0
            $Control.FlatAppearance.MouseOverBackColor = $T.AccentHover
        } elseif ($role -eq 'danger') {
            $Control.BackColor = $T.ButtonBack
            $Control.ForeColor = $T.Danger
        } elseif ($role -eq 'success') {
            $Control.BackColor = $T.ButtonBack
            $Control.ForeColor = $T.Success
        } else {
            $Control.BackColor = $T.ButtonBack
            $Control.ForeColor = $T.Text
        }

        if (-not $Control.Enabled) {
            $Control.BackColor = $T.DisabledBack
            $Control.ForeColor = $T.DisabledText
            $Control.FlatAppearance.BorderSize  = 1
            $Control.FlatAppearance.BorderColor = $T.DisabledBorder
            $Control.FlatAppearance.MouseOverBackColor = $T.DisabledBack
        }
    }
    elseif ($Control -is [System.Windows.Forms.RichTextBox]) {
        $Control.BackColor = $T.OutputBack
        $Control.ForeColor = $T.Text
    }
    elseif ($Control -is [System.Windows.Forms.TextBox]) {
        if ($role -eq 'output' -or $Control.ReadOnly -and $Control.Multiline) {
            $Control.BackColor = $T.OutputBack
        } else {
            $Control.BackColor = $T.InputBack
        }
        $Control.ForeColor = $T.Text
    }
    elseif ($Control -is [System.Windows.Forms.ListBox] -or
            $Control -is [System.Windows.Forms.ComboBox]) {
        $Control.BackColor = $T.InputBack
        $Control.ForeColor = $T.Text
    }
    elseif ($Control -is [System.Windows.Forms.StatusStrip]) {
        $Control.BackColor = $T.PanelBack
        $Control.ForeColor = $T.Text
        foreach ($item in $Control.Items) { $item.ForeColor = $T.Text }
    }
    elseif ($Control -is [System.Windows.Forms.Label]) {
        if ($Control.Parent -and ([string]$Control.Parent.Tag -eq 'TopBar')) {
            $Control.BackColor = [System.Drawing.Color]::Transparent
            if ($role -ne 'status') { $Control.ForeColor = [System.Drawing.Color]::White }
        } else {
            $Control.BackColor = [System.Drawing.Color]::Transparent
            if ($role -eq 'section') {
                $Control.ForeColor = $T.Section
            } elseif ($role -ne 'status') {
                $Control.ForeColor = $T.Text
            }
        }
    }
    elseif ($Control -is [System.Windows.Forms.CheckBox] -or
            $Control -is [System.Windows.Forms.RadioButton]) {
        $Control.BackColor = [System.Drawing.Color]::Transparent
        $Control.ForeColor = $T.Text
    }
    elseif ($Control -is [System.Windows.Forms.TabPage]) {
        $Control.BackColor = $T.MainBack
        $Control.ForeColor = $T.Text
    }
    elseif ($role -eq 'TopBar') {
        $Control.BackColor = $T.TopBar
        $Control.ForeColor = [System.Drawing.Color]::White
    }
    else {
        $Control.BackColor = $T.MainBack
        $Control.ForeColor = $T.Text
    }

    foreach ($child in $Control.Controls) { Set-ControlTheme -Control $child }
}

function Set-AppTheme {
    param([string]$Name)

    if ($Name -ne 'Dark') { $Name = 'Light' }
    $script:Theme        = Get-ThemePalette -Name $Name
    $script:Config.Theme = $Name

    if ($script:MainForm) {
        $script:MainForm.SuspendLayout()
        Set-ControlTheme -Control $script:MainForm
        $script:MainForm.ResumeLayout()
        $script:MainForm.Refresh()
    }

    if ($script:BtnTheme) {
        # the glyph shows the mode you would switch TO
        if ($Name -eq 'Dark') {
            $script:BtnTheme.Text = [string][char]0x2600      # sun
            if ($script:ThemeTip) { $script:ThemeTip.SetToolTip($script:BtnTheme, 'Switch to light mode') }
        } else {
            $script:BtnTheme.Text = [string][char]0x263E      # moon
            if ($script:ThemeTip) { $script:ThemeTip.SetToolTip($script:BtnTheme, 'Switch to dark mode') }
        }
    }

    # the Settings help pane colours each run individually, so redraw it
    if ($script:SettingsHelpBox) { Write-SettingsHelp -Box $script:SettingsHelpBox }

    Set-ConnectedState $script:IsConnected
}

function Show-MainWindow {
    # Called after the browser has had focus during sign-in. Windows blocks a
    # plain Activate() from stealing foreground, so this uses the TopMost
    # toggle plus SetForegroundWindow, and falls back quietly if either fails.
    if (-not $script:MainForm -or $script:MainForm.IsDisposed) { return }
    try {
        $minimized = ($script:MainForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized)

        if ($minimized) {
            # SW_RESTORE un-minimises. It must ONLY be used on a minimised
            # window: on a maximised one it restores to the smaller previous
            # size, which looks like the app deliberately shrinking itself.
            if ($script:NativeWindowHelper -and $script:MainForm.Handle -ne [IntPtr]::Zero) {
                [IntuneTools.Native]::ShowWindow($script:MainForm.Handle, 9) | Out-Null   # SW_RESTORE
            }
            $target = $script:LastWindowState
            if (-not $target) { $target = [System.Windows.Forms.FormWindowState]::Maximized }
            $script:MainForm.WindowState = $target
        }

        $wasTopMost = $script:MainForm.TopMost
        $script:MainForm.TopMost = $true
        $script:MainForm.Activate()
        $script:MainForm.BringToFront()
        [System.Windows.Forms.Application]::DoEvents()
        $script:MainForm.TopMost = $wasTopMost

        # bring to the foreground without touching the size
        if ($script:NativeWindowHelper -and $script:MainForm.Handle -ne [IntPtr]::Zero) {
            [IntuneTools.Native]::SetForegroundWindow($script:MainForm.Handle) | Out-Null
        }

        $script:MainForm.Focus() | Out-Null
        [System.Windows.Forms.Application]::DoEvents()
    } catch { }
}

function Set-Status {
    param([string]$Text)
    if ($script:StatusLabel) {
        $script:StatusLabel.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Start-Busy {
    param([string]$Text = 'Working...')
    $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Set-Status $Text
}

function Stop-Busy {
    param([string]$Text = 'Ready')
    $script:MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    Set-Status $Text
}

function Show-ErrorBox {
    param([string]$Message, [string]$Title = 'IntuneTools')
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, 'OK', 'Error') | Out-Null
}

function Show-InfoBox {
    param([string]$Message, [string]$Title = 'IntuneTools')
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, 'OK', 'Information') | Out-Null
}

function Confirm-Action {
    param([string]$Message, [string]$Title = 'Confirm')
    $r = [System.Windows.Forms.MessageBox]::Show($Message, $Title, 'YesNo', 'Warning', 'Button2')
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Format-GraphDate {
    param($Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    try { return ([datetime]$Value).ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
    catch { return [string]$Value }
}

function New-DataTableFromObjects {
    param([object[]]$Objects, [string[]]$Columns)

    $dt = New-Object System.Data.DataTable
    if (-not $Columns -or $Columns.Count -eq 0) {
        if ($Objects -and $Objects.Count -gt 0) {
            $Columns = @($Objects[0].PSObject.Properties.Name)
        } else {
            $Columns = @('Result')
        }
    }
    foreach ($c in $Columns) { [void]$dt.Columns.Add($c, [string]) }

    foreach ($o in $Objects) {
        if ($null -eq $o) { continue }
        $row = $dt.NewRow()
        foreach ($c in $Columns) {
            $v = $null
            try { $v = $o.$c } catch { }
            if ($null -eq $v) { $row[$c] = '' }
            elseif ($v -is [array]) { $row[$c] = ($v -join '; ') }
            else { $row[$c] = [string]$v }
        }
        $dt.Rows.Add($row)
    }
    ,$dt
}

function New-Grid {
    param(
        [int]$X, [int]$Y, [int]$W, [int]$H,
        [string]$Anchor = 'Top,Left,Right,Bottom',
        [switch]$MultiSelect
    )
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location             = New-Object System.Drawing.Point($X, $Y)
    $g.Size                 = New-Object System.Drawing.Size($W, $H)
    $g.Anchor               = $Anchor
    $g.ReadOnly             = $true
    $g.AllowUserToAddRows   = $false
    $g.AllowUserToDeleteRows= $false
    $g.AllowUserToOrderColumns = $true
    $g.SelectionMode        = 'FullRowSelect'
    $g.MultiSelect          = [bool]$MultiSelect
    $g.RowHeadersVisible    = $false
    $g.AutoSizeColumnsMode  = 'DisplayedCells'
    $g.BackgroundColor      = [System.Drawing.Color]::White
    $g.Font                 = New-Object System.Drawing.Font('Segoe UI', 9)
    return $g
}

function Set-GridData {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [object[]]$Objects,
        [string[]]$Columns,
        [string]$CacheKey,
        [string]$IdProperty = 'id'
    )
    $dt = New-DataTableFromObjects -Objects $Objects -Columns $Columns
    $Grid.DataSource = $dt

    if ($CacheKey) {
        $h = @{}
        foreach ($o in $Objects) {
            if ($null -eq $o) { continue }
            $key = $null
            try { $key = [string]$o.$IdProperty } catch { }
            if ($key) { $h[$key] = $o }
        }
        $script:Cache[$CacheKey] = $h
    }
    if ($Objects) { return $Objects.Count } else { return 0 }
}

function Get-SelectedCellValue {
    param([System.Windows.Forms.DataGridView]$Grid, [string]$Column)
    if ($Grid.SelectedRows.Count -eq 0) { return $null }
    $row = $Grid.SelectedRows[0]
    if (-not $row.DataBoundItem) { return $null }
    try { return [string]$row.DataBoundItem.Row[$Column] } catch { return $null }
}

function Get-SelectedCellValues {
    param([System.Windows.Forms.DataGridView]$Grid, [string]$Column)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($row in $Grid.SelectedRows) {
        if (-not $row.DataBoundItem) { continue }
        try {
            $v = [string]$row.DataBoundItem.Row[$Column]
            if ($v) { [void]$out.Add($v) }
        } catch { }
    }
    return $out.ToArray()
}

function Get-CachedObject {
    param([string]$CacheKey, [string]$Id)
    if (-not $Id) { return $null }
    if ($script:Cache.ContainsKey($CacheKey)) { return $script:Cache[$CacheKey][$Id] }
    return $null
}

function Export-GridToCsv {
    param([System.Windows.Forms.DataGridView]$Grid, [string]$SuggestedName = 'export')
    if (-not $Grid.DataSource) { Show-InfoBox 'Nothing to export.'; return }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV file (*.csv)|*.csv'
    $dlg.FileName = "$SuggestedName`_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $dt   = $Grid.DataSource
        $cols = @($dt.Columns | ForEach-Object { $_.ColumnName })
        $rows = foreach ($r in $dt.Rows) {
            $o = [ordered]@{}
            foreach ($c in $cols) { $o[$c] = $r[$c] }
            [pscustomobject]$o
        }
        $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Show-InfoBox "Exported $($dt.Rows.Count) rows to:`r`n$($dlg.FileName)"
    } catch {
        Show-ErrorBox "Export failed:`r`n$($_.Exception.Message)"
    }
}

function Show-JsonWindow {
    param([string]$Title, $Object)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.Size = New-Object System.Drawing.Size(900, 700)
    $f.StartPosition = 'CenterParent'

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline  = $true
    $tb.ScrollBars = 'Both'
    $tb.WordWrap   = $false
    $tb.ReadOnly   = $true
    $tb.Dock       = 'Fill'
    $tb.Font       = New-Object System.Drawing.Font('Consolas', 9)
    try { $tb.Text = ($Object | ConvertTo-Json -Depth 20) } catch { $tb.Text = [string]$Object }
    $f.Controls.Add($tb)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = 44

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy'
    $btnCopy.Location = New-Object System.Drawing.Point(10, 8)
    $btnCopy.Size = New-Object System.Drawing.Size(100, 28)
    $btnCopy.Add_Click({ if ($tb.Text) { [System.Windows.Forms.Clipboard]::SetText($tb.Text) } })
    $panel.Controls.Add($btnCopy)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Save JSON...'
    $btnSave.Location = New-Object System.Drawing.Point(120, 8)
    $btnSave.Size = New-Object System.Drawing.Size(120, 28)
    $btnSave.Add_Click({
        $sd = New-Object System.Windows.Forms.SaveFileDialog
        $sd.Filter = 'JSON file (*.json)|*.json'
        $sd.FileName = (($Title -replace '[^\w\-\. ]', '_') + '.json')
        if ($sd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Set-Content -Path $sd.FileName -Value $tb.Text -Encoding UTF8
        }
    })
    $panel.Controls.Add($btnSave)

    $f.Controls.Add($panel)
    Set-ControlTheme -Control $f
    [void]$f.ShowDialog()
    $f.Dispose()
}

function New-Button {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 120, [int]$H = 28, [string]$Anchor = 'Top,Left')
    $b = New-Object System.Windows.Forms.Button
    $b.Text     = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size     = New-Object System.Drawing.Size($W, $H)
    $b.Anchor   = $Anchor
    return $b
}

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 120, [int]$H = 20)
    $l = New-Object System.Windows.Forms.Label
    $l.Text     = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size     = New-Object System.Drawing.Size($W, $H)
    $l.TextAlign = 'MiddleLeft'
    return $l
}

function New-TextBox {
    param([int]$X, [int]$Y, [int]$W = 200, [string]$Text = '', [string]$Anchor = 'Top,Left')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size     = New-Object System.Drawing.Size($W, 24)
    $t.Text     = $Text
    $t.Anchor   = $Anchor
    return $t
}

function Escape-ODataValue {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return $Value.Replace("'", "''")
}

# ---------------------------------------------------------------------------
# Main form
# ---------------------------------------------------------------------------
Load-Config

$script:MainForm               = New-Object System.Windows.Forms.Form
$script:MainForm.Text          = 'IntuneTools'
$script:MainForm.Size          = New-Object System.Drawing.Size(1400, 900)
$script:MainForm.StartPosition = 'CenterScreen'
$script:MainForm.WindowState   = 'Maximized'
$script:MainForm.MinimumSize   = New-Object System.Drawing.Size(1100, 700)
$script:MainForm.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

# --- top bar ---------------------------------------------------------------
$topBar        = New-Object System.Windows.Forms.Panel
$topBar.Dock   = 'Top'
$topBar.Height = 44
$topBar.Tag    = 'TopBar'
$script:MainForm.Controls.Add($topBar)

$script:BtnConnect = New-Button -Text 'Connect' -X 10 -Y 8 -W 110 -H 28
$topBar.Controls.Add($script:BtnConnect)

$script:BtnSignOut = New-Button -Text 'Sign Out' -X 128 -Y 8 -W 100 -H 28
$script:BtnSignOut.Enabled = $false
$topBar.Controls.Add($script:BtnSignOut)

$script:LblAccount = New-Label -Text 'Not connected' -X 240 -Y 8 -W 700 -H 28
$script:LblAccount.Tag = 'status'
$script:LblAccount.ForeColor = [System.Drawing.Color]::DimGray
$topBar.Controls.Add($script:LblAccount)

# light / dark toggle - top right
$script:ThemeTip = New-Object System.Windows.Forms.ToolTip
$script:BtnTheme = New-Button -Text ([string][char]0x263E) -X 1330 -Y 6 -W 42 -H 30 -Anchor 'Top,Right'
$script:BtnTheme.Font      = New-Object System.Drawing.Font('Segoe UI Symbol', 12)
$script:BtnTheme.TextAlign = 'MiddleCenter'
$script:BtnTheme.Add_Click({
    if ($script:Config.Theme -eq 'Dark') { Set-AppTheme 'Light' } else { Set-AppTheme 'Dark' }
    Save-Config | Out-Null
})
$topBar.Controls.Add($script:BtnTheme)

# --- status bar ------------------------------------------------------------
$statusStrip        = New-Object System.Windows.Forms.StatusStrip
$script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusLabel.Text = 'Ready'
[void]$statusStrip.Items.Add($script:StatusLabel)
$script:MainForm.Controls.Add($statusStrip)

# --- tabs ------------------------------------------------------------------
# Tab label fonts are kept separate from the TabControl's own Font: the pages
# and their controls inherit that font, and the layouts are positioned
# absolutely, so changing it would shift them.
$script:TabFont     = New-Object System.Drawing.Font('Segoe UI', 10.5)
$script:TabFontSel  = New-Object System.Drawing.Font('Segoe UI', 10.5, [System.Drawing.FontStyle]::Bold)

$script:Tabs          = New-Object System.Windows.Forms.TabControl
$script:Tabs.Dock     = 'Fill'
$script:Tabs.DrawMode = 'OwnerDrawFixed'
$script:Tabs.SizeMode = 'Fixed'
$script:Tabs.ItemSize = New-Object System.Drawing.Size(128, 34)
$script:Tabs.Add_DrawItem({
    param($sender, $e)
    $T = $script:Theme
    if (-not $T) { $T = Get-ThemePalette -Name 'Light' }
    $selected = (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0)
    $back = if ($selected) { $T.PanelBack } else { $T.MainBack }
    $fore = if ($selected) { $T.Section }   else { $T.MutedText }
    $font = if ($selected) { $script:TabFontSel } else { $script:TabFont }

    # Settings sits apart from the working tabs, so it is labelled in green
    if ($sender.TabPages[$e.Index].Text -eq 'Settings') { $fore = $T.Success }

    $bb = New-Object System.Drawing.SolidBrush($back)
    $fb = New-Object System.Drawing.SolidBrush($fore)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment     = 'Center'
    $sf.LineAlignment = 'Center'

    $e.Graphics.FillRectangle($bb, $e.Bounds)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $rect = New-Object System.Drawing.RectangleF($e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width, $e.Bounds.Height)
    $e.Graphics.DrawString($sender.TabPages[$e.Index].Text, $font, $fb, $rect, $sf)

    # underline the active tab so the selection reads at a glance
    if ($selected) {
        $accent = New-Object System.Drawing.SolidBrush($T.Accent)
        $e.Graphics.FillRectangle($accent, $e.Bounds.X, ($e.Bounds.Bottom - 3), $e.Bounds.Width, 3)
        $accent.Dispose()
    }

    $bb.Dispose(); $fb.Dispose(); $sf.Dispose()
})
$script:MainForm.Controls.Add($script:Tabs)
$script:Tabs.BringToFront()

function New-Tab {
    param([string]$Text)
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text    = $Text
    $p.Padding = New-Object System.Windows.Forms.Padding(8)
    [void]$script:Tabs.TabPages.Add($p)
    return $p
}

$tabPackaging = New-Tab 'Packaging'
$tabDevices   = New-Tab 'Devices'
$tabAutopilot = New-Tab 'Autopilot'
$tabApps      = New-Tab 'Apps'
$tabPolicies  = New-Tab 'Policies'
$tabGroups    = New-Tab 'Group Lookup'
$tabReports   = New-Tab 'Reports'
$tabSettings  = New-Tab 'Settings'

function Get-TenantDisplayName {
    # Resolve the tenant's friendly name once per session. Falls back to the
    # initial .onmicrosoft.com domain, then to the tenant ID.
    if ($script:TenantName) { return $script:TenantName }
    try {
        $org = Invoke-Graph -Uri "/organization?`$select=id,displayName,verifiedDomains" -Raw
        $o = $null
        if ($org.value) { $o = @($org.value)[0] } else { $o = $org }

        if ($o.displayName) {
            $script:TenantName = $o.displayName
        } elseif ($o.verifiedDomains) {
            $initial = @($o.verifiedDomains | Where-Object { $_.isInitial })
            if ($initial.Count -gt 0) { $script:TenantName = $initial[0].name }
        }
    } catch {
        $script:TenantName = $null
    }
    if (-not $script:TenantName) { $script:TenantName = $script:Config.TenantId }
    return $script:TenantName
}

function Set-ConnectedState {
    param([bool]$Connected)
    $script:IsConnected = $Connected
    $accent = if ($script:Theme) { $script:Theme.Success }   else { [System.Drawing.Color]::ForestGreen }
    $dim    = if ($script:Theme) { $script:Theme.MutedText } else { [System.Drawing.Color]::DimGray }
    if ($Connected) {
        $tenant = $script:Config.TenantId
        try { $tenant = Get-TenantDisplayName } catch { }
        $script:LblAccount.Text = "Connected as $($script:Account)  |  $tenant"
        $script:LblAccount.ForeColor = $accent
        $script:BtnConnect.Text = 'Reconnect'
        $script:BtnSignOut.Enabled = $true
    } else {
        $script:LblAccount.Text = 'Not connected'
        $script:LblAccount.ForeColor = $dim
        $script:BtnConnect.Text = 'Connect'
        $script:BtnSignOut.Enabled = $false
    }
}

$script:BtnConnect.Add_Click({
    $forcePicker = ($script:BtnConnect.Text -eq 'Reconnect')
    Start-Busy 'Signing in...'
    try {
        if (Connect-IntuneGraph -ForceAccountPicker:$forcePicker) {
            Set-ConnectedState $true
            Stop-Busy "Connected as $($script:Account)"
        } else {
            Set-ConnectedState $false
            Stop-Busy 'Not connected'
        }
    } catch {
        Stop-Busy 'Sign-in failed'
        Show-ErrorBox $_.Exception.Message
    }
})

$script:BtnSignOut.Add_Click({
    Clear-TokenCache
    Set-ConnectedState $false
    Set-Status 'Signed out'
})

# ===========================================================================
# DEVICES TAB
# ===========================================================================
$devTop        = New-Object System.Windows.Forms.Panel
$devTop.Dock   = 'Top'
$devTop.Height = 76
$tabDevices.Controls.Add($devTop)

$devTop.Controls.Add((New-Label -Text 'Search:' -X 4 -Y 8 -W 55 -H 24))
$script:TxtDevSearch = New-TextBox -X 60 -Y 6 -W 260
$devTop.Controls.Add($script:TxtDevSearch)

$devTop.Controls.Add((New-Label -Text 'View:' -X 332 -Y 8 -W 40 -H 24))
$script:CmbDevFilter = New-Object System.Windows.Forms.ComboBox
$script:CmbDevFilter.Location      = New-Object System.Drawing.Point(374, 6)
$script:CmbDevFilter.Size          = New-Object System.Drawing.Size(200, 24)
$script:CmbDevFilter.DropDownStyle = 'DropDownList'
[void]$script:CmbDevFilter.Items.AddRange(@('All devices','Windows only','Non-compliant','Stale (no sync)','Personal owned'))
$script:CmbDevFilter.SelectedIndex = 0
$devTop.Controls.Add($script:CmbDevFilter)

$script:BtnDevLoad   = New-Button -Text 'Load Devices' -X 586 -Y 5 -W 120
$script:BtnDevExport = New-Button -Text 'Export CSV'   -X 714 -Y 5 -W 100

$devTop.Controls.AddRange(@($script:BtnDevLoad, $script:BtnDevExport))

# action buttons (second row)
$script:BtnDevSync     = New-Button -Text 'Sync'              -X 4   -Y 42 -W 90
$script:BtnDevRestart  = New-Button -Text 'Restart'           -X 98  -Y 42 -W 90
$script:BtnDevShutdown = New-Button -Text 'Shut Down'         -X 192 -Y 42 -W 90
$script:BtnDevScan     = New-Button -Text 'Defender Scan'     -X 286 -Y 42 -W 110
$script:BtnDevRotateBL = New-Button -Text 'Rotate BitLocker'  -X 400 -Y 42 -W 120
$script:BtnDevRotateLA = New-Button -Text 'Rotate LAPS'       -X 524 -Y 42 -W 100
$script:BtnDevJson     = New-Button -Text 'View Raw JSON'     -X 628 -Y 42 -W 110
$script:BtnDevManage   = New-Button -Text 'Manage...'         -X 742 -Y 42 -W 100
$script:BtnDevManage.Tag       = 'success'
$script:BtnDevManage.ForeColor = [System.Drawing.Color]::ForestGreen
$script:BtnDevDanger   = New-Button -Text 'Destructive...'    -X 848 -Y 42 -W 110
$script:BtnDevDanger.Tag       = 'danger'
$script:BtnDevDanger.ForeColor = [System.Drawing.Color]::Firebrick

$devTop.Controls.AddRange(@(
    $script:BtnDevSync, $script:BtnDevRestart, $script:BtnDevShutdown, $script:BtnDevScan,
    $script:BtnDevRotateBL, $script:BtnDevRotateLA, $script:BtnDevJson, $script:BtnDevManage, $script:BtnDevDanger))

$devSplit             = New-Object System.Windows.Forms.SplitContainer
$devSplit.Dock        = 'Fill'
$devSplit.Orientation = 'Horizontal'
$tabDevices.Controls.Add($devSplit)
$devSplit.BringToFront()

$script:GridDevices = New-Grid -X 0 -Y 0 -W 100 -H 100 -MultiSelect
$script:GridDevices.Dock = 'Fill'
$devSplit.Panel1.Controls.Add($script:GridDevices)

$script:DevDetailTabs      = New-Object System.Windows.Forms.TabControl
$script:DevDetailTabs.Dock = 'Fill'
$devSplit.Panel2.Controls.Add($script:DevDetailTabs)

function New-DetailTab {
    param([System.Windows.Forms.TabControl]$Parent, [string]$Text)
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text = $Text
    [void]$Parent.TabPages.Add($p)
    return $p
}

$dtOverview  = New-DetailTab $script:DevDetailTabs 'Overview'
$dtConfig    = New-DetailTab $script:DevDetailTabs 'Config Profiles'
$dtCompl     = New-DetailTab $script:DevDetailTabs 'Compliance'
$dtApps      = New-DetailTab $script:DevDetailTabs 'Detected Apps'
$dtKeys      = New-DetailTab $script:DevDetailTabs 'Recovery Keys / LAPS'

$script:TxtDevOverview            = New-Object System.Windows.Forms.TextBox
$script:TxtDevOverview.Multiline  = $true
$script:TxtDevOverview.ScrollBars = 'Vertical'
$script:TxtDevOverview.ReadOnly   = $true
$script:TxtDevOverview.Dock       = 'Fill'
$script:TxtDevOverview.Font       = New-Object System.Drawing.Font('Consolas', 9)
$script:TxtDevOverview.Tag        = 'output'
$dtOverview.Controls.Add($script:TxtDevOverview)

$script:GridDevConfig = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridDevConfig.Dock = 'Fill'
$dtConfig.Controls.Add($script:GridDevConfig)

$script:GridDevCompliance = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridDevCompliance.Dock = 'Fill'
$dtCompl.Controls.Add($script:GridDevCompliance)

$script:GridDevApps = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridDevApps.Dock = 'Fill'
$dtApps.Controls.Add($script:GridDevApps)

$keysPanel        = New-Object System.Windows.Forms.Panel
$keysPanel.Dock   = 'Top'
$keysPanel.Height = 40
$dtKeys.Controls.Add($keysPanel)

$script:BtnGetBitlocker = New-Button -Text 'Get BitLocker Keys' -X 4 -Y 6 -W 150
$script:BtnGetLaps      = New-Button -Text 'Get LAPS Password'  -X 160 -Y 6 -W 150
$script:BtnCopyLaps     = New-Button -Text 'Copy Password'      -X 316 -Y 6 -W 130
$script:BtnCopyLaps.Enabled = $false
$keysPanel.Controls.AddRange(@($script:BtnGetBitlocker, $script:BtnGetLaps, $script:BtnCopyLaps))

$script:LastLapsPassword = $null
$script:LastLapsAccount  = $null
$script:LastLapsDevice   = $null
$script:LapsCopyTimer    = $null

$script:TxtDevKeys            = New-Object System.Windows.Forms.TextBox
$script:TxtDevKeys.Multiline  = $true
$script:TxtDevKeys.ScrollBars = 'Vertical'
$script:TxtDevKeys.ReadOnly   = $true
$script:TxtDevKeys.Dock       = 'Fill'
$script:TxtDevKeys.Font       = New-Object System.Drawing.Font('Consolas', 10)
$script:TxtDevKeys.Tag        = 'output'
$dtKeys.Controls.Add($script:TxtDevKeys)
$script:TxtDevKeys.BringToFront()

$script:DeviceSelectColumns = @(
    'id','deviceName','userPrincipalName','userDisplayName','complianceState','operatingSystem',
    'osVersion','manufacturer','model','serialNumber','lastSyncDateTime','enrolledDateTime',
    'managedDeviceOwnerType','managementAgent','deviceEnrollmentType','azureADDeviceId',
    'isEncrypted','totalStorageSpaceInBytes','freeStorageSpaceInBytes','deviceRegistrationState'
) -join ','

function Load-Devices {
    if (-not (Assert-Connected)) { return }
    Start-Busy 'Loading devices...'
    try {
        $uri = "/deviceManagement/managedDevices?`$select=$script:DeviceSelectColumns&`$top=200"
        $devices = @(Invoke-Graph -Uri $uri -All)

        # client-side search: managedDevices does not support $filter on all of these properties
        $search = $script:TxtDevSearch.Text.Trim()
        if ($search) {
            $devices = @($devices | Where-Object {
                $_.deviceName        -like "*$search*" -or
                $_.userPrincipalName -like "*$search*" -or
                $_.serialNumber      -like "*$search*" -or
                $_.model             -like "*$search*"
            })
        }

        switch ($script:CmbDevFilter.SelectedItem) {
            'Windows only'    { $devices = @($devices | Where-Object { $_.operatingSystem -eq 'Windows' }) }
            'Non-compliant'   { $devices = @($devices | Where-Object { $_.complianceState -ne 'compliant' }) }
            'Personal owned'  { $devices = @($devices | Where-Object { $_.managedDeviceOwnerType -eq 'personal' }) }
            'Stale (no sync)' {
                $cut = (Get-Date).AddDays(-1 * [int]$script:Config.StaleDays)
                $devices = @($devices | Where-Object {
                    $_.lastSyncDateTime -and ([datetime]$_.lastSyncDateTime) -lt $cut
                })
            }
        }

        $view = foreach ($d in $devices) {
            [pscustomobject]@{
                DeviceName = $d.deviceName
                User       = $d.userPrincipalName
                Compliance = $d.complianceState
                OS         = $d.operatingSystem
                OSVersion  = $d.osVersion
                Model      = $d.model
                Serial     = $d.serialNumber
                Owner      = $d.managedDeviceOwnerType
                LastSync   = Format-GraphDate $d.lastSyncDateTime
                Enrolled   = Format-GraphDate $d.enrolledDateTime
                Encrypted  = $d.isEncrypted
                Id         = $d.id
            }
        }
        $view = @($view | Sort-Object DeviceName)

        Set-GridData -Grid $script:GridDevices -Objects $view `
            -Columns @('DeviceName','User','Compliance','OS','OSVersion','Model','Serial','Owner','LastSync','Enrolled','Encrypted','Id') `
            -CacheKey 'DeviceView' -IdProperty 'Id' | Out-Null

        # cache the full raw objects too
        $raw = @{}
        foreach ($d in $devices) { $raw[[string]$d.id] = $d }
        $script:Cache['DeviceRaw'] = $raw

        Stop-Busy "$($view.Count) device(s) loaded"
    } catch {
        Stop-Busy 'Load failed'
        Show-ErrorBox "Could not load devices:`r`n$($_.Exception.Message)"
    }
}

function Get-SelectedDeviceId {
    Get-SelectedCellValue -Grid $script:GridDevices -Column 'Id'
}

function Get-SelectedDeviceIds {
    Get-SelectedCellValues -Grid $script:GridDevices -Column 'Id'
}

function Show-DeviceDetail {
    $id = Get-SelectedDeviceId
    if (-not $id) { return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
    if (-not $d) { return }

    $totalGB = if ($d.totalStorageSpaceInBytes) { [math]::Round($d.totalStorageSpaceInBytes / 1GB, 1) } else { '' }
    $freeGB  = if ($d.freeStorageSpaceInBytes)  { [math]::Round($d.freeStorageSpaceInBytes  / 1GB, 1) } else { '' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Device name        : $($d.deviceName)")
    [void]$sb.AppendLine("Primary user       : $($d.userDisplayName)  <$($d.userPrincipalName)>")
    [void]$sb.AppendLine("Compliance         : $($d.complianceState)")
    [void]$sb.AppendLine("OS                 : $($d.operatingSystem) $($d.osVersion)")
    [void]$sb.AppendLine("Manufacturer/Model : $($d.manufacturer) / $($d.model)")
    [void]$sb.AppendLine("Serial number      : $($d.serialNumber)")
    [void]$sb.AppendLine("Ownership          : $($d.managedDeviceOwnerType)")
    [void]$sb.AppendLine("Management agent   : $($d.managementAgent)")
    [void]$sb.AppendLine("Enrollment type    : $($d.deviceEnrollmentType)")
    [void]$sb.AppendLine("Registration state : $($d.deviceRegistrationState)")
    [void]$sb.AppendLine("Encrypted          : $($d.isEncrypted)")
    [void]$sb.AppendLine("Storage            : $freeGB GB free of $totalGB GB")
    [void]$sb.AppendLine("Last sync          : $(Format-GraphDate $d.lastSyncDateTime)")
    [void]$sb.AppendLine("Enrolled           : $(Format-GraphDate $d.enrolledDateTime)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("Intune device ID   : $($d.id)")
    [void]$sb.AppendLine("Entra device ID    : $($d.azureADDeviceId)")
    $script:TxtDevOverview.Text = $sb.ToString()

    Clear-LapsCopyState
    $script:TxtDevKeys.Text = ''
    $script:GridDevConfig.DataSource = $null
    $script:GridDevCompliance.DataSource = $null
    $script:GridDevApps.DataSource = $null
}

function Load-DeviceSubData {
    $id = Get-SelectedDeviceId
    if (-not $id) { return }
    if (-not (Assert-Connected)) { return }

    $tabName = $script:DevDetailTabs.SelectedTab.Text
    try {
        switch ($tabName) {
            'Config Profiles' {
                if ($script:GridDevConfig.DataSource) { return }
                Start-Busy 'Loading configuration profile states...'
                $states = @(Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/deviceConfigurationStates" -All)
                $view = foreach ($s in $states) {
                    [pscustomobject]@{
                        Profile     = $s.displayName
                        State       = $s.state
                        Platform    = $s.platformType
                        Version     = $s.version
                        Settings    = if ($s.settingStates) { $s.settingStates.Count } else { '' }
                        Id          = $s.id
                    }
                }
                Set-GridData -Grid $script:GridDevConfig -Objects @($view | Sort-Object State, Profile) `
                    -Columns @('Profile','State','Platform','Version','Settings','Id') | Out-Null
                Stop-Busy "$($view.Count) profile state(s)"
            }
            'Compliance' {
                if ($script:GridDevCompliance.DataSource) { return }
                Start-Busy 'Loading compliance policy states...'
                $states = @(Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/deviceCompliancePolicyStates" -All)
                $view = foreach ($s in $states) {
                    [pscustomobject]@{
                        Policy   = $s.displayName
                        State    = $s.state
                        Platform = $s.platformType
                        Version  = $s.version
                        Id       = $s.id
                    }
                }
                Set-GridData -Grid $script:GridDevCompliance -Objects @($view | Sort-Object State, Policy) `
                    -Columns @('Policy','State','Platform','Version','Id') | Out-Null
                Stop-Busy "$($view.Count) compliance state(s)"
            }
            'Detected Apps' {
                if ($script:GridDevApps.DataSource) { return }
                Start-Busy 'Loading detected apps...'
                $apps = @(Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/detectedApps" -Beta -All)
                $view = foreach ($a in $apps) {
                    [pscustomobject]@{
                        Application = $a.displayName
                        Version     = $a.version
                        Publisher   = $a.publisher
                        SizeMB      = if ($a.sizeInByte) { [math]::Round($a.sizeInByte / 1MB, 1) } else { '' }
                        Id          = $a.id
                    }
                }
                Set-GridData -Grid $script:GridDevApps -Objects @($view | Sort-Object Application) `
                    -Columns @('Application','Version','Publisher','SizeMB','Id') | Out-Null
                Stop-Busy "$($view.Count) detected app(s)"
            }
        }
    } catch {
        Stop-Busy 'Load failed'
        Show-ErrorBox "Could not load $($tabName):`r`n$($_.Exception.Message)"
    }
}

function Invoke-DeviceAction {
    param(
        [string]$ActionPath,
        [string]$FriendlyName,
        $Body,
        [switch]$Beta,
        [switch]$RequireTypedConfirm
    )
    if (-not (Assert-Connected)) { return }

    $ids = Get-SelectedDeviceIds
    if ($ids.Count -eq 0) { Show-InfoBox 'Select one or more devices first.'; return }

    # Build an explicit target list so the action can never hit the wrong machine
    $targets = foreach ($id in $ids) {
        $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
        if ($d) { "  $($d.deviceName)  |  serial $($d.serialNumber)  |  $($d.userPrincipalName)" }
    }
    $list = ($targets -join "`r`n")

    if ($RequireTypedConfirm) {
        if (-not (Show-TypedConfirm -Action $FriendlyName -TargetList $list -Count $ids.Count)) { return }
    } else {
        if (-not (Confirm-Action -Message "$FriendlyName on $($ids.Count) device(s):`r`n`r`n$list`r`n`r`nProceed?" -Title $FriendlyName)) { return }
    }

    $ok = 0; $fail = 0; $errors = New-Object System.Collections.Generic.List[string]
    Start-Busy "$FriendlyName..."
    foreach ($id in $ids) {
        $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
        $name = if ($d) { $d.deviceName } else { $id }
        try {
            try {
                Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/$ActionPath" -Method POST -Body $Body -Beta:$Beta -Raw | Out-Null
            } catch {
                # some device actions only exist on beta - retry there before giving up
                if (-not $Beta -and $_.Exception.Message -match "not found for the segment") {
                    Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/$ActionPath" -Method POST -Body $Body -Beta -Raw | Out-Null
                } else {
                    throw
                }
            }
            $ok++
        } catch {
            $fail++
            [void]$errors.Add("$name : $($_.Exception.Message)")
        }
        Set-Status "$FriendlyName... $($ok + $fail) of $($ids.Count)"
    }
    Stop-Busy "$FriendlyName complete - $ok succeeded, $fail failed"

    $msg = "$FriendlyName`r`n`r`nSucceeded: $ok`r`nFailed: $fail"
    if ($errors.Count -gt 0) { $msg += "`r`n`r`n" + ($errors -join "`r`n") }
    Show-InfoBox $msg
}

function Show-TypedConfirm {
    param([string]$Action, [string]$TargetList, [int]$Count)

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "CONFIRM: $Action"
    $f.Size = New-Object System.Drawing.Size(640, 440)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'

    $l = New-Label -Text "This action is NOT reversible. Target device(s):" -X 12 -Y 10 -W 600 -H 20
    $l.ForeColor = [System.Drawing.Color]::Firebrick
    $l.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline  = $true
    $tb.ReadOnly   = $true
    $tb.ScrollBars = 'Vertical'
    $tb.Text       = $TargetList
    $tb.Location   = New-Object System.Drawing.Point(12, 36)
    $tb.Size       = New-Object System.Drawing.Size(600, 240)
    $tb.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $f.Controls.Add($tb)

    $f.Controls.Add((New-Label -Text "Type  $Action  below to enable the button:" -X 12 -Y 288 -W 400 -H 20))
    $txtConfirm = New-TextBox -X 12 -Y 310 -W 300
    $f.Controls.Add($txtConfirm)

    $btnGo = New-Button -Text "$Action ($Count)" -X 330 -Y 309 -W 150 -H 28
    $btnGo.Enabled = $false
    $btnGo.Tag       = 'danger'
    $btnGo.ForeColor = [System.Drawing.Color]::Firebrick
    $btnGo.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnGo)

    $btnNo = New-Button -Text 'Cancel' -X 490 -Y 309 -W 120 -H 28
    $btnNo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($btnNo)
    $f.CancelButton = $btnNo

    $txtConfirm.Add_TextChanged({
        $btnGo.Enabled = ($txtConfirm.Text.Trim().ToLower() -eq $Action.ToLower())
    })

    Set-ControlTheme -Control $f
    $r = $f.ShowDialog()
    $f.Dispose()
    return ($r -eq [System.Windows.Forms.DialogResult]::OK)
}

function Show-DestructiveMenu {
    $menu = New-Object System.Windows.Forms.ContextMenuStrip

    $i1 = $menu.Items.Add('Retire (remove company data, leave personal data)')
    $i1.Add_Click({ Invoke-DeviceAction -ActionPath 'retire' -FriendlyName 'RETIRE' -RequireTypedConfirm })

    $i2 = $menu.Items.Add('Wipe (full factory reset)')
    $i2.Add_Click({
        Invoke-DeviceAction -ActionPath 'wipe' -FriendlyName 'WIPE' -RequireTypedConfirm `
            -Body @{ keepEnrollmentData = $false; keepUserData = $false }
    })

    $i3 = $menu.Items.Add('Autopilot Reset (wipe, keep enrollment)')
    $i3.Add_Click({
        Invoke-DeviceAction -ActionPath 'wipe' -FriendlyName 'AUTOPILOT RESET' -RequireTypedConfirm `
            -Body @{ keepEnrollmentData = $true; keepUserData = $false }
    })

    $i4 = $menu.Items.Add('Fresh Start (remove apps, keep Windows)')
    $i4.Add_Click({
        Invoke-DeviceAction -ActionPath 'cleanWindowsDevice' -FriendlyName 'FRESH START' -RequireTypedConfirm `
            -Body @{ keepUserData = $true }
    })

    [void]$menu.Items.Add('-')

    $i5 = $menu.Items.Add('Delete Intune record only (device is NOT wiped)')
    $i5.Add_Click({ Remove-SelectedDeviceRecords })

    $menu.Show($script:BtnDevDanger, 0, $script:BtnDevDanger.Height)
}

function Remove-SelectedDeviceRecords {
    if (-not (Assert-Connected)) { return }
    $ids = Get-SelectedDeviceIds
    if ($ids.Count -eq 0) { Show-InfoBox 'Select one or more devices first.'; return }

    $targets = foreach ($id in $ids) {
        $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
        if ($d) { "  $($d.deviceName)  |  serial $($d.serialNumber)  |  $($d.userPrincipalName)" }
    }
    if (-not (Show-TypedConfirm -Action 'DELETE' -TargetList (($targets -join "`r`n")) -Count $ids.Count)) { return }

    $ok = 0; $fail = 0
    Start-Busy 'Deleting device records...'
    foreach ($id in $ids) {
        try { Invoke-Graph -Uri "/deviceManagement/managedDevices/$id" -Method DELETE -Raw | Out-Null; $ok++ }
        catch { $fail++ }
    }
    Stop-Busy "Deleted $ok record(s), $fail failed"
    Load-Devices
}

function Get-DeviceBitLockerKeys {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
    if (-not $d.azureADDeviceId) { Show-InfoBox 'This device has no Entra device ID.'; return }

    Start-Busy 'Retrieving BitLocker recovery keys...'
    try {
        $keys = @(Invoke-Graph -Uri "/informationProtection/bitlocker/recoveryKeys?`$filter=deviceId eq '$($d.azureADDeviceId)'" -All)
        if ($keys.Count -eq 0) {
            $script:TxtDevKeys.Text = "No BitLocker recovery keys are escrowed in Entra for $($d.deviceName)."
            Stop-Busy 'No keys found'
            return
        }
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("BitLocker recovery keys for $($d.deviceName)")
        [void]$sb.AppendLine(('-' * 70))
        foreach ($k in $keys) {
            $full = Invoke-Graph -Uri "/informationProtection/bitlocker/recoveryKeys/$($k.id)?`$select=key" -Raw
            [void]$sb.AppendLine("Key ID     : $($k.id)")
            [void]$sb.AppendLine("Volume     : $($k.volumeType)")
            [void]$sb.AppendLine("Created    : $(Format-GraphDate $k.createdDateTime)")
            [void]$sb.AppendLine("Recovery   : $($full.key)")
            [void]$sb.AppendLine('')
        }
        $script:TxtDevKeys.Text = $sb.ToString()
        Stop-Busy "$($keys.Count) key(s) retrieved"
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox "Could not retrieve BitLocker keys:`r`n$($_.Exception.Message)`r`n`r`nThis needs the BitlockerKey.Read.All delegated permission."
    }
}

function Get-DeviceLapsPassword {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
    if (-not $d.azureADDeviceId) { Show-InfoBox 'This device has no Entra device ID.'; return }

    Start-Busy 'Retrieving LAPS password...'
    try {
        $cred = Invoke-Graph -Uri "/directory/deviceLocalCredentials/$($d.azureADDeviceId)?`$select=credentials" -Beta -Raw

        # newest backup first, so the current password is always the one on top
        $creds = @($cred.credentials | Sort-Object -Property @{
            Expression = {
                if ($_.backupDateTime) { [datetime]$_.backupDateTime } else { [datetime]::MinValue }
            }
        } -Descending)

        if ($creds.Count -eq 0) {
            Clear-LapsCopyState
            $script:TxtDevKeys.Text = "No Windows LAPS credentials are backed up to Entra for $($d.deviceName)."
            Stop-Busy 'No LAPS credentials found'
            return
        }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Windows LAPS local admin credentials for $($d.deviceName)")
        [void]$sb.AppendLine(('-' * 70))

        $first = $true
        foreach ($c in $creds) {
            $plain = $c.passwordBase64
            try { $plain = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($c.passwordBase64)) } catch { }

            if ($first) {
                [void]$sb.AppendLine('*** CURRENT PASSWORD ***')
                $script:LastLapsPassword = $plain
                $script:LastLapsAccount  = $c.accountName
                $script:LastLapsDevice   = $d.deviceName
                $first = $false
            } else {
                [void]$sb.AppendLine('--- previous ---')
            }

            [void]$sb.AppendLine("Account    : $($c.accountName)")
            [void]$sb.AppendLine("Account SID: $($c.accountSid)")
            [void]$sb.AppendLine("Backed up  : $(Format-GraphDate $c.backupDateTime)")
            [void]$sb.AppendLine("Password   : $plain")
            [void]$sb.AppendLine('')
        }

        $script:TxtDevKeys.Text = $sb.ToString()
        $script:BtnCopyLaps.Enabled = $true
        if ($script:ThemeTip) {
            $script:ThemeTip.SetToolTip($script:BtnCopyLaps,
                "Copy the current password for $($script:LastLapsAccount) on $($script:LastLapsDevice)")
        }
        Stop-Busy "LAPS password retrieved for $($script:LastLapsAccount)"
    } catch {
        Clear-LapsCopyState
        Stop-Busy 'Failed'
        Show-ErrorBox "Could not retrieve LAPS password:`r`n$($_.Exception.Message)`r`n`r`nThis needs the DeviceLocalCredential.Read.All delegated permission and Windows LAPS backed up to Entra."
    }
}

function Clear-LapsCopyState {
    $script:LastLapsPassword = $null
    $script:LastLapsAccount  = $null
    $script:LastLapsDevice   = $null
    if ($script:BtnCopyLaps) {
        $script:BtnCopyLaps.Enabled = $false
        $script:BtnCopyLaps.Text    = 'Copy Password'
        if ($script:ThemeTip) { $script:ThemeTip.SetToolTip($script:BtnCopyLaps, '') }
    }
}

function Copy-LapsPassword {
    if ([string]::IsNullOrEmpty($script:LastLapsPassword)) {
        Show-InfoBox 'Retrieve a LAPS password first.'
        return
    }
    try {
        [System.Windows.Forms.Clipboard]::SetText($script:LastLapsPassword)
    } catch {
        Show-ErrorBox "Could not write to the clipboard:`r`n$($_.Exception.Message)"
        return
    }

    $script:BtnCopyLaps.Text = 'Copied'
    Set-Status "Password for $($script:LastLapsAccount) on $($script:LastLapsDevice) copied to the clipboard"

    # the timer must live at script scope: a local would be gone by the time
    # Tick fires, and calling .Stop() on it would throw from the message loop
    if ($script:LapsCopyTimer) {
        try { $script:LapsCopyTimer.Stop(); $script:LapsCopyTimer.Dispose() } catch { }
        $script:LapsCopyTimer = $null
    }
    $script:LapsCopyTimer = New-Object System.Windows.Forms.Timer
    $script:LapsCopyTimer.Interval = 1500
    $script:LapsCopyTimer.Add_Tick({
        try {
            if ($script:BtnCopyLaps -and -not $script:BtnCopyLaps.IsDisposed) {
                $script:BtnCopyLaps.Text = 'Copy Password'
            }
            if ($script:LapsCopyTimer) {
                $script:LapsCopyTimer.Stop()
                $script:LapsCopyTimer.Dispose()
                $script:LapsCopyTimer = $null
            }
        } catch { }
    })
    $script:LapsCopyTimer.Start()
}

# --- device tab wiring -----------------------------------------------------
$script:BtnDevLoad.Add_Click({ Load-Devices })
$script:TxtDevSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $_.SuppressKeyPress = $true; Load-Devices } })
$script:CmbDevFilter.Add_SelectedIndexChanged({ if ($script:GridDevices.DataSource) { Load-Devices } })
$script:BtnDevExport.Add_Click({ Export-GridToCsv -Grid $script:GridDevices -SuggestedName 'IntuneDevices' })
$script:GridDevices.Add_SelectionChanged({ Show-DeviceDetail })
$script:DevDetailTabs.Add_SelectedIndexChanged({ Load-DeviceSubData })

$script:BtnDevSync.Add_Click({     Invoke-DeviceAction -ActionPath 'syncDevice' -FriendlyName 'Sync' })
$script:BtnDevRestart.Add_Click({  Invoke-DeviceAction -ActionPath 'rebootNow'  -FriendlyName 'Restart' })
$script:BtnDevShutdown.Add_Click({ Invoke-DeviceAction -ActionPath 'shutDown'   -FriendlyName 'Shut Down' })
$script:BtnDevScan.Add_Click({     Invoke-DeviceAction -ActionPath 'windowsDefenderScan' -FriendlyName 'Defender Quick Scan' -Body @{ quickScan = $true } })
$script:BtnDevRotateBL.Add_Click({ Invoke-DeviceAction -ActionPath 'rotateBitLockerKeys' -FriendlyName 'Rotate BitLocker Keys' -Beta })
$script:BtnDevRotateLA.Add_Click({ Invoke-DeviceAction -ActionPath 'rotateLocalAdminPassword' -FriendlyName 'Rotate LAPS Password' -Beta })
$script:BtnDevManage.Add_Click({   Show-DeviceManageMenu })
$script:BtnDevDanger.Add_Click({   Show-DestructiveMenu })
$script:BtnGetBitlocker.Add_Click({ Get-DeviceBitLockerKeys })
$script:BtnGetLaps.Add_Click({ Get-DeviceLapsPassword })
$script:BtnCopyLaps.Add_Click({ Copy-LapsPassword })

$script:BtnDevJson.Add_Click({
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    if (-not (Assert-Connected)) { return }
    Start-Busy 'Loading full device record...'
    try {
        $full = Invoke-Graph -Uri "/deviceManagement/managedDevices/$id" -Beta -Raw
        Stop-Busy 'Ready'
        Show-JsonWindow -Title "Device $((Get-CachedObject -CacheKey 'DeviceRaw' -Id $id).deviceName)" -Object $full
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox $_.Exception.Message
    }
})

# ---------------------------------------------------------------------------
# Shared: resolve Entra group names (cached)
# ---------------------------------------------------------------------------
$script:GroupNameCache = @{}

function Get-GroupDisplayName {
    param([string]$GroupId)
    if (-not $GroupId) { return '' }
    if ($script:GroupNameCache.ContainsKey($GroupId)) { return $script:GroupNameCache[$GroupId] }
    try {
        $g = Invoke-Graph -Uri "/groups/$GroupId`?`$select=displayName" -Raw
        $script:GroupNameCache[$GroupId] = $g.displayName
        return $g.displayName
    } catch {
        $script:GroupNameCache[$GroupId] = "(unresolved $GroupId)"
        return $script:GroupNameCache[$GroupId]
    }
}

function Get-AssignmentTargetText {
    param($Assignment)
    $t = $Assignment.target
    if (-not $t) { return '' }
    switch -Wildcard ($t.'@odata.type') {
        '*allDevicesAssignmentTarget'      { return 'All Devices' }
        '*allLicensedUsersAssignmentTarget'{ return 'All Users' }
        '*exclusionGroupAssignmentTarget'  { return "EXCLUDE: $(Get-GroupDisplayName $t.groupId)" }
        '*groupAssignmentTarget'           { return (Get-GroupDisplayName $t.groupId) }
        default                            { return [string]$t.'@odata.type' }
    }
}

# ===========================================================================
# AUTOPILOT TAB
# ===========================================================================
$apTop        = New-Object System.Windows.Forms.Panel
$apTop.Dock   = 'Top'
$apTop.Height = 44
$tabAutopilot.Controls.Add($apTop)

$apTop.Controls.Add((New-Label -Text 'Search:' -X 4 -Y 10 -W 55 -H 24))
$script:TxtApSearch = New-TextBox -X 60 -Y 8 -W 220
$apTop.Controls.Add($script:TxtApSearch)

$script:BtnApLoad     = New-Button -Text 'Load Autopilot'   -X 290 -Y 7 -W 120
$script:BtnApGroupTag = New-Button -Text 'Set Group Tag'    -X 418 -Y 7 -W 120
$script:BtnApAssign   = New-Button -Text 'Assign User'      -X 546 -Y 7 -W 110
$script:BtnApUnassign = New-Button -Text 'Unassign User'    -X 664 -Y 7 -W 110
$script:BtnApSync     = New-Button -Text 'Sync Service'     -X 782 -Y 7 -W 110
$script:BtnApExport   = New-Button -Text 'Export CSV'       -X 900 -Y 7 -W 100
$script:BtnApDelete   = New-Button -Text 'Delete Record'    -X 1008 -Y 7 -W 120
$script:BtnApDelete.Tag       = 'danger'
$script:BtnApDelete.ForeColor = [System.Drawing.Color]::Firebrick

$apTop.Controls.AddRange(@($script:BtnApLoad, $script:BtnApGroupTag, $script:BtnApAssign,
    $script:BtnApUnassign, $script:BtnApSync, $script:BtnApExport, $script:BtnApDelete))

$script:GridAutopilot = New-Grid -X 0 -Y 0 -W 100 -H 100 -MultiSelect
$script:GridAutopilot.Dock = 'Fill'
$tabAutopilot.Controls.Add($script:GridAutopilot)
$script:GridAutopilot.BringToFront()

function Load-AutopilotDevices {
    if (-not (Assert-Connected)) { return }
    Start-Busy 'Loading Autopilot devices...'
    try {
        $devices = @(Invoke-Graph -Uri "/deviceManagement/windowsAutopilotDeviceIdentities?`$top=200" -All)

        $search = $script:TxtApSearch.Text.Trim()
        if ($search) {
            $devices = @($devices | Where-Object {
                $_.serialNumber -like "*$search*" -or
                $_.groupTag     -like "*$search*" -or
                $_.displayName  -like "*$search*" -or
                $_.userPrincipalName -like "*$search*"
            })
        }

        $view = foreach ($d in $devices) {
            [pscustomobject]@{
                Serial        = $d.serialNumber
                GroupTag      = $d.groupTag
                Model         = $d.model
                Manufacturer  = $d.manufacturer
                AssignedUser  = $d.userPrincipalName
                EnrollState   = $d.enrollmentState
                ProfileStatus = $d.deploymentProfileAssignmentStatus
                LastContact   = Format-GraphDate $d.lastContactedDateTime
                IntuneDevice  = $d.managedDeviceId
                Id            = $d.id
            }
        }
        $view = @($view | Sort-Object GroupTag, Serial)

        Set-GridData -Grid $script:GridAutopilot -Objects $view `
            -Columns @('Serial','GroupTag','Model','Manufacturer','AssignedUser','EnrollState','ProfileStatus','LastContact','IntuneDevice','Id') | Out-Null

        $raw = @{}
        foreach ($d in $devices) { $raw[[string]$d.id] = $d }
        $script:Cache['AutopilotRaw'] = $raw

        Stop-Busy "$($view.Count) Autopilot device(s)"
    } catch {
        Stop-Busy 'Load failed'
        Show-ErrorBox "Could not load Autopilot devices:`r`n$($_.Exception.Message)"
    }
}

function Show-InputDialog {
    param([string]$Title, [string]$Prompt, [string]$Default = '')
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.Size = New-Object System.Drawing.Size(460, 180)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false

    $f.Controls.Add((New-Label -Text $Prompt -X 12 -Y 15 -W 420 -H 20))
    $tb = New-TextBox -X 12 -Y 45 -W 420 -Text $Default
    $f.Controls.Add($tb)

    $ok = New-Button -Text 'OK' -X 232 -Y 95 -W 95 -H 28
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($ok); $f.AcceptButton = $ok

    $no = New-Button -Text 'Cancel' -X 337 -Y 95 -W 95 -H 28
    $no.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($no); $f.CancelButton = $no

    Set-ControlTheme -Control $f
    $r = $f.ShowDialog()
    $val = $tb.Text
    $f.Dispose()
    if ($r -eq [System.Windows.Forms.DialogResult]::OK) { return $val }
    return $null
}

function Set-AutopilotGroupTag {
    if (-not (Assert-Connected)) { return }
    $ids = Get-SelectedCellValues -Grid $script:GridAutopilot -Column 'Id'
    if ($ids.Count -eq 0) { Show-InfoBox 'Select one or more Autopilot devices first.'; return }

    $current = ''
    if ($ids.Count -eq 1) {
        $d = Get-CachedObject -CacheKey 'AutopilotRaw' -Id $ids[0]
        if ($d) { $current = $d.groupTag }
    }
    $tag = Show-InputDialog -Title 'Set Group Tag' -Prompt "New Group Tag for $($ids.Count) device(s) (blank clears it):" -Default $current
    if ($null -eq $tag) { return }

    $targets = foreach ($id in $ids) {
        $d = Get-CachedObject -CacheKey 'AutopilotRaw' -Id $id
        if ($d) { "  $($d.serialNumber)  |  current tag: '$($d.groupTag)'" }
    }
    if (-not (Confirm-Action -Message "Set Group Tag to '$tag' on:`r`n`r`n$($targets -join "`r`n")`r`n`r`nProceed?" -Title 'Set Group Tag')) { return }

    $ok = 0; $fail = 0; $errors = New-Object System.Collections.Generic.List[string]
    Start-Busy 'Updating Group Tags...'
    foreach ($id in $ids) {
        try {
            Invoke-Graph -Uri "/deviceManagement/windowsAutopilotDeviceIdentities/$id/updateDeviceProperties" `
                -Method POST -Body @{ groupTag = $tag } -Raw | Out-Null
            $ok++
        } catch { $fail++; [void]$errors.Add($_.Exception.Message) }
    }
    Stop-Busy "Group Tag updated on $ok device(s), $fail failed"
    if ($errors.Count) { Show-ErrorBox ($errors -join "`r`n") }
    Load-AutopilotDevices
}

function Set-AutopilotUser {
    if (-not (Assert-Connected)) { return }
    $ids = Get-SelectedCellValues -Grid $script:GridAutopilot -Column 'Id'
    if ($ids.Count -ne 1) { Show-InfoBox 'Select exactly one Autopilot device.'; return }

    $upn = Show-InputDialog -Title 'Assign User' -Prompt 'User principal name (user@domain):'
    if ([string]::IsNullOrWhiteSpace($upn)) { return }

    Start-Busy 'Assigning user...'
    try {
        Invoke-Graph -Uri "/deviceManagement/windowsAutopilotDeviceIdentities/$($ids[0])/assignUserToDevice" `
            -Method POST -Body @{ userPrincipalName = $upn; addressableUserName = $upn } -Raw | Out-Null
        Stop-Busy "Assigned $upn"
        Load-AutopilotDevices
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox $_.Exception.Message
    }
}

function Clear-AutopilotUser {
    if (-not (Assert-Connected)) { return }
    $ids = Get-SelectedCellValues -Grid $script:GridAutopilot -Column 'Id'
    if ($ids.Count -ne 1) { Show-InfoBox 'Select exactly one Autopilot device.'; return }
    $d = Get-CachedObject -CacheKey 'AutopilotRaw' -Id $ids[0]
    if (-not (Confirm-Action -Message "Remove assigned user from $($d.serialNumber)?" -Title 'Unassign User')) { return }

    Start-Busy 'Unassigning user...'
    try {
        Invoke-Graph -Uri "/deviceManagement/windowsAutopilotDeviceIdentities/$($ids[0])/unassignUserFromDevice" -Method POST -Raw | Out-Null
        Stop-Busy 'User unassigned'
        Load-AutopilotDevices
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox $_.Exception.Message
    }
}

function Remove-AutopilotDevices {
    if (-not (Assert-Connected)) { return }
    $ids = Get-SelectedCellValues -Grid $script:GridAutopilot -Column 'Id'
    if ($ids.Count -eq 0) { Show-InfoBox 'Select one or more Autopilot devices first.'; return }

    $targets = foreach ($id in $ids) {
        $d = Get-CachedObject -CacheKey 'AutopilotRaw' -Id $id
        if ($d) { "  $($d.serialNumber)  |  tag '$($d.groupTag)'  |  $($d.model)" }
    }
    if (-not (Show-TypedConfirm -Action 'DELETE' -TargetList (($targets -join "`r`n") + "`r`n`r`nThis deregisters the hardware hash from Autopilot.") -Count $ids.Count)) { return }

    $ok = 0; $fail = 0
    Start-Busy 'Deleting Autopilot records...'
    foreach ($id in $ids) {
        try { Invoke-Graph -Uri "/deviceManagement/windowsAutopilotDeviceIdentities/$id" -Method DELETE -Raw | Out-Null; $ok++ }
        catch { $fail++ }
    }
    Stop-Busy "Deleted $ok record(s), $fail failed"
    Load-AutopilotDevices
}

function Sync-AutopilotService {
    if (-not (Assert-Connected)) { return }
    Start-Busy 'Requesting Autopilot service sync...'
    try {
        Invoke-Graph -Uri '/deviceManagement/windowsAutopilotSettings/sync' -Method POST -Beta -Raw | Out-Null
        Stop-Busy 'Autopilot sync requested'
        Show-InfoBox 'Autopilot service sync requested. This can only run every 10 minutes.'
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox $_.Exception.Message
    }
}

$script:BtnApLoad.Add_Click({ Load-AutopilotDevices })
$script:TxtApSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $_.SuppressKeyPress = $true; Load-AutopilotDevices } })
$script:BtnApGroupTag.Add_Click({ Set-AutopilotGroupTag })
$script:BtnApAssign.Add_Click({ Set-AutopilotUser })
$script:BtnApUnassign.Add_Click({ Clear-AutopilotUser })
$script:BtnApSync.Add_Click({ Sync-AutopilotService })
$script:BtnApDelete.Add_Click({ Remove-AutopilotDevices })
$script:BtnApExport.Add_Click({ Export-GridToCsv -Grid $script:GridAutopilot -SuggestedName 'AutopilotDevices' })

# ===========================================================================
# APPS TAB
# ===========================================================================
$appTop        = New-Object System.Windows.Forms.Panel
$appTop.Dock   = 'Top'
$appTop.Height = 44
$tabApps.Controls.Add($appTop)

$appTop.Controls.Add((New-Label -Text 'Search:' -X 4 -Y 10 -W 55 -H 24))
$script:TxtAppSearch = New-TextBox -X 60 -Y 8 -W 240
$appTop.Controls.Add($script:TxtAppSearch)

$script:ChkAppWin32          = New-Object System.Windows.Forms.CheckBox
$script:ChkAppWin32.Text     = 'Win32 / MSI only'
$script:ChkAppWin32.Location = New-Object System.Drawing.Point(310, 10)
$script:ChkAppWin32.Size     = New-Object System.Drawing.Size(130, 24)
$appTop.Controls.Add($script:ChkAppWin32)

$script:BtnAppLoad    = New-Button -Text 'Load Apps'      -X 446 -Y 7 -W 110
$script:BtnAppStatus  = New-Button -Text 'Install Status' -X 564 -Y 7 -W 120
$script:BtnAppJson    = New-Button -Text 'View Raw JSON'  -X 692 -Y 7 -W 120
$script:BtnAppAssign  = New-Button -Text 'Edit Assignments' -X 820 -Y 7 -W 140
$script:BtnAppExport  = New-Button -Text 'Export CSV'     -X 968 -Y 7 -W 100
$appTop.Controls.AddRange(@($script:BtnAppLoad, $script:BtnAppStatus, $script:BtnAppJson, $script:BtnAppAssign, $script:BtnAppExport))

$appSplit             = New-Object System.Windows.Forms.SplitContainer
$appSplit.Dock        = 'Fill'
$appSplit.Orientation = 'Horizontal'
$tabApps.Controls.Add($appSplit)
$appSplit.BringToFront()

$script:GridApps = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridApps.Dock = 'Fill'
$appSplit.Panel1.Controls.Add($script:GridApps)

$script:AppDetailTabs      = New-Object System.Windows.Forms.TabControl
$script:AppDetailTabs.Dock = 'Fill'
$appSplit.Panel2.Controls.Add($script:AppDetailTabs)

$atAssign = New-DetailTab $script:AppDetailTabs 'Assignments'
$atStatus = New-DetailTab $script:AppDetailTabs 'Install Status'
$atDetect = New-DetailTab $script:AppDetailTabs 'Install / Detection'

$script:GridAppAssign = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridAppAssign.Dock = 'Fill'
$atAssign.Controls.Add($script:GridAppAssign)

$script:GridAppStatus = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridAppStatus.Dock = 'Fill'
$atStatus.Controls.Add($script:GridAppStatus)

$script:TxtAppDetect            = New-Object System.Windows.Forms.TextBox
$script:TxtAppDetect.Multiline  = $true
$script:TxtAppDetect.ScrollBars = 'Both'
$script:TxtAppDetect.WordWrap   = $false
$script:TxtAppDetect.ReadOnly   = $true
$script:TxtAppDetect.Dock       = 'Fill'
$script:TxtAppDetect.Font       = New-Object System.Drawing.Font('Consolas', 9)
$script:TxtAppDetect.Tag        = 'output'
$atDetect.Controls.Add($script:TxtAppDetect)

function Get-AppTypeShortName {
    param([string]$ODataType)
    if (-not $ODataType) { return '' }
    $t = $ODataType -replace '^#microsoft\.graph\.', ''
    return $t
}

function Load-Apps {
    if (-not (Assert-Connected)) { return }
    Start-Busy 'Loading applications...'
    try {
        $apps = @(Invoke-Graph -Uri "/deviceAppManagement/mobileApps?`$expand=assignments&`$top=200" -All)

        $search = $script:TxtAppSearch.Text.Trim()
        if ($search) { $apps = @($apps | Where-Object { $_.displayName -like "*$search*" -or $_.publisher -like "*$search*" }) }
        if ($script:ChkAppWin32.Checked) {
            $apps = @($apps | Where-Object { $_.'@odata.type' -match 'win32LobApp|windowsMobileMSI|windowsUniversalAppX|officeSuiteApp' })
        }

        $view = foreach ($a in $apps) {
            [pscustomobject]@{
                Application = $a.displayName
                Type        = Get-AppTypeShortName $a.'@odata.type'
                Publisher   = $a.publisher
                Version     = $a.displayVersion
                Assignments = if ($a.assignments) { @($a.assignments).Count } else { 0 }
                Featured    = $a.isFeatured
                Created     = Format-GraphDate $a.createdDateTime
                Modified    = Format-GraphDate $a.lastModifiedDateTime
                Id          = $a.id
            }
        }
        $view = @($view | Sort-Object Application)

        Set-GridData -Grid $script:GridApps -Objects $view `
            -Columns @('Application','Type','Publisher','Version','Assignments','Featured','Created','Modified','Id') | Out-Null

        $raw = @{}
        foreach ($a in $apps) { $raw[[string]$a.id] = $a }
        $script:Cache['AppRaw'] = $raw

        Stop-Busy "$($view.Count) app(s) loaded"
    } catch {
        Stop-Busy 'Load failed'
        Show-ErrorBox "Could not load apps:`r`n$($_.Exception.Message)"
    }
}

function Show-AppDetail {
    $id = Get-SelectedCellValue -Grid $script:GridApps -Column 'Id'
    if (-not $id) { return }
    $a = Get-CachedObject -CacheKey 'AppRaw' -Id $id
    if (-not $a) { return }

    # assignments
    $rows = foreach ($asg in @($a.assignments)) {
        [pscustomobject]@{
            Target     = Get-AssignmentTargetText $asg
            Intent     = $asg.intent
            Filter     = if ($asg.target.deviceAndAppManagementAssignmentFilterId) { $asg.target.deviceAndAppManagementAssignmentFilterId } else { '' }
            FilterType = $asg.target.deviceAndAppManagementAssignmentFilterType
            Id         = $asg.id
        }
    }
    Set-GridData -Grid $script:GridAppAssign -Objects @($rows) -Columns @('Target','Intent','Filter','FilterType','Id') | Out-Null

    # install / detection detail
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Application : $($a.displayName)")
    [void]$sb.AppendLine("Type        : $(Get-AppTypeShortName $a.'@odata.type')")
    [void]$sb.AppendLine("Publisher   : $($a.publisher)")
    [void]$sb.AppendLine("Version     : $($a.displayVersion)")
    [void]$sb.AppendLine("App ID      : $($a.id)")
    [void]$sb.AppendLine('')
    if ($a.installCommandLine)   { [void]$sb.AppendLine("Install     : $($a.installCommandLine)") }
    if ($a.uninstallCommandLine) { [void]$sb.AppendLine("Uninstall   : $($a.uninstallCommandLine)") }
    if ($a.setupFilePath)        { [void]$sb.AppendLine("Setup file  : $($a.setupFilePath)") }
    if ($a.installExperience)    { [void]$sb.AppendLine("Context     : $($a.installExperience.runAsAccount)  restart: $($a.installExperience.deviceRestartBehavior)") }
    if ($a.detectionRules) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Detection rules:')
        [void]$sb.AppendLine(($a.detectionRules | ConvertTo-Json -Depth 8))
    }
    if ($a.requirementRules) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Requirement rules:')
        [void]$sb.AppendLine(($a.requirementRules | ConvertTo-Json -Depth 8))
    }
    $script:TxtAppDetect.Text = $sb.ToString()

    $script:GridAppStatus.DataSource = $null
}

function Get-AppInstallStatusReport {
    # The per-app device status report has been renamed more than once. Try the
    # current function first, then the older name, and map the Schema/Values
    # pair generically so a column change on Microsoft's side does not break it.
    param([string]$AppId)

    $candidates = @(
        'retrieveDeviceAppInstallationStatusReport',
        'getDeviceInstallStatusReport'
    )

    $body = @{
        select  = @('DeviceId','DeviceName','UserName','UserPrincipalName','Platform',
                    'AppVersion','InstallState','InstallStateDetail','ErrorCode','LastModifiedDateTime')
        filter  = "(ApplicationId eq '$AppId')"
        orderBy = @()
        skip    = 0
        top     = 500
    }

    $r        = $null
    $lastErr  = $null
    $usedName = $null

    foreach ($fn in $candidates) {
        try {
            $r = Invoke-Graph -Uri "/deviceManagement/reports/$fn" -Method POST -Body $body -Beta -Raw
            $usedName = $fn
            break
        } catch {
            $lastErr = $_.Exception.Message
            # a rejected column list is worth one retry without the select
            if ($lastErr -notmatch 'not found for the segment') {
                try {
                    $plain = @{ filter = "(ApplicationId eq '$AppId')"; skip = 0; top = 500 }
                    $r = Invoke-Graph -Uri "/deviceManagement/reports/$fn" -Method POST -Body $plain -Beta -Raw
                    $usedName = $fn
                    break
                } catch {
                    $lastErr = $_.Exception.Message
                }
            }
        }
    }

    if (-not $r) { throw "App install status report failed. Last error: $lastErr" }

    # the report can come back as a JSON string rather than a parsed object
    if ($r -is [string]) {
        try { $r = $r | ConvertFrom-Json } catch { }
    }

    $cols = @()
    if ($r.Schema) { $cols = @($r.Schema | ForEach-Object { $_.Column }) }
    if ($cols.Count -eq 0) { return @{ Rows = @(); Columns = @(); Source = $usedName } }

    $rows = foreach ($v in @($r.Values)) {
        $o = [ordered]@{}
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $val = $null
            if ($i -lt @($v).Count) { $val = $v[$i] }
            $o[$cols[$i]] = $val
        }
        [pscustomobject]$o
    }

    return @{ Rows = @($rows); Columns = $cols; Source = $usedName }
}

function Load-AppInstallStatus {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedCellValue -Grid $script:GridApps -Column 'Id'
    if (-not $id) { Show-InfoBox 'Select an app first.'; return }

    Start-Busy 'Loading install status...'
    try {
        $view    = $null
        $columns = $null

        # preferred: the per-app device status navigation property (beta)
        try {
            $states = @(Invoke-Graph -Uri "/deviceAppManagement/mobileApps/$id/deviceStatuses" -Beta -All)
            $view = foreach ($s in $states) {
                [pscustomobject]@{
                    DeviceName = $s.deviceName
                    User       = $s.userPrincipalName
                    State      = $s.installState
                    Detail     = $s.installStateDetail
                    ErrorCode  = $s.errorCode
                    Version    = $s.displayVersion
                    LastSync   = Format-GraphDate $s.lastSyncDateTime
                    Id         = $s.id
                }
            }
            $view    = @($view | Sort-Object State, DeviceName)
            $columns = @('DeviceName','User','State','Detail','ErrorCode','Version','LastSync','Id')
        } catch {
            $view = $null
        }

        # fallback: the Intune reporting endpoint the portal itself uses
        if (-not $view -or $view.Count -eq 0) {
            Set-Status 'Loading install status from the Intune reporting endpoint...'
            $report = Get-AppInstallStatusReport -AppId $id
            $view    = $report.Rows
            $columns = $report.Columns
        }

        if (-not $view -or $view.Count -eq 0) {
            $script:GridAppStatus.DataSource = $null
            Stop-Busy 'No install status records reported for this app'
            $script:AppDetailTabs.SelectedIndex = 1
            return
        }

        Set-GridData -Grid $script:GridAppStatus -Objects @($view) -Columns $columns | Out-Null
        $script:AppDetailTabs.SelectedIndex = 1
        Stop-Busy "$(@($view).Count) device status record(s)"
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox "Could not load install status:`r`n$($_.Exception.Message)"
    }
}

$script:BtnAppLoad.Add_Click({ Load-Apps })
$script:TxtAppSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $_.SuppressKeyPress = $true; Load-Apps } })
$script:ChkAppWin32.Add_CheckedChanged({ if ($script:GridApps.DataSource) { Load-Apps } })
$script:GridApps.Add_SelectionChanged({ Show-AppDetail })
$script:BtnAppStatus.Add_Click({ Load-AppInstallStatus })
$script:BtnAppExport.Add_Click({ Export-GridToCsv -Grid $script:GridApps -SuggestedName 'IntuneApps' })
$script:BtnAppAssign.Add_Click({
    $appId = Get-SelectedCellValue -Grid $script:GridApps -Column 'Id'
    if (-not $appId) { Show-InfoBox 'Select an app first.'; return }
    $appName = Get-SelectedCellValue -Grid $script:GridApps -Column 'Application'
    Show-AssignmentEditor -Kind 'App' -ObjectId $appId -ObjectName $appName
    Load-Apps
})
$script:BtnAppJson.Add_Click({
    $id = Get-SelectedCellValue -Grid $script:GridApps -Column 'Id'
    if (-not $id) { Show-InfoBox 'Select an app first.'; return }
    $a = Get-CachedObject -CacheKey 'AppRaw' -Id $id
    Show-JsonWindow -Title "App $($a.displayName)" -Object $a
})

# ===========================================================================
# POLICIES TAB
# ===========================================================================
$polTop        = New-Object System.Windows.Forms.Panel
$polTop.Dock   = 'Top'
$polTop.Height = 44
$tabPolicies.Controls.Add($polTop)

$polTop.Controls.Add((New-Label -Text 'Type:' -X 4 -Y 10 -W 40 -H 24))
$script:CmbPolType = New-Object System.Windows.Forms.ComboBox
$script:CmbPolType.Location      = New-Object System.Drawing.Point(46, 8)
$script:CmbPolType.Size          = New-Object System.Drawing.Size(250, 24)
$script:CmbPolType.DropDownStyle = 'DropDownList'
[void]$script:CmbPolType.Items.AddRange(@(
    'Settings Catalog',
    'Device Configuration (legacy)',
    'Compliance Policies',
    'Platform Scripts',
    'Remediations',
    'Assignment Filters',
    'Enrollment Configurations (ESP etc.)'
))
$script:CmbPolType.SelectedIndex = 0
$polTop.Controls.Add($script:CmbPolType)

$script:BtnPolLoad   = New-Button -Text 'Load'            -X 306 -Y 7 -W 90
$script:BtnPolJson   = New-Button -Text 'View / Save JSON'-X 404 -Y 7 -W 140
$script:BtnPolAssign = New-Button -Text 'Edit Assignments' -X 552 -Y 7 -W 140
$script:BtnPolImport = New-Button -Text 'Import JSON...'   -X 700 -Y 7 -W 130
$script:BtnPolExport = New-Button -Text 'Export CSV'      -X 838 -Y 7 -W 100
$polTop.Controls.AddRange(@($script:BtnPolLoad, $script:BtnPolJson, $script:BtnPolAssign, $script:BtnPolImport, $script:BtnPolExport))

$polSplit             = New-Object System.Windows.Forms.SplitContainer
$polSplit.Dock        = 'Fill'
$polSplit.Orientation = 'Horizontal'
$tabPolicies.Controls.Add($polSplit)
$polSplit.BringToFront()

$script:GridPolicies = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridPolicies.Dock = 'Fill'
$polSplit.Panel1.Controls.Add($script:GridPolicies)

$script:GridPolAssign = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridPolAssign.Dock = 'Fill'
$polSplit.Panel2.Controls.Add($script:GridPolAssign)

function Get-PolicyEndpoint {
    param([string]$TypeName)
    switch ($TypeName) {
        'Settings Catalog'                      { return @{ Uri = "/deviceManagement/configurationPolicies?`$expand=assignments"; Beta = $true;  NameProp = 'name' } }
        'Device Configuration (legacy)'         { return @{ Uri = "/deviceManagement/deviceConfigurations?`$expand=assignments"; Beta = $false; NameProp = 'displayName' } }
        'Compliance Policies'                   { return @{ Uri = "/deviceManagement/deviceCompliancePolicies?`$expand=assignments"; Beta = $false; NameProp = 'displayName' } }
        'Platform Scripts'                      { return @{ Uri = "/deviceManagement/deviceManagementScripts?`$expand=assignments"; Beta = $true;  NameProp = 'displayName' } }
        'Remediations'                          { return @{ Uri = "/deviceManagement/deviceHealthScripts?`$expand=assignments"; Beta = $true;  NameProp = 'displayName' } }
        'Assignment Filters'                    { return @{ Uri = "/deviceManagement/assignmentFilters"; Beta = $true;  NameProp = 'displayName' } }
        'Enrollment Configurations (ESP etc.)'  { return @{ Uri = "/deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments"; Beta = $false; NameProp = 'displayName' } }
    }
}

function Load-Policies {
    if (-not (Assert-Connected)) { return }
    $typeName = [string]$script:CmbPolType.SelectedItem
    $ep = Get-PolicyEndpoint -TypeName $typeName

    Start-Busy "Loading $typeName..."
    try {
        $items = @(Invoke-Graph -Uri $ep.Uri -Beta:$ep.Beta -All)

        $view = foreach ($p in $items) {
            $name = $p.($ep.NameProp)
            [pscustomobject]@{
                Name        = $name
                Platform    = if ($p.platforms) { ($p.platforms -join ',') } elseif ($p.platform) { $p.platform } else { '' }
                Kind        = if ($p.'@odata.type') { ($p.'@odata.type' -replace '^#microsoft\.graph\.', '') } elseif ($p.technologies) { ($p.technologies -join ',') } else { '' }
                Assignments = if ($p.assignments) { @($p.assignments).Count } else { 0 }
                Created     = Format-GraphDate $p.createdDateTime
                Modified    = Format-GraphDate $p.lastModifiedDateTime
                Description = $p.description
                Id          = $p.id
            }
        }
        $view = @($view | Sort-Object Name)

        Set-GridData -Grid $script:GridPolicies -Objects $view `
            -Columns @('Name','Platform','Kind','Assignments','Created','Modified','Description','Id') | Out-Null

        $raw = @{}
        foreach ($p in $items) { $raw[[string]$p.id] = $p }
        $script:Cache['PolicyRaw'] = $raw
        $script:Cache['PolicyMeta'] = @{ TypeName = $typeName; Beta = $ep.Beta; NameProp = $ep.NameProp }

        $script:GridPolAssign.DataSource = $null
        Stop-Busy "$($view.Count) $typeName item(s)"
    } catch {
        Stop-Busy 'Load failed'
        Show-ErrorBox "Could not load $($typeName):`r`n$($_.Exception.Message)"
    }
}

function Show-PolicyAssignments {
    $id = Get-SelectedCellValue -Grid $script:GridPolicies -Column 'Id'
    if (-not $id) { return }
    $p = Get-CachedObject -CacheKey 'PolicyRaw' -Id $id
    if (-not $p) { return }

    $rows = foreach ($asg in @($p.assignments)) {
        [pscustomobject]@{
            Target     = Get-AssignmentTargetText $asg
            Intent     = $asg.intent
            Filter     = $asg.target.deviceAndAppManagementAssignmentFilterId
            FilterType = $asg.target.deviceAndAppManagementAssignmentFilterType
            Id         = $asg.id
        }
    }
    if (-not $rows) {
        $rows = @([pscustomobject]@{ Target='(no assignments)'; Intent=''; Filter=''; FilterType=''; Id='' })
    }
    Set-GridData -Grid $script:GridPolAssign -Objects @($rows) -Columns @('Target','Intent','Filter','FilterType','Id') | Out-Null
}

function Show-PolicyJson {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedCellValue -Grid $script:GridPolicies -Column 'Id'
    if (-not $id) { Show-InfoBox 'Select a policy first.'; return }
    $meta = $script:Cache['PolicyMeta']
    $p    = Get-CachedObject -CacheKey 'PolicyRaw' -Id $id
    $name = $p.($meta.NameProp)

    Start-Busy 'Loading full policy...'
    try {
        $full = $p
        if ($meta.TypeName -eq 'Settings Catalog') {
            $full = Invoke-Graph -Uri "/deviceManagement/configurationPolicies/$id`?`$expand=settings" -Beta -Raw
        }
        Stop-Busy 'Ready'
        Show-JsonWindow -Title "$($meta.TypeName) - $name" -Object $full
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox $_.Exception.Message
    }
}

$script:BtnPolLoad.Add_Click({ Load-Policies })
$script:CmbPolType.Add_SelectedIndexChanged({ if ($script:GridPolicies.DataSource) { Load-Policies } })
$script:GridPolicies.Add_SelectionChanged({ Show-PolicyAssignments })
$script:BtnPolJson.Add_Click({ Show-PolicyJson })
$script:BtnPolExport.Add_Click({ Export-GridToCsv -Grid $script:GridPolicies -SuggestedName 'IntunePolicies' })
$script:BtnPolImport.Add_Click({ Import-PolicyFromJson })
$script:BtnPolAssign.Add_Click({
    $polId = Get-SelectedCellValue -Grid $script:GridPolicies -Column 'Id'
    if (-not $polId) { Show-InfoBox 'Select a policy first.'; return }
    $polName = Get-SelectedCellValue -Grid $script:GridPolicies -Column 'Name'
    $meta = $script:Cache['PolicyMeta']
    $kind = switch ($meta.TypeName) {
        'Settings Catalog'              { 'SettingsCatalog' }
        'Device Configuration (legacy)' { 'DeviceConfig' }
        'Compliance Policies'           { 'Compliance' }
        'Platform Scripts'              { 'Script' }
        'Remediations'                  { 'Remediation' }
        default                         { $null }
    }
    if (-not $kind) {
        Show-InfoBox "Assignments cannot be edited for '$($meta.TypeName)' in this tool. Use the Intune portal for those."
        return
    }
    Show-AssignmentEditor -Kind $kind -ObjectId $polId -ObjectName $polName
    Load-Policies
})

# ===========================================================================
# GROUP LOOKUP TAB
#   "What is assigned to this group?" - scans apps and every policy type
# ===========================================================================
$grpTop        = New-Object System.Windows.Forms.Panel
$grpTop.Dock   = 'Top'
$grpTop.Height = 44
$tabGroups.Controls.Add($grpTop)

$grpTop.Controls.Add((New-Label -Text 'Group name:' -X 4 -Y 10 -W 80 -H 24))
$script:TxtGrpSearch = New-TextBox -X 86 -Y 8 -W 260
$grpTop.Controls.Add($script:TxtGrpSearch)

$script:BtnGrpFind    = New-Button -Text 'Find Groups'      -X 356 -Y 7 -W 110
$script:BtnGrpAssign  = New-Button -Text 'Show Assignments' -X 474 -Y 7 -W 140
$script:BtnGrpMembers = New-Button -Text 'Show Members'     -X 622 -Y 7 -W 120
$script:BtnGrpExport  = New-Button -Text 'Export CSV'       -X 750 -Y 7 -W 100
$grpTop.Controls.AddRange(@($script:BtnGrpFind, $script:BtnGrpAssign, $script:BtnGrpMembers, $script:BtnGrpExport))

$grpSplit             = New-Object System.Windows.Forms.SplitContainer
$grpSplit.Dock        = 'Fill'
$grpSplit.Orientation = 'Horizontal'
$tabGroups.Controls.Add($grpSplit)
$grpSplit.BringToFront()

$script:GridGroups = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridGroups.Dock = 'Fill'
$grpSplit.Panel1.Controls.Add($script:GridGroups)

$script:GridGrpResults = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridGrpResults.Dock = 'Fill'
$grpSplit.Panel2.Controls.Add($script:GridGrpResults)

function Find-Groups {
    if (-not (Assert-Connected)) { return }
    $s = Escape-ODataValue ($script:TxtGrpSearch.Text.Trim())
    if (-not $s) { Show-InfoBox 'Enter part of a group name.'; return }

    Start-Busy 'Searching groups...'
    try {
        $groups = @(Invoke-Graph -Uri "/groups?`$filter=startswith(displayName,'$s')&`$select=id,displayName,description,mailNickname,groupTypes,membershipRule&`$top=100" -All)
        $view = foreach ($g in $groups) {
            [pscustomobject]@{
                GroupName   = $g.displayName
                Type        = if ($g.groupTypes -contains 'DynamicMembership') { 'Dynamic' } else { 'Assigned' }
                Nickname    = $g.mailNickname
                Rule        = $g.membershipRule
                Description = $g.description
                Id          = $g.id
            }
        }
        Set-GridData -Grid $script:GridGroups -Objects @($view | Sort-Object GroupName) `
            -Columns @('GroupName','Type','Nickname','Rule','Description','Id') | Out-Null
        Stop-Busy "$($view.Count) group(s) found"
    } catch {
        Stop-Busy 'Search failed'
        Show-ErrorBox $_.Exception.Message
    }
}

function Show-GroupAssignments {
    if (-not (Assert-Connected)) { return }
    $gid  = Get-SelectedCellValue -Grid $script:GridGroups -Column 'Id'
    $gname= Get-SelectedCellValue -Grid $script:GridGroups -Column 'GroupName'
    if (-not $gid) { Show-InfoBox 'Select a group first.'; return }

    $sources = @(
        @{ Label = 'App';                   Uri = "/deviceAppManagement/mobileApps?`$expand=assignments";                 Beta = $false; NameProp = 'displayName' },
        @{ Label = 'Settings Catalog';      Uri = "/deviceManagement/configurationPolicies?`$expand=assignments";         Beta = $true;  NameProp = 'name' },
        @{ Label = 'Device Configuration';  Uri = "/deviceManagement/deviceConfigurations?`$expand=assignments";          Beta = $false; NameProp = 'displayName' },
        @{ Label = 'Compliance Policy';     Uri = "/deviceManagement/deviceCompliancePolicies?`$expand=assignments";      Beta = $false; NameProp = 'displayName' },
        @{ Label = 'Platform Script';       Uri = "/deviceManagement/deviceManagementScripts?`$expand=assignments";       Beta = $true;  NameProp = 'displayName' },
        @{ Label = 'Remediation';           Uri = "/deviceManagement/deviceHealthScripts?`$expand=assignments";           Beta = $true;  NameProp = 'displayName' },
        @{ Label = 'Enrollment Config';     Uri = "/deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments";Beta = $false; NameProp = 'displayName' }
    )

    $results = New-Object System.Collections.Generic.List[object]
    Start-Busy "Scanning assignments for '$gname'..."
    try {
        foreach ($src in $sources) {
            Set-Status "Scanning $($src.Label)..."
            try {
                $items = @(Invoke-Graph -Uri $src.Uri -Beta:$src.Beta -All)
            } catch {
                [void]$results.Add([pscustomobject]@{
                    ObjectType = $src.Label; Name = "(could not read: $($_.Exception.Message))"
                    Intent = ''; Include = ''; Filter = ''; Id = ''
                })
                continue
            }
            foreach ($item in $items) {
                foreach ($asg in @($item.assignments)) {
                    if ($asg.target.groupId -eq $gid) {
                        $isExclude = ($asg.target.'@odata.type' -like '*exclusionGroupAssignmentTarget')
                        [void]$results.Add([pscustomobject]@{
                            ObjectType = $src.Label
                            Name       = $item.($src.NameProp)
                            Intent     = $asg.intent
                            Include    = if ($isExclude) { 'EXCLUDE' } else { 'Include' }
                            Filter     = $asg.target.deviceAndAppManagementAssignmentFilterId
                            Id         = $item.id
                        })
                    }
                }
            }
        }

        if ($results.Count -eq 0) {
            [void]$results.Add([pscustomobject]@{
                ObjectType = ''; Name = "Nothing is assigned to '$gname'"; Intent = ''; Include = ''; Filter = ''; Id = ''
            })
        }

        Set-GridData -Grid $script:GridGrpResults -Objects @($results.ToArray() | Sort-Object ObjectType, Name) `
            -Columns @('ObjectType','Name','Intent','Include','Filter','Id') | Out-Null
        Stop-Busy "$($results.Count) assignment(s) found for '$gname'"
    } catch {
        Stop-Busy 'Scan failed'
        Show-ErrorBox $_.Exception.Message
    }
}

function Show-GroupMembers {
    if (-not (Assert-Connected)) { return }
    $gid   = Get-SelectedCellValue -Grid $script:GridGroups -Column 'Id'
    $gname = Get-SelectedCellValue -Grid $script:GridGroups -Column 'GroupName'
    if (-not $gid) { Show-InfoBox 'Select a group first.'; return }

    Start-Busy "Loading members of '$gname'..."
    try {
        $members = @(Invoke-Graph -Uri "/groups/$gid/members?`$top=200" -All)
        $view = foreach ($m in $members) {
            [pscustomobject]@{
                Type        = ($m.'@odata.type' -replace '^#microsoft\.graph\.', '')
                DisplayName = $m.displayName
                UPN         = $m.userPrincipalName
                DeviceOS    = $m.operatingSystem
                Enabled     = if ($null -ne $m.accountEnabled) { $m.accountEnabled } else { '' }
                Id          = $m.id
            }
        }
        Set-GridData -Grid $script:GridGrpResults -Objects @($view | Sort-Object Type, DisplayName) `
            -Columns @('Type','DisplayName','UPN','DeviceOS','Enabled','Id') | Out-Null
        Stop-Busy "$($view.Count) member(s) - direct membership only"
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox $_.Exception.Message
    }
}

$script:BtnGrpFind.Add_Click({ Find-Groups })
$script:TxtGrpSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $_.SuppressKeyPress = $true; Find-Groups } })
$script:BtnGrpAssign.Add_Click({ Show-GroupAssignments })
$script:BtnGrpMembers.Add_Click({ Show-GroupMembers })
$script:BtnGrpExport.Add_Click({ Export-GridToCsv -Grid $script:GridGrpResults -SuggestedName 'GroupLookup' })

# ===========================================================================
# REPORTS TAB
# ===========================================================================
$rptTop        = New-Object System.Windows.Forms.Panel
$rptTop.Dock   = 'Top'
$rptTop.Height = 44
$tabReports.Controls.Add($rptTop)

$rptTop.Controls.Add((New-Label -Text 'Report:' -X 4 -Y 10 -W 50 -H 24))
$script:CmbReport = New-Object System.Windows.Forms.ComboBox
$script:CmbReport.Location      = New-Object System.Drawing.Point(56, 8)
$script:CmbReport.Size          = New-Object System.Drawing.Size(330, 24)
$script:CmbReport.DropDownStyle = 'DropDownList'
[void]$script:CmbReport.Items.AddRange(@(
    'Stale devices (no sync)',
    'Non-compliant devices',
    'Windows devices not encrypted',
    'Windows build inventory (summary)',
    'Devices with no primary user',
    'Duplicate device names',
    'Autopilot devices not enrolled in Intune'
))
$script:CmbReport.SelectedIndex = 0
$rptTop.Controls.Add($script:CmbReport)

$script:BtnRptRun    = New-Button -Text 'Run Report' -X 396 -Y 7 -W 110
$script:BtnRptExport = New-Button -Text 'Export CSV' -X 514 -Y 7 -W 100
$rptTop.Controls.AddRange(@($script:BtnRptRun, $script:BtnRptExport))

$script:GridReport = New-Grid -X 0 -Y 0 -W 100 -H 100
$script:GridReport.Dock = 'Fill'
$tabReports.Controls.Add($script:GridReport)
$script:GridReport.BringToFront()

function Get-AllManagedDevices {
    Invoke-Graph -Uri "/deviceManagement/managedDevices?`$select=$script:DeviceSelectColumns&`$top=200" -All
}

function Run-Report {
    if (-not (Assert-Connected)) { return }
    $name = [string]$script:CmbReport.SelectedItem
    Start-Busy "Running report: $name..."
    try {
        switch ($name) {
            'Stale devices (no sync)' {
                $cut = (Get-Date).AddDays(-1 * [int]$script:Config.StaleDays)
                $devices = @(Get-AllManagedDevices)
                $view = foreach ($d in $devices) {
                    if ($d.lastSyncDateTime -and ([datetime]$d.lastSyncDateTime) -lt $cut) {
                        [pscustomobject]@{
                            DeviceName = $d.deviceName
                            User       = $d.userPrincipalName
                            OS         = "$($d.operatingSystem) $($d.osVersion)"
                            Serial     = $d.serialNumber
                            LastSync   = Format-GraphDate $d.lastSyncDateTime
                            DaysStale  = [int]((Get-Date) - [datetime]$d.lastSyncDateTime).TotalDays
                            Id         = $d.id
                        }
                    }
                }
                Set-GridData -Grid $script:GridReport -Objects @($view | Sort-Object { [int]$_.DaysStale } -Descending) `
                    -Columns @('DeviceName','User','OS','Serial','LastSync','DaysStale','Id') | Out-Null
                Stop-Busy "$(@($view).Count) device(s) with no sync in $($script:Config.StaleDays) days"
            }

            'Non-compliant devices' {
                $devices = @(Get-AllManagedDevices | Where-Object { $_.complianceState -ne 'compliant' })
                $view = foreach ($d in $devices) {
                    [pscustomobject]@{
                        DeviceName = $d.deviceName
                        User       = $d.userPrincipalName
                        State      = $d.complianceState
                        OS         = "$($d.operatingSystem) $($d.osVersion)"
                        Owner      = $d.managedDeviceOwnerType
                        LastSync   = Format-GraphDate $d.lastSyncDateTime
                        Id         = $d.id
                    }
                }
                Set-GridData -Grid $script:GridReport -Objects @($view | Sort-Object State, DeviceName) `
                    -Columns @('DeviceName','User','State','OS','Owner','LastSync','Id') | Out-Null
                Stop-Busy "$(@($view).Count) non-compliant device(s)"
            }

            'Windows devices not encrypted' {
                $devices = @(Get-AllManagedDevices | Where-Object { $_.operatingSystem -eq 'Windows' -and -not $_.isEncrypted })
                $view = foreach ($d in $devices) {
                    [pscustomobject]@{
                        DeviceName = $d.deviceName
                        User       = $d.userPrincipalName
                        Model      = $d.model
                        Serial     = $d.serialNumber
                        OSVersion  = $d.osVersion
                        LastSync   = Format-GraphDate $d.lastSyncDateTime
                        Id         = $d.id
                    }
                }
                Set-GridData -Grid $script:GridReport -Objects @($view | Sort-Object DeviceName) `
                    -Columns @('DeviceName','User','Model','Serial','OSVersion','LastSync','Id') | Out-Null
                Stop-Busy "$(@($view).Count) unencrypted Windows device(s)"
            }

            'Windows build inventory (summary)' {
                $devices = @(Get-AllManagedDevices | Where-Object { $_.operatingSystem -eq 'Windows' })
                $view = $devices | Group-Object osVersion | Sort-Object Name -Descending | ForEach-Object {
                    [pscustomobject]@{
                        OSVersion = $_.Name
                        Count     = $_.Count
                        Devices   = (($_.Group | Select-Object -First 5 | ForEach-Object { $_.deviceName }) -join ', ')
                        Id        = $_.Name
                    }
                }
                Set-GridData -Grid $script:GridReport -Objects @($view) -Columns @('OSVersion','Count','Devices','Id') | Out-Null
                Stop-Busy "$(@($devices).Count) Windows device(s) across $(@($view).Count) build(s)"
            }

            'Devices with no primary user' {
                $devices = @(Get-AllManagedDevices | Where-Object { [string]::IsNullOrWhiteSpace($_.userPrincipalName) })
                $view = foreach ($d in $devices) {
                    [pscustomobject]@{
                        DeviceName = $d.deviceName
                        OS         = "$($d.operatingSystem) $($d.osVersion)"
                        Serial     = $d.serialNumber
                        Owner      = $d.managedDeviceOwnerType
                        Enrolled   = Format-GraphDate $d.enrolledDateTime
                        LastSync   = Format-GraphDate $d.lastSyncDateTime
                        Id         = $d.id
                    }
                }
                Set-GridData -Grid $script:GridReport -Objects @($view | Sort-Object DeviceName) `
                    -Columns @('DeviceName','OS','Serial','Owner','Enrolled','LastSync','Id') | Out-Null
                Stop-Busy "$(@($view).Count) device(s) with no primary user"
            }

            'Duplicate device names' {
                $devices = @(Get-AllManagedDevices)
                $dupes = $devices | Group-Object deviceName | Where-Object { $_.Count -gt 1 }
                $view = foreach ($g in $dupes) {
                    foreach ($d in $g.Group) {
                        [pscustomobject]@{
                            DeviceName = $d.deviceName
                            Serial     = $d.serialNumber
                            User       = $d.userPrincipalName
                            Enrolled   = Format-GraphDate $d.enrolledDateTime
                            LastSync   = Format-GraphDate $d.lastSyncDateTime
                            Agent      = $d.managementAgent
                            Id         = $d.id
                        }
                    }
                }
                Set-GridData -Grid $script:GridReport -Objects @($view | Sort-Object DeviceName, LastSync) `
                    -Columns @('DeviceName','Serial','User','Enrolled','LastSync','Agent','Id') | Out-Null
                Stop-Busy "$(@($dupes).Count) duplicated name(s), $(@($view).Count) record(s)"
            }

            'Autopilot devices not enrolled in Intune' {
                $ap = @(Invoke-Graph -Uri "/deviceManagement/windowsAutopilotDeviceIdentities?`$top=200" -All)
                $view = foreach ($d in $ap) {
                    if ([string]::IsNullOrWhiteSpace($d.managedDeviceId) -or $d.managedDeviceId -eq '00000000-0000-0000-0000-000000000000') {
                        [pscustomobject]@{
                            Serial        = $d.serialNumber
                            GroupTag      = $d.groupTag
                            Model         = $d.model
                            EnrollState   = $d.enrollmentState
                            ProfileStatus = $d.deploymentProfileAssignmentStatus
                            LastContact   = Format-GraphDate $d.lastContactedDateTime
                            Id            = $d.id
                        }
                    }
                }
                Set-GridData -Grid $script:GridReport -Objects @($view | Sort-Object GroupTag, Serial) `
                    -Columns @('Serial','GroupTag','Model','EnrollState','ProfileStatus','LastContact','Id') | Out-Null
                Stop-Busy "$(@($view).Count) Autopilot device(s) with no Intune record"
            }
        }
    } catch {
        Stop-Busy 'Report failed'
        Show-ErrorBox "Report failed:`r`n$($_.Exception.Message)"
    }
}

$script:BtnRptRun.Add_Click({ Run-Report })
$script:BtnRptExport.Add_Click({ Export-GridToCsv -Grid $script:GridReport -SuggestedName 'IntuneReport' })

# ===========================================================================
# PACKAGING TAB  (local - IntuneWinAppUtil.exe)
# ===========================================================================
$pkgTop        = New-Object System.Windows.Forms.Panel
$pkgTop.Dock   = 'Top'
$pkgTop.Height = 108
$tabPackaging.Controls.Add($pkgTop)

$pkgTop.Controls.Add((New-Label -Text 'Packaging root:' -X 4 -Y 10 -W 95 -H 24))
$script:TxtPkgRoot = New-TextBox -X 100 -Y 8 -W 620 -Text $script:Config.PackagingRoot -Anchor 'Top,Left,Right'
$pkgTop.Controls.Add($script:TxtPkgRoot)

$script:BtnPkgBrowse  = New-Button -Text 'Browse...' -X 728 -Y 7 -W 90 -Anchor 'Top,Right'
$script:BtnPkgRefresh = New-Button -Text 'Refresh'   -X 824 -Y 7 -W 90 -Anchor 'Top,Right'
$pkgTop.Controls.AddRange(@($script:BtnPkgBrowse, $script:BtnPkgRefresh))

$script:BtnPkgCreate    = New-Button -Text 'Create .intunewin'   -X 4   -Y 44 -W 140
$script:BtnPkgCreateAll = New-Button -Text 'Create Multiple...'  -X 150 -Y 44 -W 140
$script:BtnPkgMsiInfo   = New-Button -Text 'MSI Properties'      -X 296 -Y 44 -W 120
$script:BtnPkgOpen      = New-Button -Text 'Open Folder'         -X 422 -Y 44 -W 110
$pkgTop.Controls.AddRange(@($script:BtnPkgCreate, $script:BtnPkgCreateAll, $script:BtnPkgMsiInfo, $script:BtnPkgOpen))

$script:PkgProgress          = New-Object System.Windows.Forms.ProgressBar
$script:PkgProgress.Location = New-Object System.Drawing.Point(4, 80)
$script:PkgProgress.Size     = New-Object System.Drawing.Size(910, 18)
$script:PkgProgress.Anchor   = 'Top,Left,Right'
$script:PkgProgress.Style    = 'Continuous'
$script:PkgProgress.Visible  = $false
$pkgTop.Controls.Add($script:PkgProgress)

$pkgSplit             = New-Object System.Windows.Forms.SplitContainer
$pkgSplit.Dock        = 'Fill'
$pkgSplit.Orientation = 'Horizontal'
$tabPackaging.Controls.Add($pkgSplit)
$pkgSplit.BringToFront()

$pkgLists             = New-Object System.Windows.Forms.SplitContainer
$pkgLists.Dock        = 'Fill'
$pkgLists.Orientation = 'Vertical'
$pkgSplit.Panel1.Controls.Add($pkgLists)

$script:LstPkgFolders           = New-Object System.Windows.Forms.ListBox
$script:LstPkgFolders.Dock      = 'Fill'
$script:LstPkgFolders.Font      = New-Object System.Drawing.Font('Segoe UI', 10)
$pkgLists.Panel1.Controls.Add($script:LstPkgFolders)
$lblF = New-Object System.Windows.Forms.Label
$lblF.Text = ' Application folders'
$lblF.Dock = 'Top'
$lblF.Height = 22
$lblF.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$pkgLists.Panel1.Controls.Add($lblF)

$script:LstPkgFiles           = New-Object System.Windows.Forms.ListBox
$script:LstPkgFiles.Dock      = 'Fill'
$script:LstPkgFiles.Font      = New-Object System.Drawing.Font('Segoe UI', 10)
$pkgLists.Panel2.Controls.Add($script:LstPkgFiles)
$lblS = New-Object System.Windows.Forms.Label
$lblS.Text = ' Setup files in selected folder (pick the installer)'
$lblS.Dock = 'Top'
$lblS.Height = 22
$lblS.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$pkgLists.Panel2.Controls.Add($lblS)

$script:TxtPkgLog            = New-Object System.Windows.Forms.TextBox
$script:TxtPkgLog.Multiline  = $true
$script:TxtPkgLog.ScrollBars = 'Both'
$script:TxtPkgLog.WordWrap   = $false
$script:TxtPkgLog.ReadOnly   = $true
$script:TxtPkgLog.Dock       = 'Fill'
$script:TxtPkgLog.Font       = New-Object System.Drawing.Font('Consolas', 9)
$script:TxtPkgLog.Tag        = 'output'
$script:TxtPkgLog.BackColor  = [System.Drawing.Color]::White
$pkgSplit.Panel2.Controls.Add($script:TxtPkgLog)

function Write-PkgLog {
    param([string]$Text, [switch]$Raw)
    if ($Raw) {
        $script:TxtPkgLog.AppendText($Text)
    } else {
        $script:TxtPkgLog.AppendText("$Text`r`n")
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Start-PkgProgress {
    # A real 0-100 bar. The packager reports no percentage directly, but its
    # log names each stage and the temp archive it is building, so progress is
    # derived from actual work done rather than animated for show.
    if (-not $script:PkgProgress) { return }
    $script:PkgProgress.Style   = 'Continuous'
    $script:PkgProgress.Minimum = 0
    $script:PkgProgress.Maximum = 100
    $script:PkgProgress.Value   = 0
    $script:PkgProgress.Visible = $true
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-PkgProgressPercent {
    param([int]$Percent)
    if (-not $script:PkgProgress) { return }
    if ($Percent -lt 0)   { $Percent = 0 }
    if ($Percent -gt 100) { $Percent = 100 }
    # never go backwards within one package
    if ($Percent -lt $script:PkgProgress.Value) { return }
    $script:PkgProgress.Value = $Percent
}

function Reset-PkgProgressState {
    $script:PkgTempArchive = $null
    $script:PkgTotalBytes  = 0
    $script:PkgStage       = 0
    if ($script:PkgProgress) { $script:PkgProgress.Value = 0 }
}

function Update-PkgProgressFromText {
    # Maps the packager's stage lines onto the bar. The compress stage is the
    # long one, so it is tracked by the size of the temp archive as it grows.
    param([string]$Text)
    if (-not $Text) { return }

    if ($Text -match "Compressing the source folder '[^']+' to '([^']+)'") {
        $script:PkgTempArchive = $Matches[1]
        $script:PkgStage = 5
    }
    if ($Text -match "Calculated size for folder '[^']+' is (\d+)") {
        if ($script:PkgTotalBytes -le 0) { $script:PkgTotalBytes = [int64]$Matches[1] }
    }

    if ($Text -match 'Compressed folder .* successfully') { $script:PkgStage = 70; $script:PkgTempArchive = $null }
    if ($Text -match 'Checking file type')                { $script:PkgStage = 72 }
    if ($Text -match 'Encrypting file')                   { $script:PkgStage = 75 }
    if ($Text -match 'has been encrypted successfully')   { $script:PkgStage = 80 }
    if ($Text -match 'Computing SHA256 hash')             { $script:PkgStage = 83 }
    if ($Text -match 'Copying encrypted file')            { $script:PkgStage = 86 }
    if ($Text -match 'Generating detection XML')          { $script:PkgStage = 89 }
    if ($Text -match "Compressing folder '[^']*IntuneWinPackage'") { $script:PkgStage = 92 }
    if ($Text -match 'Removing temporary files')          { $script:PkgStage = 97 }
    if ($Text -match 'has been generated successfully')   { $script:PkgStage = 99 }
    if ($Text -match 'Done!!!')                           { $script:PkgStage = 100 }

    Set-PkgProgressPercent $script:PkgStage
}

function Update-PkgProgressFromArchive {
    # During the compress stage, fill the bar from 5% to 68% based on how much
    # of the source folder has made it into the temp archive.
    if (-not $script:PkgTempArchive -or $script:PkgTotalBytes -le 0) { return }
    try {
        $fi = New-Object IO.FileInfo($script:PkgTempArchive)
        if ($fi.Exists) {
            $ratio = [double]$fi.Length / [double]$script:PkgTotalBytes
            if ($ratio -gt 1) { $ratio = 1 }
            Set-PkgProgressPercent ([int](5 + (63 * $ratio)))
        }
    } catch { }
}

function Stop-PkgProgress {
    if (-not $script:PkgProgress) { return }
    $script:PkgProgress.MarqueeAnimationSpeed = 0
    $script:PkgProgress.Style   = 'Continuous'
    $script:PkgProgress.Value   = 0
    $script:PkgProgress.Visible = $false
    [System.Windows.Forms.Application]::DoEvents()
}

function Refresh-PackagingFolders {
    $root = $script:TxtPkgRoot.Text.Trim()
    $script:LstPkgFolders.Items.Clear()
    $script:LstPkgFiles.Items.Clear()
    if (-not $root -or -not (Test-Path $root)) { return }
    Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name | ForEach-Object { [void]$script:LstPkgFolders.Items.Add($_.Name) }
    Set-Status "$($script:LstPkgFolders.Items.Count) application folder(s)"
}

function Refresh-PackagingFiles {
    $script:LstPkgFiles.Items.Clear()
    if (-not $script:LstPkgFolders.SelectedItem) { return }
    $folder = Join-Path $script:TxtPkgRoot.Text.Trim() $script:LstPkgFolders.SelectedItem
    if (-not (Test-Path $folder)) { return }
    Get-ChildItem -Path $folder -File -Include *.msi, *.exe, *.ps1, *.cmd, *.bat -Recurse -ErrorAction SilentlyContinue |
        Sort-Object Name | ForEach-Object {
            $rel = $_.FullName.Substring($folder.Length).TrimStart('\')
            [void]$script:LstPkgFiles.Items.Add($rel)
        }
    if ($script:LstPkgFiles.Items.Count -eq 1) { $script:LstPkgFiles.SelectedIndex = 0 }
}

function Invoke-IntuneWinAppUtil {
    param([string]$SourceFolder, [string]$SetupFile, [string]$OutputFolder)

    $util = $script:Config.IntuneWinAppUtil
    if (-not $util -or -not (Test-Path $util)) {
        Show-ErrorBox "IntuneWinAppUtil.exe not found. Set its path on the Settings tab."
        return $false
    }

    $outTmp = [IO.Path]::GetTempFileName()
    $errTmp = [IO.Path]::GetTempFileName()
    $argList = @('-c', "`"$SourceFolder`"", '-s', "`"$SetupFile`"", '-o', "`"$OutputFolder`"", '-q')

    Write-PkgLog ''
    Write-PkgLog ('=' * 78)
    Write-PkgLog "Source : $SourceFolder"
    Write-PkgLog "Setup  : $SetupFile"
    Write-PkgLog "Output : $OutputFolder"
    Write-PkgLog ('=' * 78)

    $proc   = $null
    $reader = $null
    $stream = $null
    $startedAt = Get-Date
    Reset-PkgProgressState

    try {
        $utilDir = Split-Path -Parent $util
        $proc = Start-Process -FilePath $util -ArgumentList $argList -NoNewWindow -PassThru `
                    -WorkingDirectory $utilDir `
                    -RedirectStandardOutput $outTmp -RedirectStandardError $errTmp

        # touching Handle caches it - without this ExitCode is unreadable once
        # the process has exited, and reads back as empty
        try { $null = $proc.Handle } catch { }

        # follow the redirected file while the packager is still writing to it,
        # so the log fills in live and the window stays responsive
        for ($try = 0; $try -lt 20 -and -not $stream; $try++) {
            try {
                $stream = New-Object IO.FileStream($outTmp, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            } catch {
                Start-Sleep -Milliseconds 25
            }
        }
        if ($stream) { $reader = New-Object IO.StreamReader($stream) }

        while (-not $proc.HasExited) {
            if ($reader) {
                $chunk = $reader.ReadToEnd()
                if ($chunk) {
                    Write-PkgLog -Text $chunk -Raw
                    Update-PkgProgressFromText -Text $chunk
                }
            }
            Update-PkgProgressFromArchive
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 75
        }

        $proc.WaitForExit()

        if ($reader) {
            $chunk = $reader.ReadToEnd()
            if ($chunk) {
                Write-PkgLog -Text $chunk -Raw
                Update-PkgProgressFromText -Text $chunk
            }
        }

        if ($reader) { $reader.Close(); $reader = $null }
        if ($stream) { $stream.Close(); $stream = $null }

        $stderr = Get-Content -Path $errTmp -Raw -ErrorAction SilentlyContinue
        if ($stderr) { Write-PkgLog "STDERR: $($stderr.TrimEnd())" }

        $exitCode = $null
        try { $exitCode = $proc.ExitCode } catch { }

        # the package file is the source of truth; the exit code is a secondary
        # signal and is not always readable
        $expected = Join-Path $OutputFolder ([IO.Path]::GetFileNameWithoutExtension($SetupFile) + '.intunewin')
        $produced = $false
        $item = $null
        if (Test-Path -LiteralPath $expected) {
            $item = Get-Item -LiteralPath $expected -ErrorAction SilentlyContinue
            if ($item -and $item.LastWriteTime -ge $startedAt.AddSeconds(-5)) { $produced = $true }
        }

        $elapsed = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 1)

        if ($produced -and ($null -eq $exitCode -or $exitCode -eq 0)) {
            $sizeMB = [math]::Round($item.Length / 1MB, 1)
            Write-PkgLog "SUCCESS -> $expected  ($sizeMB MB, $elapsed seconds)"
            return $true
        }

        if ($produced) {
            Write-PkgLog "SUCCESS (packager returned exit code $exitCode) -> $expected"
            Write-PkgLog 'Review the output above before using this package.'
            return $true
        }

        $shown = if ($null -eq $exitCode) { 'not reported' } else { $exitCode }
        Write-PkgLog "FAILED - no package produced at '$expected' (exit code: $shown)"
        return $false
    } catch {
        Write-PkgLog "ERROR: $($_.Exception.Message)"
        return $false
    } finally {
        if ($reader) { try { $reader.Close() } catch { } }
        if ($stream) { try { $stream.Close() } catch { } }
        Remove-Item $outTmp, $errTmp -Force -ErrorAction SilentlyContinue
    }
}

function New-IntunewinPackage {
    if (-not $script:LstPkgFolders.SelectedItem) { Show-InfoBox 'Select an application folder.'; return }
    if (-not $script:LstPkgFiles.SelectedItem)   { Show-InfoBox 'Select the setup file.'; return }

    $folder = Join-Path $script:TxtPkgRoot.Text.Trim() $script:LstPkgFolders.SelectedItem
    $setup  = Join-Path $folder $script:LstPkgFiles.SelectedItem

    Start-Busy "Packaging $($script:LstPkgFolders.SelectedItem)..."
    Start-PkgProgress
    try {
        $ok = Invoke-IntuneWinAppUtil -SourceFolder $folder -SetupFile $setup -OutputFolder $folder
    } finally {
        Stop-PkgProgress
    }
    Stop-Busy $(if ($ok) { 'Package created' } else { 'Packaging failed' })
}

function Get-SetupCandidates {
    # Anything that can plausibly be the setup file, including script wrappers.
    param([string]$FolderPath)
    $exts = @('.msi', '.exe', '.ps1', '.cmd', '.bat')
    $files = @(Get-ChildItem -Path $FolderPath -File -Recurse -ErrorAction SilentlyContinue |
               Where-Object { $exts -contains $_.Extension.ToLower() } |
               Sort-Object FullName)
    return $files
}

function Show-PackagingPicker {
    # Lists every application folder with a dropdown of its candidate setup
    # files. Nothing is packaged without an explicit tick and an explicit file.
    param([string]$Root)

    $folders = @(Get-ChildItem -Path $Root -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($folders.Count -eq 0) {
        Show-InfoBox "No application folders found under:`r`n$Root"
        return $null
    }

    $f               = New-Object System.Windows.Forms.Form
    $f.Text          = 'Create Multiple .intunewin Packages'
    $f.Size          = New-Object System.Drawing.Size(940, 620)
    $f.StartPosition = 'CenterParent'
    $f.MinimumSize   = New-Object System.Drawing.Size(700, 400)

    $top        = New-Object System.Windows.Forms.Panel
    $top.Dock   = 'Top'
    $top.Height = 46
    $f.Controls.Add($top)

    $lbl = New-Label -Text "Tick the folders to package and confirm the setup file for each.  Root: $Root" -X 10 -Y 12 -W 900 -H 24
    $top.Controls.Add($lbl)

    $bottom        = New-Object System.Windows.Forms.Panel
    $bottom.Dock   = 'Bottom'
    $bottom.Height = 52
    $f.Controls.Add($bottom)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock                   = 'Fill'
    $grid.AllowUserToAddRows     = $false
    $grid.AllowUserToDeleteRows  = $false
    $grid.RowHeadersVisible      = $false
    $grid.SelectionMode          = 'CellSelect'
    $grid.AutoSizeColumnsMode    = 'Fill'
    $grid.EditMode               = 'EditOnEnter'
    $grid.Font                   = New-Object System.Drawing.Font('Segoe UI', 9)
    $grid.Add_DataError({ param($sndr, $ev) $ev.ThrowException = $false })
    $f.Controls.Add($grid)
    $grid.BringToFront()

    $colChk            = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colChk.Name       = 'Package'
    $colChk.HeaderText = 'Package'
    $colChk.FillWeight = 12
    [void]$grid.Columns.Add($colChk)

    $colFolder             = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFolder.Name        = 'Folder'
    $colFolder.HeaderText  = 'Application folder'
    $colFolder.ReadOnly    = $true
    $colFolder.FillWeight  = 30
    [void]$grid.Columns.Add($colFolder)

    $colSetup             = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colSetup.Name        = 'Setup'
    $colSetup.HeaderText  = 'Setup file'
    $colSetup.FillWeight  = 45
    $colSetup.DisplayStyle = 'ComboBox'
    [void]$grid.Columns.Add($colSetup)

    $colFound             = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFound.Name        = 'Found'
    $colFound.HeaderText  = 'Candidates'
    $colFound.ReadOnly    = $true
    $colFound.FillWeight  = 13
    [void]$grid.Columns.Add($colFound)

    $pathMap = @{}   # folder name -> full folder path

    foreach ($folder in $folders) {
        $candidates = Get-SetupCandidates -FolderPath $folder.FullName
        $rel = foreach ($c in $candidates) { $c.FullName.Substring($folder.FullName.Length).TrimStart('\') }
        $rel = @($rel)

        $i = $grid.Rows.Add()
        $row = $grid.Rows[$i]
        $pathMap[$folder.Name] = $folder.FullName

        $row.Cells['Folder'].Value = $folder.Name
        $row.Cells['Found'].Value  = "$($rel.Count)"

        $cell = $row.Cells['Setup']
        [void]$cell.Items.Add('')
        foreach ($r in $rel) { [void]$cell.Items.Add($r) }

        if ($rel.Count -eq 1) {
            # unambiguous - preselect and pretick
            $cell.Value                = $rel[0]
            $row.Cells['Package'].Value = $true
        } else {
            $cell.Value = ''
            $row.Cells['Package'].Value = $false
            if ($rel.Count -eq 0) {
                $row.ReadOnly = $true
                $row.Cells['Found'].Value = 'none'
            }
        }
    }

    $btnAll = New-Button -Text 'Select All' -X 10 -Y 12 -W 100 -H 28
    $btnAll.Add_Click({
        [void]$grid.EndEdit()
        foreach ($r in $grid.Rows) {
            if ([string]$r.Cells['Setup'].Value) { $r.Cells['Package'].Value = $true }
        }
    })
    $bottom.Controls.Add($btnAll)

    $btnNone = New-Button -Text 'Select None' -X 116 -Y 12 -W 100 -H 28
    $btnNone.Add_Click({
        [void]$grid.EndEdit()
        foreach ($r in $grid.Rows) { $r.Cells['Package'].Value = $false }
    })
    $bottom.Controls.Add($btnNone)

    $btnOk = New-Button -Text 'Package Selected' -X 620 -Y 12 -W 150 -H 28 -Anchor 'Bottom,Right'
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $bottom.Controls.Add($btnOk)

    $btnCancel = New-Button -Text 'Cancel' -X 780 -Y 12 -W 110 -H 28 -Anchor 'Bottom,Right'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $bottom.Controls.Add($btnCancel)
    $f.CancelButton = $btnCancel

    Set-ControlTheme -Control $f
    $result = $f.ShowDialog()

    $selection = New-Object System.Collections.Generic.List[object]
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # commit any cell still being edited before reading the checkbox values
        [void]$grid.EndEdit()
        [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)

        foreach ($r in $grid.Rows) {
            $ticked = $false
            if ($null -ne $r.Cells['Package'].Value) { $ticked = [bool]$r.Cells['Package'].Value }
            if (-not $ticked) { continue }

            $setupRel = [string]$r.Cells['Setup'].Value
            if ([string]::IsNullOrWhiteSpace($setupRel)) { continue }

            $name = [string]$r.Cells['Folder'].Value
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $full = $pathMap[$name]
            if (-not $full) { continue }
            [void]$selection.Add([pscustomobject]@{
                FolderName = $name
                FolderPath = $full
                SetupFile  = (Join-Path $full $setupRel)
                SetupRel   = $setupRel
            })
        }
    }

    $f.Dispose()
    return $selection.ToArray()
}

function New-IntunewinPackagesFromPicker {
    $root = $script:TxtPkgRoot.Text.Trim()
    if (-not $root -or -not (Test-Path $root)) { Show-InfoBox 'Set a valid packaging root first.'; return }

    $selection = @(Show-PackagingPicker -Root $root | Where-Object { $_ -and $_.FolderPath })
    if ($selection.Count -eq 0) { return }

    $lines = foreach ($sel in $selection) { "  $($sel.FolderName)  ->  $($sel.SetupRel)" }
    if (-not (Confirm-Action -Message "Package $(@($selection).Count) folder(s)?`r`n`r`n$($lines -join "`r`n")" -Title 'Create Multiple')) { return }

    $script:TxtPkgLog.Clear()
    $done = 0; $failed = 0
    $total = $selection.Count
    Start-Busy "Packaging $total folder(s)..."
    Start-PkgProgress

    try {
        foreach ($sel in $selection) {
            Set-Status "Packaging $($sel.FolderName) ($($done + $failed + 1) of $total)..."
            if (Invoke-IntuneWinAppUtil -SourceFolder $sel.FolderPath -SetupFile $sel.SetupFile -OutputFolder $sel.FolderPath) {
                $done++
            } else {
                $failed++
            }
        }
    } finally {
        Stop-PkgProgress
    }

    Write-PkgLog ''
    Write-PkgLog "Batch complete: $done created, $failed failed."
    Stop-Busy "Batch complete: $done created, $failed failed"
}

function Get-MsiProperty {
    param([string]$Path, [string]$Property)
    try {
        $wi = New-Object -ComObject WindowsInstaller.Installer
        $db = $wi.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $wi, @($Path, 0))
        $q  = "SELECT Value FROM Property WHERE Property = '$Property'"
        $vw = $db.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $db, @($q))
        $vw.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $vw, $null) | Out-Null
        $rec = $vw.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $vw, $null)
        if ($rec) {
            return $rec.GetType().InvokeMember('StringData', 'GetProperty', $null, $rec, 1)
        }
        return ''
    } catch {
        return ''
    } finally {
        if ($wi) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wi) | Out-Null }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Show-MsiInfo {
    if (-not $script:LstPkgFiles.SelectedItem) { Show-InfoBox 'Select an MSI file.'; return }
    $folder = Join-Path $script:TxtPkgRoot.Text.Trim() $script:LstPkgFolders.SelectedItem
    $msi    = Join-Path $folder $script:LstPkgFiles.SelectedItem
    if ($msi -notmatch '\.msi$') { Show-InfoBox 'MSI properties can only be read from an .msi file.'; return }

    Start-Busy 'Reading MSI properties...'
    $code = Get-MsiProperty -Path $msi -Property 'ProductCode'
    $ver  = Get-MsiProperty -Path $msi -Property 'ProductVersion'
    $name = Get-MsiProperty -Path $msi -Property 'ProductName'
    $mfg  = Get-MsiProperty -Path $msi -Property 'Manufacturer'
    Stop-Busy 'Ready'

    $script:TxtPkgLog.AppendText(@"

MSI: $msi
  ProductName    : $name
  Manufacturer   : $mfg
  ProductVersion : $ver
  ProductCode    : $code

  Suggested Intune settings
    Install command   : msiexec /i "$([IO.Path]::GetFileName($msi))" /qn /norestart
    Uninstall command : msiexec /x $code /qn /norestart
    Detection rule    : MSI product code = $code

"@)
}

$script:BtnPkgBrowse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = 'Select the root folder that contains one subfolder per application'
    if ($script:TxtPkgRoot.Text -and (Test-Path $script:TxtPkgRoot.Text)) { $fb.SelectedPath = $script:TxtPkgRoot.Text }
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:TxtPkgRoot.Text = $fb.SelectedPath
        $script:Config.PackagingRoot = $fb.SelectedPath
        Save-Config | Out-Null
        Refresh-PackagingFolders
    }
})
$script:BtnPkgRefresh.Add_Click({ Refresh-PackagingFolders })
# Clicking the already-selected folder clears the selection, so Open Folder goes
# back to the packaging root.
#
# This compares the selection against what it was at the end of the previous
# click rather than watching event order: WinForms does not guarantee whether
# SelectedIndexChanged is raised before or after MouseUp, which made an
# event-order approach fire only sometimes. If a click leaves SelectedIndex
# exactly where the last click left it, the item was already selected.
$script:PkgPrevSelIndex = -1

$script:LstPkgFolders.Add_MouseUp({
    param($sndr, $e)

    $idx = $script:LstPkgFolders.IndexFromPoint($e.X, $e.Y)
    if ($idx -lt 0) {
        # click landed below the last item - leave the selection alone
        $script:PkgPrevSelIndex = $script:LstPkgFolders.SelectedIndex
        return
    }

    if ($script:LstPkgFolders.SelectedIndex -eq $script:PkgPrevSelIndex) {
        $script:LstPkgFolders.ClearSelected()
        $script:LstPkgFiles.Items.Clear()
        Set-Status 'Folder selection cleared'
    }

    $script:PkgPrevSelIndex = $script:LstPkgFolders.SelectedIndex
})

$script:LstPkgFolders.Add_SelectedIndexChanged({ Refresh-PackagingFiles })
$script:BtnPkgCreate.Add_Click({ New-IntunewinPackage })
$script:BtnPkgCreateAll.Add_Click({ New-IntunewinPackagesFromPicker })
$script:BtnPkgMsiInfo.Add_Click({ Show-MsiInfo })
$script:BtnPkgOpen.Add_Click({
    # the path must be quoted: these folders contain spaces and commas, and an
    # unquoted argument makes Explorer give up and open its default location
    $target = $script:TxtPkgRoot.Text.Trim()
    if ($script:LstPkgFolders.SelectedItem) {
        $candidate = Join-Path $target $script:LstPkgFolders.SelectedItem
        if (Test-Path -LiteralPath $candidate) { $target = $candidate }
    }
    if ([string]::IsNullOrWhiteSpace($target) -or -not (Test-Path -LiteralPath $target)) {
        Show-InfoBox "Folder not found:`r`n$target"
        return
    }
    Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $target + '"')
})

# ===========================================================================
# SETTINGS TAB
# ===========================================================================
$setFields            = New-Object System.Windows.Forms.Panel
$setFields.Dock       = 'Top'
$setFields.Height     = 284
$tabSettings.Controls.Add($setFields)

$y = 16
$setFields.Controls.Add((New-Label -Text 'Tenant ID (GUID or domain):' -X 12 -Y $y -W 200 -H 24))
$script:TxtTenant = New-TextBox -X 216 -Y ($y - 2) -W 420 -Text $script:Config.TenantId
$setFields.Controls.Add($script:TxtTenant)

$y += 34
$setFields.Controls.Add((New-Label -Text 'Client ID (app registration):' -X 12 -Y $y -W 200 -H 24))
$script:TxtClient = New-TextBox -X 216 -Y ($y - 2) -W 420 -Text $script:Config.ClientId
$setFields.Controls.Add($script:TxtClient)

$y += 34
$setFields.Controls.Add((New-Label -Text 'IntuneWinAppUtil.exe path:' -X 12 -Y $y -W 200 -H 24))
$script:TxtUtilPath = New-TextBox -X 216 -Y ($y - 2) -W 420 -Text $script:Config.IntuneWinAppUtil
$setFields.Controls.Add($script:TxtUtilPath)
$btnUtilBrowse = New-Button -Text 'Browse...' -X 644 -Y ($y - 3) -W 90
$btnUtilBrowse.Add_Click({
    $od = New-Object System.Windows.Forms.OpenFileDialog
    $od.Filter = 'IntuneWinAppUtil.exe|IntuneWinAppUtil.exe|Executable (*.exe)|*.exe'
    if ($od.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $script:TxtUtilPath.Text = $od.FileName }
})
$setFields.Controls.Add($btnUtilBrowse)

$y += 34
$setFields.Controls.Add((New-Label -Text 'Packaging root folder:' -X 12 -Y $y -W 200 -H 24))
$script:TxtSetPkgRoot = New-TextBox -X 216 -Y ($y - 2) -W 420 -Text $script:Config.PackagingRoot
$setFields.Controls.Add($script:TxtSetPkgRoot)

$script:BtnPkgRootBrowse = New-Button -Text 'Browse...' -X 644 -Y ($y - 3) -W 90
$script:BtnPkgRootBrowse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = 'Select the root folder that contains one subfolder per application'
    $cur = $script:TxtSetPkgRoot.Text.Trim()
    if ($cur -and (Test-Path $cur -ErrorAction SilentlyContinue)) { $fb.SelectedPath = $cur }
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $script:TxtSetPkgRoot.Text = $fb.SelectedPath }
})
$setFields.Controls.Add($script:BtnPkgRootBrowse)

$script:ChkPkgSameAsUtil          = New-Object System.Windows.Forms.CheckBox
$script:ChkPkgSameAsUtil.Text     = 'Same folder as IntuneWinAppUtil.exe'
$script:ChkPkgSameAsUtil.Location = New-Object System.Drawing.Point(744, ($y - 1))
$script:ChkPkgSameAsUtil.Size     = New-Object System.Drawing.Size(260, 24)
$script:ChkPkgSameAsUtil.Checked  = [bool]$script:Config.PkgRootFollowsUtil
$setFields.Controls.Add($script:ChkPkgSameAsUtil)

$y += 34
$setFields.Controls.Add((New-Label -Text 'Stale device threshold (days):' -X 12 -Y $y -W 200 -H 24))
$script:TxtStaleDays = New-TextBox -X 216 -Y ($y - 2) -W 80 -Text ([string]$script:Config.StaleDays)
$setFields.Controls.Add($script:TxtStaleDays)

$y += 34
$setFields.Controls.Add((New-Label -Text 'Sign-in method:' -X 12 -Y $y -W 200 -H 24))
$script:CmbSignIn = New-Object System.Windows.Forms.ComboBox
$script:CmbSignIn.Location      = New-Object System.Drawing.Point(216, ($y - 2))
$script:CmbSignIn.Size          = New-Object System.Drawing.Size(200, 24)
$script:CmbSignIn.DropDownStyle = 'DropDownList'
[void]$script:CmbSignIn.Items.AddRange(@('Browser', 'Device code'))
$script:CmbSignIn.SelectedItem  = $(if ($script:Config.SignInMethod -eq 'Device code') { 'Device code' } else { 'Browser' })
$setFields.Controls.Add($script:CmbSignIn)
$setFields.Controls.Add((New-Label -Text 'Browser = normal sign-in window. Device code = fallback if loopback is blocked.' -X 424 -Y $y -W 520 -H 24))

$y += 44
function Sync-PackagingRootWithUtil {
    if (-not $script:ChkPkgSameAsUtil -or -not $script:ChkPkgSameAsUtil.Checked) { return }
    $util = $script:TxtUtilPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($util)) { return }
    $dir = ''
    try { $dir = Split-Path -Parent $util } catch { $dir = '' }
    if ($dir) { $script:TxtSetPkgRoot.Text = $dir }
}

function Update-PackagingRootLock {
    $linked = [bool]$script:ChkPkgSameAsUtil.Checked
    $script:TxtSetPkgRoot.ReadOnly      = $linked
    $script:BtnPkgRootBrowse.Enabled    = -not $linked
    if ($linked) { Sync-PackagingRootWithUtil }
}

$script:ChkPkgSameAsUtil.Add_CheckedChanged({
    if ($script:ChkPkgSameAsUtil.Checked -and [string]::IsNullOrWhiteSpace($script:TxtUtilPath.Text.Trim())) {
        Show-InfoBox 'Set the IntuneWinAppUtil.exe path first, then tick this box.'
        $script:ChkPkgSameAsUtil.Checked = $false
        return
    }
    Update-PackagingRootLock
})

# keep the linked packaging root in step when the utility path changes
$script:TxtUtilPath.Add_TextChanged({ Sync-PackagingRootWithUtil })

Update-PackagingRootLock

$btnSaveSettings = New-Button -Text 'Save Settings' -X 216 -Y $y -W 130 -H 30
$btnSaveSettings.Add_Click({
    $script:Config.TenantId         = $script:TxtTenant.Text.Trim()
    $script:Config.ClientId         = $script:TxtClient.Text.Trim()
    $script:Config.IntuneWinAppUtil = $script:TxtUtilPath.Text.Trim()
    $script:Config.PackagingRoot    = $script:TxtSetPkgRoot.Text.Trim()
    $script:Config.SignInMethod     = [string]$script:CmbSignIn.SelectedItem
    $script:Config.PkgRootFollowsUtil = [bool]$script:ChkPkgSameAsUtil.Checked
    $d = 90
    if ([int]::TryParse($script:TxtStaleDays.Text.Trim(), [ref]$d)) { $script:Config.StaleDays = $d }
    if (Save-Config) {
        $script:TxtPkgRoot.Text = $script:Config.PackagingRoot
        Refresh-PackagingFolders
        Show-InfoBox "Settings saved to:`r`n$script:ConfigPath"
    }
})
$setFields.Controls.Add($btnSaveSettings)

$btnClearToken = New-Button -Text 'Clear Cached Token' -X 356 -Y $y -W 150 -H 30
$btnClearToken.Add_Click({
    Clear-TokenCache
    Set-ConnectedState $false
    Show-InfoBox 'Cached refresh token removed. Click Connect to sign in again.'
})
$setFields.Controls.Add($btnClearToken)

$btnResetDefaults = New-Button -Text 'Set to Org Defaults' -X 516 -Y $y -W 170 -H 30
$btnResetDefaults.Add_Click({
    $org = Get-MSToolkitOrgDefaults

    if (-not $org.TenantId -and -not $org.ClientId) {
        Show-InfoBox ("No organization defaults were found.`r`n`r`n" +
            "Enter Tenant ID and Client ID in MSToolkit Settings (Microsoft 365 and Intune), " +
            "save, then open IntuneTools from MSToolkit.")
        return
    }

    $shownTenant = if ($org.TenantId) { $org.TenantId } else { '(not set in MSToolkit - left unchanged)' }
    $shownClient = if ($org.ClientId) { $org.ClientId } else { '(not set in MSToolkit - left unchanged)' }

    $msg = "Set Tenant ID and Client ID to the values in MSToolkit Settings?`r`n`r`n" +
           "  Tenant ID : $shownTenant`r`n" +
           "  Client ID : $shownClient`r`n`r`n" +
           "From: $($org.Source)`r`n`r`n" +
           "Other settings are not changed. Nothing is saved until you click Save Settings."
    if (-not (Confirm-Action -Message $msg -Title 'Set to Org Defaults')) { return }

    if ($org.TenantId) { $script:TxtTenant.Text = $org.TenantId }
    if ($org.ClientId) { $script:TxtClient.Text = $org.ClientId }

    Show-InfoBox 'Org defaults copied into the boxes above. Click Save Settings to keep them.'
})
$setFields.Controls.Add($btnResetDefaults)

$btnClearSettings = New-Button -Text 'Clear Settings' -X 696 -Y $y -W 130 -H 30
$btnClearSettings.Add_Click({
    $msg = "Clear the IntuneTools settings and save them?`r`n`r`n" +
           "  Tenant ID                : cleared`r`n" +
           "  Client ID                : cleared`r`n" +
           "  IntuneWinAppUtil.exe path: cleared`r`n" +
           "  Packaging root folder    : cleared, and unlinked from the utility folder`r`n" +
           "  Stale device threshold   : back to 90 days`r`n" +
           "  Sign-in method           : back to Browser`r`n" +
           "  Cached sign-in token     : removed - you will be disconnected`r`n`r`n" +
           "The theme is kept. This is saved straight away to:`r`n$script:ConfigPath"
    if (-not (Confirm-Action -Message $msg -Title 'Clear Settings')) { return }

    # Boxes first, so the Settings tab shows exactly what is saved.
    $script:ChkPkgSameAsUtil.Checked = $false
    $script:TxtTenant.Text     = ''
    $script:TxtClient.Text     = ''
    $script:TxtUtilPath.Text   = ''
    $script:TxtSetPkgRoot.Text = ''
    $script:TxtStaleDays.Text  = '90'
    $script:CmbSignIn.SelectedItem = 'Browser'
    Update-PackagingRootLock

    $script:Config.TenantId            = ''
    $script:Config.ClientId            = ''
    $script:Config.IntuneWinAppUtil    = ''
    $script:Config.PackagingRoot       = ''
    $script:Config.StaleDays           = 90
    $script:Config.SignInMethod        = 'Browser'
    $script:Config.PkgRootFollowsUtil  = $false

    # The cached token belongs to the tenant and app that were just cleared.
    Clear-TokenCache
    Set-ConnectedState $false

    if (Save-Config) {
        $script:TxtPkgRoot.Text = ''
        Refresh-PackagingFolders
        Show-InfoBox "Settings cleared and saved to:`r`n$script:ConfigPath"
    }
})
$setFields.Controls.Add($btnClearSettings)

$y += 48
$setHelp                 = New-Object System.Windows.Forms.RichTextBox
$setHelp.Dock            = 'Fill'
$setHelp.ReadOnly        = $true
$setHelp.BorderStyle     = 'None'
$setHelp.BackColor       = [System.Drawing.Color]::White
$setHelp.WordWrap        = $true
$setHelp.ScrollBars      = 'Vertical'
$setHelp.DetectUrls      = $false
$setHelp.Margin          = New-Object System.Windows.Forms.Padding(0)
$tabSettings.Controls.Add($setHelp)
$setHelp.BringToFront()
$script:SettingsHelpBox = $setHelp

function Add-HelpText {
    param(
        [System.Windows.Forms.RichTextBox]$Box,
        [string]$Text = '',
        [ValidateSet('Title','Head','Body','Code','Note')][string]$Style = 'Body',
        [int]$Indent = 0
    )
    $Box.SelectionStart  = $Box.TextLength
    $Box.SelectionLength = 0
    $Box.SelectionIndent = $Indent
    $Box.SelectionHangingIndent = 0

    $T = $script:Theme
    if (-not $T) { $T = Get-ThemePalette -Name 'Light' }

    switch ($Style) {
        'Title' {
            $Box.SelectionFont  = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
            $Box.SelectionColor = $T.Section
        }
        'Head' {
            $Box.SelectionFont  = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
            $Box.SelectionColor = $T.Section
        }
        'Code' {
            $Box.SelectionFont  = New-Object System.Drawing.Font('Consolas', 9.5)
            $Box.SelectionColor = $T.Info
        }
        'Note' {
            $Box.SelectionFont  = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)
            $Box.SelectionColor = $T.MutedText
        }
        default {
            $Box.SelectionFont  = New-Object System.Drawing.Font('Segoe UI', 9.5)
            $Box.SelectionColor = $T.Text
        }
    }
    $Box.AppendText($Text + "`r`n")
}

function Write-SettingsHelp {
    param([System.Windows.Forms.RichTextBox]$Box)

    $Box.Clear()

    Add-HelpText $Box 'One-time app registration setup' 'Title'
    Add-HelpText $Box 'Microsoft Entra admin center  >  App registrations  >  New registration' 'Note'
    Add-HelpText $Box

    Add-HelpText $Box 'Register the app' 'Head'
    Add-HelpText $Box 'Name:  IntuneTools' 'Code' 24
    Add-HelpText $Box 'Supported account types:  Accounts in this organizational directory only (single tenant)' 'Code' 24
    Add-HelpText $Box 'Redirect URI:  platform "Mobile and desktop applications", value  http://localhost' 'Code' 24
    Add-HelpText $Box 'Then open Authentication and set  "Allow public client flows" = Yes.' 'Body' 24
    Add-HelpText $Box 'Copy the Application (client) ID and Directory (tenant) ID into the boxes above.' 'Body' 24
    Add-HelpText $Box

    Add-HelpText $Box 'Why http://localhost' 'Head'
    Add-HelpText $Box 'That redirect URI is what lets the Connect button open a normal sign-in window instead of asking for a copied code. Entra ignores the port on loopback redirect URIs, so the single entry http://localhost covers whichever random port this tool listens on. Use localhost, not 127.0.0.1 - the port is only ignored for localhost.' 'Body' 24
    Add-HelpText $Box 'No client secret is needed or wanted. This is a public client and the sign-in is delegated to you.' 'Body' 24
    Add-HelpText $Box

    Add-HelpText $Box 'API permissions' 'Head'
    Add-HelpText $Box 'Add a permission  >  Microsoft Graph  >  Delegated permissions' 'Note' 24
    Add-HelpText $Box

    Add-HelpText $Box 'READ - needed for every list and report' 'Body' 24
    Add-HelpText $Box 'DeviceManagementScripts.Read.All covers the Policies tab Platform Scripts and Remediations views.' 'Note' 24
    foreach ($p in @(
        'DeviceManagementManagedDevices.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'DeviceManagementApps.Read.All',
        'DeviceManagementServiceConfig.Read.All',
        'DeviceManagementScripts.Read.All',
        'Group.Read.All',
        'User.Read.All',
        'Directory.Read.All',
        'offline_access')) {
        Add-HelpText $Box $p 'Code' 48
    }
    Add-HelpText $Box

    Add-HelpText $Box 'WRITE - only if you want the action buttons to work' 'Body' 24
    Add-HelpText $Box 'DeviceManagementManagedDevices.ReadWrite.All' 'Code' 48
    Add-HelpText $Box 'sync, restart, shut down, rename, delete record' 'Note' 72
    Add-HelpText $Box 'DeviceManagementManagedDevices.PrivilegedOperations.All' 'Code' 48
    Add-HelpText $Box 'retire, wipe, Autopilot Reset, Fresh Start' 'Note' 72
    Add-HelpText $Box 'DeviceManagementServiceConfig.ReadWrite.All' 'Code' 48
    Add-HelpText $Box 'Autopilot group tag, assign user, delete record' 'Note' 72
    Add-HelpText $Box 'DeviceManagementScripts.ReadWrite.All' 'Code' 48
    Add-HelpText $Box 'editing platform scripts and remediations' 'Note' 72
    Add-HelpText $Box 'DeviceManagementApps.ReadWrite.All' 'Code' 48
    Add-HelpText $Box 'editing app assignments' 'Note' 72
    Add-HelpText $Box 'DeviceManagementConfiguration.ReadWrite.All' 'Code' 48
    Add-HelpText $Box 'editing policy assignments, importing policies' 'Note' 72
    Add-HelpText $Box 'GroupMember.ReadWrite.All' 'Code' 48
    Add-HelpText $Box 'adding or removing devices and users from Entra groups' 'Note' 72
    Add-HelpText $Box 'Your signed-in account also needs a directory role that can write group membership - Groups Administrator, User Administrator, or ownership of the group. Intune Administrator alone is not enough.' 'Note' 48
    Add-HelpText $Box

    Add-HelpText $Box 'OPTIONAL - secrets, add only if you want those tabs to work' 'Body' 24
    Add-HelpText $Box 'BitlockerKey.Read.All' 'Code' 48
    Add-HelpText $Box 'BitLocker recovery keys' 'Note' 72
    Add-HelpText $Box 'DeviceLocalCredential.Read.All' 'Code' 48
    Add-HelpText $Box 'Windows LAPS local admin password' 'Note' 72
    Add-HelpText $Box
    Add-HelpText $Box 'Finish with  "Grant admin consent for <tenant>".' 'Body' 24
    Add-HelpText $Box

    Add-HelpText $Box 'How sign-in behaves' 'Head'
    foreach ($n in @(
        'Sign-in is delegated. Your own Intune RBAC still applies on top of these permissions, and every action is recorded in the Intune audit log under your account.',
        'Start with the READ list only. Add write permissions once you trust the tool.',
        'The refresh token is cached under %APPDATA%\MSToolkit\IntuneTools encrypted with DPAPI, so it can only be read by your account on this machine. After the first sign-in the Connect button reconnects silently, with no browser window at all.',
        'Once connected the button reads Reconnect and forces the account picker, for switching accounts.',
        'If your workstation blocks the loopback listener, the tool offers device code sign-in as a fallback. You can also force that with the Sign-in method list above.')) {
        Add-HelpText $Box ('-  ' + $n) 'Body' 24
    }
    Add-HelpText $Box

    Add-HelpText $Box 'Defaults and saved settings' 'Head'
    foreach ($n in @(
        'Tenant ID and Client ID start blank. "Set to Org Defaults" copies them from MSToolkit Settings (Microsoft 365 and Intune) - passed in when MSToolkit opens this tool, or read from %APPDATA%\MSToolkit\settings.json for this account when run on its own. The boxes stay editable.',
        'Neither value is a secret - the client ID identifies a public client app.',
        'Per-user settings are saved to %APPDATA%\MSToolkit\IntuneTools\config.json when you click Save Settings.')) {
        Add-HelpText $Box ('-  ' + $n) 'Body' 24
    }

    $Box.SelectionStart = 0
    $Box.ScrollToCaret()
}

Write-SettingsHelp -Box $setHelp


# ===========================================================================
# SHARED PICKERS AND WRITE ACTIONS
# ===========================================================================

function Show-GridWindow {
    param([string]$Title, [object[]]$Objects, [string[]]$Columns, [string]$ExportName = 'export')

    $f = New-Object System.Windows.Forms.Form
    $f.Text          = $Title
    $f.Size          = New-Object System.Drawing.Size(1100, 640)
    $f.StartPosition = 'CenterParent'

    $g = New-Grid -X 0 -Y 0 -W 100 -H 100
    $g.Dock = 'Fill'
    $f.Controls.Add($g)

    $panel        = New-Object System.Windows.Forms.Panel
    $panel.Dock   = 'Bottom'
    $panel.Height = 46
    $f.Controls.Add($panel)

    $btnExport = New-Button -Text 'Export CSV' -X 10 -Y 9 -W 110 -H 28
    $btnExport.Add_Click({ Export-GridToCsv -Grid $g -SuggestedName $ExportName })
    $panel.Controls.Add($btnExport)

    $btnClose = New-Button -Text 'Close' -X 130 -Y 9 -W 100 -H 28
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $panel.Controls.Add($btnClose)
    $f.AcceptButton = $btnClose

    Set-GridData -Grid $g -Objects @($Objects) -Columns $Columns | Out-Null
    Set-ControlTheme -Control $f
    [void]$f.ShowDialog()
    $f.Dispose()
}

$script:MemberOfCache = @{}

function Get-DirectMemberOfIds {
    # Direct group membership only - nested membership is not counted here,
    # because that is what can actually be added to or removed from.
    param([string]$MemberUri)
    if (-not $MemberUri) { return @() }
    if ($script:MemberOfCache.ContainsKey($MemberUri)) { return $script:MemberOfCache[$MemberUri] }
    try {
        $groups = @(Invoke-Graph -Uri "$MemberUri/memberOf?`$select=id&`$top=200" -All)
        $ids = @($groups | ForEach-Object { $_.id } | Where-Object { $_ })
        $script:MemberOfCache[$MemberUri] = $ids
        return $ids
    } catch {
        $script:MemberOfCache[$MemberUri] = @()
        return @()
    }
}

function Clear-MemberOfCache { $script:MemberOfCache = @{} }

function Show-GroupPicker {
    # Search Entra groups and return the chosen one, or $null.
    # Members, when supplied, is a list of @{ Name = ...; Uri = '/devices/<id>' }
    # and drives the "Already a member" column.
    param(
        [string]$Prompt = 'Search for a group',
        [object[]]$Members
    )

    $f = New-Object System.Windows.Forms.Form
    $f.Text          = 'Select a group'
    $f.Size          = New-Object System.Drawing.Size(760, 500)
    $f.StartPosition = 'CenterParent'

    $top        = New-Object System.Windows.Forms.Panel
    $top.Dock   = 'Top'
    $top.Height = 44
    $f.Controls.Add($top)

    $top.Controls.Add((New-Label -Text $Prompt -X 8 -Y 12 -W 150 -H 22))
    if (@($Members).Count -gt 0) {
        $f.Text = "Select a group - 'Already a member' shows current direct membership"
    }
    $txt = New-TextBox -X 160 -Y 10 -W 350
    $top.Controls.Add($txt)

    $btnFind = New-Button -Text 'Search' -X 520 -Y 9 -W 100 -H 26
    $top.Controls.Add($btnFind)

    $grid = New-Grid -X 0 -Y 0 -W 100 -H 100
    $grid.Dock = 'Fill'
    $f.Controls.Add($grid)
    $grid.BringToFront()

    $bottom        = New-Object System.Windows.Forms.Panel
    $bottom.Dock   = 'Bottom'
    $bottom.Height = 46
    $f.Controls.Add($bottom)

    $btnOk = New-Button -Text 'Select' -X 520 -Y 9 -W 100 -H 28 -Anchor 'Bottom,Right'
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $bottom.Controls.Add($btnOk)

    $btnCancel = New-Button -Text 'Cancel' -X 628 -Y 9 -W 100 -H 28 -Anchor 'Bottom,Right'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $bottom.Controls.Add($btnCancel)
    $f.CancelButton = $btnCancel

    $doSearch = {
        $term = Escape-ODataValue ($txt.Text.Trim())
        if (-not $term) { return }
        try {
            $groups = @(Invoke-Graph -Uri "/groups?`$filter=startswith(displayName,'$term')&`$select=id,displayName,description,groupTypes&`$top=100" -All)

            # membership of the thing being added, so existing members are obvious
            $memberCount = @($Members).Count
            $memberMap   = @{}
            if ($memberCount -gt 0) {
                $n = 0
                foreach ($m in $Members) {
                    $n++
                    Set-Status "Checking current group membership ($n of $memberCount)..."
                    foreach ($gid in (Get-DirectMemberOfIds -MemberUri $m.Uri)) {
                        if ($memberMap.ContainsKey($gid)) { $memberMap[$gid] = $memberMap[$gid] + 1 }
                        else { $memberMap[$gid] = 1 }
                    }
                }
                Set-Status 'Ready'
            }

            $view = foreach ($grp in $groups) {
                $already = ''
                if ($memberCount -eq 1) {
                    if ($memberMap.ContainsKey($grp.id)) { $already = 'Yes' }
                } elseif ($memberCount -gt 1) {
                    $hit = 0
                    if ($memberMap.ContainsKey($grp.id)) { $hit = $memberMap[$grp.id] }
                    if ($hit -gt 0) { $already = "$hit of $memberCount" }
                }
                [pscustomobject]@{
                    GroupName     = $grp.displayName
                    AlreadyMember = $already
                    Type          = if ($grp.groupTypes -contains 'DynamicMembership') { 'Dynamic' } else { 'Assigned' }
                    Description   = $grp.description
                    Id            = $grp.id
                }
            }
            Set-GridData -Grid $grid -Objects @($view | Sort-Object GroupName) `
                -Columns @('GroupName','AlreadyMember','Type','Description','Id') | Out-Null
        } catch {
            Show-ErrorBox $_.Exception.Message
        }
    }
    $btnFind.Add_Click($doSearch)
    $txt.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $_.SuppressKeyPress = $true; & $doSearch } })

    Set-ControlTheme -Control $f
    $result = $f.ShowDialog()

    $picked = $null
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $gid   = Get-SelectedCellValue -Grid $grid -Column 'Id'
        $gname = Get-SelectedCellValue -Grid $grid -Column 'GroupName'
        $gtype = Get-SelectedCellValue -Grid $grid -Column 'Type'
        if ($gid) { $picked = [pscustomobject]@{ Id = $gid; Name = $gname; Type = $gtype } }
    }
    $f.Dispose()
    return $picked
}

# ---------------------------------------------------------------------------
# Assignment editor
# ---------------------------------------------------------------------------
function Get-AssignmentKindInfo {
    # Each object type has its own assign action and its own wrapper property.
    param([string]$Kind)
    switch ($Kind) {
        'App'             { return @{ Base = '/deviceAppManagement/mobileApps';          Beta = $false; Prop = 'mobileAppAssignments';           AssignType = '#microsoft.graph.mobileAppAssignment'; HasIntent = $true  } }
        'SettingsCatalog' { return @{ Base = '/deviceManagement/configurationPolicies';   Beta = $true;  Prop = 'assignments';                    AssignType = $null;                                  HasIntent = $false } }
        'DeviceConfig'    { return @{ Base = '/deviceManagement/deviceConfigurations';    Beta = $false; Prop = 'assignments';                    AssignType = '#microsoft.graph.deviceConfigurationAssignment'; HasIntent = $false } }
        'Compliance'      { return @{ Base = '/deviceManagement/deviceCompliancePolicies';Beta = $false; Prop = 'assignments';                    AssignType = $null;                                  HasIntent = $false } }
        'Script'          { return @{ Base = '/deviceManagement/deviceManagementScripts'; Beta = $true;  Prop = 'deviceManagementScriptAssignments'; AssignType = $null;                               HasIntent = $false } }
        'Remediation'     { return @{ Base = '/deviceManagement/deviceHealthScripts';     Beta = $true;  Prop = 'deviceHealthScriptAssignments';  AssignType = $null;                                  HasIntent = $false } }
    }
    return $null
}

$script:AssignIntentTarget  = $null
$script:AssignIntentRefresh = $null

function Get-AssignWriteScope {
    param([string]$Kind)
    switch ($Kind) {
        'App'             { return 'DeviceManagementApps.ReadWrite.All' }
        'SettingsCatalog' { return 'DeviceManagementConfiguration.ReadWrite.All' }
        'DeviceConfig'    { return 'DeviceManagementConfiguration.ReadWrite.All' }
        'Compliance'      { return 'DeviceManagementConfiguration.ReadWrite.All' }
        'Script'          { return 'DeviceManagementScripts.ReadWrite.All' }
        'Remediation'     { return 'DeviceManagementScripts.ReadWrite.All' }
    }
    return $null
}

function Show-AssignmentEditor {
    param([string]$Kind, [string]$ObjectId, [string]$ObjectName)

    if (-not (Assert-Connected)) { return }
    $info = Get-AssignmentKindInfo -Kind $Kind
    if (-not $info) { Show-ErrorBox "Assignments are not supported for '$Kind' in this tool."; return }

    # Load the current assignments fresh - the grid cache may be stale.
    # Use the assignments navigation property directly: $expand=assignments on a
    # single object does not reliably return them for every policy type.
    Start-Busy 'Loading current assignments...'
    $current = @()
    try {
        $current = @(Invoke-Graph -Uri "$($info.Base)/$ObjectId/assignments" -Beta:$info.Beta -All)
    } catch {
        # fall back to the expand form in case a type does not expose the
        # navigation property on its own
        try {
            $obj = Invoke-Graph -Uri "$($info.Base)/$ObjectId`?`$expand=assignments" -Beta:$info.Beta -Raw
            $current = @($obj.assignments)
        } catch {
            Stop-Busy 'Failed'
            Show-ErrorBox "Could not read assignments:`r`n$($_.Exception.Message)"
            return
        }
    }
    Stop-Busy "$($current.Count) existing assignment(s)"

    # working copy: one row per assignment
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($a in $current) {
        $t = $a.target
        $kindText = 'Group'
        $gid = $t.groupId
        switch -Wildcard ($t.'@odata.type') {
            '*allDevicesAssignmentTarget'       { $kindText = 'All Devices'; $gid = $null }
            '*allLicensedUsersAssignmentTarget' { $kindText = 'All Users';   $gid = $null }
            '*exclusionGroupAssignmentTarget'   { $kindText = 'Exclude' }
        }
        [void]$rows.Add([pscustomobject]@{
            TargetKind = $kindText
            GroupName  = if ($gid) { Get-GroupDisplayName $gid } else { '' }
            GroupId    = $gid
            Intent     = if ($a.intent) { $a.intent } else { '' }
            FilterId   = $t.deviceAndAppManagementAssignmentFilterId
            FilterMode = $t.deviceAndAppManagementAssignmentFilterType
        })
    }

    $f = New-Object System.Windows.Forms.Form
    $f.Text          = "Assignments - $ObjectName"
    $f.Size          = New-Object System.Drawing.Size(940, 560)
    $f.StartPosition = 'CenterParent'

    $top        = New-Object System.Windows.Forms.Panel
    $top.Dock   = 'Top'
    $top.Height = 82
    $f.Controls.Add($top)

    $lbl = New-Label -Text "$Kind : $ObjectName" -X 10 -Y 8 -W 880 -H 22
    $lbl.Tag = 'section'
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $top.Controls.Add($lbl)

    $btnAddInc = New-Button -Text 'Add Group'     -X 10  -Y 42 -W 110 -H 28
    $btnAddExc = New-Button -Text 'Add Exclude'   -X 126 -Y 42 -W 110 -H 28
    $btnAddAll = New-Button -Text 'Add All...'    -X 242 -Y 42 -W 100 -H 28
    $btnRemove = New-Button -Text 'Remove'        -X 348 -Y 42 -W 100 -H 28
    $btnIntent = New-Button -Text 'Set Intent'    -X 454 -Y 42 -W 110 -H 28
    $btnIntent.Enabled = [bool]$info.HasIntent
    $top.Controls.AddRange(@($btnAddInc, $btnAddExc, $btnAddAll, $btnRemove, $btnIntent))

    $grid = New-Grid -X 0 -Y 0 -W 100 -H 100
    $grid.Dock = 'Fill'
    $f.Controls.Add($grid)
    $grid.BringToFront()

    $bottom        = New-Object System.Windows.Forms.Panel
    $bottom.Dock   = 'Bottom'
    $bottom.Height = 50
    $f.Controls.Add($bottom)

    $warn = New-Label -Text 'Save replaces the whole assignment list on this object.' -X 10 -Y 14 -W 470 -H 22
    $bottom.Controls.Add($warn)

    $btnSave = New-Button -Text 'Save Assignments' -X 660 -Y 10 -W 150 -H 30 -Anchor 'Bottom,Right'
    $bottom.Controls.Add($btnSave)

    $btnCancel = New-Button -Text 'Cancel' -X 818 -Y 10 -W 100 -H 30 -Anchor 'Bottom,Right'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $bottom.Controls.Add($btnCancel)
    $f.CancelButton = $btnCancel

    $refresh = {
        Set-GridData -Grid $grid -Objects @($rows.ToArray()) `
            -Columns @('TargetKind','GroupName','Intent','FilterId','FilterMode','GroupId') | Out-Null
    }
    & $refresh

    # Match the selected row back to the working object by its values rather
    # than by row index: the grid can be sorted by clicking a column header,
    # and then the indexes no longer line up with the list.
    $findRow = {
        if ($grid.SelectedRows.Count -eq 0) { return $null }
        $sel = $grid.SelectedRows[0]
        if (-not $sel.DataBoundItem) { return $null }
        $kind = [string]$sel.DataBoundItem.Row['TargetKind']
        $gid  = [string]$sel.DataBoundItem.Row['GroupId']
        $gnam = [string]$sel.DataBoundItem.Row['GroupName']
        foreach ($item in $rows) {
            if (([string]$item.TargetKind -eq $kind) -and
                ([string]$item.GroupId   -eq $gid)  -and
                ([string]$item.GroupName -eq $gnam)) {
                return $item
            }
        }
        return $null
    }

    $addGroup = {
        param([bool]$Exclude)
        $picked = Show-GroupPicker -Prompt 'Group name starts with'
        if (-not $picked) { return }
        foreach ($r in $rows) {
            if ($r.GroupId -eq $picked.Id -and
                (($r.TargetKind -eq 'Exclude') -eq $Exclude)) {
                Show-InfoBox "'$($picked.Name)' is already in the list."
                return
            }
        }
        $intent = ''
        if ($info.HasIntent) { $intent = 'required' }
        [void]$rows.Add([pscustomobject]@{
            TargetKind = $(if ($Exclude) { 'Exclude' } else { 'Group' })
            GroupName  = $picked.Name
            GroupId    = $picked.Id
            Intent     = $intent
            FilterId   = ''
            FilterMode = ''
        })
        & $refresh
    }

    $btnAddInc.Add_Click({ & $addGroup $false })
    $btnAddExc.Add_Click({ & $addGroup $true })

    $btnAddAll.Add_Click({
        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        $i1 = $menu.Items.Add('All Devices')
        $i1.Add_Click({
            [void]$rows.Add([pscustomobject]@{
                TargetKind='All Devices'; GroupName=''; GroupId=$null
                Intent=$(if ($info.HasIntent) { 'required' } else { '' }); FilterId=''; FilterMode='' })
            & $refresh
        })
        $i2 = $menu.Items.Add('All Users')
        $i2.Add_Click({
            [void]$rows.Add([pscustomobject]@{
                TargetKind='All Users'; GroupName=''; GroupId=$null
                Intent=$(if ($info.HasIntent) { 'required' } else { '' }); FilterId=''; FilterMode='' })
            & $refresh
        })
        $menu.Show($btnAddAll, 0, $btnAddAll.Height)
    })

    $btnRemove.Add_Click({
        $target = & $findRow
        if (-not $target) { Show-InfoBox 'Select an assignment row first.'; return }
        [void]$rows.Remove($target)
        & $refresh
    })

    $btnIntent.Add_Click({
        if (-not $info.HasIntent) { return }
        $target = & $findRow
        if (-not $target) { Show-InfoBox 'Select an assignment row first.'; return }

        # ContextMenuStrip.Show returns straight away, so the row being edited
        # and the redraw have to survive this handler returning
        $script:AssignIntentTarget  = $target
        $script:AssignIntentRefresh = $refresh

        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        foreach ($choice in @('required','available','uninstall','availableWithoutEnrollment')) {
            $item = $menu.Items.Add($choice)
            $item.Tag = $choice
            $item.Add_Click({
                param($sndr, $ev)
                if ($script:AssignIntentTarget) {
                    $script:AssignIntentTarget.Intent = [string]$sndr.Tag
                    if ($script:AssignIntentRefresh) { & $script:AssignIntentRefresh }
                }
            })
        }
        $menu.Show($btnIntent, 0, $btnIntent.Height)
    })

    $btnSave.Add_Click({
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows) {
            $target = @{}
            switch ($r.TargetKind) {
                'All Devices' { $target['@odata.type'] = '#microsoft.graph.allDevicesAssignmentTarget' }
                'All Users'   { $target['@odata.type'] = '#microsoft.graph.allLicensedUsersAssignmentTarget' }
                'Exclude'     { $target['@odata.type'] = '#microsoft.graph.exclusionGroupAssignmentTarget'; $target['groupId'] = $r.GroupId }
                default       { $target['@odata.type'] = '#microsoft.graph.groupAssignmentTarget';          $target['groupId'] = $r.GroupId }
            }
            if ($r.FilterId) {
                $target['deviceAndAppManagementAssignmentFilterId']   = $r.FilterId
                $target['deviceAndAppManagementAssignmentFilterType'] = $(if ($r.FilterMode) { $r.FilterMode } else { 'include' })
            }

            $entry = @{ target = $target }
            if ($info.AssignType) { $entry['@odata.type'] = $info.AssignType }
            if ($info.HasIntent)  { $entry['intent'] = $(if ($r.Intent) { $r.Intent } else { 'required' }) }
            [void]$list.Add($entry)
        }

        $summary = if ($list.Count -eq 0) {
            "Remove ALL assignments from '$ObjectName'?"
        } else {
            $desc = foreach ($r in $rows) {
                $who = if ($r.GroupName) { $r.GroupName } else { $r.TargetKind }
                if ($r.Intent) { "  $($r.TargetKind): $who ($($r.Intent))" } else { "  $($r.TargetKind): $who" }
            }
            "Replace the assignments on '$ObjectName' with:`r`n`r`n$($desc -join "`r`n")"
        }
        if (-not (Confirm-Action -Message $summary -Title 'Save Assignments')) { return }

        $body = @{}
        $body[$info.Prop] = $list.ToArray()

        Start-Busy 'Saving assignments...'
        try {
            Invoke-Graph -Uri "$($info.Base)/$ObjectId/assign" -Method POST -Body $body -Beta:$info.Beta -Raw | Out-Null
            Stop-Busy 'Assignments saved'
            Show-InfoBox "Assignments updated on '$ObjectName'."
            $f.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $f.Close()
        } catch {
            Stop-Busy 'Save failed'
            $msg = $_.Exception.Message
            if ($msg -match 'Forbidden|not authorized|scopes') {
                $scope = Get-AssignWriteScope -Kind $Kind
                $hint  = "Saving assignments on a $Kind needs a write scope this sign-in does not have."
                if ($scope) {
                    $hint += "`r`n`r`nAdd this delegated permission to the app registration and grant admin consent:`r`n  $scope`r`n`r`nThen use Settings > Clear Cached Token and sign in again, or the old token keeps the old scopes."
                }
                Show-ErrorBox "$hint`r`n`r`nGraph returned:`r`n$msg"
            } else {
                Show-ErrorBox "Could not save assignments:`r`n$msg"
            }
        }
    })

    Set-ControlTheme -Control $f
    [void]$f.ShowDialog()
    $f.Dispose()
}

# ---------------------------------------------------------------------------
# Device write actions
# ---------------------------------------------------------------------------
function Get-MembershipErrorHint {
    param([string]$Message)
    if ($Message -match 'Authorization_RequestDenied|Insufficient privileges') {
        return "`r`n`r`nThis needs the GroupMember.ReadWrite.All delegated permission on the app registration, granted with admin consent - then Settings > Clear Cached Token and sign in again.`r`n`r`nYour signed-in account also needs a directory role that can write group membership (Groups Administrator, User Administrator, or ownership of this group). Intune Administrator on its own is not enough."
    }
    return ''
}

function Test-GroupIsAssigned {
    # Dynamic groups are rule-driven; members cannot be added or removed by hand.
    param($Picked)
    if ($Picked -and $Picked.Type -eq 'Dynamic') {
        Show-InfoBox "'$($Picked.Name)' is a dynamic group. Its membership comes from its membership rule, so members cannot be added or removed directly. Change the rule in Entra instead."
        return $false
    }
    return $true
}

function Get-EntraDeviceObjectId {
    # The Entra directory object id, which is what group membership uses.
    # This is NOT the Intune managed device id or the azureADDeviceId.
    param([string]$AzureAdDeviceId)
    if (-not $AzureAdDeviceId) { return $null }
    try {
        $d = @(Invoke-Graph -Uri "/devices?`$filter=deviceId eq '$AzureAdDeviceId'&`$select=id,displayName" -All)
        if ($d.Count -gt 0) { return $d[0].id }
    } catch { }
    return $null
}

function Set-DevicePrimaryUser {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id

    $upn = Show-InputDialog -Title 'Set Primary User' `
            -Prompt "User principal name for $($d.deviceName) (blank removes the primary user):" `
            -Default $d.userPrincipalName
    if ($null -eq $upn) { return }
    $upn = $upn.Trim()

    Start-Busy 'Updating primary user...'
    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            if (-not (Confirm-Action -Message "Remove the primary user from $($d.deviceName)?" -Title 'Set Primary User')) { Stop-Busy 'Ready'; return }
            Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/users/`$ref" -Method DELETE -Beta -Raw | Out-Null
            Stop-Busy 'Primary user removed'
        } else {
            $u = @(Invoke-Graph -Uri "/users?`$filter=userPrincipalName eq '$(Escape-ODataValue $upn)'&`$select=id,displayName,userPrincipalName" -All)
            if ($u.Count -eq 0) { Stop-Busy 'Not found'; Show-ErrorBox "No user found with UPN '$upn'."; return }
            $body = @{ '@odata.id' = "https://graph.microsoft.com/beta/users/$($u[0].id)" }
            Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/users/`$ref" -Method POST -Body $body -Beta -Raw | Out-Null
            Stop-Busy "Primary user set to $($u[0].userPrincipalName)"
        }
        Load-Devices
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox "Could not change the primary user:`r`n$($_.Exception.Message)"
    }
}

function Rename-ManagedDevice {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id

    $newName = Show-InputDialog -Title 'Rename Device' `
                -Prompt "New device name (serial $($d.serialNumber)):" -Default $d.deviceName
    if ([string]::IsNullOrWhiteSpace($newName)) { return }
    $newName = $newName.Trim()
    if ($newName -eq $d.deviceName) { return }

    if (-not (Confirm-Action -Message "Rename`r`n  $($d.deviceName)  (serial $($d.serialNumber))`r`nto`r`n  $newName`r`n`r`nThe device applies the new name on its next check-in and may need a restart." -Title 'Rename Device')) { return }

    Start-Busy 'Renaming device...'
    try {
        Invoke-Graph -Uri "/deviceManagement/managedDevices/$id/setDeviceName" -Method POST -Body @{ deviceName = $newName } -Raw | Out-Null
        Stop-Busy "Rename requested - $($d.deviceName) -> $newName"
        Show-InfoBox "Rename requested. The device picks up '$newName' at its next check-in."
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox "Could not rename the device:`r`n$($_.Exception.Message)"
    }
}

$script:PendingCategoryDeviceId = $null
$script:PendingCategoryDevice    = $null

function Set-ManagedDeviceCategory {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id

    $cats = Get-DeviceCategories
    if ($cats.Count -eq 0) { Show-InfoBox 'No device categories are defined in this tenant.'; return }

    # the menu is not modal and this function returns before anything is
    # clicked, so the target has to live at script scope
    $script:PendingCategoryDeviceId = $id
    $script:PendingCategoryDevice   = $d

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    foreach ($c in $cats) {
        $item = $menu.Items.Add($c.displayName)
        $item.Tag = $c
        $item.Add_Click({
            param($sndr, $ev)
            $cat    = $sndr.Tag
            $devId  = $script:PendingCategoryDeviceId
            $dev    = $script:PendingCategoryDevice
            if (-not $devId -or -not $cat) { return }
            if (-not (Confirm-Action -Message "Set the category on $($dev.deviceName) to '$($cat.displayName)'?" -Title 'Set Device Category')) { return }
            Start-Busy 'Setting device category...'
            try {
                Set-DeviceCategoryById -ManagedDeviceId $devId -Category $cat
                Stop-Busy "Category set to $($cat.displayName)"
                Load-Devices
            } catch {
                Stop-Busy 'Failed'
                Show-ErrorBox "Could not set the category:`r`n$($_.Exception.Message)"
            }
        })
    }
    $menu.Show($script:BtnDevManage, 0, $script:BtnDevManage.Height)
}

$script:DeviceCategoryCache = $null

function Get-DeviceCategories {
    if ($script:DeviceCategoryCache) { return $script:DeviceCategoryCache }
    try {
        $script:DeviceCategoryCache = @(Invoke-Graph -Uri '/deviceManagement/deviceCategories' -All | Sort-Object displayName)
    } catch {
        $script:DeviceCategoryCache = @()
    }
    return $script:DeviceCategoryCache
}

function Set-DeviceCategoryById {
    # Prefer the reference endpoint; fall back to patching the display name,
    # because tenants differ on which one is accepted.
    param([string]$ManagedDeviceId, $Category)
    try {
        $body = @{ '@odata.id' = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$($Category.id)" }
        Invoke-Graph -Uri "/deviceManagement/managedDevices/$ManagedDeviceId/deviceCategory/`$ref" -Method PUT -Body $body -Beta -Raw | Out-Null
    } catch {
        Invoke-Graph -Uri "/deviceManagement/managedDevices/$ManagedDeviceId" -Method PATCH `
            -Body @{ deviceCategoryDisplayName = $Category.displayName } -Beta -Raw | Out-Null
    }
}

function Edit-DeviceNotes {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id

    $current = ''
    try {
        $full = Invoke-Graph -Uri "/deviceManagement/managedDevices/$id`?`$select=id,notes" -Beta -Raw
        $current = [string]$full.notes
    } catch { }

    $f = New-Object System.Windows.Forms.Form
    $f.Text            = "Notes - $($d.deviceName)"
    $f.Size            = New-Object System.Drawing.Size(620, 380)
    $f.StartPosition   = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'

    $f.Controls.Add((New-Label -Text "Notes for $($d.deviceName) (serial $($d.serialNumber))" -X 12 -Y 10 -W 580 -H 22))

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline  = $true
    $tb.ScrollBars = 'Vertical'
    $tb.Location   = New-Object System.Drawing.Point(12, 38)
    $tb.Size       = New-Object System.Drawing.Size(580, 240)
    $tb.Text       = $current
    $f.Controls.Add($tb)

    $btnOk = New-Button -Text 'Save' -X 380 -Y 292 -W 100 -H 30
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($btnOk)

    $btnNo = New-Button -Text 'Cancel' -X 490 -Y 292 -W 100 -H 30
    $btnNo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($btnNo)
    $f.CancelButton = $btnNo

    Set-ControlTheme -Control $f
    $r = $f.ShowDialog()
    $text = $tb.Text
    $f.Dispose()
    if ($r -ne [System.Windows.Forms.DialogResult]::OK) { return }
    if ($text -eq $current) { return }

    Start-Busy 'Saving notes...'
    try {
        Invoke-Graph -Uri "/deviceManagement/managedDevices/$id" -Method PATCH -Body @{ notes = $text } -Beta -Raw | Out-Null
        Stop-Busy 'Notes saved'
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox "Could not save notes:`r`n$($_.Exception.Message)"
    }
}

function Edit-DeviceGroupMembership {
    param([bool]$Remove)

    if (-not (Assert-Connected)) { return }
    $ids = Get-SelectedDeviceIds
    if ($ids.Count -eq 0) { Show-InfoBox 'Select one or more devices first.'; return }

    # resolve the Entra objects first, so the picker can show existing membership
    Start-Busy 'Resolving devices in Entra...'
    $resolved = New-Object System.Collections.Generic.List[object]
    $unresolved = New-Object System.Collections.Generic.List[string]
    foreach ($id in $ids) {
        $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
        if (-not $d) { continue }
        $objId = Get-EntraDeviceObjectId -AzureAdDeviceId $d.azureADDeviceId
        if ($objId) {
            [void]$resolved.Add([pscustomobject]@{
                Name   = $d.deviceName
                Serial = $d.serialNumber
                ObjId  = $objId
                Uri    = "/devices/$objId"
            })
        } else {
            [void]$unresolved.Add("$($d.deviceName) (no matching Entra device object)")
        }
    }
    Stop-Busy 'Ready'

    if ($resolved.Count -eq 0) {
        Show-ErrorBox "None of the selected devices resolved to an Entra device object.`r`n`r`n$($unresolved -join "`r`n")"
        return
    }

    $members = foreach ($r in $resolved) { @{ Name = $r.Name; Uri = $r.Uri } }
    $picked = Show-GroupPicker -Prompt 'Group name starts with' -Members @($members)
    if (-not $picked) { return }
    if (-not (Test-GroupIsAssigned -Picked $picked)) { return }

    # split into work and no-ops so the confirmation reflects what will happen
    $todo = New-Object System.Collections.Generic.List[object]
    $skip = New-Object System.Collections.Generic.List[string]
    foreach ($r in $resolved) {
        $isMember = ((Get-DirectMemberOfIds -MemberUri $r.Uri) -contains $picked.Id)
        if ($Remove -and -not $isMember) {
            [void]$skip.Add("  $($r.Name) - not a member")
        } elseif ((-not $Remove) -and $isMember) {
            [void]$skip.Add("  $($r.Name) - already a member")
        } else {
            [void]$todo.Add($r)
        }
    }

    $verb = if ($Remove) { 'Remove' } else { 'Add' }
    $prep = if ($Remove) { 'from' } else { 'to' }

    if ($todo.Count -eq 0) {
        Show-InfoBox "Nothing to do for '$($picked.Name)'.`r`n`r`n$($skip -join "`r`n")"
        return
    }

    $targets = foreach ($r in $todo) { "  $($r.Name)  |  serial $($r.Serial)" }
    $msg = "$verb $($todo.Count) device(s) $prep group '$($picked.Name)'?`r`n`r`n$($targets -join "`r`n")"
    if ($skip.Count -gt 0)       { $msg += "`r`n`r`nSkipped:`r`n$($skip -join "`r`n")" }
    if ($unresolved.Count -gt 0) { $msg += "`r`n`r`nNot resolved:`r`n  $($unresolved -join "`r`n  ")" }
    if (-not (Confirm-Action -Message $msg -Title "$verb Group Member")) { return }

    $ok = 0; $fail = 0; $errors = New-Object System.Collections.Generic.List[string]
    Start-Busy "$verb group membership..."
    foreach ($r in $todo) {
        try {
            if ($Remove) {
                Invoke-Graph -Uri "/groups/$($picked.Id)/members/$($r.ObjId)/`$ref" -Method DELETE -Raw | Out-Null
            } else {
                $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($r.ObjId)" }
                Invoke-Graph -Uri "/groups/$($picked.Id)/members/`$ref" -Method POST -Body $body -Raw | Out-Null
            }
            $ok++
            $script:MemberOfCache.Remove($r.Uri) | Out-Null
        } catch {
            $fail++
            [void]$errors.Add("$($r.Name) : $($_.Exception.Message)")
        }
    }
    Stop-Busy "$verb complete - $ok succeeded, $fail failed"

    $result = "$verb group '$($picked.Name)'`r`n`r`nSucceeded: $ok`r`nFailed: $fail"
    if ($skip.Count -gt 0) { $result += "`r`nSkipped: $($skip.Count)" }
    if ($errors.Count -gt 0) {
        $result += "`r`n`r`n" + ($errors -join "`r`n")
        $result += Get-MembershipErrorHint -Message ($errors -join ' ')
    }
    Show-InfoBox $result
}

function Edit-UserGroupMembership {
    param([bool]$Remove)

    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id
    if ([string]::IsNullOrWhiteSpace($d.userPrincipalName)) {
        Show-InfoBox "$($d.deviceName) has no primary user."
        return
    }

    Start-Busy 'Resolving user in Entra...'
    try {
        $u = @(Invoke-Graph -Uri "/users?`$filter=userPrincipalName eq '$(Escape-ODataValue $d.userPrincipalName)'&`$select=id" -All)
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox $_.Exception.Message
        return
    }
    Stop-Busy 'Ready'
    if ($u.Count -eq 0) { Show-ErrorBox "No user found with UPN $($d.userPrincipalName)."; return }
    $userUri = "/users/$($u[0].id)"

    $picked = Show-GroupPicker -Prompt 'Group name starts with' `
                -Members @(@{ Name = $d.userPrincipalName; Uri = $userUri })
    if (-not $picked) { return }
    if (-not (Test-GroupIsAssigned -Picked $picked)) { return }

    $isMember = ((Get-DirectMemberOfIds -MemberUri $userUri) -contains $picked.Id)
    if ($Remove -and -not $isMember) {
        Show-InfoBox "$($d.userPrincipalName) is not a direct member of '$($picked.Name)'."
        return
    }
    if ((-not $Remove) -and $isMember) {
        Show-InfoBox "$($d.userPrincipalName) is already a member of '$($picked.Name)'."
        return
    }

    $verb = if ($Remove) { 'Remove' } else { 'Add' }
    $prep = if ($Remove) { 'from' } else { 'to' }
    if (-not (Confirm-Action -Message "$verb user $($d.userPrincipalName) $prep group '$($picked.Name)'?" -Title "$verb Group Member")) { return }

    Start-Busy "$verb user group membership..."
    try {
        if ($Remove) {
            Invoke-Graph -Uri "/groups/$($picked.Id)/members/$($u[0].id)/`$ref" -Method DELETE -Raw | Out-Null
        } else {
            $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($u[0].id)" }
            Invoke-Graph -Uri "/groups/$($picked.Id)/members/`$ref" -Method POST -Body $body -Raw | Out-Null
        }
        $script:MemberOfCache.Remove($userUri) | Out-Null
        Stop-Busy "$verb complete"
        Show-InfoBox "$($d.userPrincipalName) $(if ($Remove) { 'removed from' } else { 'added to' }) '$($picked.Name)'."
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox ("Could not change group membership:`r`n$($_.Exception.Message)" +
                       (Get-MembershipErrorHint -Message $_.Exception.Message))
    }
}

# ---------------------------------------------------------------------------
# What is assigned to this device
# ---------------------------------------------------------------------------
function Show-DeviceAssignments {
    if (-not (Assert-Connected)) { return }
    $id = Get-SelectedDeviceId
    if (-not $id) { Show-InfoBox 'Select a device first.'; return }
    $d = Get-CachedObject -CacheKey 'DeviceRaw' -Id $id

    Start-Busy "Resolving group membership for $($d.deviceName)..."
    try {
        $deviceGroups = @{}
        $userGroups   = @{}

        $objId = Get-EntraDeviceObjectId -AzureAdDeviceId $d.azureADDeviceId
        if ($objId) {
            $memberOf = @(Invoke-Graph -Uri "/devices/$objId/transitiveMemberOf?`$select=id,displayName&`$top=200" -All)
            foreach ($g in $memberOf) { if ($g.id) { $deviceGroups[$g.id] = $g.displayName } }
        }

        if ($d.userPrincipalName) {
            $u = @(Invoke-Graph -Uri "/users?`$filter=userPrincipalName eq '$(Escape-ODataValue $d.userPrincipalName)'&`$select=id" -All)
            if ($u.Count -gt 0) {
                $umem = @(Invoke-Graph -Uri "/users/$($u[0].id)/transitiveMemberOf?`$select=id,displayName&`$top=200" -All)
                foreach ($g in $umem) { if ($g.id) { $userGroups[$g.id] = $g.displayName } }
            }
        }

        $sources = @(
            @{ Label = 'App';                  Uri = "/deviceAppManagement/mobileApps?`$expand=assignments";                  Beta = $false; NameProp = 'displayName' },
            @{ Label = 'Settings Catalog';     Uri = "/deviceManagement/configurationPolicies?`$expand=assignments";          Beta = $true;  NameProp = 'name' },
            @{ Label = 'Device Configuration'; Uri = "/deviceManagement/deviceConfigurations?`$expand=assignments";           Beta = $false; NameProp = 'displayName' },
            @{ Label = 'Compliance Policy';    Uri = "/deviceManagement/deviceCompliancePolicies?`$expand=assignments";       Beta = $false; NameProp = 'displayName' },
            @{ Label = 'Platform Script';      Uri = "/deviceManagement/deviceManagementScripts?`$expand=assignments";        Beta = $true;  NameProp = 'displayName' },
            @{ Label = 'Remediation';          Uri = "/deviceManagement/deviceHealthScripts?`$expand=assignments";            Beta = $true;  NameProp = 'displayName' }
        )

        $results = New-Object System.Collections.Generic.List[object]

        foreach ($src in $sources) {
            Set-Status "Checking $($src.Label)..."
            try {
                $items = @(Invoke-Graph -Uri $src.Uri -Beta:$src.Beta -All)
            } catch {
                [void]$results.Add([pscustomobject]@{
                    ObjectType = $src.Label; Name = "(could not read: $($_.Exception.Message))"
                    Applies = ''; Via = ''; MatchedOn = ''; Intent = ''
                })
                continue
            }

            foreach ($item in $items) {
                $includes = New-Object System.Collections.Generic.List[string]
                $excludes = New-Object System.Collections.Generic.List[string]
                $intent   = ''

                foreach ($a in @($item.assignments)) {
                    $t = $a.target
                    if ($a.intent) { $intent = $a.intent }
                    $gid = $t.groupId

                    switch -Wildcard ($t.'@odata.type') {
                        '*allDevicesAssignmentTarget' {
                            [void]$includes.Add('All Devices|All Devices')
                        }
                        '*allLicensedUsersAssignmentTarget' {
                            if ($d.userPrincipalName) { [void]$includes.Add('All Users|All Users') }
                        }
                        '*exclusionGroupAssignmentTarget' {
                            if ($gid -and ($deviceGroups.ContainsKey($gid) -or $userGroups.ContainsKey($gid))) {
                                $n = if ($deviceGroups.ContainsKey($gid)) { $deviceGroups[$gid] } else { $userGroups[$gid] }
                                [void]$excludes.Add($n)
                            }
                        }
                        default {
                            if ($gid -and $deviceGroups.ContainsKey($gid)) { [void]$includes.Add("Device group|$($deviceGroups[$gid])") }
                            elseif ($gid -and $userGroups.ContainsKey($gid)) { [void]$includes.Add("User group|$($userGroups[$gid])") }
                        }
                    }
                }

                if ($includes.Count -eq 0) { continue }

                $viaList  = foreach ($inc in $includes) { ($inc -split '\|')[0] }
                $nameList = foreach ($inc in $includes) { ($inc -split '\|')[1] }

                [void]$results.Add([pscustomobject]@{
                    ObjectType = $src.Label
                    Name       = $item.($src.NameProp)
                    Applies    = $(if ($excludes.Count -gt 0) { 'EXCLUDED' } else { 'Yes' })
                    Via        = (($viaList | Select-Object -Unique) -join ', ')
                    MatchedOn  = (($nameList | Select-Object -Unique) -join '; ')
                    Intent     = $intent
                    ExcludedBy = (($excludes | Select-Object -Unique) -join '; ')
                })
            }
        }

        Stop-Busy "$($results.Count) object(s) target $($d.deviceName)"

        if ($results.Count -eq 0) {
            Show-InfoBox "Nothing targets $($d.deviceName) through its $($deviceGroups.Count) device group(s) and $($userGroups.Count) user group(s)."
            return
        }

        Show-GridWindow -Title "Assigned to $($d.deviceName)  -  $($deviceGroups.Count) device group(s), $($userGroups.Count) user group(s)" `
            -Objects @($results.ToArray() | Sort-Object Applies, ObjectType, Name) `
            -Columns @('ObjectType','Name','Applies','Via','MatchedOn','Intent','ExcludedBy') `
            -ExportName "AssignedTo_$($d.deviceName)"
    } catch {
        Stop-Busy 'Failed'
        Show-ErrorBox "Could not resolve assignments:`r`n$($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Policy import from JSON
# ---------------------------------------------------------------------------
function Import-PolicyFromJson {
    if (-not (Assert-Connected)) { return }

    $od = New-Object System.Windows.Forms.OpenFileDialog
    $od.Filter = 'JSON policy export (*.json)|*.json|All files (*.*)|*.*'
    $od.Title  = 'Select a policy JSON file'
    if ($od.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $json = Get-Content -LiteralPath $od.FileName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Show-ErrorBox "Could not read that file as JSON:`r`n$($_.Exception.Message)"
        return
    }

    # work out what kind of policy this is from its own shape
    $kind = $null
    if ($json.PSObject.Properties.Name -contains 'technologies' -or
        $json.PSObject.Properties.Name -contains 'settings') {
        $kind = 'SettingsCatalog'
    } elseif ($json.'@odata.type' -match 'CompliancePolicy') {
        $kind = 'Compliance'
    } elseif ($json.'@odata.type' -match 'Configuration') {
        $kind = 'DeviceConfig'
    }

    if (-not $kind) {
        Show-ErrorBox "Could not tell what kind of policy this is.`r`n`r`nSupported: Settings Catalog exports, and legacy device configuration or compliance policies that carry an @odata.type."
        return
    }

    $nameProp = if ($kind -eq 'SettingsCatalog') { 'name' } else { 'displayName' }
    $original = [string]$json.$nameProp
    $newName  = Show-InputDialog -Title 'Import Policy' `
                 -Prompt 'Name for the imported policy:' `
                 -Default $(if ($original) { "$original - imported" } else { 'Imported policy' })
    if ([string]::IsNullOrWhiteSpace($newName)) { return }

    # strip anything the service assigns itself
    $strip = @('id','createdDateTime','lastModifiedDateTime','version','@odata.context',
               'supportsScopeTags','assignments','assignments@odata.context','isAssigned',
               'settingCount','creationSource','priorityMetaData','deviceStatusOverview',
               'userStatusOverview','deviceStatuses','userStatuses','scheduledActionsForRule')

    $body = @{}
    foreach ($prop in $json.PSObject.Properties) {
        if ($strip -contains $prop.Name) { continue }
        if ($prop.Name -like '*@odata.count') { continue }
        $body[$prop.Name] = $prop.Value
    }
    $body[$nameProp] = $newName.Trim()

    $info = Get-AssignmentKindInfo -Kind $kind
    $target = $info.Base

    $detail = "Create a new $kind policy named:`r`n  $($newName.Trim())`r`n`r`nSource file:`r`n  $($od.FileName)`r`n`r`nAssignments are NOT imported - set them afterwards with Edit Assignments."
    if (-not (Confirm-Action -Message $detail -Title 'Import Policy')) { return }

    Start-Busy 'Importing policy...'
    try {
        $created = Invoke-Graph -Uri $target -Method POST -Body $body -Beta:$info.Beta -Raw
        Stop-Busy 'Policy imported'
        Show-InfoBox "Created '$($newName.Trim())'.`r`n`r`nId: $($created.id)`r`n`r`nIt has no assignments yet."
        Load-Policies
    } catch {
        Stop-Busy 'Import failed'
        Show-ErrorBox "Could not import the policy:`r`n$($_.Exception.Message)`r`n`r`nSettings Catalog exports import most reliably. Legacy profiles often carry properties the create call rejects."
    }
}

function Show-DeviceManageMenu {
    $menu = New-Object System.Windows.Forms.ContextMenuStrip

    $i1 = $menu.Items.Add('Set primary user...')
    $i1.Add_Click({ Set-DevicePrimaryUser })

    $i2 = $menu.Items.Add('Rename device...')
    $i2.Add_Click({ Rename-ManagedDevice })

    $i3 = $menu.Items.Add('Set device category...')
    $i3.Add_Click({ Set-ManagedDeviceCategory })

    $i4 = $menu.Items.Add('Edit notes...')
    $i4.Add_Click({ Edit-DeviceNotes })

    [void]$menu.Items.Add('-')

    $i5 = $menu.Items.Add('Add device to group...')
    $i5.Add_Click({ Edit-DeviceGroupMembership -Remove $false })

    $i6 = $menu.Items.Add('Remove device from group...')
    $i6.Add_Click({ Edit-DeviceGroupMembership -Remove $true })

    $i7 = $menu.Items.Add('Add primary user to group...')
    $i7.Add_Click({ Edit-UserGroupMembership -Remove $false })

    $i8 = $menu.Items.Add('Remove primary user from group...')
    $i8.Add_Click({ Edit-UserGroupMembership -Remove $true })

    [void]$menu.Items.Add('-')

    $i9 = $menu.Items.Add("What's assigned to this device...")
    $i9.Add_Click({ Show-DeviceAssignments })

    $menu.Show($script:BtnDevManage, 0, $script:BtnDevManage.Height)
}

# ===========================================================================
# Startup
# ===========================================================================
if ($ThemeMode -eq 'Light' -or $ThemeMode -eq 'Dark') {
    # Launched from MSToolkit: follow its theme. Nothing is written to config.json
    # unless the theme button is used or settings are saved.
    Set-AppTheme $ThemeMode
}
else {
    Set-AppTheme $script:Config.Theme
}

Set-ConnectedState $false
Refresh-PackagingFolders

# Try a silent reconnect if a refresh token is already cached
if ($script:Config.TenantId -and $script:Config.ClientId -and (Test-Path $script:TokenPath)) {
    if (Invoke-TokenRefresh) { Set-ConnectedState $true; Set-Status "Connected as $($script:Account)" }
}

# Open on Settings when anything required is missing, or on a genuine first run
# (no saved config yet). Once everything is set and saved, the tool opens on Devices.
$missingSettings = New-Object System.Collections.Generic.List[string]

if ([string]::IsNullOrWhiteSpace($script:Config.TenantId)) { [void]$missingSettings.Add('Tenant ID') }
if ([string]::IsNullOrWhiteSpace($script:Config.ClientId)) { [void]$missingSettings.Add('Client ID') }
if ([string]::IsNullOrWhiteSpace($script:Config.IntuneWinAppUtil) -or
    -not (Test-Path $script:Config.IntuneWinAppUtil -ErrorAction SilentlyContinue)) {
    [void]$missingSettings.Add('IntuneWinAppUtil.exe path')
}
if ([string]::IsNullOrWhiteSpace($script:Config.PackagingRoot) -or
    -not (Test-Path $script:Config.PackagingRoot -ErrorAction SilentlyContinue)) {
    [void]$missingSettings.Add('Packaging root folder')
}

$firstRun = -not (Test-Path $script:ConfigPath)

if ($missingSettings.Count -gt 0 -or $firstRun) {
    $script:Tabs.SelectedTab = $tabSettings
    if ($missingSettings.Count -gt 0) {
        Set-Status ('Settings needed: ' + ($missingSettings.ToArray() -join ', ') + ' - fill these in, then Save Settings')
    } else {
        Set-Status 'First run - review the settings below, then click Save Settings'
    }
} else {
    $script:Tabs.SelectedTab = $tabPackaging
}

function Set-SplitterSafe {
    param([System.Windows.Forms.SplitContainer]$Split, [int]$Distance)
    try {
        $max = if ($Split.Orientation -eq 'Horizontal') { $Split.Height } else { $Split.Width }
        $d = [Math]::Max($Split.Panel1MinSize + 1, [Math]::Min($Distance, $max - $Split.Panel2MinSize - 1))
        if ($d -gt 0) { $Split.SplitterDistance = $d }
    } catch { }
}

$script:MainForm.Add_Resize({
    if ($script:MainForm.WindowState -ne [System.Windows.Forms.FormWindowState]::Minimized) {
        $script:LastWindowState = $script:MainForm.WindowState
    }
})

$script:MainForm.Add_FormClosed({
    # a WinForms timer outlives the script if it is left running, and keeps
    # firing inside the host session - always tear ours down on exit
    if ($script:LapsCopyTimer) {
        try { $script:LapsCopyTimer.Stop(); $script:LapsCopyTimer.Dispose() } catch { }
        $script:LapsCopyTimer = $null
    }
    if ($script:ThemeTip) {
        try { $script:ThemeTip.Dispose() } catch { }
        $script:ThemeTip = $null
    }
})

$script:MainForm.Add_Shown({
    Set-SplitterSafe -Split $devSplit  -Distance 380
    Set-SplitterSafe -Split $appSplit  -Distance 380
    Set-SplitterSafe -Split $polSplit  -Distance 400
    Set-SplitterSafe -Split $grpSplit  -Distance 260
    Set-SplitterSafe -Split $pkgSplit  -Distance 300
    Set-SplitterSafe -Split $pkgLists  -Distance 400
    $script:MainForm.Activate()
})

[void]$script:MainForm.ShowDialog()
$script:MainForm.Dispose()
