<#
.SYNOPSIS
    Teams Call Blocking - tenant-wide inbound PSTN number block (MSToolkit child-style GUI).

.DESCRIPTION
    Prompts for a phone number and adds a tenant-wide inbound blocked number pattern in
    Microsoft Teams with New-CsInboundBlockedNumberPattern.

    Flow:
      1. Connect-MicrosoftTeams (interactive sign-in).
      2. Tenant check: the connected tenant must list -ExpectedTenantDomain as a verified
         domain (and match -TenantId if supplied). Blocking stays disabled if it doesn't.
      3. Read-only checks: tenant blocking on/off, existing blocked and exempt patterns,
         rule-name collision, Test-CsInboundBlockedNumberPattern.
      4. Confirmation dialog naming the exact number, pattern, rule name and tenant.
      5. New-CsInboundBlockedNumberPattern, then re-read and re-test to verify.

    Unblock: the pattern list shows each blocked rule and the number it covers. Select one
    or more Blocked rows (Ctrl/Shift-click) and click Unblock Selected. After a confirmation
    naming each rule, it runs Remove-CsInboundBlockedNumberPattern for those rules only,
    then re-reads and re-tests. Exempt patterns are listed but never changed.

    Module install: if MicrosoftTeams is not found at launch, the tool installs it for the
    current user (CurrentUser scope, PSGallery) in separate powershell.exe processes and
    streams its output to Activity Output. Install Module retries it on demand.

    Never calls Set-CsTenantBlockedCallingNumbers -InboundBlockedNumberPatterns (that
    replaces the whole pattern list). The only tenant-level change it can make is
    Set-CsTenantBlockedCallingNumbers -Enabled $true, and only after a separate prompt.

    Stand-alone: run as the Windows user signed in to this PC (use the launcher .cmd).
    Later MSToolkit attachment: launch it the same way as the other M365-* child tools and
    pass -ThemeMode.

.PARAMETER ThemeMode
    Light or Dark. If omitted, reads Theme from %APPDATA%\MSToolkit\settings.json,
    otherwise Light. When MSToolkit passes -ThemeMode it always wins at launch. The header
    toggle switches live and saves Theme back to settings.json (other settings kept).

.PARAMETER ExpectedTenantDomain
    A verified domain that must exist in the connected tenant before blocking is allowed.
    MSToolkit passes its "Expected tenant domain" setting. If omitted, it is read from
    %APPDATA%\MSToolkit\settings.json for the account running this tool. Blank keeps
    blocking disabled.

.PARAMETER TenantId
    Optional. Passed to Connect-MicrosoftTeams -TenantId and also required to match the
    connected tenant. MSToolkit passes its "Tenant ID" setting; if omitted, it is read
    from %APPDATA%\MSToolkit\settings.json the same way. Blank skips this check.

.PARAMETER ShowConsole
    Keep the PowerShell console visible (troubleshooting, or when run from an open console).
#>
[CmdletBinding()]
param(
    [ValidateSet('Light','Dark')]
    [string]$ThemeMode,

    [string]$ExpectedTenantDomain = '',

    [string]$TenantId,

    [switch]$ShowConsole
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region ---------- Tenant settings from MSToolkit ----------
# Nothing organization-specific is stored in this script. When MSToolkit launches it,
# -ExpectedTenantDomain and -TenantId come from MSToolkit Settings. Run on its own, the
# same values are read from %APPDATA%\MSToolkit\settings.json for this account.
$script:TenantSettingSource = 'not set'
if ($PSBoundParameters.ContainsKey('ExpectedTenantDomain') -or $PSBoundParameters.ContainsKey('TenantId')) {
    $script:TenantSettingSource = 'MSToolkit (passed on launch)'
} else {
    try {
        $mstkSettings = Join-Path (Join-Path $env:APPDATA 'MSToolkit') 'settings.json'
        if (Test-Path -LiteralPath $mstkSettings) {
            $mstk = Get-Content -LiteralPath $mstkSettings -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($mstk.PSObject.Properties['ExpectedTenantDomain']) { $ExpectedTenantDomain = ([string]$mstk.ExpectedTenantDomain).Trim() }
            if ($mstk.PSObject.Properties['TenantId'])             { $TenantId             = ([string]$mstk.TenantId).Trim() }
            if ($ExpectedTenantDomain -or $TenantId) { $script:TenantSettingSource = $mstkSettings }
        }
    } catch {
        # Unreadable settings: the tenant check below fails closed.
    }
}
if ($null -eq $ExpectedTenantDomain) { $ExpectedTenantDomain = '' }
#endregion

#region ---------- Script state ----------
$script:ToolTitle        = 'Teams Call Blocking'
$script:IsConnected      = $false
$script:TenantVerified   = $false
$script:IsBusy           = $false
$script:BlockingEnabled  = $null
$script:BlockedPatterns  = @()
$script:ExemptPatterns   = @()
$script:ConnectedAccount = ''
$script:TenantInfo       = $null
$script:CurrentTarget    = $null
$script:ModuleAvailable  = $null
$script:ConsoleScreen    = ''
$script:ConsoleShownForSignIn = $false
$script:LogEntries       = New-Object System.Collections.Generic.List[object]

if ($PSScriptRoot) { $script:LogDir = Join-Path $PSScriptRoot 'Logs' } else { $script:LogDir = Join-Path $env:TEMP 'MSToolkit-Teams-Block-Number' }
try {
    if (-not (Test-Path -LiteralPath $script:LogDir)) { New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null }
} catch {
    $script:LogDir = Join-Path $env:TEMP 'MSToolkit-Teams-Block-Number'
    if (-not (Test-Path -LiteralPath $script:LogDir)) { New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null }
}
$script:LogFile = Join-Path $script:LogDir ('M365-Teams-Block-Number_{0}.log' -f (Get-Date -Format 'yyyyMMdd'))
#endregion

#region ---------- Console hide ----------
if (-not ('MSToolkitTeamsBlock.NativeMethods' -as [type])) {
    Add-Type -Namespace MSToolkitTeamsBlock -Name NativeMethods -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@
}

function Hide-PowerShellConsole {
    try {
        $h = [MSToolkitTeamsBlock.NativeMethods]::GetConsoleWindow()
        if ($h -ne [IntPtr]::Zero) { [void][MSToolkitTeamsBlock.NativeMethods]::ShowWindow($h, 0) }
    } catch { }
}
#endregion

#region ---------- Sign-in window placement ----------
# Keeps the Microsoft sign-in window on the same monitor as this tool (same approach as
# MSToolkit: EnumWindows / GetWindowRect / SetWindowPos). A background thread watches for NEW
# top-level windows from this process or from the sign-in host processes while
# Connect-MicrosoftTeams runs, and centers any that open on a different monitor over the tool.
if (-not ('MSToolkitTeamsBlock.SignInWindowMover' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace MSToolkitTeamsBlock
{
    public static class CursorProbe
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X; public int Y; }
        [StructLayout(LayoutKind.Sequential)]
        private struct CURSORINFO { public int cbSize; public int flags; public IntPtr hCursor; public POINT pt; }
        [DllImport("user32.dll")] private static extern bool GetCursorInfo(ref CURSORINFO ci);
        [DllImport("user32.dll")] private static extern IntPtr LoadCursor(IntPtr hInstance, IntPtr id);

        // Name of the cursor Windows is showing right now (system-wide).
        public static string CurrentCursorName()
        {
            CURSORINFO ci = new CURSORINFO();
            ci.cbSize = Marshal.SizeOf(typeof(CURSORINFO));
            if (!GetCursorInfo(ref ci)) { return "unknown"; }
            if (ci.hCursor == LoadCursor(IntPtr.Zero, new IntPtr(32650))) { return "AppStarting (arrow + spinning circle)"; }
            if (ci.hCursor == LoadCursor(IntPtr.Zero, new IntPtr(32514))) { return "Wait (spinning circle)"; }
            if (ci.hCursor == LoadCursor(IntPtr.Zero, new IntPtr(32512))) { return "Arrow"; }
            if (ci.hCursor == LoadCursor(IntPtr.Zero, new IntPtr(32513))) { return "IBeam"; }
            if (ci.hCursor == LoadCursor(IntPtr.Zero, new IntPtr(32649))) { return "Hand"; }
            return "other";
        }
    }

    public static class SignInWindowMover
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags; }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr hWnd, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool GetMonitorInfo(IntPtr hMon, ref MONITORINFO mi);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);

        // Top-level owner of a window (GA_ROOTOWNER). For a Windows Terminal tab this is the
        // terminal window rather than the hidden pseudo-console window.
        public static IntPtr GetRootOwner(IntPtr hWnd)
        {
            return GetAncestor(hWnd, 3);
        }

        // Shows hWnd (without activating it) directly BEHIND target in z-order, centered on
        // target and resized to w x h, so it is fully covered by the tool.
        public static bool ShowBehind(IntPtr hWnd, IntPtr target, int w, int h)
        {
            RECT tr;
            if (!GetWindowRect(target, out tr)) { return false; }
            int x = ((tr.Left + tr.Right) / 2) - (w / 2);
            int y = ((tr.Top + tr.Bottom) / 2) - (h / 2);
            const uint SWP_SHOWWINDOW = 0x0040;
            return SetWindowPos(hWnd, target, x, y, w, h, SWP_NOACTIVATE | SWP_SHOWWINDOW);
        }

        public static string ClassOfWindow(IntPtr hWnd)
        {
            return ClassOf(hWnd);
        }

        // Same placement MSToolkit uses (WindowHelper.CenterWindowInWorkingArea), without the
        // "already on that monitor" skip, because hidden windows must always follow the tool.
        public static bool CenterInArea(IntPtr hWnd, int areaLeft, int areaTop, int areaWidth, int areaHeight)
        {
            RECT rect;
            if (!GetWindowRect(hWnd, out rect)) { return false; }
            int w = Math.Max(1, rect.Right - rect.Left);
            int h = Math.Max(1, rect.Bottom - rect.Top);
            int x = areaLeft + Math.Max(0, (areaWidth - w) / 2);
            int y = areaTop + Math.Max(0, (areaHeight - h) / 2);
            return SetWindowPos(hWnd, IntPtr.Zero, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
        [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr ctx);

        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint MONITOR_DEFAULTTONEAREST = 2;

        public static string[] WatchProcessNames = new string[] {
            "Microsoft.AAD.BrokerPlugin", "ApplicationFrameHost", "CredentialUIBroker",
            "msedgewebview2", "msedge", "chrome", "firefox" };

        // Title fallback for sign-in hosts that aren't in WatchProcessNames
        public static string TitlePattern = "sign in|signed in|sign-in|pick an account|work or school|authenticat|log ?in|windows security|verify your identity";

        private static Thread _thread;
        private static volatile bool _stop;
        private static readonly object _lock = new object();
        private static readonly List<string> _moved = new List<string>();
        private static readonly Dictionary<IntPtr, string> _seen = new Dictionary<IntPtr, string>();
        private static readonly Dictionary<IntPtr, string> _status = new Dictionary<IntPtr, string>();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);

        private static string TitleOf(IntPtr h)
        {
            StringBuilder sb = new StringBuilder(256);
            GetWindowText(h, sb, sb.Capacity);
            return sb.ToString();
        }

        private static string ClassOf(IntPtr h)
        {
            StringBuilder sb = new StringBuilder(256);
            GetClassName(h, sb, sb.Capacity);
            return sb.ToString();
        }

        private static List<IntPtr> VisibleTopWindows()
        {
            List<IntPtr> list = new List<IntPtr>();
            EnumWindowsProc cb = delegate(IntPtr h, IntPtr l) { if (IsWindowVisible(h)) { list.Add(h); } return true; };
            EnumWindows(cb, IntPtr.Zero);
            GC.KeepAlive(cb);
            return list;
        }

        private static RECT WorkArea(IntPtr hwnd)
        {
            MONITORINFO mi = new MONITORINFO();
            mi.cbSize = Marshal.SizeOf(typeof(MONITORINFO));
            GetMonitorInfo(MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST), ref mi);
            return mi.rcWork;
        }

        public static bool SameMonitor(IntPtr a, IntPtr b)
        {
            return MonitorFromWindow(a, MONITOR_DEFAULTTONEAREST) == MonitorFromWindow(b, MONITOR_DEFAULTTONEAREST);
        }

        // Centers hwnd over target, clamped to target's monitor work area. Size is not changed.
        public static bool CenterOnWindow(IntPtr hwnd, IntPtr target)
        {
            if (hwnd == IntPtr.Zero || target == IntPtr.Zero) { return false; }
            RECT tr; RECT wr;
            if (!GetWindowRect(target, out tr) || !GetWindowRect(hwnd, out wr)) { return false; }
            RECT work = WorkArea(target);
            int w = wr.Right - wr.Left;
            int h = wr.Bottom - wr.Top;
            int x = ((tr.Left + tr.Right) / 2) - (w / 2);
            int y = ((tr.Top + tr.Bottom) / 2) - (h / 2);
            if (x + w > work.Right) { x = work.Right - w; }
            if (y + h > work.Bottom) { y = work.Bottom - h; }
            if (x < work.Left) { x = work.Left; }
            if (y < work.Top) { y = work.Top; }
            return SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        public static void Start(IntPtr target, int seconds)
        {
            Stop();
            lock (_lock) { _moved.Clear(); _seen.Clear(); _status.Clear(); }
            _stop = false;

            HashSet<IntPtr> known = new HashSet<IntPtr>(VisibleTopWindows());
            int myPid = Process.GetCurrentProcess().Id;
            HashSet<string> names = new HashSet<string>(WatchProcessNames, StringComparer.OrdinalIgnoreCase);
            Regex titleRx = new Regex(TitlePattern, RegexOptions.IgnoreCase);

            _thread = new Thread(delegate()
            {
                try { SetThreadDpiAwarenessContext(new IntPtr(-4)); } catch { }   // per-monitor v2, if supported
                Dictionary<IntPtr, int> moves = new Dictionary<IntPtr, int>();
                HashSet<IntPtr> raised = new HashSet<IntPtr>();
                Dictionary<uint, string> pidNames = new Dictionary<uint, string>();
                Stopwatch sw = Stopwatch.StartNew();
                while (!_stop && sw.Elapsed.TotalSeconds < seconds)
                {
                    try
                    {
                        foreach (IntPtr h in VisibleTopWindows())
                        {
                            if (h == target || known.Contains(h)) { continue; }
                            uint pid;
                            GetWindowThreadProcessId(h, out pid);
                            string pname;
                            if (!pidNames.TryGetValue(pid, out pname))
                            {
                                try { pname = Process.GetProcessById((int)pid).ProcessName; } catch { pname = "?"; }
                                pidNames[pid] = pname;
                            }
                            string title = TitleOf(h);
                            lock (_lock) { _seen[h] = pname + " | " + ClassOf(h) + " | '" + title + "'"; }

                            // Titles can change as the sign-in page loads, so re-check every pass.
                            bool candidate = (pid == (uint)myPid) || names.Contains(pname) || titleRx.IsMatch(title);
                            if (!candidate) { SetStatus(h, "not a sign-in window"); continue; }

                            RECT r;
                            if (!GetWindowRect(h, out r)) { continue; }
                            if ((r.Right - r.Left) < 150 || (r.Bottom - r.Top) < 100) { SetStatus(h, "too small / not drawn yet"); continue; }
                            if (SameMonitor(h, target))
                            {
                                if (!moves.ContainsKey(h)) { SetStatus(h, "already on this screen"); }
                            }
                            else
                            {
                                // Re-apply if the sign-in host re-centers itself after loading (max 10 moves per window).
                                int n;
                                moves.TryGetValue(h, out n);
                                if (n >= 10) { SetStatus(h, "moved 10 times, window kept moving back"); continue; }
                                if (CenterOnWindow(h, target))
                                {
                                    moves[h] = n + 1;
                                    SetStatus(h, "MOVED to this screen (" + (n + 1) + "x)");
                                    if (n == 0) { lock (_lock) { _moved.Add(pname + ": " + title); } }
                                }
                                else
                                {
                                    SetStatus(h, "move FAILED (SetWindowPos error " + Marshal.GetLastWin32Error() + ")");
                                    continue;
                                }
                            }

                            // Now that it shares a screen with the tool it can open BEHIND the tool,
                            // so bring it in front once.
                            if (!raised.Contains(h))
                            {
                                RaiseWindow(h);
                                raised.Add(h);
                                string st;
                                lock (_lock) { if (!_status.TryGetValue(h, out st)) { st = ""; } }
                                SetStatus(h, st + " + brought to front");
                            }
                        }
                    }
                    catch { }
                    Thread.Sleep(200);
                }
            });
            _thread.IsBackground = true;
            _thread.Start();
        }

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);

        // Puts a window above all normal windows (topmost on, then off) and activates it.
        private static void RaiseWindow(IntPtr h)
        {
            const uint SWP_NOMOVE = 0x0002;
            SetWindowPos(h, new IntPtr(-1), 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);   // HWND_TOPMOST
            SetWindowPos(h, new IntPtr(-2), 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);   // HWND_NOTOPMOST
            SetForegroundWindow(h);
        }

        private static void SetStatus(IntPtr h, string status)
        {
            lock (_lock) { _status[h] = status; }
        }

        // One line per new window seen while the watch ran: process | class | 'title' -> result
        public static string[] GetDiagnostics()
        {
            lock (_lock)
            {
                List<string> lines = new List<string>();
                foreach (KeyValuePair<IntPtr, string> kv in _seen)
                {
                    string st;
                    if (!_status.TryGetValue(kv.Key, out st)) { st = "no action"; }
                    lines.Add(kv.Value + " -> " + st);
                }
                return lines.ToArray();
            }
        }

        public static void Stop()
        {
            _stop = true;
            Thread t = _thread;
            if (t != null && t.IsAlive) { t.Join(1000); }
            _thread = null;
        }

        public static string[] GetMovedWindows()
        {
            lock (_lock) { return _moved.ToArray(); }
        }
    }
}
'@
}

function Move-MSToolkitConsoleToToolMonitor {
    # Microsoft sign-in opens over this process's console window, even while it is hidden.
    # MSToolkit' child tools get this for free: MSToolkit moves each child's console onto its own
    # monitor before the child hides it. Stand-alone, this does the same thing: keeps the
    # hidden console (and its root owner) centered on whichever monitor the tool is on.
    param([switch]$Force)
    if ($ShowConsole) { return }
    try {
        $screen = [System.Windows.Forms.Screen]::FromControl($script:Form)
        if (-not $Force -and $screen.DeviceName -eq $script:ConsoleScreen) { return }
        $area = $screen.WorkingArea
        $console = [MSToolkitTeamsBlock.NativeMethods]::GetConsoleWindow()
        if ($console -eq [IntPtr]::Zero) { return }
        $handles = New-Object System.Collections.Generic.List[IntPtr]
        [void]$handles.Add($console)
        $root = [MSToolkitTeamsBlock.SignInWindowMover]::GetRootOwner($console)
        if ($root -ne [IntPtr]::Zero -and $root -ne $console) { [void]$handles.Add($root) }
        foreach ($h in $handles) {
            [void][MSToolkitTeamsBlock.SignInWindowMover]::CenterInArea($h, $area.Left, $area.Top, $area.Width, $area.Height)
        }
        $script:ConsoleScreen = $screen.DeviceName
    } catch { }
}

function Write-MSToolkitConsoleInfo {
    # One startup line showing what kind of console this process has.
    try {
        $console = [MSToolkitTeamsBlock.NativeMethods]::GetConsoleWindow()
        if ($console -eq [IntPtr]::Zero) { Write-Log 'Console window: none.' Info; return }
        $cls  = [MSToolkitTeamsBlock.SignInWindowMover]::ClassOfWindow($console)
        $root = [MSToolkitTeamsBlock.SignInWindowMover]::GetRootOwner($console)
        $rcls = [MSToolkitTeamsBlock.SignInWindowMover]::ClassOfWindow($root)
        Write-Log ('Console window: {0}; root owner: {1}.' -f $cls, $rcls) Info
        if ($cls -eq 'PseudoConsoleWindow') {
            Write-Log 'This window is running inside Windows Terminal. Start it with Launch-M365-Teams-Block-Number.cmd so it gets a classic console and the sign-in window follows the tool.' Warning
        }
    } catch { }
}

function Show-MSToolkitConsoleForSignIn {
    # Microsoft sign-in is parented to this process's console window. With the console
    # hidden, the sign-in window did not appear until the tool closed. During sign-in the
    # console is shown small and directly behind the tool (fully covered by it), which is
    # the same situation as running Connect-MicrosoftTeams from a normal PowerShell window.
    if ($ShowConsole) { return }
    try {
        $console = [MSToolkitTeamsBlock.NativeMethods]::GetConsoleWindow()
        if ($console -eq [IntPtr]::Zero) { return }
        if ([MSToolkitTeamsBlock.SignInWindowMover]::ShowBehind($console, $script:Form.Handle, 420, 160)) {
            $script:ConsoleShownForSignIn = $true
        }
    } catch {
        Write-Log ('Could not show the console for sign-in: ' + $_.Exception.Message) Warning
    }
}

function Hide-MSToolkitConsoleAfterSignIn {
    if (-not $script:ConsoleShownForSignIn) { return }
    $script:ConsoleShownForSignIn = $false
    Hide-PowerShellConsole
    Move-MSToolkitConsoleToToolMonitor -Force
}

function Start-MSToolkitSignInWindowWatch {
    try {
        $target = $script:Form.Handle
        Move-MSToolkitConsoleToToolMonitor -Force
        [MSToolkitTeamsBlock.SignInWindowMover]::Start($target, 180)
        Write-Log 'Watching for the sign-in window to move it to this screen.' Info
    } catch {
        Write-Log ('Sign-in window placement unavailable: ' + $_.Exception.Message) Warning
    }
}

function Stop-MSToolkitSignInWindowWatch {
    try {
        [MSToolkitTeamsBlock.SignInWindowMover]::Stop()
        foreach ($w in @([MSToolkitTeamsBlock.SignInWindowMover]::GetMovedWindows())) {
            Write-Log ('Moved sign-in window to this screen: ' + $w) Info
        }
        # Diagnostics: every new window that appeared during sign-in and what was done with it
        $diag = @([MSToolkitTeamsBlock.SignInWindowMover]::GetDiagnostics())
        if ($diag.Count -eq 0) {
            Write-Log 'Sign-in window check: no new windows were detected during sign-in.' Info
        } else {
            foreach ($d in $diag) { Write-Log ('Sign-in window check: ' + $d) Info }
        }
    } catch {
        Write-Log ('Sign-in window check failed: ' + $_.Exception.Message) Warning
    }
}
#endregion

#region ---------- Theme ----------
function ConvertFrom-MSToolkitHex { param([string]$Hex) return [System.Drawing.ColorTranslator]::FromHtml($Hex) }

function Resolve-MSToolkitThemeMode {
    param([string]$Requested)
    if ($Requested -eq 'Light' -or $Requested -eq 'Dark') { return $Requested }
    try {
        $settingsPath = Join-Path $env:APPDATA 'MSToolkit\settings.json'
        if (Test-Path -LiteralPath $settingsPath) {
            $s = Get-Content -LiteralPath $settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($s.Theme -eq 'Light' -or $s.Theme -eq 'Dark') { return [string]$s.Theme }
        }
    } catch { }
    return 'Light'
}

function Get-MSToolkitThemePalette {
    param([string]$Mode = 'Light')
    if ($Mode -eq 'Dark') {
        return @{
            FormBack    = ConvertFrom-MSToolkitHex '#1E1E1E'
            PanelBack   = ConvertFrom-MSToolkitHex '#2B2B2B'
            InputBack   = ConvertFrom-MSToolkitHex '#333333'
            Text        = ConvertFrom-MSToolkitHex '#EDEDED'
            SubText     = ConvertFrom-MSToolkitHex '#A8A8A8'
            Header      = ConvertFrom-MSToolkitHex '#0F2440'
            HeaderText  = ConvertFrom-MSToolkitHex '#FFFFFF'
            HeaderSub   = ConvertFrom-MSToolkitHex '#C9D6EA'
            Button      = ConvertFrom-MSToolkitHex '#2B4F7E'
            ButtonHover = ConvertFrom-MSToolkitHex '#3A659C'
            ButtonOff   = ConvertFrom-MSToolkitHex '#4A4A4A'
            ButtonText  = ConvertFrom-MSToolkitHex '#FFFFFF'
            Border      = ConvertFrom-MSToolkitHex '#444444'
            GridHeader  = ConvertFrom-MSToolkitHex '#0F2440'
            GridAlt     = ConvertFrom-MSToolkitHex '#303030'
            GridSelect  = ConvertFrom-MSToolkitHex '#3A659C'
            Good        = ConvertFrom-MSToolkitHex '#6CCB7F'
            Warn        = ConvertFrom-MSToolkitHex '#E3B341'
            Bad         = ConvertFrom-MSToolkitHex '#FF7A70'
            Action      = ConvertFrom-MSToolkitHex '#7FB3FF'
        }
    }
    return @{
        FormBack    = ConvertFrom-MSToolkitHex '#F0F2F5'
        PanelBack   = ConvertFrom-MSToolkitHex '#FFFFFF'
        InputBack   = ConvertFrom-MSToolkitHex '#FFFFFF'
        Text        = ConvertFrom-MSToolkitHex '#1E1E1E'
        SubText     = ConvertFrom-MSToolkitHex '#5F6B7A'
        Header      = ConvertFrom-MSToolkitHex '#1F3A5F'
        HeaderText  = ConvertFrom-MSToolkitHex '#FFFFFF'
        HeaderSub   = ConvertFrom-MSToolkitHex '#C9D6EA'
        Button      = ConvertFrom-MSToolkitHex '#1F3A5F'
        ButtonHover = ConvertFrom-MSToolkitHex '#2B4F7E'
        ButtonOff   = ConvertFrom-MSToolkitHex '#A7B0BC'
        ButtonText  = ConvertFrom-MSToolkitHex '#FFFFFF'
        Border      = ConvertFrom-MSToolkitHex '#D0D5DD'
        GridHeader  = ConvertFrom-MSToolkitHex '#1F3A5F'
        GridAlt     = ConvertFrom-MSToolkitHex '#F5F7FA'
        GridSelect  = ConvertFrom-MSToolkitHex '#CFE0F5'
        Good        = ConvertFrom-MSToolkitHex '#1E7B34'
        Warn        = ConvertFrom-MSToolkitHex '#9A6700'
        Bad         = ConvertFrom-MSToolkitHex '#B3261E'
        Action      = ConvertFrom-MSToolkitHex '#1F3A5F'
    }
}

function Set-MSToolkitButtonColor {
    param([System.Windows.Forms.Button]$Button)
    if ($Button.Enabled) { $Button.BackColor = $script:Palette.Button } else { $Button.BackColor = $script:Palette.ButtonOff }
}

function Apply-MSToolkitSharedTheme {
    param([System.Windows.Forms.Control]$Control, [hashtable]$Palette)
    foreach ($c in $Control.Controls) {
        $tag = [string]$c.Tag
        switch ($c.GetType().Name) {
            'Panel' {
                if ($tag -eq 'Header') { $c.BackColor = $Palette.Header } else { $c.BackColor = $Palette.FormBack }
            }
            'GroupBox' {
                $c.BackColor = $Palette.PanelBack
                $c.ForeColor = $Palette.Text
            }
            'Label' {
                $c.BackColor = [System.Drawing.Color]::Transparent
                switch ($tag) {
                    'HeaderTitle' { $c.ForeColor = $Palette.HeaderText }
                    'HeaderSub'   { $c.ForeColor = $Palette.HeaderSub }
                    'Sub'         { $c.ForeColor = $Palette.SubText }
                    default       { $c.ForeColor = $Palette.Text }
                }
            }
            'TextBox' {
                $c.BackColor = $Palette.InputBack
                $c.ForeColor = $Palette.Text
                $c.BorderStyle = 'FixedSingle'
            }
            'RichTextBox' {
                $c.BackColor = $Palette.InputBack
                $c.ForeColor = $Palette.Text
                $c.BorderStyle = 'FixedSingle'
            }
            'Button' {
                $c.FlatStyle = 'Flat'
                $c.FlatAppearance.BorderSize = 0
                $c.FlatAppearance.MouseOverBackColor = $Palette.ButtonHover
                if ($tag -eq 'HeaderButton') {
                    $c.BackColor = $Palette.Header
                    $c.ForeColor = $Palette.HeaderText
                } else {
                    $c.ForeColor = $Palette.ButtonText
                    Set-MSToolkitButtonColor -Button $c
                }
            }
            'DataGridView' {
                $c.BackgroundColor = $Palette.PanelBack
                $c.GridColor = $Palette.Border
                $c.BorderStyle = 'FixedSingle'
                $c.EnableHeadersVisualStyles = $false
                $c.ColumnHeadersDefaultCellStyle.BackColor = $Palette.GridHeader
                $c.ColumnHeadersDefaultCellStyle.ForeColor = $Palette.HeaderText
                $c.ColumnHeadersDefaultCellStyle.SelectionBackColor = $Palette.GridHeader
                $c.DefaultCellStyle.BackColor = $Palette.PanelBack
                $c.DefaultCellStyle.ForeColor = $Palette.Text
                $c.DefaultCellStyle.SelectionBackColor = $Palette.GridSelect
                $c.DefaultCellStyle.SelectionForeColor = $Palette.Text
                $c.AlternatingRowsDefaultCellStyle.BackColor = $Palette.GridAlt
            }
        }
        if ($c.HasChildren) { Apply-MSToolkitSharedTheme -Control $c -Palette $Palette }
    }
}

$script:ThemeMode = Resolve-MSToolkitThemeMode -Requested $ThemeMode
$script:Palette   = Get-MSToolkitThemePalette -Mode $script:ThemeMode
if ($ThemeMode) { $script:ThemeSource = '-ThemeMode parameter' } else { $script:ThemeSource = 'settings.json / default' }

function Save-MSToolkitThemeSetting {
    # Writes Theme into %APPDATA%\MSToolkit\settings.json, keeping every other setting
    # (SelectedDC, etc.). A file that exists but can't be parsed is left untouched.
    param([string]$Mode)
    $dir  = Join-Path $env:APPDATA 'MSToolkit'
    $path = Join-Path $dir 'settings.json'
    try {
        $settings = $null
        if (Test-Path -LiteralPath $path) {
            $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                try {
                    $settings = $raw | ConvertFrom-Json -ErrorAction Stop
                } catch {
                    Write-Log ('Theme not saved: {0} is not valid JSON, so it was left unchanged.' -f $path) Warning
                    return
                }
            }
        } elseif (-not (Test-Path -LiteralPath $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
        if ($null -eq $settings) { $settings = New-Object psobject }
        $settings | Add-Member -NotePropertyName 'Theme' -NotePropertyValue $Mode -Force
        ($settings | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $path -Encoding UTF8 -ErrorAction Stop
        Write-Log ('Theme set to {0} and saved to {1}.' -f $Mode, $path) Info
    } catch {
        Write-Log ('Theme set to {0} for this window, but could not be saved: {1}' -f $Mode, $_.Exception.Message) Warning
    }
}

function Get-MSToolkitPaletteKey {
    # Which palette role (Good/Bad/Warn/...) a status color came from, so it can be re-mapped.
    param([hashtable]$Palette, [System.Drawing.Color]$Color)
    foreach ($k in @('Good', 'Bad', 'Warn', 'SubText', 'Action', 'Text')) {
        if ($Palette[$k].ToArgb() -eq $Color.ToArgb()) { return $k }
    }
    return $null
}

function Get-MSToolkitValueLabels {
    param([System.Windows.Forms.Control]$Control)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($c in $Control.Controls) {
        if ($c -is [System.Windows.Forms.Label] -and [string]$c.Tag -eq 'Value') { [void]$list.Add($c) }
        if ($c.HasChildren) { foreach ($x in @(Get-MSToolkitValueLabels -Control $c)) { [void]$list.Add($x) } }
    }
    return $list.ToArray()
}

function Update-MSToolkitThemeButton {
    if (-not $script:btnTheme) { return }
    if ($script:ThemeMode -eq 'Dark') {
        $script:btnTheme.Text = $script:ThemeGlyphLight
        $script:ThemeToolTip.SetToolTip($script:btnTheme, 'Switch to light mode')
    } else {
        $script:btnTheme.Text = $script:ThemeGlyphDark
        $script:ThemeToolTip.SetToolTip($script:btnTheme, 'Switch to dark mode')
    }
}

function Switch-MSToolkitTheme {
    $newMode = 'Dark'
    if ($script:ThemeMode -eq 'Dark') { $newMode = 'Light' }
    $old = $script:Palette
    $new = Get-MSToolkitThemePalette -Mode $newMode

    # Remember the meaning of each status label's color before the theme resets it
    $labels = @(Get-MSToolkitValueLabels -Control $script:Form)
    $keys = @{}
    foreach ($l in $labels) { $keys[$l.Name + '|' + $l.GetHashCode()] = Get-MSToolkitPaletteKey -Palette $old -Color $l.ForeColor }

    $script:ThemeMode = $newMode
    $script:Palette   = $new
    $script:Form.SuspendLayout()
    try {
        $script:Form.BackColor = $new.FormBack
        Apply-MSToolkitSharedTheme -Control $script:Form -Palette $new
        foreach ($l in $labels) {
            $k = $keys[$l.Name + '|' + $l.GetHashCode()]
            if ($k) { $l.ForeColor = $new[$k] }
        }
        Update-MSToolkitThemeButton
        Update-MSToolkitLogTheme
    } finally {
        $script:Form.ResumeLayout()
    }
    Save-MSToolkitThemeSetting -Mode $newMode
}
#endregion

#region ---------- Logging ----------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error','Action')]
        [string]$Level = 'Info'
    )
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.ToUpper(), $Message

    [void]$script:LogEntries.Add([pscustomobject]@{ Line = $line; Level = $Level })
    if ($script:rtbLog) {
        Add-MSToolkitLogLine -Line $line -Level $Level
        $script:rtbLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch { }
}

function Add-MSToolkitLogLine {
    param([string]$Line, [string]$Level)
    switch ($Level) {
        'Success' { $color = $script:Palette.Good }
        'Warning' { $color = $script:Palette.Warn }
        'Error'   { $color = $script:Palette.Bad }
        'Action'  { $color = $script:Palette.Action }
        default   { $color = $script:Palette.Text }
    }
    $script:rtbLog.SelectionStart  = $script:rtbLog.TextLength
    $script:rtbLog.SelectionLength = 0
    $script:rtbLog.SelectionColor  = $color
    $script:rtbLog.AppendText($Line + "`r`n")
    $script:rtbLog.SelectionColor  = $script:Palette.Text
}

function Update-MSToolkitLogTheme {
    # Re-draws Activity Output so earlier lines use the new theme's colors
    if (-not $script:rtbLog) { return }
    $script:rtbLog.Clear()
    foreach ($e in $script:LogEntries) { Add-MSToolkitLogLine -Line $e.Line -Level $e.Level }
    $script:rtbLog.SelectionStart = $script:rtbLog.TextLength
    $script:rtbLog.ScrollToCaret()
}

function Write-MSToolkitErrorHint {
    param($ErrorRecord)
    $msg = [string]$ErrorRecord.Exception.Message
    if ($ErrorRecord.Exception.InnerException) { $msg += ' ' + [string]$ErrorRecord.Exception.InnerException.Message }

    switch -Regex ($msg) {
        '0x80070520|80070520|logon session' {
            Write-Log 'Hint: 0x80070520 means Windows sign-in (WAM) found no logon session for the account running this window. Run the tool as the Windows user signed in to this PC (use the launcher). Not Run as different user, and not from the Domain Admin MSToolkit session.' Warning
            break
        }
        'is not recognized as the name' {
            Write-Log 'Hint: the MicrosoftTeams module is not loaded in this session. See the Module status line and the install steps logged at startup.' Warning
            break
        }
        'already exist' {
            Write-Log 'Hint: a pattern with that rule name already exists. Refresh Tenant and look for it in the pattern list.' Warning
            break
        }
        'Access.?Denied|Forbidden|Unauthorized|not authorized|\b403\b|\b401\b' {
            Write-Log 'Hint: the signed-in account needs the Teams Administrator or Teams Communications Administrator role (active, if assigned through PIM).' Warning
            break
        }
        'AADSTS\d+' {
            Write-Log ('Hint: Microsoft sign-in returned {0}. Search that code on Microsoft Learn; AADSTS53003 is a Conditional Access block.' -f $Matches[0]) Warning
            break
        }
    }
}
#endregion

#region ---------- Number handling ----------
function ConvertTo-MSToolkitPhoneTarget {
    param([string]$InputText)

    $raw = ([string]$InputText).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }

    if ($raw -match '[A-Za-z]') {
        return [pscustomobject]@{ Valid = $false; Reason = 'Letters are not allowed. Remove extensions (x123) or words.' }
    }
    if ($raw -notmatch '^[\d\s\-\.\(\)\+]+$') {
        return [pscustomobject]@{ Valid = $false; Reason = 'Use only digits, spaces, and + - . ( )' }
    }
    if ($raw.IndexOf('+') -gt 0 -or ([regex]::Matches($raw, '\+')).Count -gt 1) {
        return [pscustomobject]@{ Valid = $false; Reason = 'A + is only allowed once, at the start.' }
    }

    $hasPlus = $raw.StartsWith('+')
    $digits  = $raw -replace '\D', ''
    $e164    = $null

    if (-not $hasPlus -and $digits.Length -eq 10) {
        $e164 = '1' + $digits
    } elseif ($digits.Length -eq 11 -and $digits.StartsWith('1')) {
        $e164 = $digits
    } elseif ($hasPlus -and -not $digits.StartsWith('1') -and $digits.Length -ge 8 -and $digits.Length -le 15) {
        $e164 = $digits
    } else {
        return [pscustomobject]@{ Valid = $false; Reason = 'Enter a 10-digit US number, 1 + 10 digits, or a full international number starting with +.' }
    }

    if ($e164.StartsWith('1')) {
        $npa = $e164.Substring(1, 3)
        $nxx = $e164.Substring(4, 3)
        if ('01'.Contains($npa.Substring(0, 1)) -or '01'.Contains($nxx.Substring(0, 1))) {
            return [pscustomobject]@{ Valid = $false; Reason = 'Not a valid North American number (area code and prefix cannot start with 0 or 1).' }
        }
        $display  = '+1 ({0}) {1}-{2}' -f $npa, $nxx, $e164.Substring(7, 4)
        $ruleName = 'Block-' + $e164.Substring(1)
    } else {
        $display  = '+' + $e164
        $ruleName = 'Block-' + $e164
    }

    return [pscustomobject]@{
        Valid    = $true
        Reason   = ''
        Digits   = $e164
        E164     = '+' + $e164
        Display  = $display
        Pattern  = '^\+?' + $e164 + '$'
        RuleName = $ruleName
    }
}

function Get-MSToolkitPatternId {
    param($PatternObject)
    if ($PatternObject.PSObject.Properties['Identity'] -and $PatternObject.Identity) { return [string]$PatternObject.Identity }
    if ($PatternObject.PSObject.Properties['Name'] -and $PatternObject.Name) { return [string]$PatternObject.Name }
    return ''
}

function Get-MSToolkitPatternNumber {
    # Returns the single number a pattern covers, or $null for ranges / other regex.
    param([string]$Pattern)
    $m = [regex]::Match([string]$Pattern, '^\^?(?:\\\+\??)?(\d{8,15})(\$?)$')
    if (-not $m.Success) { return $null }
    $digits = $m.Groups[1].Value
    if ($digits.Length -eq 11 -and $digits.StartsWith('1')) {
        $display = '+1 ({0}) {1}-{2}' -f $digits.Substring(1, 3), $digits.Substring(4, 3), $digits.Substring(7, 4)
    } else {
        $display = '+' + $digits
    }
    if (-not $m.Groups[2].Value) { $display += ' (prefix match)' }
    return [pscustomobject]@{ Digits = $digits; E164 = '+' + $digits; Display = $display }
}

function Get-MSToolkitMatchingPatterns {
    param($Patterns, [string]$E164Digits)
    $withPlus = '+' + $E164Digits
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($p in @($Patterns)) {
        if ($null -eq $p -or [string]::IsNullOrEmpty([string]$p.Pattern)) { continue }
        try {
            $mPlus = [regex]::IsMatch($withPlus, [string]$p.Pattern)
            $mBare = [regex]::IsMatch($E164Digits, [string]$p.Pattern)
        } catch {
            Write-Log ("Pattern '{0}' could not be evaluated locally ({1}). Relying on the Microsoft test for it." -f (Get-MSToolkitPatternId $p), $p.Pattern) Warning
            continue
        }
        if ($mPlus -or $mBare) {
            [void]$result.Add([pscustomobject]@{
                Identity    = Get-MSToolkitPatternId $p
                Enabled     = [bool]$p.Enabled
                Pattern     = [string]$p.Pattern
                MatchesPlus = $mPlus
                MatchesBare = $mBare
            })
        }
    }
    return $result.ToArray()
}

function Get-MSToolkitTestResultValue {
    param($TestResult)
    if ($null -eq $TestResult) { return $null }
    if ($TestResult.PSObject.Properties['IsNumberBlocked']) { return [bool]$TestResult.IsNumberBlocked }
    if ($TestResult.PSObject.Properties['IsMatch'])         { return [bool]$TestResult.IsMatch }
    return $null
}

function Get-MSToolkitNumberAnalysis {
    param($Target)

    $serviceBlocked = $null
    try {
        $r = Test-CsInboundBlockedNumberPattern -PhoneNumber $Target.E164 -ErrorAction Stop | Select-Object -First 1
        $serviceBlocked = Get-MSToolkitTestResultValue -TestResult $r
    } catch {
        Write-Log ('Test-CsInboundBlockedNumberPattern failed: ' + $_.Exception.Message) Warning
    }

    $blockedMatches = @(Get-MSToolkitMatchingPatterns -Patterns $script:BlockedPatterns -E164Digits $Target.Digits)
    $exemptMatches  = @(Get-MSToolkitMatchingPatterns -Patterns $script:ExemptPatterns  -E164Digits $Target.Digits)
    $enabledBlocked = @($blockedMatches | Where-Object { $_.Enabled })
    $enabledExempt  = @($exemptMatches  | Where-Object { $_.Enabled })
    $coversPlus     = @($enabledBlocked | Where-Object { $_.MatchesPlus }).Count -gt 0
    $coversBare     = @($enabledBlocked | Where-Object { $_.MatchesBare }).Count -gt 0
    $existingRule   = @($script:BlockedPatterns | Where-Object { (Get-MSToolkitPatternId $_) -eq $Target.RuleName }) | Select-Object -First 1

    return [pscustomobject]@{
        ServiceSaysBlocked = $serviceBlocked
        BlockedMatches     = $blockedMatches
        EnabledExempt      = $enabledExempt
        ExemptMatches      = $exemptMatches
        CoversPlus         = $coversPlus
        CoversBare         = $coversBare
        FullyCovered       = ($coversPlus -and $coversBare)
        ExistingRule       = $existingRule
    }
}

function Write-MSToolkitAnalysis {
    param($Target, $Analysis)

    Write-Log ('Number: {0}   Rule name: {1}   Pattern: {2}' -f $Target.Display, $Target.RuleName, $Target.Pattern) Info

    if ($Analysis.ServiceSaysBlocked -eq $true) {
        Write-Log 'Microsoft test (Test-CsInboundBlockedNumberPattern): BLOCKED' Success
    } elseif ($Analysis.ServiceSaysBlocked -eq $false) {
        Write-Log 'Microsoft test (Test-CsInboundBlockedNumberPattern): NOT blocked' Info
    } else {
        Write-Log 'Microsoft test (Test-CsInboundBlockedNumberPattern): no result' Warning
    }

    foreach ($m in $Analysis.BlockedMatches) {
        $lvl = 'Info'; if (-not $m.Enabled) { $lvl = 'Warning' }
        Write-Log ("Matches blocked pattern '{0}' (Enabled: {1}, Pattern: {2}) - with +: {3}, without +: {4}" -f $m.Identity, $m.Enabled, $m.Pattern, $m.MatchesPlus, $m.MatchesBare) $lvl
    }
    foreach ($m in $Analysis.ExemptMatches) {
        Write-Log ("Matches EXEMPT pattern '{0}' (Enabled: {1}, Pattern: {2}). An enabled exempt pattern overrides any block." -f $m.Identity, $m.Enabled, $m.Pattern) Warning
    }
    if ($Analysis.BlockedMatches.Count -eq 0) { Write-Log 'No existing blocked pattern matches this number.' Info }
    if ($Analysis.FullyCovered) {
        Write-Log 'Existing enabled patterns already cover this number with and without the +.' Success
    } elseif ($Analysis.CoversPlus -or $Analysis.CoversBare) {
        Write-Log 'Partial coverage: an existing pattern matches only one caller ID format (with or without +).' Warning
    }
    if ($script:BlockingEnabled -eq $false) {
        Write-Log 'Tenant call blocking is OFF. No blocked pattern is being enforced right now.' Warning
    }
}
#endregion

#region ---------- Tenant / Teams operations ----------
function Test-MSToolkitTeamsModule {
    param([switch]$Quiet)
    $mods = @(Get-Module -Name MicrosoftTeams -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending)
    if ($mods.Count -eq 0) {
        $script:ModuleAvailable = $false
        $script:lblModuleVal.Text = 'Not installed'
        $script:lblModuleVal.ForeColor = $script:Palette.Bad
        if (-not $Quiet) {
            Write-Log 'MicrosoftTeams module is not installed for this user. Click Install Module, or install it manually from a new PowerShell window running as this same user, then relaunch:' Error
            Write-Log '  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12' Info
            Write-Log '  Import-Module PackageManagement -MinimumVersion 1.4.7 -Force' Info
            Write-Log '  Import-Module PowerShellGet -MinimumVersion 2.2.5 -Force' Info
            Write-Log '  Install-Module MicrosoftTeams -Scope CurrentUser -Repository PSGallery' Info
        }
        Update-MSToolkitActionState
        return $false
    }
    $script:ModuleAvailable = $true
    $script:lblModuleVal.Text = 'MicrosoftTeams ' + [string]$mods[0].Version
    $script:lblModuleVal.ForeColor = $script:Palette.Good
    Update-MSToolkitActionState
    return $true
}

# Install runs in separate powershell.exe (-NonInteractive) processes so a prompt fails instead
# of hanging behind the hidden console. Stage 1 updates PackageManagement/PowerShellGet only if
# needed. Stage 2 always runs in a FRESH process: a PackageManagement assembly that is already
# loaded cannot be unloaded, and mixing it with a newer PowerShellGet in the same process causes
# "Unable to find module providers (PowerShellGet)".
$script:TeamsBootstrapScript = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Write-Output ('Running as ' + [Security.Principal.WindowsIdentity]::GetCurrent().Name)
    $pm = Get-Module PackageManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    $pg = Get-Module PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    Write-Output ('Found PackageManagement {0}, PowerShellGet {1}' -f $pm.Version, $pg.Version)
    if ($pm.Version -ge [version]'1.4.7' -and $pg.Version -ge [version]'2.2.5') {
        Write-Output 'PackageManagement and PowerShellGet are current. No update needed.'
        exit 0
    }
    Write-Output 'Updating PowerShellGet for this user (one-time)...'
    $nuget = Get-PackageProvider -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'NuGet' -and $_.Version -ge [version]'2.8.5.201' }
    if (-not $nuget) {
        Write-Output 'Installing NuGet package provider (CurrentUser)...'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
    }
    Install-Module -Name PowerShellGet -MinimumVersion 2.2.5 -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
    $pm = Get-Module PackageManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    $pg = Get-Module PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    if ($pm.Version -lt [version]'1.4.7' -or $pg.Version -lt [version]'2.2.5') { throw ('Update finished but found PackageManagement {0}, PowerShellGet {1}.' -f $pm.Version, $pg.Version) }
    Write-Output ('Updated: PackageManagement {0}, PowerShellGet {1}' -f $pm.Version, $pg.Version)
    exit 0
} catch {
    Write-Output ('ERROR: ' + $_.Exception.Message)
    exit 1
}
'@

$script:TeamsInstallScript = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Install-TeamsFromGalleryPackage {
    # Fallback that bypasses PackageManagement: download the .nupkg from PSGallery and
    # extract it into this user's module folder (same folder Install-Module -Scope CurrentUser uses).
    $userModules = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
    $work = Join-Path $env:TEMP ('MSToolkit-MicrosoftTeams-' + [guid]::NewGuid().ToString('N'))
    New-Item -Path $work -ItemType Directory -Force | Out-Null
    try {
        $zip = Join-Path $work 'MicrosoftTeams.zip'
        $ex  = Join-Path $work 'x'
        Write-Output 'Fallback: downloading MicrosoftTeams package directly from www.powershellgallery.com...'
        Invoke-WebRequest -Uri 'https://www.powershellgallery.com/api/v2/package/MicrosoftTeams' -OutFile $zip -UseBasicParsing
        Write-Output 'Extracting package...'
        Expand-Archive -LiteralPath $zip -DestinationPath $ex -Force
        $nuspec = Get-ChildItem -LiteralPath $ex -Filter '*.nuspec' | Select-Object -First 1
        if (-not $nuspec) { throw 'Downloaded package has no .nuspec file.' }
        [xml]$spec = Get-Content -LiteralPath $nuspec.FullName -Raw
        $ver = [string]$spec.package.metadata.version
        if (-not $ver) { throw 'Could not read the package version.' }
        $dest = Join-Path $userModules ('MicrosoftTeams\' + $ver)
        if (Test-Path -LiteralPath $dest) { throw ('Folder already exists, not overwriting: ' + $dest) }
        New-Item -Path $dest -ItemType Directory -Force | Out-Null
        Get-ChildItem -LiteralPath $ex -Force |
            Where-Object { $_.Name -ne '_rels' -and $_.Name -ne 'package' -and $_.Name -ne '[Content_Types].xml' -and $_.Extension -ne '.nuspec' } |
            ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $dest -Recurse -Force }
        Get-ChildItem -LiteralPath $dest -Recurse -File | Unblock-File
        Write-Output ('Extracted MicrosoftTeams {0} to {1}' -f $ver, $dest)
    } finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $pm = Get-Module PackageManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    $pg = Get-Module PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    $viaGallery = $false
    try {
        Import-Module PackageManagement -RequiredVersion $pm.Version -Force
        Import-Module PowerShellGet -RequiredVersion $pg.Version -Force
        Write-Output ('Loaded: ' + ((Get-Module PackageManagement, PowerShellGet | ForEach-Object { '{0} {1}' -f $_.Name, $_.Version }) -join ', '))
        Write-Output 'Installing MicrosoftTeams from PSGallery (CurrentUser)...'
        Install-Module -Name MicrosoftTeams -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        $viaGallery = $true
    } catch {
        Write-Output ('WARNING: Install-Module failed: ' + $_.Exception.Message)
    }
    if (-not $viaGallery) { Install-TeamsFromGalleryPackage }
    $m = Get-Module MicrosoftTeams -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $m) { throw 'MicrosoftTeams was not found after install.' }
    Write-Output ('INSTALLED: MicrosoftTeams {0} at {1}' -f $m.Version, $m.ModuleBase)
    exit 0
} catch {
    Write-Output ('ERROR: ' + $_.Exception.Message)
    exit 1
}
'@
function Write-MSToolkitInstallLine {
    param([string]$Line, [switch]$FromStdErr)
    $l = $Line.TrimEnd()
    if ([string]::IsNullOrWhiteSpace($l)) { return }
    if ($l.StartsWith('#< CLIXML') -or $l.StartsWith('<Objs')) { return }
    if ($l.StartsWith('ERROR:'))        { Write-Log ('  [install] ' + $l) Error;   return }
    if ($l.StartsWith('INSTALLED:'))    { Write-Log ('  [install] ' + $l) Success; return }
    if ($FromStdErr -or $l -match '^WARNING') { Write-Log ('  [install] ' + $l) Warning; return }
    Write-Log ('  [install] ' + $l) Info
}

function Read-MSToolkitInstallOutput {
    # Reads text appended to a redirect file since the last read and logs complete lines.
    param([hashtable]$State, [string]$Key, [string]$Path, [switch]$Flush)
    if (Test-Path -LiteralPath $Path) {
        $fs = $null
        try {
            $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            if ($fs.Length -gt $State[$Key + 'Pos']) {
                [void]$fs.Seek($State[$Key + 'Pos'], [System.IO.SeekOrigin]::Begin)
                $bytes = New-Object byte[] ($fs.Length - $State[$Key + 'Pos'])
                $read = $fs.Read($bytes, 0, $bytes.Length)
                $State[$Key + 'Pos'] += $read
                $State[$Key + 'Buf'] += [System.Text.Encoding]::Default.GetString($bytes, 0, $read)
            }
        } catch { } finally { if ($fs) { $fs.Dispose() } }
    }
    $parts = $State[$Key + 'Buf'] -split "`r?`n"
    $State[$Key + 'Buf'] = $parts[-1]
    for ($i = 0; $i -lt $parts.Count - 1; $i++) { Write-MSToolkitInstallLine -Line $parts[$i] -FromStdErr:($Key -eq 'Err') }
    if ($Flush -and $State[$Key + 'Buf']) {
        Write-MSToolkitInstallLine -Line $State[$Key + 'Buf'] -FromStdErr:($Key -eq 'Err')
        $State[$Key + 'Buf'] = ''
    }
}

function Invoke-MSToolkitInstallStage {
    # Runs one install script in its own powershell.exe and streams its output. Returns the exit code.
    param([string]$Name, [string]$ScriptText)
    $stamp   = [guid]::NewGuid().ToString('N')
    $outFile = Join-Path $env:TEMP ('MSToolkit-TeamsModuleInstall-{0}.out.txt' -f $stamp)
    $errFile = Join-Path $env:TEMP ('MSToolkit-TeamsModuleInstall-{0}.err.txt' -f $stamp)
    try {
        Write-Log ('Install stage: ' + $Name) Info
        $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($ScriptText))
        $psExe   = Join-Path $PSHOME 'powershell.exe'
        $proc = Start-Process -FilePath $psExe -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encoded) `
            -NoNewWindow -PassThru -RedirectStandardOutput $outFile -RedirectStandardError $errFile
        $null = $proc.Handle   # needed so ExitCode is populated on Windows PowerShell 5.1

        $state = @{ OutPos = [long]0; OutBuf = ''; ErrPos = [long]0; ErrBuf = '' }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $timedOut = $false
        while (-not $proc.HasExited) {
            Read-MSToolkitInstallOutput -State $state -Key 'Out' -Path $outFile
            Read-MSToolkitInstallOutput -State $state -Key 'Err' -Path $errFile
            if ($sw.Elapsed.TotalMinutes -ge 10) {
                try { $proc.Kill() } catch { }
                Write-Log ('Install stage "{0}" timed out after 10 minutes and was stopped.' -f $Name) Error
                $timedOut = $true
                break
            }
            for ($i = 0; $i -lt 5; $i++) { Start-Sleep -Milliseconds 50; [System.Windows.Forms.Application]::DoEvents() }
        }
        $proc.WaitForExit()
        Read-MSToolkitInstallOutput -State $state -Key 'Out' -Path $outFile -Flush
        Read-MSToolkitInstallOutput -State $state -Key 'Err' -Path $errFile -Flush
        if ($timedOut) { return -1 }
        return [int]$proc.ExitCode
    } finally {
        Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-MSToolkitInstallTeamsModule {
    Set-MSToolkitBusy $true
    try {
        $script:lblModuleVal.Text = 'Installing...'
        $script:lblModuleVal.ForeColor = $script:Palette.Warn
        Write-Log ('Installing MicrosoftTeams for {0}\{1} (CurrentUser scope, PSGallery). This can take a few minutes...' -f $env:USERDOMAIN, $env:USERNAME) Action

        $code = Invoke-MSToolkitInstallStage -Name 'PackageManagement / PowerShellGet check' -ScriptText $script:TeamsBootstrapScript
        if ($code -eq 0) {
            $code = Invoke-MSToolkitInstallStage -Name 'MicrosoftTeams install (fresh process)' -ScriptText $script:TeamsInstallScript
        }

        if ($code -eq 0 -and (Test-MSToolkitTeamsModule -Quiet)) {
            Write-Log 'MicrosoftTeams module installed. Click Connect to sign in.' Success
        } else {
            Write-Log ('Install did not complete (exit code {0}).' -f $code) Error
            Write-Log 'Common causes: PSGallery blocked by proxy/web filter (powershellgallery.com), or OneDrive Files On-Demand on Documents\WindowsPowerShell\Modules (set that folder to Always keep on this device).' Warning
            [void](Test-MSToolkitTeamsModule)
        }
    } catch {
        Write-Log ('Install failed to start: ' + $_.Exception.Message) Error
        [void](Test-MSToolkitTeamsModule)
    } finally {
        Set-MSToolkitBusy $false
    }
}
function Test-MSToolkitInteractiveUser {
    try {
        $me = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME
        $owners = New-Object System.Collections.Generic.List[string]
        foreach ($p in @(Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop)) {
            $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction SilentlyContinue
            if ($o -and $o.User) {
                $n = '{0}\{1}' -f $o.Domain, $o.User
                if (-not $owners.Contains($n)) { [void]$owners.Add($n) }
            }
        }
        if ($owners.Count -gt 0 -and -not ($owners -contains $me)) {
            Write-Log ('This window runs as {0}, but the signed-in Windows user is {1}. Microsoft sign-in will likely fail with 0x80070520. Relaunch as the signed-in user.' -f $me, ($owners.ToArray() -join ', ')) Warning
        }
    } catch { }
}

function Get-MSToolkitTenantDomainNames {
    param($Tenant)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($propName in @('VerifiedDomains', 'Domains')) {
        if ($names.Count -gt 0) { break }
        if (-not $Tenant.PSObject.Properties[$propName]) { continue }
        foreach ($d in @($Tenant.$propName)) {
            if ($null -eq $d) { continue }
            if ($d.PSObject.Properties['Name']) { $n = [string]$d.Name } else { $n = [string]$d }
            $n = $n.Trim().ToLowerInvariant()
            if ($n -and -not $names.Contains($n)) { [void]$names.Add($n) }
        }
    }
    return $names.ToArray()
}

function Confirm-MSToolkitTenant {
    $script:TenantVerified = $false
    $tenant = Get-CsTenant -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -First 1
    $script:TenantInfo = $tenant

    $script:lblTenantVal.Text   = [string]$tenant.DisplayName
    $script:lblTenantIdVal.Text = [string]$tenant.TenantId

    $domains  = @(Get-MSToolkitTenantDomainNames -Tenant $tenant)
    $expected = $ExpectedTenantDomain.Trim().ToLowerInvariant()
    $domainOk = ($domains -contains $expected)
    $idOk     = $true
    if ($TenantId) { $idOk = ([string]$tenant.TenantId -eq $TenantId.Trim()) }

    if ($domainOk -and $idOk) {
        $script:TenantVerified = $true
        $script:lblCheckVal.Text = 'Verified (' + $expected + ')'
        $script:lblCheckVal.ForeColor = $script:Palette.Good
        Write-Log ("Tenant verified: {0} ({1}) lists verified domain {2}." -f $tenant.DisplayName, $tenant.TenantId, $expected) Success
    } else {
        $script:lblCheckVal.Text = 'FAILED - blocking disabled'
        $script:lblCheckVal.ForeColor = $script:Palette.Bad
        if (-not $domainOk) {
            if (-not $expected) {
                Write-Log 'Tenant check FAILED: no expected tenant domain is set. Enter one of your verified domains under Expected tenant domain in MSToolkit Settings (Microsoft 365 and Intune), then reopen this tool.' Error
            } elseif ($domains.Count -eq 0) {
                Write-Log 'Tenant check FAILED: no verified domains were returned by Get-CsTenant.' Error
            } else {
                Write-Log ("Tenant check FAILED: '{0}' is not a verified domain here. Domains found: {1}" -f $expected, ($domains -join ', ')) Error
            }
        }
        if (-not $idOk) { Write-Log ("Tenant check FAILED: connected TenantId {0} does not match the Tenant ID setting {1}." -f $tenant.TenantId, $TenantId) Error }
        Write-Log 'Disconnect and sign in with an admin account from the correct tenant. If this IS the correct tenant, set Expected tenant domain (and Tenant ID, if used) in MSToolkit Settings to match it, then reopen this tool.' Warning
    }
}

function Update-MSToolkitTenantState {
    $cfg = Get-CsTenantBlockedCallingNumbers -ErrorAction Stop | Select-Object -First 1
    $script:BlockingEnabled = [bool]$cfg.Enabled
    $script:BlockedPatterns = @(Get-CsInboundBlockedNumberPattern -ErrorAction Stop)
    $script:ExemptPatterns  = @(Get-CsInboundExemptNumberPattern  -ErrorAction Stop)

    if ($script:BlockingEnabled) {
        $script:lblBlockingVal.Text = 'ON'
        $script:lblBlockingVal.ForeColor = $script:Palette.Good
    } else {
        $script:lblBlockingVal.Text = 'OFF (patterns not enforced)'
        $script:lblBlockingVal.ForeColor = $script:Palette.Warn
    }

    $script:grid.Rows.Clear()
    foreach ($p in $script:BlockedPatterns) {
        $num = Get-MSToolkitPatternNumber -Pattern ([string]$p.Pattern)
        if ($num) { $numText = $num.Display } else { $numText = '(range / custom pattern)' }
        [void]$script:grid.Rows.Add('Blocked', (Get-MSToolkitPatternId $p), $numText, [string]$p.Enabled, [string]$p.Pattern, [string]$p.Description)
    }
    foreach ($p in $script:ExemptPatterns) {
        $num = Get-MSToolkitPatternNumber -Pattern ([string]$p.Pattern)
        if ($num) { $numText = $num.Display } else { $numText = '(range / custom pattern)' }
        [void]$script:grid.Rows.Add('Exempt', (Get-MSToolkitPatternId $p), $numText, [string]$p.Enabled, [string]$p.Pattern, [string]$p.Description)
    }
    $script:grid.ClearSelection()

    Write-Log ('Tenant call blocking: {0}. Blocked patterns: {1}. Exempt patterns: {2}.' -f $(if ($script:BlockingEnabled) { 'ON' } else { 'OFF' }), $script:BlockedPatterns.Count, $script:ExemptPatterns.Count) Info
}

function Clear-MSToolkitConnectionDisplay {
    $script:lblAccountVal.Text  = '-'
    $script:lblTenantVal.Text   = '-'
    $script:lblTenantIdVal.Text = '-'
    $script:lblCheckVal.Text    = 'Not connected'
    $script:lblCheckVal.ForeColor = $script:Palette.SubText
    $script:lblBlockingVal.Text = '-'
    $script:lblBlockingVal.ForeColor = $script:Palette.SubText
    $script:grid.Rows.Clear()
}
#endregion

#region ---------- UI state ----------
function Update-MSToolkitActionState {
    $hasTarget = ($null -ne $script:CurrentTarget)
    $script:btnConnect.Enabled    = (-not $script:IsBusy) -and (-not $script:IsConnected)
    $script:btnInstall.Enabled    = (-not $script:IsBusy) -and ($script:ModuleAvailable -eq $false)
    $script:btnDisconnect.Enabled = (-not $script:IsBusy) -and $script:IsConnected
    $script:btnRefresh.Enabled    = (-not $script:IsBusy) -and $script:IsConnected
    $script:btnTest.Enabled       = (-not $script:IsBusy) -and $script:IsConnected -and $hasTarget
    $script:btnBlock.Enabled      = (-not $script:IsBusy) -and $script:IsConnected -and $script:TenantVerified -and $hasTarget
    $hasBlockedSel = (@(Get-MSToolkitSelectedBlockedIds).Count -gt 0)
    $script:btnUnblock.Enabled    = (-not $script:IsBusy) -and $script:IsConnected -and $script:TenantVerified -and $hasBlockedSel
}

function Get-MSToolkitSelectedBlockedIds {
    $ids = New-Object System.Collections.Generic.List[string]
    if (-not $script:grid) { return $ids.ToArray() }
    foreach ($row in $script:grid.SelectedRows) {
        if ([string]$row.Cells['Type'].Value -eq 'Blocked') {
            $id = [string]$row.Cells['Identity'].Value
            if ($id -and -not $ids.Contains($id)) { [void]$ids.Add($id) }
        }
    }
    return $ids.ToArray()
}

function Select-MSToolkitBlockedRows {
    param([string[]]$Identities)
    $script:grid.ClearSelection()
    $first = $null
    foreach ($row in $script:grid.Rows) {
        if ([string]$row.Cells['Type'].Value -eq 'Blocked' -and (@($Identities) -contains [string]$row.Cells['Identity'].Value)) {
            $row.Selected = $true
            if ($null -eq $first) { $first = $row.Index }
        }
    }
    if ($null -ne $first) { $script:grid.FirstDisplayedScrollingRowIndex = $first }
}

function Reset-MSToolkitCursor {
    # Clears the wait cursor and makes Windows re-evaluate the pointer now,
    # instead of waiting for the next mouse move.
    try {
        $script:Form.UseWaitCursor = $false
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        $pos = [System.Windows.Forms.Cursor]::Position
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($pos.X + 1), $pos.Y)
        [System.Windows.Forms.Cursor]::Position = $pos
    } catch { }
}

function Set-MSToolkitBusy {
    param([bool]$Busy)
    $script:IsBusy = $Busy
    if ($Busy) { $script:Form.UseWaitCursor = $true } else { Reset-MSToolkitCursor }
    Update-MSToolkitActionState
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-MSToolkitPreview {
    $t = ConvertTo-MSToolkitPhoneTarget -InputText $script:txtPhone.Text
    $script:CurrentTarget = $null
    if ($null -eq $t) {
        $script:lblNormVal.Text = '-'
        $script:lblNormVal.ForeColor = $script:Palette.Text
        $script:lblPatVal.Text  = '-'
        $script:lblRuleVal.Text = '-'
    } elseif (-not $t.Valid) {
        $script:lblNormVal.Text = $t.Reason
        $script:lblNormVal.ForeColor = $script:Palette.Warn
        $script:lblPatVal.Text  = '-'
        $script:lblRuleVal.Text = '-'
    } else {
        $script:CurrentTarget = $t
        $script:lblNormVal.Text = $t.Display + '   (' + $t.E164 + ')'
        $script:lblNormVal.ForeColor = $script:Palette.Text
        $script:lblPatVal.Text  = $t.Pattern
        $script:lblRuleVal.Text = $t.RuleName
    }
    Update-MSToolkitActionState
}

function Get-MSToolkitDescription {
    $ticket = ([string]$script:txtTicket.Text -replace '[\r\n\t]', ' ').Trim()
    $by = $script:ConnectedAccount
    if (-not $by) { $by = $env:USERNAME }
    $d = 'Blocked tenant-wide via Teams Call Blocking by {0} on {1}' -f $by, (Get-Date -Format 'yyyy-MM-dd')
    if ($ticket) { $d += ' - ' + $ticket }
    return $d
}
#endregion

#region ---------- Button actions ----------
function Invoke-MSToolkitConnect {
    if (-not (Test-MSToolkitTeamsModule)) { return }
    Set-MSToolkitBusy $true
    try {
        if (-not (Get-Module -Name MicrosoftTeams)) {
            Write-Log 'Loading MicrosoftTeams module...' Info
            Import-Module MicrosoftTeams -ErrorAction Stop -WarningAction SilentlyContinue
        }
        Write-Log 'Opening Microsoft sign-in. Sign in with your Teams admin account.' Action
        Write-Log 'If no sign-in window appears within a few seconds, check the taskbar or Alt+Tab for it. This window waits until sign-in finishes or is cancelled.' Info
        $connectParams = @{ ErrorAction = 'Stop' }
        if ($TenantId) { $connectParams['TenantId'] = $TenantId.Trim() }
        Start-MSToolkitSignInWindowWatch
        Show-MSToolkitConsoleForSignIn
        try {
            $conn = Connect-MicrosoftTeams @connectParams | Select-Object -First 1
        } finally {
            Stop-MSToolkitSignInWindowWatch
            Hide-MSToolkitConsoleAfterSignIn
        }

        $script:IsConnected = $true
        $script:ConnectedAccount = ''
        if ($conn -and $conn.Account) { $script:ConnectedAccount = [string]$conn.Account }
        if ($script:ConnectedAccount) { $script:lblAccountVal.Text = $script:ConnectedAccount } else { $script:lblAccountVal.Text = '(connected)' }
        Write-Log ('Connected to Teams as {0}.' -f $script:lblAccountVal.Text) Success

        Confirm-MSToolkitTenant
        Update-MSToolkitTenantState
    } catch {
        Write-Log ('Connect failed: ' + $_.Exception.Message) Error
        Write-MSToolkitErrorHint $_
        if ($script:IsConnected) {
            Write-Log 'Connected, but the tenant read failed. Blocking stays disabled.' Warning
            $script:TenantVerified = $false
        }
    } finally {
        Set-MSToolkitBusy $false
    }
}

function Invoke-MSToolkitDisconnect {
    Set-MSToolkitBusy $true
    try {
        Disconnect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Log 'Disconnected from Microsoft Teams.' Info
    } catch {
        Write-Log ('Disconnect returned: ' + $_.Exception.Message) Warning
    } finally {
        $script:IsConnected      = $false
        $script:TenantVerified   = $false
        $script:ConnectedAccount = ''
        $script:BlockingEnabled  = $null
        $script:BlockedPatterns  = @()
        $script:ExemptPatterns   = @()
        Clear-MSToolkitConnectionDisplay
        Set-MSToolkitBusy $false
    }
}

function Invoke-MSToolkitRefresh {
    Set-MSToolkitBusy $true
    try {
        Confirm-MSToolkitTenant
        Update-MSToolkitTenantState
    } catch {
        Write-Log ('Refresh failed: ' + $_.Exception.Message) Error
        Write-MSToolkitErrorHint $_
    } finally {
        Set-MSToolkitBusy $false
    }
}

function Invoke-MSToolkitTestNumber {
    $t = $script:CurrentTarget
    if ($null -eq $t) { return }
    Set-MSToolkitBusy $true
    try {
        Write-Log ('--- Read-only test: {0} ---' -f $t.Display) Action
        Update-MSToolkitTenantState
        $a = Get-MSToolkitNumberAnalysis -Target $t
        Write-MSToolkitAnalysis -Target $t -Analysis $a
        if ($a.ExistingRule) {
            Write-Log ("A rule named '{0}' already exists (Enabled: {1}, Pattern: {2})." -f $t.RuleName, $a.ExistingRule.Enabled, $a.ExistingRule.Pattern) Info
        }
        $matchIds = @($a.BlockedMatches | ForEach-Object { $_.Identity })
        if ($matchIds.Count -gt 0) {
            Select-MSToolkitBlockedRows -Identities $matchIds
            Write-Log 'Matching blocked rule(s) are selected in the pattern list. Use Unblock Selected to remove them.' Info
        }
        Write-Log 'Test complete. No changes were made.' Info
    } catch {
        Write-Log ('Test failed: ' + $_.Exception.Message) Error
        Write-MSToolkitErrorHint $_
    } finally {
        Set-MSToolkitBusy $false
    }
}

function Invoke-MSToolkitBlockNumber {
    $t = $script:CurrentTarget
    if ($null -eq $t) { return }
    if (-not $script:TenantVerified) {
        Write-Log 'Block refused: the connected tenant has not passed the tenant check.' Error
        return
    }

    Set-MSToolkitBusy $true
    try {
        Write-Log ('=== Block request: {0} ===' -f $t.Display) Action

        # Read-only pre-checks against current tenant state
        Update-MSToolkitTenantState
        $a = Get-MSToolkitNumberAnalysis -Target $t
        Write-MSToolkitAnalysis -Target $t -Analysis $a

        if ($a.ExistingRule) {
            Write-Log ("No change made: a rule named '{0}' already exists (Enabled: {1}, Pattern: {2}). Review it in the pattern list." -f $t.RuleName, $a.ExistingRule.Enabled, $a.ExistingRule.Pattern) Warning
            return
        }
        if ($a.FullyCovered) {
            $names = (@($a.BlockedMatches | Where-Object { $_.Enabled } | ForEach-Object { $_.Identity }) -join ', ')
            Write-Log ('No change made: already blocked by existing enabled pattern(s): {0}.' -f $names) Success
            return
        }

        $desc = Get-MSToolkitDescription
        $tenantName = [string]$script:TenantInfo.DisplayName
        $tenantIdTx = [string]$script:TenantInfo.TenantId

        $msg  = "Block this number for the ENTIRE tenant?`r`n`r`n"
        $msg += "Number:       $($t.Display)`r`n"
        $msg += "Rule name:    $($t.RuleName)`r`n"
        $msg += "Pattern:      $($t.Pattern)`r`n"
        $msg += "Description:  $desc`r`n"
        $msg += "Tenant:       $tenantName ($tenantIdTx)`r`n`r`n"
        $msg += "Effect: inbound PSTN calls from this number are rejected for every user, call queue, and auto attendant in the tenant. Teams-to-Teams and federated calls are not affected. Existing patterns are not changed."
        if ($a.EnabledExempt.Count -gt 0) {
            $msg += "`r`n`r`nWARNING: enabled exempt pattern(s) match this number: " + (@($a.EnabledExempt | ForEach-Object { $_.Identity }) -join ', ') + ". The exempt pattern overrides the block until it is removed or disabled."
        }
        if ($script:BlockingEnabled -eq $false) {
            $msg += "`r`n`r`nNOTE: tenant call blocking is currently OFF. You will be asked separately whether to turn it on."
        }

        $answer = [System.Windows.Forms.MessageBox]::Show($script:Form, $msg, 'Confirm tenant-wide block',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log 'Cancelled by operator. No change made.' Warning
            return
        }

        Write-Log ("Running: New-CsInboundBlockedNumberPattern -Identity '{0}' -Pattern '{1}' -Enabled `$true -Description '{2}'" -f $t.RuleName, $t.Pattern, $desc) Action
        New-CsInboundBlockedNumberPattern -Identity $t.RuleName -Pattern $t.Pattern -Description $desc -Enabled $true -ErrorAction Stop | Out-Null
        Write-Log ("Created blocked number pattern '{0}'." -f $t.RuleName) Success

        if ($script:BlockingEnabled -eq $false) {
            $enabledCount = @($script:BlockedPatterns | Where-Object { $_.Enabled }).Count + 1
            $m2  = "Tenant call blocking is OFF for $tenantName, so the new rule is not enforced yet.`r`n`r`n"
            $m2 += "Turn tenant call blocking ON now?`r`n`r`n"
            $m2 += "Target: tenant-wide setting (Set-CsTenantBlockedCallingNumbers -Enabled `$true)`r`n"
            $m2 += "Effect: every ENABLED blocked pattern in the tenant starts rejecting calls ($enabledCount pattern(s), including the new one). Existing patterns are not modified."
            $answer2 = [System.Windows.Forms.MessageBox]::Show($script:Form, $m2, 'Turn on tenant call blocking',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
            if ($answer2 -eq [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log 'Running: Set-CsTenantBlockedCallingNumbers -Enabled $true' Action
                Set-CsTenantBlockedCallingNumbers -Enabled $true -ErrorAction Stop | Out-Null
                Write-Log 'Tenant call blocking turned ON.' Success
            } else {
                Write-Log 'Tenant call blocking left OFF. The new rule exists but is NOT enforced.' Warning
            }
        }

        # Verify
        Write-Log 'Verifying...' Info
        for ($i = 0; $i -lt 6; $i++) { Start-Sleep -Milliseconds 500; [System.Windows.Forms.Application]::DoEvents() }
        Update-MSToolkitTenantState
        $created = @($script:BlockedPatterns | Where-Object { (Get-MSToolkitPatternId $_) -eq $t.RuleName }) | Select-Object -First 1
        if ($created) {
            Write-Log ("Rule read back: '{0}'  Enabled: {1}  Pattern: {2}" -f (Get-MSToolkitPatternId $created), $created.Enabled, $created.Pattern) Success
        } else {
            Write-Log ("Rule '{0}' was not returned by Get-CsInboundBlockedNumberPattern yet. Click Refresh Tenant in a minute." -f $t.RuleName) Warning
        }

        $v = Get-MSToolkitNumberAnalysis -Target $t
        if ($v.ServiceSaysBlocked -eq $true -and $v.FullyCovered -and $script:BlockingEnabled -and $v.EnabledExempt.Count -eq 0) {
            Write-Log ('VERIFIED: {0} is blocked tenant-wide (Test-CsInboundBlockedNumberPattern = True).' -f $t.Display) Success
        } else {
            if ($v.ServiceSaysBlocked -ne $true) { Write-Log 'Verification: Microsoft test did not return True yet. Use Test Number again in a few minutes.' Warning }
            if (-not $v.FullyCovered)            { Write-Log 'Verification: the rule list does not yet show full coverage for this number.' Warning }
            if (-not $script:BlockingEnabled)    { Write-Log 'Verification: tenant call blocking is OFF, so the block is not enforced.' Warning }
            if ($v.EnabledExempt.Count -gt 0)    { Write-Log 'Verification: an enabled exempt pattern still overrides this block.' Warning }
        }
        Write-Log ("To reverse this block later: Remove-CsInboundBlockedNumberPattern -Identity '{0}'" -f $t.RuleName) Info
    } catch {
        Write-Log ('Block failed: ' + $_.Exception.Message) Error
        Write-MSToolkitErrorHint $_
    } finally {
        Set-MSToolkitBusy $false
    }
}
function Invoke-MSToolkitUnblockSelected {
    $ids = @(Get-MSToolkitSelectedBlockedIds)
    if ($ids.Count -eq 0) { return }
    if (-not $script:TenantVerified) {
        Write-Log 'Unblock refused: the connected tenant has not passed the tenant check.' Error
        return
    }
    $exemptSel = @($script:grid.SelectedRows | Where-Object { [string]$_.Cells['Type'].Value -eq 'Exempt' }).Count
    if ($exemptSel -gt 0) { Write-Log ('{0} selected Exempt row(s) skipped. This tool does not change exempt patterns.' -f $exemptSel) Warning }

    Set-MSToolkitBusy $true
    try {
        Write-Log ('=== Unblock request: {0} rule(s) ===' -f $ids.Count) Action

        # Re-read so the confirmation reflects the tenant as it is now
        Update-MSToolkitTenantState
        $targets = New-Object System.Collections.Generic.List[object]
        foreach ($id in $ids) {
            $p = @($script:BlockedPatterns | Where-Object { (Get-MSToolkitPatternId $_) -eq $id }) | Select-Object -First 1
            if (-not $p) {
                Write-Log ("Rule '{0}' no longer exists in the tenant. Skipped." -f $id) Warning
                continue
            }
            [void]$targets.Add([pscustomobject]@{
                Identity    = $id
                Pattern     = [string]$p.Pattern
                Enabled     = [bool]$p.Enabled
                Description = [string]$p.Description
                Number      = Get-MSToolkitPatternNumber -Pattern ([string]$p.Pattern)
            })
        }
        if ($targets.Count -eq 0) {
            Write-Log 'No change made: none of the selected rules exist any more.' Warning
            return
        }

        $tenantName = [string]$script:TenantInfo.DisplayName
        $tenantIdTx = [string]$script:TenantInfo.TenantId
        $hasRange = $false
        $msg = "Remove these blocked number rule(s) from the ENTIRE tenant?`r`n`r`n"
        foreach ($tg in $targets) {
            if ($tg.Number) { $numText = $tg.Number.Display } else { $numText = '(range / custom pattern)'; $hasRange = $true }
            $msg += "Rule:     $($tg.Identity)`r`n"
            $msg += "Number:   $numText`r`n"
            $msg += "Pattern:  $($tg.Pattern)   (Enabled: $($tg.Enabled))`r`n"
            if ($tg.Description) { $msg += "Desc:     $($tg.Description)`r`n" }
            $msg += "`r`n"
        }
        $msg += "Tenant:   $tenantName ($tenantIdTx)`r`n`r`n"
        $msg += "Effect: inbound PSTN calls from numbers matching these rules will ring through again for every user, call queue, and auto attendant. Other rules and the tenant on/off setting are not changed."
        if ($hasRange) {
            $msg += "`r`n`r`nWARNING: at least one rule is a range or custom pattern and may cover many numbers."
        }

        $answer = [System.Windows.Forms.MessageBox]::Show($script:Form, $msg, 'Confirm tenant-wide unblock',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log 'Cancelled by operator. No change made.' Warning
            return
        }

        $removed = New-Object System.Collections.Generic.List[object]
        foreach ($tg in $targets) {
            try {
                Write-Log ("Running: Remove-CsInboundBlockedNumberPattern -Identity '{0}'" -f $tg.Identity) Action
                # This cmdlet has no -Confirm parameter. $ConfirmPreference = 'None' (scoped to this
                # call) keeps any confirmation prompt from waiting behind the hidden console;
                # the tool's own confirmation dialog above is the approval step.
                & {
                    $ConfirmPreference = 'None'
                    Remove-CsInboundBlockedNumberPattern -Identity $tg.Identity -ErrorAction Stop | Out-Null
                }
                Write-Log ("Removed blocked number pattern '{0}' (Pattern: {1})." -f $tg.Identity, $tg.Pattern) Success
                Write-Log ("To re-block: New-CsInboundBlockedNumberPattern -Identity '{0}' -Pattern '{1}' -Enabled `${2} -Description '{3}'" -f $tg.Identity, $tg.Pattern, $tg.Enabled.ToString().ToLower(), ($tg.Description -replace "'", "''")) Info
                [void]$removed.Add($tg)
            } catch {
                Write-Log ("Remove failed for '{0}': {1}" -f $tg.Identity, $_.Exception.Message) Error
                Write-MSToolkitErrorHint $_
            }
        }
        if ($removed.Count -eq 0) { return }

        # Verify
        Write-Log 'Verifying...' Info
        for ($i = 0; $i -lt 6; $i++) { Start-Sleep -Milliseconds 500; [System.Windows.Forms.Application]::DoEvents() }
        Update-MSToolkitTenantState
        foreach ($tg in $removed) {
            $still = @($script:BlockedPatterns | Where-Object { (Get-MSToolkitPatternId $_) -eq $tg.Identity }).Count -gt 0
            if ($still) {
                Write-Log ("Rule '{0}' is still returned by Get-CsInboundBlockedNumberPattern. Click Refresh Tenant in a minute." -f $tg.Identity) Warning
                continue
            }
            Write-Log ("Rule '{0}' no longer listed." -f $tg.Identity) Success
            if (-not $tg.Number) {
                Write-Log ("Rule '{0}' was a range / custom pattern; test specific numbers with Test Number." -f $tg.Identity) Info
                continue
            }
            $blockedNow = $null
            try {
                $r = Test-CsInboundBlockedNumberPattern -PhoneNumber $tg.Number.E164 -ErrorAction Stop | Select-Object -First 1
                $blockedNow = Get-MSToolkitTestResultValue -TestResult $r
            } catch {
                Write-Log ('Test-CsInboundBlockedNumberPattern failed: ' + $_.Exception.Message) Warning
            }
            $others = @(Get-MSToolkitMatchingPatterns -Patterns $script:BlockedPatterns -E164Digits $tg.Number.Digits | Where-Object { $_.Enabled })
            if ($others.Count -gt 0) {
                Write-Log ('{0} is STILL blocked by other enabled rule(s): {1}' -f $tg.Number.Display, (@($others | ForEach-Object { $_.Identity }) -join ', ')) Warning
            } elseif ($blockedNow -eq $true) {
                Write-Log ('{0}: Microsoft test still returns blocked. Use Test Number again in a few minutes.' -f $tg.Number.Display) Warning
            } else {
                Write-Log ('VERIFIED: {0} is no longer blocked tenant-wide.' -f $tg.Number.Display) Success
            }
        }
    } catch {
        Write-Log ('Unblock failed: ' + $_.Exception.Message) Error
        Write-MSToolkitErrorHint $_
    } finally {
        Set-MSToolkitBusy $false
    }
}
#endregion

#region ---------- Build form ----------
$fontMain  = New-Object System.Drawing.Font('Segoe UI', 9)
$fontBold  = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$fontInput = New-Object System.Drawing.Font('Segoe UI', 10)
$fontTitle = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$fontLog   = New-Object System.Drawing.Font('Consolas', 9)

function New-MSToolkitLabel {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 110, [string]$Tag = '', [System.Drawing.Font]$Font = $fontMain)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 18)
    $l.Font = $Font
    $l.Tag = $Tag
    $l.AutoEllipsis = $true
    return $l
}

function New-MSToolkitButton {
    param([string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    $b.Font = $fontBold
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.Add_EnabledChanged({ param($s, $e) Set-MSToolkitButtonColor -Button $s })
    return $b
}

$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = 'MSToolkit - ' + $script:ToolTitle
$script:Form.Size = New-Object System.Drawing.Size(944, 910)
$script:Form.MinimumSize = New-Object System.Drawing.Size(944, 790)
$script:Form.StartPosition = 'CenterScreen'
$script:Form.Font = $fontMain
$script:Form.BackColor = $script:Palette.FormBack

# Header
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = 'Top'
$pnlHeader.Height = 64
$pnlHeader.Tag = 'Header'
$lblTitle = New-MSToolkitLabel -Text $script:ToolTitle -X 16 -Y 6 -W 600 -Tag 'HeaderTitle' -Font $fontTitle
$lblTitle.Height = 30
$lblSub = New-MSToolkitLabel -Text 'Tenant-wide inbound PSTN number blocking - Microsoft Teams' -X 18 -Y 38 -W 700 -Tag 'HeaderSub'

# Theme switcher (top right of the header)
$probeFont = New-Object System.Drawing.Font('Segoe MDL2 Assets', 14)
if ($probeFont.Name -eq 'Segoe MDL2 Assets') {
    $themeFont = $probeFont
    $script:ThemeGlyphDark  = [string][char]0xE708   # moon  (shown in Light mode)
    $script:ThemeGlyphLight = [string][char]0xE706   # sun   (shown in Dark mode)
} else {
    $probeFont.Dispose()
    $themeFont = New-Object System.Drawing.Font('Segoe UI Symbol', 14)
    $script:ThemeGlyphDark  = [string][char]0x263E
    $script:ThemeGlyphLight = [string][char]0x2600
}
$script:ThemeToolTip = New-Object System.Windows.Forms.ToolTip
$script:btnTheme = New-Object System.Windows.Forms.Button
$script:btnTheme.Size = New-Object System.Drawing.Size(40, 40)
$script:btnTheme.Location = New-Object System.Drawing.Point(880, 12)
$script:btnTheme.Font = $themeFont
$script:btnTheme.Tag = 'HeaderButton'
$script:btnTheme.TabStop = $false
$script:btnTheme.Cursor = [System.Windows.Forms.Cursors]::Hand
$script:btnTheme.TextAlign = 'MiddleCenter'
$pnlHeader.Controls.AddRange(@($lblTitle, $lblSub, $script:btnTheme))
$pnlHeader.Add_Resize({ $script:btnTheme.Left = $pnlHeader.ClientSize.Width - $script:btnTheme.Width - 14 })

# Tenant connection group
$grpConn = New-Object System.Windows.Forms.GroupBox
$grpConn.Text = 'Tenant Connection'
$grpConn.Location = New-Object System.Drawing.Point(12, 74)
$grpConn.Size = New-Object System.Drawing.Size(904, 134)
$grpConn.Anchor = 'Top, Left, Right'

$script:btnConnect    = New-MSToolkitButton -Text 'Connect'        -X 14  -Y 24 -W 130 -H 30
$script:btnDisconnect = New-MSToolkitButton -Text 'Disconnect'     -X 150 -Y 24 -W 110 -H 30
$script:btnRefresh    = New-MSToolkitButton -Text 'Refresh Tenant' -X 266 -Y 24 -W 130 -H 30
$script:btnInstall    = New-MSToolkitButton -Text 'Install Module' -X 402 -Y 24 -W 130 -H 30

$lblAccount  = New-MSToolkitLabel -Text 'Signed in as:'  -X 14 -Y 64
$lblTenant   = New-MSToolkitLabel -Text 'Tenant:'        -X 14 -Y 84
$lblTenantId = New-MSToolkitLabel -Text 'Tenant ID:'     -X 14 -Y 104
$script:lblAccountVal  = New-MSToolkitLabel -Text '-' -X 124 -Y 64  -W 320 -Tag 'Value' -Font $fontBold
$script:lblTenantVal   = New-MSToolkitLabel -Text '-' -X 124 -Y 84  -W 320 -Tag 'Value' -Font $fontBold
$script:lblTenantIdVal = New-MSToolkitLabel -Text '-' -X 124 -Y 104 -W 320 -Tag 'Value'

$lblCheck    = New-MSToolkitLabel -Text 'Tenant check:'  -X 470 -Y 64
$lblBlocking = New-MSToolkitLabel -Text 'Call blocking:' -X 470 -Y 84
$lblModule   = New-MSToolkitLabel -Text 'Module:'        -X 470 -Y 104
$script:lblCheckVal    = New-MSToolkitLabel -Text 'Not connected' -X 580 -Y 64  -W 310 -Tag 'Value' -Font $fontBold
$script:lblBlockingVal = New-MSToolkitLabel -Text '-'             -X 580 -Y 84  -W 310 -Tag 'Value' -Font $fontBold
$script:lblModuleVal   = New-MSToolkitLabel -Text 'Checking...'   -X 580 -Y 104 -W 310 -Tag 'Value'

$grpConn.Controls.AddRange(@($script:btnConnect, $script:btnDisconnect, $script:btnRefresh, $script:btnInstall,
    $lblAccount, $lblTenant, $lblTenantId, $script:lblAccountVal, $script:lblTenantVal, $script:lblTenantIdVal,
    $lblCheck, $lblBlocking, $lblModule, $script:lblCheckVal, $script:lblBlockingVal, $script:lblModuleVal))

# Block group
$grpBlock = New-Object System.Windows.Forms.GroupBox
$grpBlock.Text = 'Block a Number (tenant-wide, inbound PSTN)'
$grpBlock.Location = New-Object System.Drawing.Point(12, 216)
$grpBlock.Size = New-Object System.Drawing.Size(904, 176)
$grpBlock.Anchor = 'Top, Left, Right'

$lblPhone = New-MSToolkitLabel -Text 'Phone number:' -X 14 -Y 30
$script:txtPhone = New-Object System.Windows.Forms.TextBox
$script:txtPhone.Location = New-Object System.Drawing.Point(124, 27)
$script:txtPhone.Size = New-Object System.Drawing.Size(230, 24)
$script:txtPhone.Font = $fontInput
$script:txtPhone.MaxLength = 30
$lblPhoneHint = New-MSToolkitLabel -Text 'e.g. 202-555-0100  or  +44 20 7946 0000' -X 364 -Y 30 -W 270 -Tag 'Sub'

$lblTicket = New-MSToolkitLabel -Text 'Ticket / reason:' -X 14 -Y 64
$script:txtTicket = New-Object System.Windows.Forms.TextBox
$script:txtTicket.Location = New-Object System.Drawing.Point(124, 61)
$script:txtTicket.Size = New-Object System.Drawing.Size(510, 24)
$script:txtTicket.Font = $fontInput
$script:txtTicket.MaxLength = 150

$lblNorm = New-MSToolkitLabel -Text 'Normalized:' -X 14 -Y 100
$lblPat  = New-MSToolkitLabel -Text 'Pattern:'    -X 14 -Y 122
$lblRule = New-MSToolkitLabel -Text 'Rule name:'  -X 14 -Y 144
$script:lblNormVal = New-MSToolkitLabel -Text '-' -X 124 -Y 100 -W 510 -Tag 'Value' -Font $fontBold
$script:lblPatVal  = New-MSToolkitLabel -Text '-' -X 124 -Y 122 -W 510 -Tag 'Value'
$script:lblRuleVal = New-MSToolkitLabel -Text '-' -X 124 -Y 144 -W 510 -Tag 'Value'

$script:btnTest  = New-MSToolkitButton -Text 'Test Number (read-only)'  -X 666 -Y 24 -W 222 -H 34
$script:btnBlock = New-MSToolkitButton -Text 'Block Number Tenant-Wide' -X 666 -Y 66 -W 222 -H 42
$script:btnTest.Anchor  = 'Top, Right'
$script:btnBlock.Anchor = 'Top, Right'

$grpBlock.Controls.AddRange(@($lblPhone, $script:txtPhone, $lblPhoneHint, $lblTicket, $script:txtTicket,
    $lblNorm, $lblPat, $lblRule, $script:lblNormVal, $script:lblPatVal, $script:lblRuleVal,
    $script:btnTest, $script:btnBlock))

# Patterns group
$grpPatterns = New-Object System.Windows.Forms.GroupBox
$grpPatterns.Text = 'Current Tenant Patterns'
$grpPatterns.Location = New-Object System.Drawing.Point(12, 400)
$grpPatterns.Size = New-Object System.Drawing.Size(904, 240)
$grpPatterns.Anchor = 'Top, Left, Right'

$lblPatHint = New-MSToolkitLabel -Text 'Select one or more Blocked rows (Ctrl/Shift-click), then Unblock Selected. Exempt rows are view-only.' -X 12 -Y 30 -W 690 -Tag 'Sub'
$script:btnUnblock = New-MSToolkitButton -Text 'Unblock Selected' -X 722 -Y 22 -W 172 -H 30
$script:btnUnblock.Anchor = 'Top, Right'

$script:grid = New-Object System.Windows.Forms.DataGridView
$script:grid.Location = New-Object System.Drawing.Point(10, 60)
$script:grid.Size = New-Object System.Drawing.Size(884, 170)
$script:grid.Anchor = 'Top, Left, Right, Bottom'
$script:grid.ReadOnly = $true
$script:grid.AllowUserToAddRows = $false
$script:grid.AllowUserToDeleteRows = $false
$script:grid.AllowUserToResizeRows = $false
$script:grid.RowHeadersVisible = $false
$script:grid.SelectionMode = 'FullRowSelect'
$script:grid.MultiSelect = $true
$script:grid.AutoSizeColumnsMode = 'Fill'
$script:grid.ColumnHeadersHeightSizeMode = 'AutoSize'
[void]$script:grid.Columns.Add('Type', 'Type')
[void]$script:grid.Columns.Add('Identity', 'Rule name')
[void]$script:grid.Columns.Add('Number', 'Number')
[void]$script:grid.Columns.Add('Enabled', 'Enabled')
[void]$script:grid.Columns.Add('Pattern', 'Pattern')
[void]$script:grid.Columns.Add('Description', 'Description')
$script:grid.Columns['Type'].FillWeight = 11
$script:grid.Columns['Identity'].FillWeight = 22
$script:grid.Columns['Number'].FillWeight = 24
$script:grid.Columns['Enabled'].FillWeight = 10
$script:grid.Columns['Pattern'].FillWeight = 22
$script:grid.Columns['Description'].FillWeight = 40
$grpPatterns.Controls.AddRange(@($lblPatHint, $script:btnUnblock, $script:grid))

# Activity Output
$lblLog = New-MSToolkitLabel -Text 'Activity Output' -X 12 -Y 648 -W 300 -Font $fontBold
$script:rtbLog = New-Object System.Windows.Forms.RichTextBox
$script:rtbLog.Location = New-Object System.Drawing.Point(12, 668)
$script:rtbLog.Size = New-Object System.Drawing.Size(904, 190)
$script:rtbLog.Anchor = 'Top, Bottom, Left, Right'
$script:rtbLog.ReadOnly = $true
$script:rtbLog.Font = $fontLog
$script:rtbLog.WordWrap = $true
$script:rtbLog.DetectUrls = $false

$script:Form.Controls.AddRange(@($grpConn, $grpBlock, $grpPatterns, $lblLog, $script:rtbLog, $pnlHeader))

Apply-MSToolkitSharedTheme -Control $script:Form -Palette $script:Palette
Update-MSToolkitThemeButton
$script:lblCheckVal.ForeColor    = $script:Palette.SubText
$script:lblBlockingVal.ForeColor = $script:Palette.SubText
#endregion

#region ---------- Events ----------
$script:btnConnect.Add_Click({ Invoke-MSToolkitConnect })
$script:btnDisconnect.Add_Click({ Invoke-MSToolkitDisconnect })
$script:btnRefresh.Add_Click({ Invoke-MSToolkitRefresh })
$script:btnInstall.Add_Click({ Invoke-MSToolkitInstallTeamsModule })
$script:btnTheme.Add_Click({ Switch-MSToolkitTheme })
$script:btnTest.Add_Click({ Invoke-MSToolkitTestNumber })
$script:btnBlock.Add_Click({ Invoke-MSToolkitBlockNumber })
$script:btnUnblock.Add_Click({ Invoke-MSToolkitUnblockSelected })
$script:grid.Add_SelectionChanged({ if (-not $script:IsBusy) { Update-MSToolkitActionState } })

$script:txtPhone.Add_TextChanged({ Update-MSToolkitPreview })
$script:txtPhone.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        if ($script:btnTest.Enabled) { Invoke-MSToolkitTestNumber }
    }
})

# Startup checks run from a one-shot timer so the window finishes painting and goes idle first.
# That ends the Windows "app starting" cursor from the launch, and the checks then run under
# the normal busy cursor, which is cleared when they finish.
# One-shot check 2 seconds after startup finishes: records which cursor Windows is showing,
# clears a leftover wait cursor, and logs the result so a stuck cursor can be identified.
$script:CursorCheckTimer = New-Object System.Windows.Forms.Timer
$script:CursorCheckTimer.Interval = 2000
$script:CursorCheckTimer.Add_Tick({
    $script:CursorCheckTimer.Stop()
    try {
        $name = [MSToolkitTeamsBlock.CursorProbe]::CurrentCursorName()
        $over = $script:Form.Bounds.Contains([System.Windows.Forms.Cursor]::Position)
        Write-Log ('Cursor check: Windows is showing {0}; mouse over this window: {1}; busy: {2}; UseWaitCursor: {3}.' -f $name, $over, $script:IsBusy, $script:Form.UseWaitCursor) Info
        if (-not $script:IsBusy -and $name -like 'Wait*') { Reset-MSToolkitCursor }
    } catch { }
})

$script:StartupTimer = New-Object System.Windows.Forms.Timer
$script:StartupTimer.Interval = 300
$script:StartupTimer.Add_Tick({
    $script:StartupTimer.Stop()
    Set-MSToolkitBusy $true
    try {
        Test-MSToolkitInteractiveUser
        if (Test-MSToolkitTeamsModule -Quiet) {
            Write-Log 'Click Connect to sign in to Microsoft Teams.' Info
        } else {
            Write-Log 'MicrosoftTeams module not detected for this user. Starting install automatically.' Warning
            Invoke-MSToolkitInstallTeamsModule
        }
    } finally {
        Set-MSToolkitBusy $false
        $script:txtPhone.Focus() | Out-Null
        $script:CursorCheckTimer.Start()
    }
})

$script:Form.Add_Shown({
    if (-not $ShowConsole) { Hide-PowerShellConsole }
    Move-MSToolkitConsoleToToolMonitor -Force
    $script:Form.Activate()
    Update-MSToolkitActionState
    $shownDomain = if ($ExpectedTenantDomain) { $ExpectedTenantDomain } else { '(not set - blocking stays disabled)' }
    $shownTenant = if ($TenantId) { $TenantId } else { '(not set - not checked)' }
    Write-Log ('{0} started. Running as {1}\{2}. Theme: {3} (from {4}). Expected tenant domain: {5}. Tenant ID: {6}. Tenant settings from: {7}.' -f $script:ToolTitle, $env:USERDOMAIN, $env:USERNAME, $script:ThemeMode, $script:ThemeSource, $shownDomain, $shownTenant, $script:TenantSettingSource) Info
    Write-Log ('Log file: ' + $script:LogFile) Info
    Write-MSToolkitConsoleInfo
    $script:txtPhone.Focus() | Out-Null
    $script:StartupTimer.Start()
})

# Keep the hidden console on the tool's monitor whenever the tool is dragged to another screen
$script:Form.Add_LocationChanged({ Move-MSToolkitConsoleToToolMonitor })

$script:Form.Add_FormClosing({
    if ($script:IsConnected) {
        try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
})
#endregion

Update-MSToolkitPreview
Clear-MSToolkitConnectionDisplay
[void]$script:Form.ShowDialog()
$script:Form.Dispose()
