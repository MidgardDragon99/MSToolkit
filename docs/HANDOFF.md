# MSToolkit - Handoff notes

Background that cannot be read from the code alone. Facts only. Anything marked
**unverified - check the code** may have changed since this was written
(2026-11-18).

Everywhere below, `contoso` names are placeholders. Nothing in this repository may
contain a real organisation's name, domain, tenant ID, client ID, hostname, OU
path, person, username or naming convention.

---

## 1. What MSToolkit is

MSToolkit is a PowerShell 5.1 / WinForms desktop toolkit for day-to-day Active
Directory and Microsoft 365 administration: a main console (`MSToolkit.ps1`) for
common AD tasks, plus companion tools launched from it.

It is a generalised branch of an internal toolkit built for one organisation. The
differences:

- **Every site-specific value moved into a settings file.** The internal version
  hardcoded its domain, OU paths, sync server, tenant and client IDs, UPN and mail
  domains, company name, product names and document locations. MSToolkit ships
  with all of these blank.
- **Renamed throughout.** Script, launcher, EXE, icon, installer, install folder,
  runtime folder, settings folder and the internal `<prefix>`-named functions all
  use `MSToolkit`. Child tools keep descriptive names (`NewADUser.ps1`,
  `M365-*.ps1`, `IntuneTools.ps1`).
- **Different storage locations,** so both can be installed on one PC without
  colliding.
- **No automatic domain prefix** on any username (section 5).
- **Security hardening the internal version did not originally have:** scripts
  run from Program Files, not a user-writable folder (section 9); the Offboard
  User undo journals live in the admin's own profile, and Undo refuses SID
  mismatches and protected groups.

**Branch rule:** changes made here are never copied back automatically. The two
codebases are developed separately; a fix that belongs in both is applied to each
deliberately.

### Accounts

The console and the AD child tools run as an **admin account** (via `runas`).
The Microsoft 365 tools, IntuneTools and the launchers run as the **standard
signed-in Windows account**. Settings are per account, in each profile's
`%APPDATA%\MSToolkit`.

---

## 2. Settings model

### Where it lives

- `%APPDATA%\MSToolkit\settings.json`, in the profile of the account running the
  tool. For the console and AD tools that is the **admin** account's profile.
- Only the console's **Settings** window writes the site values.
- The same file also holds `Theme`, `SelectedDC` and each child tool's `Theme_*`
  key. Every save is **read-modify-write** so those survive.
- The AD child tools carry their own read-only copy of the settings readers so
  they work standalone. They never write site values.
- Cloud tools run as the standard account, whose profile has a different
  `settings.json`. The console therefore **passes the values they need as
  parameters at launch** (see each tool in section 4). Run standalone, some fall
  back to reading their own account's file.

### Key functions (in `MSToolkit.ps1`)

| Function | Purpose |
| --- | --- |
| `Get-MSToolkitSettingDefaults` | All keys, in Settings-window order. Every default is `""`. |
| `Get-MSToolkitSettings` / `Save-MSToolkitSettings` | Load (cached; `-Reload`) and read-modify-write save |
| `Get-MSToolkitSetting -Name X -Default Y` | One value, trimmed; blank returns `-Default` |
| `Get-MSToolkitOU` | A configured OU, or a fallback |
| `Get-MSToolkitDomainRootDN` | `DC=...` built from the DNS name, no directory call |
| `Test-MSToolkitConfigured` | True only once at least one site value is non-blank (the file alone doesn't count). Settings opens automatically until then. |
| `Get-MSToolkitDomainPrefix` | Exists but is not used for sign-in (section 5). **Unverified - check the code** whether anything still calls it. |

### Every key, and what happens when it's blank

| Key | Used by | If blank |
| --- | --- | --- |
| `Domain` | Console | Read from the computer's domain |
| `NetBiosName` | Nothing currently (the hint says so) | - |
| `EntraSyncServer` | Delta Sync button; Offboard User delta-sync step | Delta Sync reports no server configured; the Offboard step is hidden and Do All skips it |
| `OUUsers` | User pickers, report scopes, NewADUser default OU | Domain root. NewADUser starts at the root and **refuses to create an account in the root**. |
| `OUAdmins` | Added to user pickers | Nothing added |
| `OUComputers` | Computer OU lists | One level under the root, not all levels |
| `OUServers`, `OUDisabled`, `OUSecurityGroups`, `OUDistributionGroups` | OU lists, group lists, Offboard disabled-OU reminder | Domain root, with an orange "OU not set" reminder in the output |
| `OUServiceAccounts` | Get Special Function Accounts | Button refuses until set |
| `UpnDomain`, `MailDomain` | NewADUser | **Required.** Account creation is refused before anything is written. |
| `OnMicrosoftDomain`, `OnMicrosoftMailDomain` | NewADUser proxy aliases; SharePoint admin URL fallback | Aliases skipped |
| `CompanyName` | NewADUser `company` attribute | Not written |
| `TenantId`, `ClientId` | IntuneTools (passed as `-OrgTenantId` / `-OrgClientId`); Teams tool (`-TenantId`) | IntuneTools must be configured in its own Settings tab; Teams skips the tenant ID check |
| `ExpectedTenantDomain` | Phone tool (block and forward) | **Changes stay disabled** - read-only |
| `SharePointAdminUrl` | OneDrive / SharePoint (`-AdminUrl`) | Built as `https://<name>-admin.sharepoint.com` from `OnMicrosoftDomain`; if that's blank too, typed in the tool |
| `OffboardMfaName` / `OffboardMfaUrl` | Offboard User | MFA step hidden (both blank) |
| `OffboardAccessName` / `OffboardAccessHost` | Offboard User | Physical-access step hidden. A host starting `http(s)://` opens in the browser; anything else opens Remote Desktop. |
| `OffboardRmmName` / `OffboardRmmReport` | Offboard User | The report sub-steps are dropped (both needed) |
| `OffboardArchiveName` | Offboard User | Archive sub-steps dropped |
| `OffboardLobSystems` | Offboard User | Multi-line, `Name \| address \| note` per line. Blank hides the step; buttons for the first five addresses. |
| `OffboardDeviceDocument`, `OffboardChecklistDocument` | Offboard User | No document button. Value may be a full path, a URL or a file name in the install folder. |
| `OffboardChecklistLocation` | Offboard User | Free text; the final step just says to keep a record |
| `OffboardExportFolder` | Offboard User group export | `C:\ProgramData\MSToolkit\Logs` |

### Other per-account storage

| Path (per account) | Contents |
| --- | --- |
| `%APPDATA%\MSToolkit\m365-credential.xml` (admin profile) | The standard account's saved password, DPAPI-encrypted, for launching M365 tools |
| `%APPDATA%\MSToolkit\signin-auto.flag` (admin profile) | "Sign in automatically" opt-in |
| `%APPDATA%\MSToolkit\launcher-admin-username.txt` (standard profile) | The admin username the `runas` launchers remember. Never a password. |
| `%APPDATA%\MSToolkit\IntuneTools\config.json`, `token.dat` (standard profile) | IntuneTools settings and cached refresh token |
| `%LOCALAPPDATA%\MSToolkit\Offboard-Undo\` (admin profile) | Offboard User undo journals |

**Clear Settings** blanks every site key, and can optionally forget the saved M365
password and the launcher username (the latter runs as the signed-in Windows
account, because the admin can't reach that profile).

---

## 3. Safety model

### Critical-object protection

Destructive and sensitive operations call a `Get-MSToolkitCritical*Reason`
function first; a non-null reason goes to `Stop-MSToolkitCriticalOperation`,
which blocks with a dialog and an output line. Matching is on **SID / RID**, never
display names, with one documented exception below.

| Object | Blocked when |
| --- | --- |
| User | RID 500 (Administrator), 501 (Guest), 502 (KRBTGT), or `isCriticalSystemObject` |
| Group | Builtin SIDs `S-1-5-32-544` / `548` / `549` / `550` / `551` / `552`; domain RIDs 512-522, 526, 527; `isCriticalSystemObject`; **DnsAdmins by name** (it has no well-known RID) |
| Computer | `userAccountControl` has `SERVER_TRUST_ACCOUNT` (8192) or `PrimaryGroupID` 516 (domain controller), or `isCriticalSystemObject` |

The same functions are copied into child tools that need them (Offboard User,
Compare-UserGroups). Group membership is treated as **direct** unless nested is
explicitly asked for.

### Deletion protection

Deleting a user, group or computer that has **Protect object from accidental
deletion** set asks first (`Request-MSToolkitAccidentalDeletionOverride`): Yes
clears the flag for that one deletion, No cancels. If the deletion then fails,
`Restore-MSToolkitAccidentalDeletionProtection` sets the flag back. Critical
objects can never be overridden here.

### Confirmations and verification

- Destructive actions name the resolved object (name, username, DN) and default
  to **No**.
- Writes are **read back** before success is reported or a checklist box ticks.
- Offboard User and pickers guard against acting on the wrong record: clear before
  loading, confirm the match, re-check the picker against the loaded object at
  action time.

### Reports

CSV reports (90-Day Inactive Users/Computers, Password Expiring Soon) and logs go
to `C:\ProgramData\MSToolkit\Logs` (fallback `%LOCALAPPDATA%\MSToolkit\Logs`
outside an install). The **Logs** button opens that folder as the signed-in user.
**Unverified - check the code** for any report that opens its own child window
rather than writing to the output box.

### How child tools launch

| Tool type | Runs as | How |
| --- | --- | --- |
| AD tools (NewADUser, Offboard User, Compare, Lockout) | The admin account (same as console) | `Start-Process powershell.exe -File ...` from `$PSScriptRoot`, passing `-ThemeMode` and the selected DC (`-Server` / `-DefaultServer`) |
| M365 tools, IntuneTools | The standard signed-in account | `Start-MSToolkitM365Tool` -> `Invoke-MSToolkitAsSignedInUser` -> `Start-Process -Credential` |

Windows are moved to the console's monitor (`Move-ProcessWindowsToMSToolkitMonitor`).

**Parent window link.** The console and Offboard User add `-ParentWindow <handle
as a decimal number> -ParentPid <process id>` to every tool they start
(`Get-MSToolkitParentArguments` in the console; inline in Offboard's
`Start-MSToolkitM365Tool`). Every child declares both (`[long]`/`[int]`, default
0) - an unknown parameter stops a tool starting, so a new child must declare them
too. The same "Parent window link" block is in each child: a small C# class
(`MSToolkitParentLink.Native`) plus three functions, and three top-level handlers
registered just before `ShowDialog`:

- Shown and Activated: `SetWindowPos(parent, child, SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE)`
  puts the parent directly behind the child. A minimised parent stays minimised.
- FormClosing (skipped if an earlier handler cancelled): `AllowSetForegroundWindow(ASFW_ANY)`,
  restore the parent if minimised, `SetForegroundWindow(parent)`. Done in
  FormClosing because the closing window still has the foreground then; the
  parent calling `Activate()` after the child exits is refused by the foreground
  lock.
- Before every call the handle is checked with `IsWindow` and
  `GetWindowThreadProcessId` against `-ParentPid`, so a closed parent or a reused
  handle makes it a no-op. A tool started on its own has no parent.

After a Microsoft sign-in each M365 tool calls `Invoke-MSToolkitRaiseAfterSignIn`
(after its sign-in window watcher stops): if it isn't in front, wait 300 ms, then
attach to the foreground thread + `BringWindowToTop`/`SetForegroundWindow`, then a
synthetic Alt key, then TopMost on/off; then the parent goes behind it again.
IntuneTools keeps its own `Show-MainWindow` (the same steps).

`Start-Process -Credential` (CreateProcessWithLogonW) fails past about 1024
characters with only "The parameter is incorrect", so both launchers refuse a
command line over 1000 characters with a clear message. Pass paths and ids, never
data, on a command line.

### M365 sign-in

The console runs as admin, but M365 work must run as the standard account.
`Invoke-MSToolkitAsSignedInUser` shows a sign-in box for that account (empty
unless a sign-in was saved; entered as `DOMAIN\username` or a UPN, nothing
added). Options:

- **Remember password** - DPAPI-encrypted to `m365-credential.xml`, readable only
  by that admin account on that PC.
- **Sign in automatically** - skips the box. **Hold Shift** while clicking to
  bring it back.
- A wrong or outdated saved password is detected because `Start-Process
  -Credential` throws; the box reappears.

Offboard User contains a copy of this flow and shares the same files.

Each cloud tool then signs in to its own service interactively: Exchange
(`Connect-ExchangeOnline`), SharePoint (`Connect-SPOService`), Teams
(`Connect-MicrosoftTeams`), Microsoft Graph PowerShell for Group Compare and
Conditional Access, and IntuneTools' own OAuth flow (browser loopback or device
code) against the tenant's app registration. All are **delegated** - the
signed-in admin's own roles apply.

---

## 4. The tools

### Console - `MSToolkit.ps1`

Sections: Users, Groups, Computers, OUs, M365, Reports, plus a top bar (DC
selector, replication, Delta Sync, DC Health, ADUC, GPMC, Print Management, RDP to
selected DC, PowerShell/ISE, Logs, theme, Settings).

Notable:

- **DC Health** - replication partners with `FailureCount` 0 show amber
  "Recovered", confirmed by `Get-ADReplicationPartnerMetadata` matched on the
  partner's NTDS Settings DN. Partners failing now but not yet in the failure
  cache are reported red. **Untested** (section 8).
- **Remove User From All Sessions** - scans enabled Windows computers active in
  the last 30 days (TCP 445 check, then `quser`), confirms, then `logoff`. Never
  logs off sessions on the local PC if the target is the account running the
  console or the signed-in Windows user. No SMB sessions.
- **Unforce Password Change** - acts only when `pwdLastSet` is 0; notes the
  password age restarts and whether Password never expires is set.
- **Add Phone Number** - shows current value; blank clears (confirmation says
  CLEAR); reads back.

### NewADUser.ps1

Parameters: `-ThemeMode`, `-DefaultServer`. Creates an account from settings
(UPN and SMTP domains required; onmicrosoft aliases optional; `company` from
`CompanyName`). Optional cloning of groups, address (including
`physicalDeliveryOfficeName`, the Office field) and manager, then a post-creation
check of every copied attribute. OU picker starts at `OUUsers`; the domain root is
refused.

- Mirror Groups From refuses a critical mirror user (SID / RID,
  `isCriticalSystemObject`) before the confirmation, and nothing is created.
- Each mirrored group is read with SID and `isCriticalSystemObject` and checked.
  The confirmation lists every group by name - Add, BLOCKED with the reason, NOT
  READ - only allowed groups are added, and the BLOCKED and NOT READ lists are
  repeated in the output after the run.
- "User must change password at next logon" is ticked every time the form opens
  and after each account is created (never remembered), and can be unticked.
  `New-ADUser -ChangePasswordAtLogon` uses its value; the confirmation and the
  output show "Must change password at next logon: Yes/No", and the final
  verification checks it against `pwdLastSet` (0 = must change).

### Compare-UserGroups.ps1

Parameters: `-Server`, `-ThemeMode`. Compares two users' direct AD group
memberships and adds missing ones; Manage AD User Groups dialog adds/removes.
Users from `OUUsers` + `OUAdmins`.

- Critical users (SID / RID, `isCriticalSystemObject`) are blocked on every
  path - Compare (reference and target), Compare -> Add, Manage -> Load, Add,
  Remove - before any grid, picker or confirmation.
- Critical groups are hard-blocked on both Add paths and on Remove: an OK-only
  Stop box listing them, a BLOCKED log line each, and they leave the selection.
  The name-pattern HIGH RISK flags stay warnings.

### Investigate-AccountLockout.ps1

Parameter: `-ThemeMode`. Lockout investigation across DCs. Users from `OUUsers`
+ `OUAdmins`.

### OffboardUser.ps1

Parameters: `-Server`, `-ThemeMode`. Departing-employee checklist, 18 steps with
every setting filled; steps hide when their setting is blank, and numbering is
assigned at creation so there are no gaps.

- **Tool-performed, auto-ticking only when verified:** disable, export and remove
  security groups (protected groups never removed), clear `telephoneNumber` /
  `mobile`, replicate selected DC (`repadmin /syncall /AdeP`), Entra Connect delta
  sync (recognises "already running").
- **Do All Local Actions** - one confirmation; exports groups first; doesn't stop
  at the first failure; names anything outstanding.
- **Step order** (as created in the code; numbers are assigned at creation):
  1 disable (AD), 2 security groups (AD), 3 clear phone (AD), 4 replicate,
  5 delta sync (1-5 are what Do All Local Actions runs), 6 Entra and Intune
  (opens IntuneTools in user mode, `-ReuseWindow`; straight after the sync so
  sessions are revoked as soon as the disable reaches the cloud), 7 distro groups
  and 8 M365 groups (Distro / Group Compare with the user loaded), 9 MFA,
  10 physical access, 11 update the ticket and
  12 PC decommission (both **Done / Deferred**), 13 Mailbox handover, 14 OneDrive
  handover, 15 Retention/phone/licences (each opens its tool with `-ManageUser`,
  `-AutoConnect` and the manager's UPN as `-Delegate`), 16 disabled-OU reminder,
  17 line-of-business systems, 18 file the checklist. MFA and physical access
  come after the local AD actions, matching the internal build. Nothing in the
  code refers to a step by number.
- **Untick log:** a real click that unticks a step (or a sub-step click that
  unticks its parent) logs `Step N unticked by hand: <title>`. It hangs off the
  `Click` event, which code changes to `.Checked` never raise, so Undo's own
  "step unticked" lines are not duplicated. Number and title are NoteProperties
  on the checkbox (`StepNumber`, `StepTitle`).
- **LOADED:** line; a disabled account shows **DISABLED** in a separate red label
  placed with `TextRenderer.MeasureText`. Redrawn from a fresh AD read after
  Disable, Clear Phone, Do All Local Actions (`Update-OffboardCurrentUser`) and
  Undo.
- **Undo...** - JSON journal per run in `%LOCALAPPDATA%\MSToolkit\Offboard-Undo`,
  flushed after each action. Reverses groups, then attributes, then re-enables
  (unless already disabled before the run). AD only. Before the confirmation
  (`Get-OffboardUndoPreflight`) it refuses a record with no SID, resolves the
  account **by SID** (refused if not found; a moved account is named in the
  confirmation and its current DN is used), refuses a protected account, and
  resolves every group in the record - refusing the whole record if any is
  protected, since offboarding never removes one. Groups that can't be found are
  listed and skipped. During the undo it re-checks the SID and each group again.
  New records store each group's SID; older ones are resolved by DN.
  Several records can be selected (Ctrl/Shift-click) and undone together
  (`Invoke-OffboardUndoSelection`): every record is checked first, one
  confirmation lists what will be undone and what was refused and why (nothing
  but the refused list when all are refused), records are undone newest first
  by `StartedUtc`, each on its own, and a summary lists undone / refused /
  failed per record.
- **Copy Checklist / Finish** - text record with `[x]` / `[ ]`, sub-steps, and
  `- Done` / `- DEFERRED, to be done later`.
- `-WithDelegate` is opt-in per launch: the group tools don't accept `-Delegate`
  and an unknown parameter stops a tool starting.

### M365-Group-Compare.ps1

Parameters: `-ThemeMode`, `-ManageUser`, `-AutoConnect`. Direct, static Entra
security-group compare/add; Manage User Groups removes; Export Security Groups.
Graph (delegated, Graph PowerShell app): `User.Read.All`, `Group.Read.All`,
`GroupMember.ReadWrite.All`. Role: Groups Administrator or User Administrator, or
group owner. Role-assignable groups are labelled only (backlog).

### M365-Distribution-Group-Compare.ps1

Parameters: `-ThemeMode`, `-ManageUser`, `-AutoConnect`. Exchange Online
distribution group compare/add/remove; Export Distribution Groups. Module
`ExchangeOnlineManagement`. Role: Recipient Management or Exchange Administrator.

### M365-Conditional-Access-User-Manager.ps1

Parameter: `-ThemeMode`. Shows a user's targeting across CA policies; adds or
removes **direct** include/exclude only; group and All Users assignments are never
changed. Export Conditional Access. Graph: `Policy.Read.All`,
`Policy.ReadWrite.ConditionalAccess`, `Application.Read.All`, `User.Read.All`,
`Group.Read.All`, `Directory.Read.All`. Role: Conditional Access Administrator or
Security Administrator.

### M365-Exchange-Online-Tools.ps1

Parameters: `-ThemeMode`, `-ManageUser`, `-Delegate`, `-AutoConnect`. Tabs: Inbox
Rules, Message Trace, Group Delivery, Mailbox Report, Quarantine, Header Analyser,
Mail Flow Rules, Folder Permissions, **Delegation**.

Delegation: lists Full Access / Send As / Send on Behalf (inherited and
`NT AUTHORITY\SELF` excluded); Grant Full Access (AutoMapping default on), Grant
Send As, Remove Selected, Convert to Shared. Convert states the licence outcome
first (Plan 2 needed if > 50 GB, archive, or litigation hold; size parsed from the
byte count in `TotalItemSize`) and reads the type back. With `-ManageUser` it
opens on Delegation with the mailbox loaded. Role: Exchange Administrator (or
Recipient Management for delegation; a quarantine-capable role for release).

### M365-OneDrive-SharePoint-Tools.ps1

Parameters: `-ThemeMode`, `-AdminUrl`, `-ManageUser`, `-Delegate`, `-AutoConnect`.
OneDrive access grant/remove/audit, offboarding handover, deleted-OneDrive
restore, storage/sharing report. Module `Microsoft.Online.SharePoint.PowerShell`;
role SharePoint Administrator.

- An access list that can't be read (no site-collection admin) is a **gap, not a
  failure**: placeholder row, orange summary, Grant still usable. Other errors
  still raise.
- Picker lists OneDrive owners only; any account can be typed.
- With `-ManageUser`: opens Offboarding, fills owner and manager, builds the
  report. **Grant Handover is never automatic.**

### M365-Teams-Block-Number.ps1 - "Block and Forward Phone Numbers"

Filename unchanged; window title changed; console button **Block/Forward Phone**.
Parameters: `-ThemeMode`, `-ExpectedTenantDomain`, `-TenantId`, `-ShowConsole`,
`-ManageUser`, `-Delegate`, `-AutoConnect`.

- Tenant-wide inbound blocked-number patterns. Changes need the connected tenant
  to list `ExpectedTenantDomain` as verified (and match `TenantId` if set).
- **Forward a User's Calls:** Load, then Forward Calls / Stop Forwarding, enabled
  only for the same loaded user. Writes the whole forwarding group together
  (`IsForwardingEnabled`, `Immediate`, `SingleTarget`, target); user targets by SIP
  address via `Get-CsOnlineUser`, numbers in E.164; reads back with
  `Get-CsUserCallingSettings`. Needs MicrosoftTeams **4.0.0+** and the user's Phone
  System licence; outside numbers need PSTN calling.
- Role: Teams Administrator or Teams Communications Administrator.

### IntuneTools.ps1

Parameters: `-ThemeMode`, `-OrgTenantId`, `-OrgClientId`, `-ManageUser`,
`-AutoConnect`, `-Delegate` (accepted and ignored). Tabs: Packaging, Devices,
Autopilot (incl. Groups menu), Apps, Policies, Group Lookup (Add/Remove Members),
Reports, Settings. Signs in with the tenant's own app registration, delegated
`.default`.

- **Offboarding user mode** (`-ManageUser`): strip under the top bar with Revoke
  Sign-in Sessions (`POST /users/{id}/revokeSignInSessions`), App Selective Wipe
  (registrations from `managedAppRegistrations` filtered client-side by userId,
  grouped by deviceTag, wiped via beta `wipeManagedAppRegistrationByDeviceTag`),
  Retire (refuses devices owned by anyone else). A 403 names the missing permission.
- **Set Primary User** refuses non-Windows and non-Entra/hybrid-joined devices
  before asking for a user.
- Permissions (delegated, admin-consented; sign out and back in after adding):

| Permission | For |
| --- | --- |
| DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All, DeviceManagementApps.Read.All, DeviceManagementServiceConfig.Read.All, DeviceManagementScripts.Read.All, Group.Read.All, User.Read.All, Directory.Read.All, offline_access | Read |
| DeviceManagementManagedDevices.ReadWrite.All | Sync, restart, rename, delete record |
| DeviceManagementManagedDevices.PrivilegedOperations.All | Retire, wipe, Autopilot Reset, Fresh Start |
| DeviceManagementServiceConfig.ReadWrite.All | Autopilot group tag, assign user |
| DeviceManagementScripts.ReadWrite.All, DeviceManagementApps.ReadWrite.All, DeviceManagementConfiguration.ReadWrite.All | Script, app and policy edits; App Selective Wipe |
| GroupMember.ReadWrite.All | Group membership changes |
| User.RevokeSessions.All | Revoke Sign-in Sessions |
| BitlockerKey.Read.All, DeviceLocalCredential.Read.All | Optional: BitLocker keys, LAPS |

### Packaging

- `InstallMSToolkit.ps1` installs to `%ProgramFiles(x86)%\MSToolkit`, resets that
  folder's ACL to inherited, creates `C:\ProgramData\MSToolkit` (SYSTEM/Admins
  full, Users RX) with only `Logs` writable by Users, removes old staged scripts,
  and adds an Apps & Features entry.
- `UninstallMSToolkit.ps1` self-relocates to TEMP, elevates, removes everything,
  removes the Apps & Features entry last.
- `MSToolkit.EXE` is compiled by `Build\Build-All.ps1` from
  `Build\Launchers\MSToolkitLauncher.cs` and embeds nothing: it runs
  `Launch-MSToolkit.bat` from its own folder in a visible console. It replaced
  an IExpress EXE that embedded `Launcher.ps1` and so ran a stale copy after
  every update. `Launcher.ps1` is still packaged but no longer used.
- **Version:** `VERSION` in the repo root is the only place it is set. The build
  stamps it into `MSToolkit.EXE`; the installer reads the EXE's product version
  and writes `DisplayVersion` to the Apps & Features key (removed with the key on
  uninstall). Intune detects on that value (README 4.5), so an older install -
  including any from before 1.1.0, which has no `DisplayVersion` - is upgraded.
- `Build-Install-Uninstall-Only.cmd` builds `MSToolkit-Install.exe` /
  `MSToolkit-Uninstall.exe` (7-Zip SFX, running `install.cmd` / `uninstall.cmd`).
  Its validation and archive lists must match the installer's `$PayloadFiles`.
- Filenames are canonical - no version or date suffixes.
- Intune: install `install.cmd` / uninstall `uninstall.cmd` (as packaged), System
  context. Upgrades need detection that fails while `.ps1` files remain in
  `C:\ProgramData\MSToolkit`, or old installs keep the writable layout; the
  `DisplayVersion` rule does that, because those old installs have no version.

---

## 5. UI rules

### No automatic domain prefix

No launcher or sign-in box adds, strips or changes a domain. What is typed - domain
included - is used and remembered exactly.

| Right | Wrong |
| --- | --- |
| Prompt: "Enter as DOMAIN\username or username@contoso.com" | Prompt: "CONTOSO\ is added automatically" |
| `runas /user:%ADMINUSER%` with the value as typed | `runas /user:%USERDOMAIN%\%ADMINUSER%` |
| No domain typed -> refuse and ask again (runas would sign in to a local account) | Silently prepend `%USERDOMAIN%` or a stored NetBIOS name |
| Remember `CONTOSO\admin.jdoe` as entered | Strip `CONTOSO\` before saving |

### Generic, specific UI hints

Hints must not reveal a naming convention, but must say exactly what value is
expected and which attribute or source it maps to.

| Right | Wrong |
| --- | --- |
| "Primary SMTP domain, for example contoso.com - the part after @ in the mail attribute." | "Used for first.last@ addresses." |
| "Written exactly as typed to the Company attribute (company)." | "Our company name." |
| Example user: `Jane Doe (jdoe)`, `jane.doe@contoso.com` | Any real person or address |
| Example phone: `202-555-0100` (fiction range) | Any real number |

Before publishing, search every Settings hint, tooltip and example for anything
that describes a naming pattern or a real value.

### Scrolling containers that hold clickable controls

A stock AutoScroll panel or form scrolls the clicked control into view, so a
click near an edge moves the page under the pointer. Offboard User's step list
(`MSToolkitScroll.StablePanel`) and the Create New User form
(`MSToolkitScroll.StableForm`) use subclasses that: return
`DisplayRectangle.Location` from `ScrollToControl` while a mouse button is down
(Tab and the keyboard still scroll a focused control into view); and save
`AutoScrollPosition` on every child's MouseDown and put it back on MouseUp if it
moved, at once and again via `BeginInvoke`. `OnControlAdded` hooks controls added
later. The override alone was not enough: clicks still moved the page. Use these
for any new long scrolling window.

### Waiting loops inside a Shown handler

A `while (...) { DoEvents }` wait inside a Shown handler blocks the close: Windows
only completes it once the handler returns. Set a flag in FormClosing and leave
the loop on it (IntuneTools: `$script:CloseRequested` / `$script:SignInCancelled`
end the browser sign-in wait; the listener closes in `finally`). Likewise a timer
fires inside any `DoEvents`: anything a timer could start must check a busy flag
set by long actions that pump messages.

---

## 6. Known traps

1. **Log level must match the logger's own `ValidateSet`.** A mismatch throws
   *after* the UI has updated, so a working feature shows an error. Current sets:
   - `WARN`: M365-Group-Compare, M365-Distribution-Group-Compare,
     M365-Exchange-Online-Tools, M365-OneDrive-SharePoint-Tools
   - `WARNING`: Compare-UserGroups, Investigate-AccountLockout,
     M365-Conditional-Access-User-Manager, OffboardUser
   - Teams `Write-Log`: `Info, Success, Warning, Error, Action`
   Audit every call against its own file's set; code copied between tools is the
   usual source.
2. **`-f` inside a method call** swallows the following arguments:
   `[MessageBox]::Show(("...{0}") -f $a, 'Title', ...)` formats the title in.
   Build the string into a variable first.
3. **Body-less Graph POST** - `Invoke-RestMethod` sends it as a web form and Graph
   rejects it. IntuneTools' `Invoke-Graph` now sends `'{}'` for POST/PATCH/PUT with
   no body; keep it that way.
4. **`Add_Shown` placement.** Start-up behaviour gets its own top-level
   `Add_Shown`, with any condition inside it. Combo-filter registration was once
   folded into a conditional handler and never ran on a normal launch. Verify the
   call site (column 0, own handler, walker defined earlier, state hashtables
   assigned at top level), not just presence.
5. **Picker (combo filter):** `AutoCompleteMode/Source = None`; never
   `DroppedDown = $true` from a text handler; filter on `TextUpdate` not
   `TextChanged`; own ListBox under the combo; close the real dropdown on
   `KeyDown` for printable keys (setting the skip-hide flag first) or the first
   keystroke is lost.
6. **Generic `List` in `@()`** throws "Argument types do not match" in PS 5.1. Use
   `.ToArray()` or walk by index.
7. **WinForms swallows handler exceptions** - wrap clicks, log type/message/line,
   and keep the busy-state function incapable of throwing.
8. **Layout:** Top-docked controls follow z-order (`SetChildIndex`); measure
   wrapped text against a narrower width than the label and round up; check new
   button labels against the real width in a screenshot (the longest sidebar label
   wrapped); controls laid out from a running offset must be created in display
   order.
9. **Radio buttons as sub-steps** need `AutoCheck = $false`; Done/Deferred pairs
   need their own panel.
10. **Check what a helper already returns** - the phone parser's `E164` already
    includes `+`.
11. **Elevated tools can't use the desktop shell.** `Start-Process <url>` silently
    does nothing from the admin account: resolve the default browser from the
    registry and start its EXE; open documents as the signed-in user.
12. **The admin account can't read the standard user's profile** (including
    OneDrive). A test copy kept there needs an explicit read grant for the admin
    account - never `Users`.
13. **CRLF vs LF** - files differ; match each file's existing line endings when
    editing.
14. **AD module start-up:** set `$env:ADPS_LoadDefaultDrive = 0` before
    `Import-Module ActiveDirectory`, and don't rely on automatic DC discovery at
    start-up.
15. **`Get-ADReplicationFailure` keeps old errors** after recovery; only
    `FailureCount` resets.

---

## 7. Port status

### Ported from the internal version

- Offboard User tool, generalised to Settings.
- Fixes round: filtered picker in all 8 tools with typeable user combos; click
  wrapper and safe busy state; read-back verification; clear-before-load, match
  confirmation, action-time re-check, LOADED line; persistent Undo (with
  MSToolkit's extra safeguards); input box sized from measured text; Add Phone
  Number; Unforce Password Change; Office copied in NewADUser address clone; export
  buttons and `-ManageUser` / `-AutoConnect` on Group / Distro Compare; CA export.
- OneDrive access-list-not-readable handling and three hint lines.
- IntuneTools Autopilot Groups menu and Group Lookup Add/Remove Members.
- Exchange Online and OneDrive tools, blue Connect buttons, unified theme toggle.

### Offboarding round 2 - implemented from a written brief, not ported code

Offboard User 18-step order, Done/Deferred, red DISABLED, manager UPN /
`-WithDelegate` / `-ReuseWindow`; Exchange Delegation tab; OneDrive auto-open;
call forwarding; IntuneTools user mode, body-less POST fix, Set Primary User
check; DC Health recovered partners; Unforce "never expires" note; log-level audit
across all files. **All untested** (section 8).

### Pending

- Rebuild `MSToolkit-Install.exe`, `MSToolkit-Uninstall.exe` and the `.intunewin`
  after the latest changes.
- Code signing.
- Test on a clean PC with no previous install.
- A later ticketing/inventory tool in the internal version is **not ported** -
  don't start unless asked.

---

## 8. Untested, and deferred backlog

### Untested - written but never run against a real tenant/domain

- Everything in round 2 (section 7).
- Exchange Convert to Shared licence verdict on a real mailbox over 50 GB / with
  archive / on hold.
- Call forwarding to a user and to an outside number.
- App Selective Wipe on an unenrolled phone; Retire owner guard.
- DC Health "Recovered" after a real partner reboot.
- Two standard users writing the same daily log in `Logs`.
- Intune upgrade from an old (writable-runtime) install via the detection script.

### Deferred backlog - **don't start unless asked**

- NewADUser: optionally create-disabled -> configure -> verify -> enable.
- M365-Group-Compare: hard-block add/remove on `isAssignableToRole` groups
  (currently labelled only).
- Remove User From All Sessions: optional file-server setting for SMB sessions.

---

## 9. Things not to trust

- **Anything that ran from `C:\ProgramData\MSToolkit` outside `Logs`.** Older
  versions staged scripts there with `Users:Modify`; any standard user could have
  edited them, and they ran as the admin account. Treat those copies as
  untrusted; the installer deletes them.
- **Old EXEs and `.intunewin` packages** built before the runtime-folder fix or
  before the latest changes. Rebuild from the repository.
- **`MSToolkit.lnk` in old packages** carries the SID of the account that created
  it. It is no longer shipped or committed (the installer creates the shortcut);
  don't publish one made on a domain account.
- **Copies in someone's profile, OneDrive or a share** - may be stale and are
  user-writable. Edit in the repository, install from a fresh build.
- **Memory of earlier versions.** The repository is authoritative; uploaded or
  committed files beat any description of them, including this one.
- **Any file containing an organisation's name, domain, tenant/client ID, host,
  OU path, person or naming convention.** Remove it before it reaches the public
  repository; check history as well as the current tree.
