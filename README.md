# MSToolkit

MSToolkit is a Windows desktop toolkit for day-to-day Active Directory and Microsoft 365 administration. A main console handles common AD tasks (users, groups, computers, OUs, replication, reports) and launches a set of companion tools for account creation, group comparison, lockout investigation, Microsoft 365 groups, Conditional Access, Teams call blocking and Intune.

Nothing organization-specific is built in. Domains, OUs, tenant details and naming values are entered once in **Settings** and shared by every tool.

---

## Contents

1. [What's in the package](#1-whats-in-the-package)
2. [Requirements](#2-requirements)
3. [Installation](#3-installation)
4. [Deploying with Microsoft Intune](#4-deploying-with-microsoft-intune)
5. [Deploying with NinjaOne](#5-deploying-with-ninjaone)
6. [Uninstalling](#6-uninstalling)
7. [Starting MSToolkit](#7-starting-mstoolkit)
8. [First-run setup: Settings](#8-first-run-setup-settings)
9. [The main console](#9-the-main-console)
10. [Companion tools](#10-companion-tools)
11. [IntuneTools and its app registration](#11-intunetools-and-its-app-registration)
12. [Where MSToolkit keeps its files](#12-where-mstoolkit-keeps-its-files)
13. [Troubleshooting](#13-troubleshooting)
14. [Security notes](#14-security-notes)

---

## 1. What's in the package

### Files you receive

| File | Purpose |
| --- | --- |
| `MSToolkit-Install.exe` | Installs or updates MSToolkit. Self-contained: everything the install needs is inside it. |
| `MSToolkit-Uninstall.exe` | Removes MSToolkit. Optional - the installer also adds an **MSToolkit** entry to Apps & Features that does the same job. |
| `install.intunewin` | Ready-made Microsoft Intune package containing the full set of install files. Optional - only needed to deploy through Intune ([section 4](#4-deploying-with-microsoft-intune)). |
| `README.md` | This document. |

### What the installer puts on the PC

Everything below is installed to `C:\Program Files (x86)\MSToolkit`. The files MSToolkit runs from are also copied to `C:\ProgramData\MSToolkit`.

| File | What it is |
| --- | --- |
| `MSToolkit.EXE` | Starts MSToolkit (the desktop shortcut points here). It refreshes the runtime copy of the scripts, then runs the launcher. |
| `MSToolkit.ps1` | The main console |
| `Launch-MSToolkit.bat` | Asks for the admin account and starts the console |
| `NewADUser.ps1` / `Launch-NewADUser.bat` | Create New AD User, and its standalone launcher |
| `Compare-UserGroups.ps1` | Compare and manage AD group memberships |
| `Investigate-AccountLockout.ps1` | Account lockout investigation |
| `M365-Group-Compare.ps1` | Microsoft 365 security group compare and add |
| `M365-Distribution-Group-Compare.ps1` | Exchange Online distribution group compare and add |
| `M365-Conditional-Access-User-Manager.ps1` | Conditional Access user include/exclude manager |
| `M365-Teams-Block-Number.ps1` / `Launch-M365-Teams-Block-Number.cmd` | Teams inbound number blocking, and its standalone launcher |
| `IntuneTools.ps1` | Intune and Entra device, app and policy management |
| `Launcher.ps1` | Used by `MSToolkit.EXE` to refresh the runtime copy |
| `MSToolkit.ico` / `MSToolkit.lnk` | Icon and shortcut |
| `InstallMSToolkit.ps1` / `Launch-InstallMSToolkit.bat` | Copies of the installer script and its launcher. Not used day to day; to reinstall or repair, run `MSToolkit-Install.exe` again. |
| `UninstallMSToolkit.ps1` / `Launch-UninstallMSToolkit.bat` | The uninstaller script used by the Apps & Features entry; the `.bat` runs it directly |

---

## 2. Requirements

### The PC

| Requirement | Notes |
| --- | --- |
| Windows 10 or Windows 11, 64-bit, Pro / Enterprise / Education | RSAT is not available on Home editions. |
| Joined to an Active Directory domain | The console discovers the domain and its domain controllers from the PC's own membership. |
| Windows PowerShell 5.1 | Built into Windows 10 and 11. |
| Local administrator rights to install | The installer writes to Program Files, ProgramData and HKLM. |
| Access to Windows Update, or a Features on Demand source | The installer adds four RSAT components if they are missing (below). If the PC gets updates from WSUS without Features on Demand content, this step can fail - see [Troubleshooting](#13-troubleshooting). |

RSAT components the installer adds when missing:

- Active Directory Domain Services and Lightweight Directory Services Tools
- Group Policy Management Tools
- Print Management Console
- DNS Server Tools

### Network

| From the PC to | Port / service | Used by |
| --- | --- | --- |
| Domain controllers | LDAP 389, Active Directory Web Services 9389 | Everything AD |
| Domain controllers | Remote Event Log (RPC) | Investigate Account Lockout |
| The Entra Connect server | PowerShell remoting (WinRM 5985) | Delta Sync |
| Microsoft 365 (`login.microsoftonline.com`, `graph.microsoft.com`, Exchange Online, Teams) | HTTPS 443 | The Microsoft 365 tools and IntuneTools |
| PowerShell Gallery (`www.powershellgallery.com`) | HTTPS 443 | One-time install of the Microsoft 365 PowerShell modules |

### Accounts

MSToolkit deliberately uses two accounts:

| Account | Used for |
| --- | --- |
| **An AD admin account** (for example a separate domain admin account) | The main console and the AD tools. You enter it when MSToolkit starts. |
| **Your normal Windows account** - the one signed in to the PC | The Microsoft 365 tools, IntuneTools and opening log files. Admin accounts are often not licensed or allowed for Microsoft 365, so these tools run as you. |

Microsoft 365 roles needed per tool are listed in [Companion tools](#10-companion-tools).

---

## 3. Installation

### Install on one PC

1. Copy `MSToolkit-Install.exe` to the PC.
2. Run it and approve the administrator (UAC) prompt.
3. Wait for it to finish. The RSAT step can take several minutes the first time.

The installer asks no questions. Running it again later updates an existing install in place; settings are kept.

### What the installer creates

| Item | Location |
| --- | --- |
| Program files | `C:\Program Files (x86)\MSToolkit` |
| Runtime files (what actually runs) | `C:\ProgramData\MSToolkit` - local Users are given Modify rights so launches don't need elevation |
| Desktop shortcut for all users | `C:\Users\Public\Desktop\MSToolkit.lnk` |
| Apps & Features entry | **MSToolkit** (registry: `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSToolkit`) |

### If Windows warns about the file

The installer and scripts are not code-signed, so SmartScreen or antivirus may warn about or block them. If you distribute MSToolkit widely, sign the files with your organization's code-signing certificate, or allow them through your security tooling.

---

## 4. Deploying with Microsoft Intune

MSToolkit deploys as a **Windows app (Win32)**. Use the included `install.intunewin`, add it to Intune and assign it.

### 4.1 The .intunewin file

The included `install.intunewin` is ready to upload - skip to [4.2](#42-add-the-app). Its setup file is `install.cmd`, and it contains the complete set of install files, including `install.cmd` and `uninstall.cmd`.

Only if you need to build your own package instead (for example after receiving updated EXEs), wrap the two EXEs:

1. Download the **Microsoft Win32 Content Prep Tool** (`IntuneWinAppUtil.exe`) from Microsoft.
2. Make a source folder that contains **only**:
   - `MSToolkit-Install.exe`
   - `MSToolkit-Uninstall.exe`
3. Run:

   ```
   IntuneWinAppUtil.exe -c "C:\Packages\MSToolkit\Source" -s MSToolkit-Install.exe -o "C:\Packages\MSToolkit\Output" -q
   ```

   This creates `MSToolkit-Install.intunewin` in the output folder. For a package built this way, use `MSToolkit-Install.exe` as the install command and `MSToolkit-Uninstall.exe` as the uninstall command in [4.3](#43-program).

### 4.2 Add the app

1. Sign in to the **Microsoft Intune admin center**.
2. Go to **Apps** > **All Apps** > **Create** (shown as **Add** in some tenants).
3. Platform **Windows**, app type **Windows app (Win32)**, then **Select**.
4. **App information**: select `install.intunewin`, then fill in:
   - **Name**: `MSToolkit`
   - **Publisher**: your organization or team
   - Description, category and logo as you like

### 4.3 Program

| Field | Value |
| --- | --- |
| Installer type | Command line |
| **Install command** | `install.cmd` |
| **Uninstall command** | `uninstall.cmd` |
| Installation time required | `60` (minutes). The RSAT step can be slow on a first install. |
| Allow available uninstall | Your choice |
| **Install behavior** | **System** |
| Device restart behavior | Determine behavior based on return codes |
| Return codes | Leave the defaults |

Both commands run silently and switch to 64-bit PowerShell themselves, so no extra switches or wrapper scripts are needed. They run from Intune's copy of the package, so no paths on the PC are involved - Intune's uninstall command doesn't expand environment variables.

### 4.4 Requirements

| Field | Value |
| --- | --- |
| Operating system architecture | 64-bit (x64) |
| Minimum operating system | The oldest Windows 10 or 11 release you support |

### 4.5 Detection rule

Choose **Manually configure detection rules** > **Add** and use this rule:

| Field | Value |
| --- | --- |
| Rule type | **File** |
| **Path** | `C:\Program Files (x86)\MSToolkit` |
| **File or folder** | `MSToolkit.EXE` |
| Detection method | **File or folder exists** |
| Associated with a 32-bit app on 64-bit clients | **No** |

Registry alternative (use one rule or the other, not both):

| Field | Value |
| --- | --- |
| Rule type | Registry |
| Key path | `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\MSToolkit` |
| Value name | *(leave empty)* |
| Detection method | Key exists |
| Associated with a 32-bit app on 64-bit clients | **No** |

### 4.6 Assign and monitor

1. Skip **Dependencies** and **Supersedence** unless you need them.
2. **Assignments**: assign as **Required** to a group of admin devices, or **Available for enrolled devices** so admins install it from the Company Portal.
3. **Review + create**.
4. Check progress under the app's **Device install status**. The detection rule is what confirms success.

To update MSToolkit later, upload a new `.intunewin` into the same app (**Properties** > **App information** > **Edit**). The installer updates an existing install in place.

---

## 5. Deploying with NinjaOne

### 5.1 Add the installer to the Automation Library

1. In NinjaOne go to **Administration** > **Library** > **Automation**.
2. Click **Add automation** (or **+ Add**) and choose **Installation**.
3. In the **Install application** dialog:

   | Field | Value |
   | --- | --- |
   | Name | `MSToolkit - Install` |
   | Description | Optional |
   | Operating system | Windows |
   | Architecture | All (or 64-bit) |
   | Installer | Upload `MSToolkit-Install.exe` |
   | Run As | **System** |
   | Preset parameters | Leave empty - the installer is silent |

4. Click **Submit**. The automation goes to **Under Review** while NinjaOne scans it, then appears under **Automations**.

### 5.2 Install on devices

- **One-off**: select the device(s) > **Run** > **Run Automation** > **Install Application** > **MSToolkit - Install** > Run As **System**.
- **By policy**: open the policy > **Scheduled Automations** > add **MSToolkit - Install** with a **Run once** schedule, Run As **System**.

### 5.3 Uninstalling through NinjaOne (optional)

Uploading the uninstaller isn't required, because MSToolkit adds its own Apps & Features entry and that entry runs silently when started by a management tool as SYSTEM. If you want a one-click uninstall automation anyway:

1. **Administration** > **Library** > **Automation** > **Add automation** > **Run**.
2. Name `MSToolkit - Uninstall`, operating system Windows, upload `MSToolkit-Uninstall.exe`, Run As **System**, no parameters.
3. Run it the same way as the install.

---

## 6. Uninstalling

Any of these removes the program folder, the runtime folder, the desktop shortcut and the Apps & Features entry:

| Method | Notes |
| --- | --- |
| **Settings** > **Apps** > **Installed apps** > **MSToolkit** > **Uninstall** | Asks for administrator approval, then shows a result message |
| Run `MSToolkit-Uninstall.exe` | Approve the administrator prompt |
| Intune | Assign the app as **Uninstall** |
| NinjaOne | See [5.3](#53-uninstalling-through-ninjaone-optional) |

Close MSToolkit windows first; the uninstaller closes any it finds, but only ones started from MSToolkit's own folders.

**Kept after uninstalling:** each user's settings in `%APPDATA%\MSToolkit` (see [section 12](#12-where-mstoolkit-keeps-its-files)). To clear them, use **Clear Settings** in MSToolkit before uninstalling, or run this in PowerShell as each account that used MSToolkit:

```powershell
Remove-Item -LiteralPath (Join-Path $env:APPDATA 'MSToolkit') -Recurse -Force
```

---

## 7. Starting MSToolkit

1. Open the **MSToolkit** desktop shortcut.
2. A command window asks for the admin account. Enter it **with the domain**:
   - `DOMAIN\username`, or
   - `username@domain.com`

   A name without a domain is refused, because Windows would otherwise look for a local account.
3. For a new account you're asked whether to remember it. The answer is saved only once the sign-in works; the password is never saved - Windows asks for it every time.
4. Windows asks for the password, then the console opens.

Next time, press **Enter** at the prompt to reuse the remembered account, or type a different one.

Two companion tools also have their own launchers in `C:\ProgramData\MSToolkit`, for use without the console:

- `Launch-NewADUser.bat` - Create New AD User
- `Launch-M365-Teams-Block-Number.cmd` - Teams call blocking (runs as your normal Windows account)

---

## 8. First-run setup: Settings

The first time an account opens MSToolkit, **Settings** opens by itself, and keeps doing so on each launch until at least one value is saved. After that, open it with the **Settings** button on the top bar.

**Everything is optional.** Blank OU settings fall back to the domain root, and the tools say so clearly in orange whenever that happens. Fill in what applies to your environment; everything else can wait.

Each field has a hint underneath it describing exactly what it expects and which tool uses it. Summary:

### Directory

| Setting | What to enter | If blank |
| --- | --- | --- |
| DNS domain | The AD domain's DNS name, e.g. `contoso.local` | The domain this PC is joined to |
| Short domain name | The pre-Windows 2000 name, e.g. `CONTOSO` | Not currently used by any tool |
| Entra Connect server | Name of the server running Microsoft Entra Connect Sync | **Delta Sync** only reports that no server is set |

### Organizational units

Enter each as a distinguished name (`OU=...,DC=...`), or click **Browse...** to pick it from the directory.

| Setting | Used by | If blank |
| --- | --- | --- |
| Users | Get Employee OUs; the default OU in Create New AD User; the user lists in Compare User Groups and Investigate Account Lockout | Domain root. Create New AD User then requires you to pick an OU - it never creates accounts in the domain root. |
| Admin accounts | Added to the user lists in Compare User Groups and Investigate Account Lockout when admin accounts sit outside the Users OU | Nothing is added |
| Computers | Get Computer OUs (all levels under it) | Top-level OUs of the domain root |
| Servers | Get Server OUs | Domain root |
| Disabled accounts | Get Disabled OUs | Domain root |
| Security groups | Get Security Groups | Every security group in the domain |
| Distribution groups | Get Distribution Groups | Every distribution group in the domain |
| Service accounts | Get Special Function Accounts | **Required** - that button asks for it |

When some OUs are blank, the console lists them once at startup.

### New account naming

Used by **Create New AD User**.

| Setting | What to enter | If blank |
| --- | --- | --- |
| UPN domain | The UPN suffix for new accounts, e.g. `contoso.com`. Must be a UPN suffix in AD and a verified domain in Microsoft 365. | **Required** - accounts can't be created |
| Primary SMTP domain | Domain for the `mail` attribute and primary SMTP address, e.g. `contoso.com` | **Required** - accounts can't be created |
| onmicrosoft domain | The tenant's initial domain, e.g. `contoso.onmicrosoft.com` | That proxy address is skipped |
| mail.onmicrosoft domain | The Exchange Online routing domain, e.g. `contoso.mail.onmicrosoft.com` | That proxy address is skipped |
| Company | Text written to the `company` attribute, exactly as typed | Company is left blank |

### Microsoft 365 and Intune

| Setting | What to enter | If blank |
| --- | --- | --- |
| Tenant ID | Directory (tenant) ID of your Entra tenant | IntuneTools' **Set to Org Defaults** has nothing to copy; Teams call blocking skips its tenant ID check |
| Client ID | Application (client) ID of the IntuneTools app registration ([section 11](#11-intunetools-and-its-app-registration)) | **Set to Org Defaults** has nothing to copy - enter it in IntuneTools directly instead |
| Expected tenant domain | Any verified domain of your tenant, e.g. `contoso.com` | Teams call blocking stays read-only |

### Buttons

| Button | What it does |
| --- | --- |
| Save | Saves and closes |
| Cancel | Closes without saving |
| Clear Settings | Blanks and saves every value after an "Are you sure?" prompt. Two optional tick boxes also forget the saved Microsoft 365 sign-in password and the admin username remembered by the launcher. The theme and selected domain controller are kept. |

Settings are saved per Windows account in `%APPDATA%\MSToolkit\settings.json`. Run MSToolkit as the same admin account each time, or its settings will look empty.

---

## 9. The main console

### Top bar

| Control | What it does |
| --- | --- |
| **AD Server** | The domain controller every AD action uses, chosen from the DCs AD reports. The status bar shows whether it answers on port 9389. |
| REPL Selected DC / REPL All DCs | Forces AD replication with `repadmin` |
| Delta Sync | Runs an Entra Connect delta sync on the server set in Settings |
| DC Health | Checks every domain controller |
| ADAC, ADUC, DNS, GPO, Print MGMT, Computer MGMT | Open the standard Windows consoles; ADUC and Computer Management target the selected DC |
| PowerShell, ISE | Open a PowerShell window or ISE in the MSToolkit folder, as the admin account |
| RDP | Remote Desktop to the selected DC |
| Logs | Opens the logs and reports folder |
| Settings | See [section 8](#8-first-run-setup-settings) |
| Sun / moon | Light or dark theme, remembered per account |

### Sections

| Section | Buttons |
| --- | --- |
| Users | Create New User, Get User, Get User Groups, Get User OU, Unlock User, Reset Password, Force Password Change, Enable User, Disable User, Investigate Lockout, Delete User |
| Service Accounts | Get Managed Service Accounts, Get Special Function Accounts |
| Groups | Get Group, Get Group Members, Get Security Groups, Get Distribution Groups, Add User to Group, Remove User from Group, Create Group, Compare/Manage User Groups, Delete Group |
| M365 | M365 Group Compare/Add, M365 Distro Compare/Add, M365 Conditional Access, Block Teams Numbers, Intune Tools |
| Computers | Get Computer, Get Computer OU, Enable Computer, Disable Computer, Delete Computer, Reset Computer Account |
| OUs | Get Top-Level OUs, Get Employee / Computer / Server / Disabled OUs, Get All Common OUs, Create OU, Move Object to OU |
| Reports | 90-Day Inactive Users, 90-Day Inactive Computers, Password Expiring Soon |

Results appear in **Activity Output** (use **Copy Output** / **Clear Output**). Reports are also saved as CSV in the logs folder (`C:\ProgramData\MSToolkit\Logs`).

### Built-in protections

Some changes are refused outright for critical directory objects, identified by their well-known SID/RID or by AD's own critical-object flag rather than by name. These include the built-in Administrator, Guest and KRBTGT accounts, domain controllers, and built-in privileged groups such as Domain Admins, Enterprise Admins and Schema Admins. Objects protected from accidental deletion need an explicit, per-deletion confirmation. Use ADUC or your normal change process when such a change really is intended.

### Signing in for the Microsoft 365 tools

The **M365** buttons, **Intune Tools** and the **Logs** file actions run as your normal Windows account. The first time, a sign-in box asks for it:

- Enter it as `DOMAIN\username`.
- **Remember password on this computer** stores it encrypted (Windows DPAPI), readable only by the admin account on this PC.
- **Sign in automatically next time** skips the box afterwards. **Hold Shift** while clicking a button to show the box again.

---

## 10. Companion tools

### Create New AD User

Creates an AD user and fills in the Microsoft 365-related attributes.

- **Needs:** UPN domain and Primary SMTP domain in Settings.
- **Username:** generated as first name + last initial (editable before creating). Email: `firstname.lastname@<Primary SMTP domain>`.
- **Sets:** UPN, `mail`, primary and secondary proxy addresses (including the onmicrosoft addresses if set), Company, Title, Department, Description; the password must be changed at next logon.
- **Optional copying:** group memberships, address and manager from existing users.
- **OU:** pick it in the Department OU picker (starts at the Users OU), type a child OU name or path, or paste a full DN. Accounts are never created in the domain root.
- A confirmation screen shows everything before the account is created, and the result is verified afterwards.

### Compare User Groups (AD)

Compares the direct AD group memberships of a reference user and a target user, and adds selected missing groups to the target. **Manage Groups** loads one user to review all groups and add or remove memberships. Nested memberships can be shown, but changes are always made as direct memberships.

### Investigate Account Lockout

Searches every domain controller for lockout events (4740) and, optionally, failed sign-ins (4625) for a user, and shows bad-password counts per DC and the lockout source.

- **Needs:** permission to read the Security log on the DCs (for example Event Log Readers or domain admin), and auditing of account lockouts and logon failures enabled on the DCs.

### M365 Group Compare/Add

Compares direct, static Entra security-group memberships of two users and adds missing groups; **Manage User Groups** removes memberships.

- **Module:** `Microsoft.Graph.Authentication` - the tool offers to install it for your account.
- **Sign-in permissions (delegated):** User.Read.All, Group.Read.All, GroupMember.ReadWrite.All. The first sign-in shows a consent prompt for Microsoft's Graph command-line app; depending on your tenant's consent settings an administrator may need to approve it.
- **Role:** Groups Administrator or User Administrator, or ownership of the groups being changed.

### M365 Distro Compare/Add

Compares Exchange Online distribution-group memberships of two users and adds or removes memberships.

- **Module:** `ExchangeOnlineManagement` - the tool offers to install it for your account.
- **Role:** an Exchange role that can manage distribution groups, e.g. Recipient Management or Exchange Administrator.

### M365 Conditional Access

Shows how a user is targeted across Conditional Access policies and adds or removes the user as a **direct** include or exclude. Group-based and All Users assignments are never changed.

- **Module:** `Microsoft.Graph.Authentication` - install it once for your account if it isn't already there (M365 Group Compare offers to, or run `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`).
- **Sign-in permissions (delegated):** Policy.Read.All, Policy.ReadWrite.ConditionalAccess, Application.Read.All, User.Read.All, Group.Read.All, Directory.Read.All.
- **Role:** Conditional Access Administrator or Security Administrator.

### Block Teams Numbers

Adds and removes tenant-wide inbound blocked-number patterns in Microsoft Teams, one number at a time, with checks and confirmation before any change.

- **Needs:** **Expected tenant domain** in Settings - without it the tool stays read-only. **Tenant ID** is optional; when set, sign-in is pinned to that tenant and checked.
- **Module:** `MicrosoftTeams` - installed automatically for your account on first use.
- **Role:** Teams Administrator or Teams Communications Administrator (activated, if assigned through PIM).

### Intune Tools

See the next section.

---

## 11. IntuneTools and its app registration

IntuneTools manages Intune and Entra devices, Autopilot, apps, policies, groups and reports, and packages Win32 apps. It signs in through **your own app registration** as you (delegated), so your Intune role still applies and every action is audited under your account.

### 11.1 Create the app registration (one time, per tenant)

1. **Microsoft Entra admin center** > **App registrations** (under **Entra ID**, or **Identity** > **Applications** in older menus) > **New registration**.
2. **Name:** anything recognisable, e.g. `IntuneTools`.
3. **Supported account types:** *Accounts in this organizational directory only* (single tenant).
4. **Redirect URI:** platform **Public client/native (mobile & desktop)**, value `http://localhost`.
   - Use `localhost`, not `127.0.0.1`. Entra ignores the port for localhost, so this one entry covers the random port IntuneTools listens on.
5. **Register**.
6. Open **Authentication** and set **Allow public client flows** to **Yes** (in the newer Authentication view this is on the **Settings** tab). Save.
7. From **Overview**, copy the **Application (client) ID** and **Directory (tenant) ID**.

No client secret or certificate is needed - IntuneTools is a public client and never stores one.

### 11.2 API permissions

**API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**. Start with the read permissions; add write permissions only for the buttons you want to use.

**Read (needed for all lists and reports):**

- DeviceManagementManagedDevices.Read.All
- DeviceManagementConfiguration.Read.All
- DeviceManagementApps.Read.All
- DeviceManagementServiceConfig.Read.All
- DeviceManagementScripts.Read.All
- Group.Read.All
- User.Read.All
- Directory.Read.All
- offline_access

**Write (only for the action buttons you need):**

| Permission | Enables |
| --- | --- |
| DeviceManagementManagedDevices.ReadWrite.All | Sync, restart, shut down, rename, delete record |
| DeviceManagementManagedDevices.PrivilegedOperations.All | Retire, wipe, Autopilot Reset, Fresh Start |
| DeviceManagementServiceConfig.ReadWrite.All | Autopilot group tag, assign user, delete Autopilot record |
| DeviceManagementScripts.ReadWrite.All | Editing platform scripts and remediations |
| DeviceManagementApps.ReadWrite.All | Editing app assignments |
| DeviceManagementConfiguration.ReadWrite.All | Editing policy assignments, importing policies |
| GroupMember.ReadWrite.All | Adding and removing devices and users in Entra groups |

**Optional (for the secrets tabs):**

| Permission | Enables |
| --- | --- |
| BitlockerKey.Read.All | BitLocker recovery keys |
| DeviceLocalCredential.Read.All | Windows LAPS local administrator passwords |

Then click **Grant admin consent for <your tenant>**.

### 11.3 Lock the app down (recommended)

The registration also appears under **Enterprise applications** with the same name. To limit who can use it:

1. **Enterprise applications** > your app > **Properties** > **Assignment required?** = **Yes** > Save.
2. **Users and groups** > **Add user/group** > add the admins (or an admin group) who should use IntuneTools.

Sign-ins through this app are also subject to your Conditional Access policies.

### 11.4 Roles for the people using it

The app's permissions are only a ceiling - each user also needs a matching role:

- **Intune:** Intune Administrator, or a custom Intune role covering the actions they'll use.
- **Group membership changes:** Groups Administrator or User Administrator, or ownership of the group. Intune Administrator alone is not enough.
- **BitLocker keys / LAPS:** a role allowed to read them (for example Cloud Device Administrator for LAPS passwords).

### 11.5 Configure IntuneTools

1. Enter **Tenant ID** and **Client ID** in MSToolkit **Settings** (optional, but it lets every admin fill them in with one click).
2. Open **Intune Tools** from the console. On first use it opens on its **Settings** tab.
3. Click **Set to Org Defaults** to copy Tenant ID and Client ID from MSToolkit Settings - or type them in.
4. Optional: **IntuneWinAppUtil.exe path** (the Microsoft Win32 Content Prep Tool) and **Packaging root folder** for the Packaging tab; **Stale device threshold** (days); **Sign-in method** (Browser, or Device code if the loopback sign-in is blocked).
5. **Save Settings**, then **Connect**.

After the first sign-in, **Connect** reconnects silently from a cached sign-in. **Reconnect** forces the account picker. **Clear Cached Token** forces a full sign-in next time, and **Clear Settings** blanks everything and signs out.

Tabs: Packaging, Devices, Autopilot, Apps, Policies, Groups, Reports, Settings.

---

## 12. Where MSToolkit keeps its files

| Location | Contents | Account |
| --- | --- | --- |
| `C:\Program Files (x86)\MSToolkit` | Installed program files, the uninstaller | All users |
| `C:\ProgramData\MSToolkit` | Runtime copy of the scripts; `Logs` folder with logs and CSV reports | All users |
| `%APPDATA%\MSToolkit\settings.json` | Settings, theme, selected DC | Admin account |
| `%APPDATA%\MSToolkit\m365-credential.xml`, `signin-auto.flag` | Remembered Microsoft 365 sign-in (encrypted) | Admin account |
| `%APPDATA%\MSToolkit\launcher-admin-username.txt` | Admin username remembered by the launcher | Your Windows account |
| `%APPDATA%\MSToolkit\settings.json` | Theme choices of the Microsoft 365 tools | Your Windows account |
| `%APPDATA%\MSToolkit\IntuneTools\config.json`, `token.dat` | IntuneTools settings and cached sign-in (encrypted) | Your Windows account |

---

## 13. Troubleshooting

| Symptom | Cause and fix |
| --- | --- |
| Install fails at an RSAT component | The PC can't reach Windows Update or a Features on Demand source (common with WSUS). Allow Features on Demand from Windows Update, or install the four RSAT components another way first - the installer skips components that are already present. |
| Launcher says the name "has no domain" | Enter the admin account as `DOMAIN\username` or `username@domain.com`. |
| Settings look empty after saving | MSToolkit was started as a different account. Settings are per account; the startup output shows which account and settings file are in use. |
| Status bar shows "ADWS unreachable (port 9389)" | Pick another DC in **AD Server**, or check the firewall between the PC and that DC. |
| A console (ADUC, DNS, GPO, Print MGMT) says it isn't installed | Its RSAT component is missing - rerun the installer, or add it under **Settings** > **System** > **Optional features**. |
| An OU button shows "- domain root" in orange | That OU isn't set in Settings, so the domain root was used. Set it under **Settings** > **Organizational units**. |
| Delta Sync fails | Check the Entra Connect server name, that PowerShell remoting is enabled on it, and that your admin account may run `Start-ADSyncSyncCycle` there. |
| A Microsoft 365 module won't install | The PC can't reach the PowerShell Gallery. Allow `www.powershellgallery.com`, or install the module from a PowerShell window running as your Windows account. |
| IntuneTools sign-in never returns | The loopback listener is blocked. Set **Sign-in method** to **Device code**. |
| IntuneTools says a permission is missing | Add that Graph permission to the app registration and grant admin consent again ([11.2](#112-api-permissions)). |
| Teams sign-in error `0x80070520` | Windows sign-in found no session for the account running the tool. Run `C:\ProgramData\MSToolkit\Launch-M365-Teams-Block-Number.cmd` as the Windows user signed in to the PC. |
| Old icon after an update | Windows caches icons. Sign out and back in, or run `ie4uinit.exe -show`. |

---

## 14. Security notes

- Passwords are never stored in plain text. Remembered sign-ins and IntuneTools' cached sign-in are encrypted with Windows DPAPI and can only be read by the same account on the same PC.
- The launcher remembers only a username, never a password.
- IntuneTools uses delegated sign-in through your own single-tenant app registration, with no client secret. Actions run with the signed-in user's rights and appear in the Intune and Entra audit logs under that user.
- Every change the tools make is confirmed first, naming the exact object and the effect.
- Keep the admin account used for MSToolkit separate from your everyday account, and consider requiring app assignment for the IntuneTools registration ([11.3](#113-lock-the-app-down-recommended)).
