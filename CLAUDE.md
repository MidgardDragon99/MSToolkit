# MSToolkit - notes for Claude Code

Read this first. Background on why things are the way they are (the settings
model, the safety model, each tool's parameters, known traps, port status) is
in [docs/HANDOFF.md](docs/HANDOFF.md). Where this file and that one disagree on a
rule, this file wins. Where either disagrees with the code, the code wins.

## 1. What it is

MSToolkit is a PowerShell 5.1 / WinForms admin toolkit for Active Directory,
Microsoft 365 and Intune: a main console (`MSToolkit.ps1`) plus companion tools
launched from it. MIT licence, intended for public release.

**Every organisation-specific value comes from the settings file**
(`%APPDATA%\MSToolkit\settings.json`, written only by the console's Settings
window). Nothing in this repository may contain a real organisation's name,
domain, tenant or client ID, hostname, OU path, person, username or naming
convention - in code, comments, docs, build scripts or assembly metadata.
Examples use `contoso`, `example.com`, `Jane Doe (jdoe)`, `202-555-0100`.

MSToolkit is a generalised branch of an internal toolkit. Changes here never flow
back to it, and changes from it only come in when re-done generically.

## 2. Layout

| File | What it is |
| --- | --- |
| `MSToolkit.ps1` | Main console: Users, Groups, Computers, OUs, M365, Reports, Settings. Runs as the admin account. |
| `NewADUser.ps1` | Create an AD account from Settings, optionally cloning groups, address and manager |
| `Compare-UserGroups.ps1` | Compare two users' direct AD groups; add or remove |
| `Investigate-AccountLockout.ps1` | Lockout investigation across DCs |
| `OffboardUser.ps1` | Departing-user checklist with verified steps and Undo |
| `M365-Group-Compare.ps1` | Entra security-group compare/add (Graph) |
| `M365-Distribution-Group-Compare.ps1` | Exchange Online distribution-group compare/add/remove |
| `M365-Conditional-Access-User-Manager.ps1` | A user's direct include/exclude across CA policies |
| `M365-Exchange-Online-Tools.ps1` | Inbox rules, message trace, delegation, quarantine, reports |
| `M365-OneDrive-SharePoint-Tools.ps1` | OneDrive access, handover, restore, reports |
| `M365-Teams-Block-Number.ps1` | Block and Forward Phone Numbers (Teams) |
| `IntuneTools.ps1` | Intune / Entra devices, Autopilot, apps, policies, groups, reports |
| `Launcher.ps1` | Starts `Launch-MSToolkit.bat` from the install folder |
| `Launch-*.bat`, `Launch-M365-Teams-Block-Number.cmd`, `Start-IntuneTools.bat` | Launchers. The AD ones use `runas` with the admin username exactly as typed. |
| `InstallMSToolkit.ps1`, `install.cmd`, `install-config.txt` | Installer, its 64-bit PowerShell wrapper, 7-Zip SFX config |
| `UninstallMSToolkit.ps1`, `uninstall.cmd`, `uninstall-config.txt` | Uninstaller (self-relocates to TEMP, elevates), wrapper, SFX config |
| `Build-Install-Uninstall-Only.cmd` | Builds `MSToolkit-Install.exe` / `MSToolkit-Uninstall.exe` (7-Zip SFX) |
| `MSToolkit.ico` | Icon for the EXE, the desktop shortcut and Apps & Features |
| `README.md` | User and deployment documentation |
| `docs/HANDOFF.md` | Design background and history |
| `Build\` | **Planned:** `Build-All.ps1` / `.bat` and `Launchers\MSToolkitLauncher.cs` (section 3) |

## 3. Building

**Today:** run `Build-Install-Uninstall-Only.cmd` from the repo root. It needs
7-Zip at `C:\Program Files (x86)\7-Zip` (`7z.exe` and `7zSD.sfx`) and an existing
`MSToolkit.EXE` (currently an IExpress package built by hand).

**Planned - one command, `Build\Build-All.bat`** (passes arguments through to
`Build\Build-All.ps1`, PowerShell 5.1, and pauses at the end). Parameters
`-SkipLauncher`, `-SkipIntuneWin`, `-IntuneWinAppUtil <path>`.

1. **Launcher EXE.** Compile `Build\Launchers\MSToolkitLauncher.cs` with the
   `csc.exe` that ships with Windows
   (`%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe`, C# 5 only: no `$""`,
   `=>` members, `nameof` or `?.`):
   `csc.exe /nologo /target:winexe /platform:anycpu /optimize+ /out:MSToolkit.EXE /win32icon:MSToolkit.ico /reference:System.Windows.Forms.dll Build\Launchers\MSToolkitLauncher.cs`.
   The launcher embeds nothing, resolves its own folder, starts what
   `Launcher.ps1` starts, uses the 64-bit system folder (`Sysnative` from a
   32-bit process), shows a message box and exits 1 if the target is missing,
   and carries neutral metadata (`AssemblyTitle`/`AssemblyProduct` "MSToolkit",
   no company). It adds no username or domain. Fails the build if `csc` returns
   non-zero or the EXE is missing.
2. **7-Zip SFX.** Runs `Build-Install-Uninstall-Only.cmd` unchanged with stdin
   redirected from an empty file (so its closing `pause` falls through) and
   checks its exit code.
3. **.intunewin.** Stages the package folder in `%TEMP%` (leaving out `*.7z`,
   `*.intunewin`, the builder and `Logs`) and runs
   `IntuneWinAppUtil.exe -c <staging> -s install.cmd -o <repo root> -q`. The
   output is `install.intunewin`; the Intune app's install command depends on
   that name.
4. **Summary** of OK / FAILED per step; exit 1 on any failure.

**IntuneWinAppUtil.exe** is Microsoft's tool and is never committed. Download it
from Microsoft's `Microsoft-Win32-Content-Prep-Tool` repository on GitHub, and
put it in the repo root (git-ignored) or pass `-IntuneWinAppUtil <path>`.

Build outputs (`*.7z`, `*-Install.exe`, `*-Uninstall.exe`, `MSToolkit.EXE`,
`*.intunewin`, `*.zip`) are git-ignored and never committed.

## 4. Install layout and security

- Installs to `%ProgramFiles(x86)%\MSToolkit` and runs only from there. The
  installer resets that folder's ACL to inherited.
- `C:\ProgramData\MSToolkit` is SYSTEM/Administrators full, Users read; only its
  `Logs` subfolder is writable by Users. Nothing runs from it.
- Never run or stage scripts in a user-writable path.
- The installer creates the all-users desktop shortcut itself
  (`WScript.Shell.CreateShortcut`). No `.lnk` is shipped or committed: a
  shortcut file carries the SID of the account that made it.
- Settings and other per-account files live under `MSToolkit`-named folders, apart
  from the internal toolkit's, so both can be installed side by side.

## 5. Rules

- **PowerShell 5.1.** Use `$list.ToArray()`, not `return @($list)` (a generic
  `List` inside `@()` throws). Commit a DataGridView edit before reading a
  checkbox cell.
- **Safety logic uses SID / RID or resolved AD properties, never display names**
  (one documented exception: DnsAdmins, which has no well-known RID).
- **Group membership is direct** unless nested is explicitly asked for.
- **Preserve DN escaping.**
- **No automatic domain prefix** on any username, anywhere. What is typed is used
  and remembered exactly.
- **UI hints are generic but exact:** no naming convention or real value, but name
  the exact value expected and the AD attribute or source it maps to.
- **Every Graph POST sends a body** (`-Body '{}'` when empty).
- **Check the log-level `ValidateSet`** of the file you are in before adding a log
  call (`WARN` in some tools, `WARNING` in others - see HANDOFF section 6).
- **Start-up behaviour goes in its own top-level `Add_Shown` handler**, with any
  condition inside it.
- **Minimal, targeted changes.** No redesigns, refactors or renames unless asked.
- **Canonical filenames**, no version or date suffixes.
- **Line endings are set by `.gitattributes`:** text is stored as LF and checked
  out as CRLF everywhere (`* text=auto eol=crlf`); binaries are never converted.
  Write CRLF in new and edited files. `.bat`/`.cmd` must be CRLF. A file
  missing from the binary list in `.gitattributes` gets added there before it is
  committed.
- **ASCII-only `.ps1` source** unless the file already has a UTF-8 BOM (5.1 reads
  BOM-less files as ANSI).

## 6. Packaging rule

Adding, removing or renaming a packaged file means updating, together:

- the builder's validation list **and** its archive list
  (`Build-Install-Uninstall-Only.cmd`),
- the installer's `$PayloadFiles` list (copy and validation),
- the uninstaller, if the file lives outside the install folder,
- `README.md` (section 1, "What's in the package").

## 7. How to work in this repo (git and rollback)

The main branch is `main`. **There is no remote yet.** The `git pull` and
`git push` rules below apply once one is added; until then, skip those steps.
"The maintainer" below means the person who owns this repository.

**Before a change**

1. Run `git status`. If there are uncommitted changes Claude didn't make, stop
   and ask. They may be the maintainer's work in progress.
2. Run `git pull --ff-only` on `main` so the work starts from what is on GitHub.
   If that fails, stop and ask. Don't merge or rebase to force it.
3. Make a branch for the change: `git switch -c <short-topic>`. Never work
   directly on `main` beyond a one-line fix the maintainer has approved.

**While changing**

- Edit the files in place.
- Parse-check every edited `.ps1`. Commit in small, working steps with a message
  saying what changed and why.
- Show the maintainer `git diff` (or a summary of it) before each commit that
  changes behaviour.
- Before every commit, check for anything org-specific or secret (names, domains,
  tenant/client IDs, hosts, usernames, SIDs, tokens, logs), including UTF-16 text
  and binaries. The repo is heading public, and once something is pushed, assume
  it can't be taken back.

**When it's done**

- Tell the maintainer exactly what changed and what to test in the GUI.
- Merge to `main` only after they say it tested OK:
  `git switch main`, `git merge --no-ff <branch>`, `git branch -d <branch>`.
- **Push only when the maintainer says to**: `git push origin main`. Never push a
  branch or tag they haven't asked for.
- Tag each release they publish or deploy: `git tag v<version>` or
  `released-YYYY-MM-DD`, pushed with `git push origin <tag>` when they say.

**Rolling back**

| Need | Command |
| --- | --- |
| See what changed but isn't committed | `git status`, `git diff` |
| Undo uncommitted edits to one file | `git restore <path>` |
| Undo all uncommitted edits | `git restore .` (ask first) |
| See a file's history | `git log --oneline -- <path>` |
| Put one file back to an earlier version | `git restore --source <commit-or-tag> -- <path>`, then commit |
| Undo a committed change, keeping history (safe after a push) | `git revert <commit>` |
| Abandon an unmerged local branch | `git switch main`, then `git branch -D <branch>` (ask first) |
| Rebuild a released version | `git switch --detach <tag>`, `Build\Build-All.bat`, `git switch main` |

Once a commit has been pushed, roll it back **only with `git revert`**. Never
rewrite pushed history.

**Never without the maintainer's explicit OK in the same conversation**

- `git push --force` (or `--force-with-lease`), `git reset --hard`, `git clean`,
  `git rebase`, amending or rewriting commits, deleting tags or branches on
  GitHub.
- Changing the repo's visibility, licence, settings or collaborators.
- Deleting or renaming files in the repo.
- Running MSToolkit or any AD / Graph / Exchange / Teams / Intune command that
  changes something. Claude Code edits, parse-checks and builds; the maintainer
  runs the tool and does all live testing.
- Installing a built package.

Read-only commands (`git status`, `git log`, `git diff`, `git fetch`,
`Get-FileHash`, `Get-ChildItem`, the parser check, `Build-All.bat`) are fine
without asking.

## 8. Testing

There are no automated tests. After a change:

- Parse-check every edited script:
  `[System.Management.Automation.Language.Parser]::ParseFile(<path>, [ref]$null, [ref]$errors)`
  with zero errors.
- State exactly what changed and what to test in the GUI, step by step. The
  maintainer does all live testing.

## 9. Open items

- **Remove `MSToolkit.lnk` from the package** (installer `$PayloadFiles`, builder
  validation and archive lists, README). The installer already creates the
  shortcut.
- **Compiled launcher and `Build\Build-All`** (section 3) to replace the hand-built
  IExpress `MSToolkit.EXE`.
- **Neutral sample name** in the ComboBox prefix-matching comment in the eight
  picker tools.
- Add `LICENSE` (MIT) before publishing.
- Rebuild the installers and `.intunewin` from this repo after the changes above.
- Code signing.
- Test on a clean PC with no previous install, and an Intune upgrade from an old
  (writable-runtime) install.
- Offboarding round 2 is implemented but **untested** against a real tenant or
  domain; see HANDOFF sections 7 and 8 for the list and the deferred backlog
  (don't start backlog items unless asked).
