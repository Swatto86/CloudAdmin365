# CloudAdmin365 — AI Progress Document

**Project path:** `c:\Users\Steve Watson\ExchangeAnalyzer\`  
**Product name:** CloudAdmin365 (renamed from ExchangeAnalyzer)  
**Framework:** .NET 8.0-windows, WinForms, framework-dependent deployment  
**Namespace:** `CloudAdmin365`  
**Project file:** `CloudAdmin365.csproj`

---

## Current State (Phase 4 complete)

Build: **0 errors, 0 warnings**  
Modules registered: **8**  
PS modules checked at runtime: **ExchangeOnlineManagement**, **MicrosoftTeams**

---

## Completed Phases

### Phase 1 — Lightweight deployment + dependency checking
- Removed bundled DLLs: `SelfContained=false`, `PublishSingleFile=false`
- `DependencyManager` checks for required PowerShell modules at startup
- `setup.bat` / `setup.ps1` launchers
- `AppLogger` with file rotation and `CLOUDADMIN365_DEBUG` env var

### Phase 2 — Rename + icon
- Bulk rename across all `.cs` files, `.csproj`, `.ico`: ExchangeAnalyzer → CloudAdmin365
- Programmatic cloud+cog icon (`IconGenerator.cs`) — no external `.ico` dependency
- All 22 source files updated; zero stray references

### Phase 3 — Left nav UI + Teams module
- `MainForm` redesigned: `SplitContainer` (210 px fixed nav + right content)
- `AuditTabBase` changed from `TabPage` to `UserControl`
- Teams Explorer module added: `ITeamsExplorerService`, `TeamsExplorerService`, `TeamsExplorerTab`
- Teams tab has two-pane drill-down (teams list → member detail)
- `MicrosoftTeams` added to dependency checks

### Phase 4 — Extensibility hub + nav disabling
- **`ModuleRegistry.cs`** (new) — single file to change when adding a module:
  - `_descriptors`: list of `(ServiceFactory, TabFactory)` pairs
  - `RegisterAll(provider, ps)` — populates audit provider + tab factory index
  - `CreateTab(service)` — runtime tab creation via factory
- `IAuditService.RequiredPowerShellModules` added to base interface
- All 8 services implement `RequiredPowerShellModules`
- `IAuditServiceProvider` / `AuditServiceProvider` gain non-generic `RegisterAudit(IAuditService)`
- `DependencyManager` completely rewritten:
  - Derives required modules from registered services (no hardcoded list)
  - Returns `IReadOnlyDictionary<string, bool>` (module → installed)
  - Custom `DependencyDialog` form: per-module ✓/✗, tick-boxes to install, two action buttons
- `MainForm` updated: nav items greyed + non-interactive when PS module unavailable; tooltip + click gives install instructions
- `Program.cs` simplified: 8 `RegisterAudit` calls → `ModuleRegistry.RegisterAll()`

---

## Debug Logging

**Enable:** set env var `CLOUDADMIN365_DEBUG=1` (or `true` / `yes`) before launching.

```powershell
$env:CLOUDADMIN365_DEBUG = "1"
.\CloudAdmin365.exe
```

**Log file:** `%LocalAppData%\CloudAdmin365\logs\app.log`  
**Rotation:** file rotated to `app.log.old` at 2 MB.  
**Secrets policy:** tokens, passwords, PII are never written to the log at any level.

### What gets logged at DEBUG level

| Component | What is logged |
|-----------|----------------|
| `AppLogger` | "Debug logging enabled" on startup |
| `ModuleRegistry.RegisterAll` | Each module registered: ServiceId, DisplayName, RequiredPSModules |
| `ModuleRegistry.CreateTab` | Which factory was found (or not) for a ServiceId |
| `DependencyManager` | Derived module list, per-module check result (FOUND/NOT FOUND), missing list, install command |
| `PowerShellHelper` | Command name before each PS execution; window handle for auth |

### What gets logged at INFO level

- Application start, executable path, version
- Logger initialized
- Each PS module found / not found
- Dependency installation start and completion
- Login cancelled
- Session logout

### What gets logged at ERROR level

- All unhandled exceptions (with full stack trace)
- PS module install failures
- Fatal startup errors

---

## Architecture snapshot

```
Program.cs
  ModuleRegistry.RegisterAll(provider, ps)   ← 8 modules wired here
  DependencyManager.CheckAndInstall(services) ← derives PS module list from services
  MainForm(auth, provider, availabilityMap)

ModuleRegistry.cs                            ← ONLY file to change for a new module
  _descriptors: [ (ServiceFactory, TabFactory), ... ]

Services/IAuditService.cs
  + string[] RequiredPowerShellModules       ← drives dependency dialog + nav disabling

Services/Implementations/
  AllAuditServices.cs     — 8 service impls; each declares RequiredPowerShellModules
  AuditServiceProvider.cs — RegisterAudit<T> + RegisterAudit(IAuditService)
  DependencyManager.cs    — per-module availability map, DependencyDialog inner form
  PowerShellHelper.cs     — PS runspace, EXO connection

UI/
  MainForm.cs             — nav disabled+grey when module unavailable
  AuditTabBase.cs         — UserControl base (Input/Buttons/Grid/Status)
  AllAuditTabs.cs         — 8 tab implementations
```

---

## How to add a new module (checklist for AI and humans)

- [ ] Add interface to `Services/IAuditServices.cs`
- [ ] Add implementation to `Services/Implementations/AllAuditServices.cs`  
      — set `RequiredPowerShellModules` to the needed PS module(s)
- [ ] Add tab to `UI/AllAuditTabs.cs`
- [ ] Add one entry to `ModuleRegistry._descriptors` in `ModuleRegistry.cs`
- [ ] Build — if 0 errors, done.

No changes needed in `Program.cs`, `DependencyManager`, or `MainForm`.

---

## Known gaps / future work

| Item | Notes |
|------|-------|
| E2E test suite | No automated tests yet; manual testing only |
| Regression tests | Needed before large refactors |
| SharePoint module | Interface + service not yet implemented |
| Security / Compliance module | Not started |
| Audit history / SQLite | Not started |
| Report generation (PDF/Excel) | Not started |
| Scheduled/automated audits | Not started |
| update-application release script | Required by DevWorkflow Rule A18 — not yet created |
| CI/CD pipeline | Not configured |

---

## NuGet packages

| Package | Version | Purpose |
|---------|---------|---------|
| `Azure.Identity` | 1.13.1 | MSAL interactive auth |
| `Microsoft.Graph` | 5.62.0 | Graph API client |
| `Microsoft.Identity.Client` | 4.66.2 | Token caching |
| `Microsoft.PowerShell.SDK` | 7.4.6 | In-process PS runspace |

---

## Build commands

```powershell
# Debug
dotnet build

# Release (framework-dependent, requires .NET 8 runtime on target)
dotnet publish -c Release -r win-x64 --self-contained false

# Self-contained single EXE (larger, no runtime requirement)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

---

*Last updated: Phase 4 complete — February 18, 2026*
