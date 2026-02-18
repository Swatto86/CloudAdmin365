# CloudAdmin365 — Project Atlas

**Microsoft 365 Administration Platform**  
A lightweight, extensible WinForms application for auditing and administering Microsoft 365 tenants via PowerShell and Microsoft Graph.

---

## Project Status

| Phase | Description | Status |
|-------|-------------|--------|
| 1 | Framework-dependent deployment, DependencyManager, setup scripts | ✅ Complete |
| 2 | Rename ExchangeAnalyzer → CloudAdmin365, cloud+cog icon | ✅ Complete |
| 3 | Left nav panel UI, AuditTabBase → UserControl, Teams Explorer module | ✅ Complete |
| 4 | ModuleRegistry extensibility hub, per-module availability, nav disabling | ✅ Complete |

Build status: **0 errors, 0 warnings** (verified after each phase).

---

## System Purpose

CloudAdmin365 allows IT administrators to:
- Audit Exchange Online room/calendar permissions and booking policies
- Inspect mailbox-level permissions, mail forwarding, and shared mailboxes
- Explore Microsoft 365 Group membership
- Browse Teams and their members

Modules are PowerShell-backed; the app checks for required PS modules at startup and can install them automatically. Nav items for modules whose PS modules are missing are greyed out and non-interactive.

---

## Architecture

```
┌─────────────────────────────────────────────┐
│  Entry Point                                │
│  Program.cs                                 │
│  └─ ModuleRegistry.RegisterAll()            │  ← only file to change for new module
│  └─ DependencyManager: availability map     │
│  └─ MainForm(auth, provider, availability)  │
├─────────────────────────────────────────────┤
│  ModuleRegistry  (root namespace)           │
│  One entry per module:                      │
│    ServiceFactory + TabFactory              │
│  RegisterAll() → populates IAuditService-   │
│    Provider and builds tab factory index    │
│  CreateTab(IAuditService) → UserControl     │
├─────────────────────────────────────────────┤
│  UI Layer  (CloudAdmin365.UI)               │
│  ├─ MainForm: left nav + right content      │
│  │   Nav items disabled when PS module      │
│  │   unavailable; tooltip shows requirement │
│  ├─ AuditTabBase (UserControl base class)   │
│  │   Standard layout: Input/Buttons/Grid/   │
│  │   Status                                 │
│  ├─ LoginDialog                             │
│  └─ Module tabs: *Tab.cs                   │
├─────────────────────────────────────────────┤
│  Service Layer  (CloudAdmin365.Services)    │
│  ├─ IAuditService: DisplayName, ServiceId,  │
│  │   Category, RequiredScopes,              │
│  │   RequiredPowerShellModules,             │
│  │   IsAvailableAsync(), GetDescription()  │
│  ├─ IAuditServiceProvider: RegisterAudit,  │
│  │   GetAllAudits, GetAudit,               │
│  │   GetAuditsByCategory                   │
│  └─ Implementations/                        │
│     ├─ AuditServiceProvider                 │
│     ├─ AllAuditServices (8 services)        │
│     ├─ DependencyManager                    │
│     └─ PowerShellHelper                     │
├─────────────────────────────────────────────┤
│  Authentication  (CloudAdmin365.Services)   │
│  ├─ IAuthService                            │
│  └─ AzureIdentityAuthService (MSAL)         │
├─────────────────────────────────────────────┤
│  Utilities  (CloudAdmin365.Utilities)       │
│  ├─ AppLogger (file + debug logging)        │
│  ├─ AppTheme  (colors, fonts)               │
│  ├─ IconGenerator (cloud+cog programmatic)  │
│  └─ UiHelpers (DataGrid factory, CSV)       │
└─────────────────────────────────────────────┘
```

---

## Repository Structure

```
CloudAdmin365.csproj              # .NET 8.0-windows, framework-dependent
Program.cs                        # Entry point, DI wiring, app lifecycle
ModuleRegistry.cs                 # ← ONLY file to touch when adding a module

Models/
  AuditModels.cs                  # Audit progress, result base models
  RoomPermission.cs               # Room permission models

Services/
  IAuthService.cs                 # Authentication contract
  IAuditService.cs                # Audit base contract + IAuditServiceProvider
  IAuditServices.cs               # Per-module interfaces + result models

  Implementations/
    AllAuditServices.cs           # 8 service implementations
    AuditServiceProvider.cs       # Registry implementation
    AzureIdentityAuthService.cs   # MSAL interactive auth
    DependencyManager.cs          # PS module checks, install dialog
    PowerShellHelper.cs           # Runspace management, EXO connection

UI/
  AuditTabBase.cs                 # UserControl base: Input/Button/Grid/Status
  AllAuditTabs.cs                 # All 8 module tab implementations
  LoginDialog.cs                  # Authentication UI
  MainForm.cs                     # Left nav + right content shell
  RoomPermissionsTab.cs           # (legacy; merged into AllAuditTabs)

Utilities/
  AppLogger.cs                    # Structured file logging, debug mode
  AppTheme.cs                     # Color/font constants
  IconGenerator.cs                # Programmatic cloud+cog icon generation
  UiHelpers.cs                    # DataGrid factory, CSV export

create-embedded-setup.ps1         # Single-file setup builder
generate-icon.ps1                 # Icon generation helper
```

---

## How to Add a New Module

Adding a module touches exactly **one code location** (`ModuleRegistry.cs`) plus new files for the service and tab.

### Steps

1. **Define the service interface** in `Services/IAuditServices.cs`:
   ```csharp
   public interface IMyNewService : IAuditService
   {
       Task<MyResult> RunAsync(/* params */, CancellationToken ct = default);
   }
   ```

2. **Implement the service** in `Services/Implementations/AllAuditServices.cs`:
   ```csharp
   public sealed class MyNewService : IMyNewService
   {
       public string DisplayName => "My New Feature";
       public string ServiceId   => "my-new";
       public string Category    => "Exchange";  // or Teams, Security, etc.
       public string[] RequiredScopes             => ["User.Read.All"];
       public string[] RequiredPowerShellModules  => ["ExchangeOnlineManagement"];
       // ... implement RunAsync, IsAvailableAsync, GetDescription
   }
   ```

3. **Create the UI tab** in `UI/AllAuditTabs.cs`:
   ```csharp
   public class MyNewTab : AuditTabBase
   {
       private readonly IMyNewService _svc;
       public MyNewTab(IMyNewService svc) { _svc = svc; Text = "My New Feature"; }
       protected override void SetupInputPanel() { /* inputs */ }
       protected override async Task RunAuditAsync(CancellationToken ct) { /* run */ }
       protected override void RenderResults(object data) { /* grid */ }
   }
   ```

4. **Register in `ModuleRegistry.cs`** — add one entry to `_descriptors`:
   ```csharp
   new(
       ps  => new MyNewService(ps),
       svc => new MyNewTab((IMyNewService)svc)),
   ```

That's it. The nav item appears automatically. If the required PS module is missing, the nav item is disabled with a tooltip pointing to the install command.

---

## Service Interfaces

### IAuditService
```csharp
public interface IAuditService
{
    string DisplayName { get; }
    string ServiceId   { get; }
    string Category    { get; }
    string[] RequiredScopes            { get; }
    string[] RequiredPowerShellModules { get; }
    Task<bool> IsAvailableAsync(CancellationToken ct = default);
    string GetDescription();
}
```

### IAuditServiceProvider
```csharp
public interface IAuditServiceProvider
{
    void RegisterAudit<T>(T service) where T : class, IAuditService;
    void RegisterAudit(IAuditService service);          // runtime registration
    IReadOnlyList<IAuditService> GetAllAudits();
    IAuditService? GetAudit(string serviceId);
    IReadOnlyList<IAuditService> GetAuditsByCategory(string category);
}
```

---

## Registered Modules (Phase 4 baseline)

| ServiceId | DisplayName | Category | RequiredPSModules |
|-----------|-------------|----------|-------------------|
| `room-permissions` | Room Calendar Permissions | Rooms | ExchangeOnlineManagement |
| `room-booking` | Room Booking Audit | Rooms | ExchangeOnlineManagement |
| `calendar-diagnostic` | Calendar Diagnostic Logs | Calendar | ExchangeOnlineManagement |
| `mailbox-permissions` | Mailbox Permissions | Mailbox | ExchangeOnlineManagement |
| `mail-forwarding` | Mail Forwarding Audit | Mailbox | ExchangeOnlineManagement |
| `shared-mailbox` | Shared Mailbox Explorer | Mailbox | ExchangeOnlineManagement |
| `group-explorer` | Group Membership Explorer | Groups | ExchangeOnlineManagement |
| `teams-explorer` | Teams Explorer | Teams | MicrosoftTeams |

---

## Startup Sequence

```
Program.Main()
  │
  ├─ DPI awareness
  ├─ AppLogger.Initialize
  ├─ ApplicationConfiguration.Initialize
  │
  ├─ AzureIdentityAuthService created
  ├─ LoginDialog.ShowDialog()  (MSAL interactive)
  │
  ├─ PowerShellHelper.InitializeAsync()
  │
  ├─ AuditServiceProvider created
  ├─ ModuleRegistry.RegisterAll(provider, psHelper)
  │   └─ Creates 8 services, stores tab factories by ServiceId
  │
  ├─ DependencyManager.CheckAndInstallDependenciesAsync(allServices)
  │   ├─ Derives unique PS modules from RequiredPowerShellModules
  │   ├─ Checks each module via Get-Module -ListAvailable
  │   ├─ Shows DependencyDialog if any are missing
  │   │   ├─ Checklist with ✓/✗ and "Install automatically" checkboxes
  │   │   └─ Installs selected modules via Install-Module -Scope CurrentUser
  │   └─ Returns IReadOnlyDictionary<string, bool> (module → installed)
  │
  └─ MainForm(auth, provider, moduleAvailability)
      └─ Nav items disabled+greyed if RequiredPowerShellModules unavailable
```

---

## Build & Run

### Debug
```powershell
cd "c:\Users\Steve Watson\ExchangeAnalyzer"
dotnet build
dotnet run
```

### Release (framework-dependent)
```powershell
dotnet publish -c Release -r win-x64 --self-contained false
```
Output: `bin\Release\net8.0-windows\win-x64\publish\CloudAdmin365.exe`  
Requires .NET 8.0 Runtime on target machine.

### Standalone (self-contained, larger)
```powershell
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

---

## Debug Logging

Enable verbose logging:
```powershell
$env:CLOUDADMIN365_DEBUG = "1"
.\CloudAdmin365.exe
```

Log file: `%LocalAppData%\CloudAdmin365\logs\app.log`  
Log level `DEBUG` includes function entry/exit, PS command invocations, state transitions.  
Secrets, tokens, and PII are never logged.

---

## Configuration

All configuration is external to code:
- **Azure AD**: `clientId` / `tenantId` passed at runtime via auth service
- **PS Module timeouts**: `CommandTimeoutSeconds = 30`, `ModuleInstallTimeoutSeconds = 180` in `DependencyManager`
- **Nav categories**: `CategoryOrder[]` in `MainForm` controls display order

---

## Critical Invariants

1. `ModuleRegistry._descriptors` is the single source of truth for registered modules.
2. `DependencyManager` derives required PS modules from services; there is no separate hardcoded list.
3. All nav items for modules whose PS modules are missing are non-clickable with informational tooltips.
4. `AuditTabBase` must remain a `UserControl` (not `TabPage`).
5. `PowerShellHelper` owns the PS runspace; only one instance exists per session.
6. Auth tokens are never written to logs.

---

## DevWorkflow Rule Compliance

| Rule | Compliance |
|------|-----------|
| A1 — Architectural boundaries | Services layer has no UI/WinForms dependency |
| A2 — Code correctness | All code compiles 0 errors, 0 warnings |
| A3 — Testing | E2E test suite planned (auth, audit, nav disabling) |
| A7 — Incremental development | Each phase left system in buildable state |
| A10 — Debug mode | `CLOUDADMIN365_DEBUG=1` env var activates verbose logging |
| A11 — Resilience | PS module checks retry, DependencyDialog accumulates all errors |
| A12 — Security | Tokens/PII never logged |
| A16 — UI resilience | Nav items disabled when module unavailable |
| A17 — Pre-flight validation | DependencyManager validates all preconditions before opening MainForm |
| B — Atlas | This document; updated with every structural change |

---

**Last Updated**: Phase 4 complete  
**Framework**: .NET 8.0-windows (framework-dependent deployment)  
**Namespace**: `CloudAdmin365`  
**Project file**: `CloudAdmin365.csproj`  

