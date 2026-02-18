# CloudAdmin365

A lightweight, extensible Microsoft 365 administration platform built with C# and WinForms.  
Connects to Exchange Online and Microsoft Teams via PowerShell; uses Microsoft Graph / MSAL for authentication.

---

## Features

| Module | Category | Requires |
|--------|----------|---------|
| Room Calendar Permissions | Rooms | ExchangeOnlineManagement |
| Room Booking Audit | Rooms | ExchangeOnlineManagement |
| Calendar Diagnostic Logs | Calendar | ExchangeOnlineManagement |
| Mailbox Permissions | Mailbox | ExchangeOnlineManagement |
| Mail Forwarding Audit | Mailbox | ExchangeOnlineManagement |
| Shared Mailbox Explorer | Mailbox | ExchangeOnlineManagement |
| Group Membership Explorer | Groups | ExchangeOnlineManagement |
| Teams Explorer | Teams | MicrosoftTeams |

On first launch the app checks for the required PowerShell modules and offers to install any that are missing. Modules that are not installed have their nav items disabled; clicking them shows the manual install command.

---

## Prerequisites

| Requirement | Version |
|-------------|---------|
| .NET Runtime | 8.0 or later |
| Windows | 10 / 11 / Server 2019+ |
| PowerShell | 5.1+ (included with Windows) |

PowerShell modules (`ExchangeOnlineManagement`, `MicrosoftTeams`) are checked at startup and can be installed automatically.

---

## Running from source

```powershell
git clone <repo>
cd ExchangeAnalyzer
dotnet run
```

Or build a framework-dependent release (requires .NET 8 on target machine):

```powershell
dotnet publish -c Release -r win-x64 --self-contained false
# Output: bin\Release\net8.0-windows\win-x64\publish\CloudAdmin365.exe
```

For a fully self-contained single EXE (no runtime required on target):

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

---

## Running the published EXE

If distributing the framework-dependent build, use the included launcher:

```
setup.bat          # checks .NET runtime, then launches the app
setup.ps1          # same, PowerShell version with coloured output
```

---

## Adding a new module

Only **one file** needs to change for wiring — `ModuleRegistry.cs`.

1. **Interface** — add to `Services/IAuditServices.cs`
2. **Service** — add to `Services/Implementations/AllAuditServices.cs`  
   Set `RequiredPowerShellModules` to the PS module(s) needed.
3. **Tab** — add to `UI/AllAuditTabs.cs`
4. **Register** — add one entry to `_descriptors` in `ModuleRegistry.cs`:

```csharp
new(
    ps  => new MyNewService(ps),
    svc => new MyNewTab((IMyNewService)svc)),
```

The nav item, dependency check, and tab creation are all automatic.

---

## Debug logging

Set the environment variable before launching to enable verbose output:

```powershell
$env:CLOUDADMIN365_DEBUG = "1"
.\CloudAdmin365.exe
```

Log file: `%LocalAppData%\CloudAdmin365\logs\app.log`

Logged at `DEBUG` level: module registration, PS module check results, tab factory lookups, PS command names.  
Secrets and tokens are never logged at any level.

---

## Project structure

```
ModuleRegistry.cs          ← add new modules here
Program.cs                 ← entry point
CloudAdmin365.csproj

Models/                    ← POCO data contracts
Services/
  IAuditService.cs         ← base interface (RequiredPowerShellModules etc.)
  IAuditServices.cs        ← per-module interfaces + result models
  Implementations/
    AllAuditServices.cs    ← 8 service implementations
    AuditServiceProvider.cs
    AzureIdentityAuthService.cs
    DependencyManager.cs   ← startup PS module checks + install dialog
    PowerShellHelper.cs    ← PS runspace management
UI/
  MainForm.cs              ← left nav + right content shell
  AuditTabBase.cs          ← UserControl base class
  AllAuditTabs.cs          ← 8 tab implementations
  LoginDialog.cs
Utilities/
  AppLogger.cs             ← file logger with rotation
  AppTheme.cs
  IconGenerator.cs         ← programmatic cloud+cog icon
  UiHelpers.cs
```

---

## Documentation

- [PROJECT_ATLAS.md](PROJECT_ATLAS.md) — architecture, invariants, startup sequence, module table
- [AI_PROGRESS.md](AI_PROGRESS.md) — development history, phase log, future work backlog
