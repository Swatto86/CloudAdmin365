namespace CloudAdmin365;

using System.Collections.Generic;
using System.Windows.Forms;
using CloudAdmin365.Services;
using CloudAdmin365.Services.Implementations;
using CloudAdmin365.UI;
using CloudAdmin365.Utilities;

/// <summary>
/// Central module registry for CloudAdmin365.
/// This is the ONLY file that needs to be changed when adding a new module.
///
/// To add a new module:
///   1. Create your service interface in Services/IAuditServices.cs.
///   2. Implement the service in Services/Implementations/AllAuditServices.cs.
///   3. Create the UI tab as a UserControl in UI/.
///   4. Add a single entry in the _descriptors list below.
///
/// Everything else (nav items, dependency checks, tab activation) is automatic.
/// </summary>
public static class ModuleRegistry
{
    // ── Module descriptor ─────────────────────────────────────────────────
    // Each entry maps a service factory (PowerShellHelper → IAuditService)
    // to a tab factory (IAuditService → UserControl).
    // ─────────────────────────────────────────────────────────────────────

    private sealed record ModuleDescriptor(
        Func<PowerShellHelper, IAuditService> ServiceFactory,
        Func<IAuditService, UserControl>      TabFactory);

    // ADD NEW MODULES HERE ↓ ───────────────────────────────────────────────
    private static readonly IReadOnlyList<ModuleDescriptor> _descriptors =
    [
        new(
            ps  => new RoomPermissionsService(ps),
            svc => new RoomPermissionsTab((IRoomPermissionsService)svc)),

        new(
            ps  => new RoomBookingAuditService(ps),
            svc => new RoomBookingTab((IRoomBookingAuditService)svc)),

        new(
            ps  => new CalendarDiagnosticService(ps),
            svc => new CalendarDiagnosticTab((ICalendarDiagnosticService)svc)),

        new(
            ps  => new MailboxPermissionsService(ps),
            svc => new MailboxPermissionsTab((IMailboxPermissionsService)svc)),

        new(
            ps  => new MailForwardingAuditService(ps),
            svc => new MailForwardingTab((IMailForwardingAuditService)svc)),

        new(
            ps  => new SharedMailboxService(ps),
            svc => new SharedMailboxTab((ISharedMailboxService)svc)),

        new(
            ps  => new GroupExplorerService(ps),
            svc => new GroupExplorerTab((IGroupExplorerService)svc)),

        new(
            ps  => new TeamsExplorerService(ps),
            svc => new TeamsExplorerTab((ITeamsExplorerService)svc)),
    ];
    // ADD NEW MODULES HERE ↑ ───────────────────────────────────────────────

    // Tab factory index, populated by RegisterAll and used by CreateTab.
    private static readonly Dictionary<string, Func<IAuditService, UserControl>> _tabFactories = [];

    // ── Public API ────────────────────────────────────────────────────────

    /// <summary>
    /// Creates and registers all known services with the provider.
    /// Must be called once after authentication, before opening MainForm.
    /// </summary>
    public static void RegisterAll(IAuditServiceProvider provider, PowerShellHelper powerShellHelper)
    {
        ArgumentNullException.ThrowIfNull(provider);
        ArgumentNullException.ThrowIfNull(powerShellHelper);

        _tabFactories.Clear();

        AppLogger.WriteDebug($"ModuleRegistry.RegisterAll: registering {_descriptors.Count} module(s).");

        foreach (var descriptor in _descriptors)
        {
            var service = descriptor.ServiceFactory(powerShellHelper);
            provider.RegisterAudit(service);
            _tabFactories[service.ServiceId] = descriptor.TabFactory;
            AppLogger.WriteDebug($"  Registered: [{service.ServiceId}] '{service.DisplayName}' (PSModules: [{string.Join(", ", service.RequiredPowerShellModules)}])");
        }

        AppLogger.WriteInfo($"ModuleRegistry: {_tabFactories.Count} module(s) registered.");
    }

    /// <summary>
    /// Creates the UserControl tab for the given service.
    /// Returns null if the service was not registered via RegisterAll.
    /// </summary>
    public static UserControl? CreateTab(IAuditService service)
    {
        ArgumentNullException.ThrowIfNull(service);
        var found = _tabFactories.TryGetValue(service.ServiceId, out var factory);
        AppLogger.WriteDebug($"ModuleRegistry.CreateTab: [{service.ServiceId}] — {(found ? "factory found" : "NO factory, returning null")}");
        return found ? factory!(service) : null;
    }
}
