namespace CloudAdmin365.Services.Implementations;

using CloudAdmin365.Services;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// Default implementation of audit service provider.
/// Manages registration and retrieval of audit services.
/// </summary>
public class AuditServiceProvider : IAuditServiceProvider
{
    private readonly Dictionary<string, IAuditService> _audits = [];

    /// <summary>
    /// Register an audit service by its unique ID (generic overload).
    /// </summary>
    public void RegisterAudit<T>(T service) where T : class, IAuditService
    {
        ArgumentNullException.ThrowIfNull(service);
        _audits[service.ServiceId] = service;
    }

    /// <summary>
    /// Register an audit service when the concrete type is only known at runtime.
    /// </summary>
    public void RegisterAudit(IAuditService service)
    {
        ArgumentNullException.ThrowIfNull(service);
        _audits[service.ServiceId] = service;
    }

    /// <summary>
    /// Get all registered audit services.
    /// </summary>
    public IReadOnlyList<IAuditService> GetAllAudits() => _audits.Values.ToList().AsReadOnly();

    /// <summary>
    /// Get audit by ID.
    /// </summary>
    public IAuditService? GetAudit(string serviceId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(serviceId);
        _audits.TryGetValue(serviceId, out var service);
        return service;
    }

    /// <summary>
    /// Get audits by category (e.g., "Exchange", "Teams").
    /// </summary>
    public IReadOnlyList<IAuditService> GetAuditsByCategory(string category)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(category);
        return _audits.Values
            .Where(a => a.Category.Equals(category, StringComparison.OrdinalIgnoreCase))
            .ToList()
            .AsReadOnly();
    }
}
