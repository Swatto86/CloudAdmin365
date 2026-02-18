namespace CloudAdmin365.Services;

/// <summary>
/// Base interface for all audit services.
/// Allows extensibility: add RoomPermissions, MailboxPermissions, TeamOwners, etc.
/// </summary>
public interface IAuditService
{
    /// <summary>
    /// Human-readable name of this audit (e.g., "Room Calendar Permissions").
    /// </summary>
    string DisplayName { get; }

    /// <summary>
    /// Unique identifier for this audit service.
    /// </summary>
    string ServiceId { get; }

    /// <summary>
    /// Category of audit (e.g., "Exchange", "Teams", "SharePoint").
    /// </summary>
    string Category { get; }

    /// <summary>
    /// Required scopes to run this audit (e.g., User.Read.All, Group.Read.All).
    /// </summary>
    string[] RequiredScopes { get; }

    /// <summary>
    /// PowerShell modules that must be installed for this service to function.
    /// e.g. ["ExchangeOnlineManagement"] or ["MicrosoftTeams"].
    /// </summary>
    string[] RequiredPowerShellModules { get; }

    /// <summary>
    /// Check if this audit is available (auth, permissions, etc).
    /// </summary>
    Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Get description shown in UI.
    /// </summary>
    string GetDescription();
}

/// <summary>
/// Audit provider pattern: allows registering multiple audit implementations.
/// </summary>
public interface IAuditServiceProvider
{
    /// <summary>
    /// Register an audit service (generic overload — type-safe at call site).
    /// </summary>
    void RegisterAudit<T>(T service) where T : class, IAuditService;

    /// <summary>
    /// Register an audit service when the concrete type is only known at runtime.
    /// </summary>
    void RegisterAudit(IAuditService service);

    /// <summary>
    /// Get all registered audit services.
    /// </summary>
    IReadOnlyList<IAuditService> GetAllAudits();

    /// <summary>
    /// Get audit by ID.
    /// </summary>
    IAuditService? GetAudit(string serviceId);

    /// <summary>
    /// Get audits by category (e.g., "Exchange").
    /// </summary>
    IReadOnlyList<IAuditService> GetAuditsByCategory(string category);
}

/// <summary>
/// Specific audit for room calendar permissions.
/// </summary>
public interface IRoomPermissionsService : IAuditService
{
    /// <summary>
    /// Run audit for a specific room email.
    /// </summary>
    Task<RoomPermissionResult> AuditRoomAsync(string roomEmail, IProgress<AuditProgress>? progress = null, CancellationToken cancellationToken = default);
}

/// <summary>
/// Result of room permission audit.
/// </summary>
public class RoomPermissionResult
{
    public required string Room { get; set; }
    public required string RoomEmail { get; set; }
    public List<RoomPermissionEntry> Permissions { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

/// <summary>
/// Individual permission entry in room audit.
/// </summary>
public class RoomPermissionEntry
{
    public required string PermissionType { get; set; }  // "Calendar", "Booking Agent", "Creator", etc
    public required string Delegate { get; set; }
    public required string AccessLevel { get; set; }
    public string? GrantedVia { get; set; }  // Direct, Group, etc
}

/// <summary>
/// Progress information during audit.
/// </summary>
public class AuditProgress
{
    public int Current { get; set; }
    public int Total { get; set; }
    public string? CurrentItem { get; set; }
    public string? Status { get; set; }
}
