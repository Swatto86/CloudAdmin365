namespace CloudAdmin365.Services;

/// <summary>
/// Audit service for room mailbox booking history.
/// </summary>
public interface IRoomBookingAuditService : IAuditService
{
    Task<RoomBookingAuditResult> AuditRoomBookingsAsync(
        string roomEmail,
        DateTime? startDate = null,
        DateTime? endDate = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public class RoomBookingAuditResult
{
    public required string Room { get; set; }
    public List<RoomBookingEntry> Bookings { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class RoomBookingEntry
{
    public required string Subject { get; set; }
    public required string Organizer { get; set; }
    public DateTime? StartTime { get; set; }
    public DateTime? EndTime { get; set; }
    public required string Status { get; set; }  // Accepted, Tentative, Declined
}

/// <summary>
/// Audit service for calendar diagnostic logs.
/// </summary>
public interface ICalendarDiagnosticService : IAuditService
{
    Task<CalendarDiagnosticResult> GetDiagnosticLogsAsync(
        string mailboxEmail,
        int hoursBack = 24,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public class CalendarDiagnosticResult
{
    public required string Mailbox { get; set; }
    public List<CalendarDiagnosticEntry> Logs { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class CalendarDiagnosticEntry
{
    public DateTime? Timestamp { get; set; }
    public required string Operation { get; set; }  // Create, Modify, Delete, Accept, etc
    public required string ItemSubject { get; set; }
    public required string ModifiedBy { get; set; }
    public string? Details { get; set; }
}

/// <summary>
/// Audit service for mailbox permissions.
/// </summary>
public interface IMailboxPermissionsService : IAuditService
{
    Task<MailboxPermissionsResult> AuditMailboxAsync(
        string mailboxEmail,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public class MailboxPermissionsResult
{
    public required string Mailbox { get; set; }
    public List<MailboxPermissionEntry> Permissions { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class MailboxPermissionEntry
{
    public required string PermissionType { get; set; }  // Full Access, Send As, Send on Behalf
    public required string User { get; set; }
    public required string Rights { get; set; }
    public string? AutoMap { get; set; }
}

/// <summary>
/// Audit service for mail forwarding rules.
/// </summary>
public interface IMailForwardingAuditService : IAuditService
{
    Task<MailForwardingAuditResult> AuditForwardingRulesAsync(
        string? mailboxEmail = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public class MailForwardingAuditResult
{
    public List<MailForwardingEntry> ForwardingRules { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class MailForwardingEntry
{
    public required string Mailbox { get; set; }
    public required string ForwardingAddress { get; set; }
    public bool DeliverToMailbox { get; set; }
    public DateTime? CreatedDate { get; set; }
    public string? CreatedBy { get; set; }
}

/// <summary>
/// Audit service for shared mailboxes.
/// </summary>
public interface ISharedMailboxService : IAuditService
{
    Task<SharedMailboxAuditResult> GetSharedMailboxesAsync(
        string? searchTerm = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<SharedMailboxDetailsResult> GetMailboxDetailsAsync(
        string mailboxEmail,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public class SharedMailboxAuditResult
{
    public List<SharedMailboxInfo> Mailboxes { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class SharedMailboxInfo
{
    public required string DisplayName { get; set; }
    public required string Email { get; set; }
    public string? SizeMB { get; set; }
    public string? ItemCount { get; set; }
    public bool ArchiveActive { get; set; }
}

public class SharedMailboxDetailsResult
{
    public required string Email { get; set; }
    public List<SharedMailboxPermission> Permissions { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class SharedMailboxPermission
{
    public required string PermissionType { get; set; }  // Full Access, Send As, Send on Behalf
    public required string User { get; set; }
    public required string Rights { get; set; }
    public string? AutoMap { get; set; }
}

/// <summary>
/// Audit service for distribution groups and members.
/// </summary>
public interface IGroupExplorerService : IAuditService
{
    Task<GroupExplorerResult> SearchGroupsAsync(
        string searchTerm,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<GroupMembersResult> GetGroupMembersAsync(
        string groupIdentity,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public class GroupExplorerResult
{
    public List<GroupInfo> Groups { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class GroupInfo
{
    public required string DisplayName { get; set; }
    public required string Email { get; set; }
    public required string GroupType { get; set; }  // Distribution, Security, M365, etc
    public int MemberCount { get; set; }
}

public class GroupMembersResult
{
    public required string GroupName { get; set; }
    public List<GroupMember> Members { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class GroupMember
{
    public required string DisplayName { get; set; }
    public required string Email { get; set; }
    public required string RecipientType { get; set; }
    public bool IsNestedGroup { get; set; }
}

/// <summary>
/// Audit service for Microsoft Teams — lists teams and their membership.
/// Requires the MicrosoftTeams PowerShell module.
/// </summary>
public interface ITeamsExplorerService : IAuditService
{
    Task<TeamsListResult> GetTeamsAsync(
        string? filterByName = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<TeamMembersResult> GetTeamMembersAsync(
        string groupId,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public class TeamsListResult
{
    public List<TeamInfo> Teams { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class TeamInfo
{
    public required string DisplayName { get; set; }
    public required string GroupId { get; set; }
    public string? Description { get; set; }
    public string? Visibility { get; set; }   // Public / Private / HiddenMembership
    public bool IsArchived { get; set; }
    public int? MemberCount { get; set; }
}

public class TeamMembersResult
{
    public required string TeamName { get; set; }
    public List<TeamMemberEntry> Members { get; set; } = [];
    public string? Error { get; set; }
    public bool Success => Error == null;
}

public class TeamMemberEntry
{
    public required string DisplayName { get; set; }
    public required string UserPrincipalName { get; set; }
    public required string Role { get; set; }   // Owner / Member / Guest
}
