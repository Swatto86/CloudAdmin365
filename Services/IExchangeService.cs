using CloudAdmin365.Models;

namespace CloudAdmin365.Services;

/// <summary>
/// Interface for Exchange Online operations.
/// Abstracts the PowerShell/EXO SDK interaction behind a clean contract.
/// </summary>
public interface IExchangeService
{
    Task<ConnectionState> ConnectAsync(CancellationToken cancellationToken = default);
    Task DisconnectAsync();
    bool IsConnected { get; }
    
    // Room calendar permissions
    Task<AuditResult> GetRoomCalendarPermissionsAsync(string roomEmail, CancellationToken cancellationToken = default);
    
    // Shared mailboxes
    Task<List<SharedMailboxInfo>> GetSharedMailboxesAsync(CancellationToken cancellationToken = default);
    Task<List<DirectoryEntry>> GetSharedMailboxPermissionsAsync(string mailboxEmail, CancellationToken cancellationToken = default);
    
    // Groups
    Task<List<DistributionGroupInfo>> SearchGroupsAsync(string filter, CancellationToken cancellationToken = default);
    Task<List<GroupMember>> GetGroupMembersAsync(string groupIdentity, CancellationToken cancellationToken = default);
}

/// <summary>
/// Generic directory entry (permission, delegate, etc).
/// </summary>
public class DirectoryEntry
{
    public required string Type { get; set; }              // "Full Access", "Send As", "Send on Behalf"
    public required string User { get; set; }
    public required string Rights { get; set; }
    public string? AutoMap { get; set; }
}

/// <summary>
/// Represents connection state to Exchange and Graph services.
/// </summary>
public class ConnectionState
{
    public bool ExchangeConnected { get; set; }
    public bool GraphConnected { get; set; }
    public string? CurrentUser { get; set; }
}
