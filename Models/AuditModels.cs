using CloudAdmin365.Services;

namespace CloudAdmin365.Models;

/// <summary>
/// Represents application state and audit result data.
/// </summary>
public class AuditResult
{
    public bool Success { get; set; }
    public string? Error { get; set; }
    public int TotalItems { get; set; }
    public List<object> Items { get; set; } = [];
    public DateTime? CompletedAt { get; set; }
}

/// <summary>
/// Represents a shared mailbox with statistics.
/// </summary>
public class SharedMailboxInfo
{
    public required string DisplayName { get; set; }
    public required string PrimarySmtpAddress { get; set; }
    public string? SizeMB { get; set; }
    public string? ItemCount { get; set; }
    public bool ArchiveActive { get; set; }
}

/// <summary>
/// Represents a distribution group with members.
/// </summary>
public class DistributionGroupInfo
{
    public required string DisplayName { get; set; }
    public required string PrimarySmtpAddress { get; set; }
    public required string RecipientType { get; set; }
    public int MemberCount { get; set; }
}

/// <summary>
/// Represents a group member.
/// </summary>
public class GroupMember
{
    public required string DisplayName { get; set; }
    public required string Email { get; set; }
    public required string RecipientType { get; set; }
    public string? NestedGroupInfo { get; set; }
}
