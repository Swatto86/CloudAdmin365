namespace CloudAdmin365.Models;

/// <summary>
/// Represents a calendar permission entry for a room mailbox.
/// </summary>
public class RoomPermission
{
    public required string Room { get; set; }
    public required string RoomEmail { get; set; }
    public required string Permission { get; set; }
    public required string Delegate { get; set; }
    public required string AccessLevel { get; set; }
    public string? GrantedVia { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}
