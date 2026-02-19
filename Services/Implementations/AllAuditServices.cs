namespace CloudAdmin365.Services.Implementations;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using CloudAdmin365.Utilities;
using System.Management.Automation;

/// <summary>
/// Exchange Online audit services using PowerShell cmdlets.
/// </summary>
/// <summary>
/// Audits calendar permissions for room mailboxes.
/// </summary>
public sealed class RoomPermissionsService : IRoomPermissionsService
{
    private const int MaxEmailLength = 256;
    private const int MaxResults = 1000;

    private readonly PowerShellHelper _powerShell;

    public string DisplayName => "Room Calendar Permissions";
    public string ServiceId => "room-permissions";
    public string Category => "Exchange";
    public string[] RequiredScopes => new[] { "Calendars.Read.All", "User.Read.All" };
    public string[] RequiredPowerShellModules => ["ExchangeOnlineManagement"];

    public RoomPermissionsService(PowerShellHelper powerShell)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
    }

    /// <summary>
    /// Check whether the PowerShell runspace is ready for use.
    /// </summary>
    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    /// <summary>
    /// Returns a short description used in the UI.
    /// </summary>
    public string GetDescription() => "Analyze room resource mailbox calendar permissions and delegations.";

    /// <summary>
    /// Audit room calendar permissions and return matching entries.
    /// </summary>
    public async Task<RoomPermissionResult> AuditRoomAsync(
        string roomEmail,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = AuditValidation.ValidateEmail(roomEmail, MaxEmailLength, "Room email");
        if (validationErrors.Count > 0)
        {
            return new RoomPermissionResult
            {
                Room = roomEmail ?? string.Empty,
                RoomEmail = roomEmail ?? string.Empty,
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 1, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Fetching calendar permissions..." });
            var identity = roomEmail + ":\\Calendar";

            var results = await _powerShell.ExecuteCommandAsync(
                "Get-MailboxFolderPermission",
                new Dictionary<string, object?>
                {
                    ["Identity"] = identity
                },
                cancellationToken).ConfigureAwait(false);

            var result = new RoomPermissionResult { Room = roomEmail, RoomEmail = roomEmail };
            var errors = new List<string>();
            foreach (var entry in results.Take(MaxResults))
            {
                try
                {
                    var user = entry.Properties["User"]?.Value?.ToString() ?? "Unknown";
                    var accessRights = entry.Properties["AccessRights"]?.Value?.ToString() ?? "";

                    result.Permissions.Add(new RoomPermissionEntry
                    {
                        PermissionType = "Calendar",
                        Delegate = user,
                        AccessLevel = accessRights,
                        GrantedVia = "Direct"
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some permission entries failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Room permission audit failed.", ex);
            return new RoomPermissionResult
            {
                Room = roomEmail ?? string.Empty,
                RoomEmail = roomEmail ?? string.Empty,
                Error = "Room permission audit failed. See log for details."
            };
        }
    }

}

/// <summary>
/// Audits room booking activity via the unified audit log.
/// </summary>
public sealed class RoomBookingAuditService : IRoomBookingAuditService
{
    private const int MaxEmailLength = 256;
    private const int MaxResults = 1000;

    private readonly PowerShellHelper _powerShell;

    public string DisplayName => "Room Booking Audit";
    public string ServiceId => "room-booking";
    public string Category => "Exchange";
    public string[] RequiredScopes => new[] { "Calendars.Read.All", "AuditLog.Read.All" };
    public string[] RequiredPowerShellModules => ["ExchangeOnlineManagement"];

    public RoomBookingAuditService(PowerShellHelper powerShell)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
    }

    /// <summary>
    /// Check whether the PowerShell runspace is ready for use.
    /// </summary>
    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    /// <summary>
    /// Returns a short description used in the UI.
    /// </summary>
    public string GetDescription() => "Review room resource mailbox bookings and calendar utilization.";

    /// <summary>
    /// Audit room bookings over a date range.
    /// </summary>
    public async Task<RoomBookingAuditResult> AuditRoomBookingsAsync(
        string roomEmail,
        DateTime? startDate = null,
        DateTime? endDate = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = AuditValidation.ValidateEmail(roomEmail, MaxEmailLength, "Room email");
        validationErrors.AddRange(AuditValidation.ValidateDateRange(startDate, endDate));

        var safeRoomEmail = roomEmail ?? string.Empty;
        if (validationErrors.Count > 0)
        {
            return new RoomBookingAuditResult
            {
                Room = safeRoomEmail,
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        var start = startDate ?? DateTime.UtcNow.AddDays(-7);
        var end = endDate ?? DateTime.UtcNow;

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Querying audit log..." });

            var results = await _powerShell.ExecuteCommandAsync(
                "Search-UnifiedAuditLog",
                new Dictionary<string, object?>
                {
                    ["StartDate"] = start,
                    ["EndDate"] = end,
                    ["ObjectIds"] = new[] { roomEmail ?? string.Empty },
                    ["Operations"] = new[] { "New-CalendarEvent", "Set-CalendarEvent", "Update-CalendarEvent" },
                    ["ResultSize"] = MaxResults
                },
                cancellationToken).ConfigureAwait(false);

            var result = new RoomBookingAuditResult { Room = roomEmail ?? string.Empty };
            var errors = new List<string>();
            foreach (var entry in results.Take(MaxResults))
            {
                try
                {
                    var auditData = entry.Properties["AuditData"]?.Value?.ToString();
                    var data = AuditValidation.ParseAuditData(auditData);

                    var subject = AuditValidation.GetJsonString(data, "Subject") ?? "(No Subject)";
                    var organizer = AuditValidation.GetJsonString(data, "UserId") ?? "Unknown";
                    var startTime = AuditValidation.GetJsonDateTime(data, "StartTime");
                    var endTime = AuditValidation.GetJsonDateTime(data, "EndTime");
                    var operation = AuditValidation.GetJsonString(data, "Operation") ?? "Update";

                    result.Bookings.Add(new RoomBookingEntry
                    {
                        Subject = subject,
                        Organizer = organizer,
                        StartTime = startTime,
                        EndTime = endTime,
                        Status = operation.Contains("New", StringComparison.OrdinalIgnoreCase) ? "Created" : "Updated"
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some booking entries failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Room booking audit failed.", ex);
            return new RoomBookingAuditResult
            {
                Room = safeRoomEmail,
                Error = "Room booking audit failed. See log for details."
            };
        }
    }
}

/// <summary>
/// Audits calendar diagnostic logs via the unified audit log.
/// </summary>
public sealed class CalendarDiagnosticService : ICalendarDiagnosticService
{
    private const int MaxEmailLength = 256;
    private const int MaxHoursBack = 720;
    private const int MaxResults = 1000;

    private readonly PowerShellHelper _powerShell;

    public string DisplayName => "Calendar Diagnostic Logs";
    public string ServiceId => "calendar-diagnostic";
    public string Category => "Exchange";
    public string[] RequiredScopes => new[] { "AuditLog.Read.All" };
    public string[] RequiredPowerShellModules => ["ExchangeOnlineManagement"];

    public CalendarDiagnosticService(PowerShellHelper powerShell)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
    }

    /// <summary>
    /// Check whether the PowerShell runspace is ready for use.
    /// </summary>
    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    /// <summary>
    /// Returns a short description used in the UI.
    /// </summary>
    public string GetDescription() => "Analyze calendar operations and event changes from unified audit logs.";

    /// <summary>
    /// Retrieve calendar diagnostic events for a mailbox.
    /// </summary>
    public async Task<CalendarDiagnosticResult> GetDiagnosticLogsAsync(
        string mailboxEmail,
        int hoursBack = 24,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = AuditValidation.ValidateEmail(mailboxEmail, MaxEmailLength, "Mailbox email");
        if (hoursBack <= 0 || hoursBack > MaxHoursBack)
        {
            validationErrors.Add($"Hours back must be between 1 and {MaxHoursBack}.");
        }

        var safeMailbox = mailboxEmail ?? string.Empty;
        if (validationErrors.Count > 0)
        {
            return new CalendarDiagnosticResult
            {
                Mailbox = safeMailbox,
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        var start = DateTime.UtcNow.AddHours(-hoursBack);
        var end = DateTime.UtcNow;

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Querying audit log..." });

            var results = await _powerShell.ExecuteCommandAsync(
                "Search-UnifiedAuditLog",
                new Dictionary<string, object?>
                {
                    ["StartDate"] = start,
                    ["EndDate"] = end,
                    ["UserIds"] = new[] { mailboxEmail ?? string.Empty },
                    ["Operations"] = new[] { "New-CalendarEvent", "Set-CalendarEvent", "Update-CalendarEvent", "Remove-CalendarEvent" },
                    ["ResultSize"] = MaxResults
                },
                cancellationToken).ConfigureAwait(false);

            var result = new CalendarDiagnosticResult { Mailbox = mailboxEmail ?? string.Empty };
            var errors = new List<string>();
            foreach (var entry in results.Take(MaxResults))
            {
                try
                {
                    var auditData = entry.Properties["AuditData"]?.Value?.ToString();
                    var data = AuditValidation.ParseAuditData(auditData);

                    result.Logs.Add(new CalendarDiagnosticEntry
                    {
                        Timestamp = AuditValidation.GetJsonDateTime(data, "CreationTime"),
                        Operation = AuditValidation.GetJsonString(data, "Operation") ?? "Unknown",
                        ItemSubject = AuditValidation.GetJsonString(data, "Subject") ?? "(No Subject)",
                        ModifiedBy = AuditValidation.GetJsonString(data, "UserId") ?? "Unknown",
                        Details = AuditValidation.GetJsonString(data, "ClientIP")
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some diagnostic entries failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Calendar diagnostic audit failed.", ex);
            return new CalendarDiagnosticResult
            {
                Mailbox = safeMailbox,
                Error = "Calendar diagnostic audit failed. See log for details."
            };
        }
    }
}

/// <summary>
/// Audits mailbox permissions and delegations.
/// </summary>
public sealed class MailboxPermissionsService : IMailboxPermissionsService
{
    private const int MaxEmailLength = 256;
    private const int MaxResults = 1000;

    private readonly PowerShellHelper _powerShell;

    public string DisplayName => "Mailbox Permissions";
    public string ServiceId => "mailbox-permissions";
    public string Category => "Exchange";
    public string[] RequiredScopes => new[] { "Mail.Read.All", "User.Read.All" };
    public string[] RequiredPowerShellModules => ["ExchangeOnlineManagement"];

    public MailboxPermissionsService(PowerShellHelper powerShell)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
    }

    /// <summary>
    /// Check whether the PowerShell runspace is ready for use.
    /// </summary>
    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    /// <summary>
    /// Returns a short description used in the UI.
    /// </summary>
    public string GetDescription() => "Analyze mailbox delegations, send-on-behalf, and inbox rule permissions.";

    /// <summary>
    /// Audit mailbox permissions and send-as/send-on-behalf delegations.
    /// </summary>
    public async Task<MailboxPermissionsResult> AuditMailboxAsync(
        string mailboxEmail,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = AuditValidation.ValidateEmail(mailboxEmail, MaxEmailLength, "Mailbox email");
        var safeMailbox = mailboxEmail ?? string.Empty;
        if (validationErrors.Count > 0)
        {
            return new MailboxPermissionsResult
            {
                Mailbox = safeMailbox,
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 3, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 3, Status = "Fetching mailbox permissions..." });
            var permissions = await _powerShell.ExecuteCommandAsync(
                "Get-MailboxPermission",
                new Dictionary<string, object?>
                {
                    ["Identity"] = mailboxEmail
                },
                cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 2, Total = 3, Status = "Fetching send-as permissions..." });
            var sendAsPermissions = await _powerShell.ExecuteCommandAsync(
                "Get-RecipientPermission",
                new Dictionary<string, object?>
                {
                    ["Identity"] = mailboxEmail
                },
                cancellationToken).ConfigureAwait(false);

            var sendOnBehalf = await _powerShell.ExecuteCommandAsync(
                "Get-Mailbox",
                new Dictionary<string, object?>
                {
                    ["Identity"] = mailboxEmail
                },
                cancellationToken).ConfigureAwait(false);

            var result = new MailboxPermissionsResult { Mailbox = mailboxEmail ?? string.Empty };
            var errors = new List<string>();
            foreach (var entry in permissions.Take(MaxResults))
            {
                try
                {
                    result.Permissions.Add(new MailboxPermissionEntry
                    {
                        PermissionType = "MailboxPermission",
                        User = entry.Properties["User"]?.Value?.ToString() ?? "Unknown",
                        Rights = entry.Properties["AccessRights"]?.Value?.ToString() ?? "",
                        AutoMap = entry.Properties["AutoMapping"]?.Value?.ToString()
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            foreach (var entry in sendAsPermissions.Take(MaxResults))
            {
                try
                {
                    result.Permissions.Add(new MailboxPermissionEntry
                    {
                        PermissionType = "Send As",
                        User = entry.Properties["Trustee"]?.Value?.ToString() ?? "Unknown",
                        Rights = entry.Properties["AccessRights"]?.Value?.ToString() ?? "",
                        AutoMap = null
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            var sendOnBehalfUsers = sendOnBehalf.FirstOrDefault()?.Properties["GrantSendOnBehalfTo"]?.Value;
            if (sendOnBehalfUsers is IEnumerable<object> sendOnBehalfList)
            {
                foreach (var user in sendOnBehalfList)
                {
                    try
                    {
                        result.Permissions.Add(new MailboxPermissionEntry
                        {
                            PermissionType = "Send On Behalf",
                            User = user?.ToString() ?? "Unknown",
                            Rights = "SendOnBehalf",
                            AutoMap = null
                        });
                    }
                    catch (Exception ex)
                    {
                        AuditValidation.TryAddError(errors, ex.Message);
                    }
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some permission entries failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 3, Total = 3, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Mailbox permission audit failed.", ex);
            return new MailboxPermissionsResult
            {
                Mailbox = safeMailbox,
                Error = "Mailbox permission audit failed. See log for details."
            };
        }
    }
}

/// <summary>
/// Audits mailbox forwarding configuration.
/// </summary>
public sealed class MailForwardingAuditService : IMailForwardingAuditService
{
    private const int MaxEmailLength = 256;
    private const int MaxResults = 1000;

    private readonly PowerShellHelper _powerShell;

    public string DisplayName => "Mail Forwarding Audit";
    public string ServiceId => "mail-forwarding";
    public string Category => "Exchange";
    public string[] RequiredScopes => new[] { "Mail.Read.All" };
    public string[] RequiredPowerShellModules => ["ExchangeOnlineManagement"];

    public MailForwardingAuditService(PowerShellHelper powerShell)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
    }

    /// <summary>
    /// Check whether the PowerShell runspace is ready for use.
    /// </summary>
    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    /// <summary>
    /// Returns a short description used in the UI.
    /// </summary>
    public string GetDescription() => "Identify mail forwarding rules and delivery settings.";

    /// <summary>
    /// Audit mail forwarding rules for a mailbox or all mailboxes.
    /// </summary>
    public async Task<MailForwardingAuditResult> AuditForwardingRulesAsync(
        string? mailboxEmail = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = new List<string>();
        if (!string.IsNullOrWhiteSpace(mailboxEmail))
        {
            validationErrors.AddRange(AuditValidation.ValidateEmail(mailboxEmail, MaxEmailLength, "Mailbox email"));
        }

        if (validationErrors.Count > 0)
        {
            return new MailForwardingAuditResult
            {
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Querying mailboxes..." });

            var parameters = new Dictionary<string, object?>
            {
                ["ResultSize"] = MaxResults
            };
            if (!string.IsNullOrWhiteSpace(mailboxEmail))
            {
                parameters["Identity"] = mailboxEmail;
                parameters.Remove("ResultSize");
            }

            var results = await _powerShell.ExecuteCommandAsync(
                "Get-Mailbox",
                parameters,
                cancellationToken).ConfigureAwait(false);

            var result = new MailForwardingAuditResult();
            var errors = new List<string>();
            foreach (var entry in results.Take(MaxResults))
            {
                try
                {
                    var identity = entry.Properties["PrimarySmtpAddress"]?.Value?.ToString()
                        ?? entry.Properties["Identity"]?.Value?.ToString()
                        ?? "Unknown";

                    var forwardingAddress = entry.Properties["ForwardingAddress"]?.Value?.ToString();
                    var smtpForwarding = entry.Properties["ForwardingSmtpAddress"]?.Value?.ToString();
                    var deliverToMailbox = entry.Properties["DeliverToMailboxAndForward"]?.Value as bool? ?? false;

                    if (string.IsNullOrWhiteSpace(forwardingAddress) && string.IsNullOrWhiteSpace(smtpForwarding))
                    {
                        continue;
                    }

                    result.ForwardingRules.Add(new MailForwardingEntry
                    {
                        Mailbox = identity,
                        ForwardingAddress = forwardingAddress ?? smtpForwarding ?? "",
                        DeliverToMailbox = deliverToMailbox,
                        CreatedDate = null,
                        CreatedBy = null
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some forwarding entries failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Mail forwarding audit failed.", ex);
            return new MailForwardingAuditResult
            {
                Error = "Mail forwarding audit failed. See log for details."
            };
        }
    }
}

/// <summary>
/// Audits shared mailboxes and their permissions.
/// </summary>
public sealed class SharedMailboxService : ISharedMailboxService
{
    private const int MaxEmailLength = 256;
    private const int MaxResults = 1000;

    private readonly PowerShellHelper _powerShell;

    public string DisplayName => "Shared Mailbox Explorer";
    public string ServiceId => "shared-mailbox";
    public string Category => "Exchange";
    public string[] RequiredScopes => new[] { "Mail.Read.All", "User.Read.All" };
    public string[] RequiredPowerShellModules => ["ExchangeOnlineManagement"];

    public SharedMailboxService(PowerShellHelper powerShell)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
    }

    /// <summary>
    /// Check whether the PowerShell runspace is ready for use.
    /// </summary>
    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    /// <summary>
    /// Returns a short description used in the UI.
    /// </summary>
    public string GetDescription() => "Discover shared mailboxes and their member access.";

    /// <summary>
    /// Retrieve shared mailboxes, optionally filtered by search term.
    /// </summary>
    public async Task<SharedMailboxAuditResult> GetSharedMailboxesAsync(
        string? searchTerm = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (!string.IsNullOrWhiteSpace(searchTerm) && searchTerm.Length > MaxEmailLength)
        {
            return new SharedMailboxAuditResult
            {
                Error = $"Search term exceeds max length {MaxEmailLength}."
            };
        }

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Querying shared mailboxes..." });

            var parameters = new Dictionary<string, object?>
            {
                ["RecipientTypeDetails"] = "SharedMailbox",
                ["ResultSize"] = MaxResults
            };

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                var filterValue = AuditValidation.EscapeFilterValue(searchTerm);
                parameters["Filter"] = $"DisplayName -like '*{filterValue}*'";
            }

            var results = await _powerShell.ExecuteCommandAsync(
                "Get-Mailbox",
                parameters,
                cancellationToken).ConfigureAwait(false);

            var result = new SharedMailboxAuditResult();
            var errors = new List<string>();
            foreach (var entry in results.Take(MaxResults))
            {
                try
                {
                    result.Mailboxes.Add(new SharedMailboxInfo
                    {
                        DisplayName = entry.Properties["DisplayName"]?.Value?.ToString() ?? "Unknown",
                        Email = entry.Properties["PrimarySmtpAddress"]?.Value?.ToString() ?? "",
                        SizeMB = null,
                        ItemCount = null,
                        ArchiveActive = string.Equals(entry.Properties["ArchiveStatus"]?.Value?.ToString(), "Active", StringComparison.OrdinalIgnoreCase)
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some shared mailboxes failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Shared mailbox audit failed.", ex);
            return new SharedMailboxAuditResult
            {
                Error = "Shared mailbox audit failed. See log for details."
            };
        }
    }

    /// <summary>
    /// Retrieve permissions for a specific shared mailbox.
    /// </summary>
    public async Task<SharedMailboxDetailsResult> GetMailboxDetailsAsync(
        string mailboxEmail,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = AuditValidation.ValidateEmail(mailboxEmail, MaxEmailLength, "Mailbox email");
        if (validationErrors.Count > 0)
        {
            return new SharedMailboxDetailsResult
            {
                Email = mailboxEmail ?? string.Empty,
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Fetching mailbox permissions..." });

            var permissions = await _powerShell.ExecuteCommandAsync(
                "Get-MailboxPermission",
                new Dictionary<string, object?>
                {
                    ["Identity"] = mailboxEmail
                },
                cancellationToken).ConfigureAwait(false);

            var sendAs = await _powerShell.ExecuteCommandAsync(
                "Get-RecipientPermission",
                new Dictionary<string, object?>
                {
                    ["Identity"] = mailboxEmail
                },
                cancellationToken).ConfigureAwait(false);

            var result = new SharedMailboxDetailsResult { Email = mailboxEmail };
            var errors = new List<string>();
            foreach (var entry in permissions.Take(MaxResults))
            {
                try
                {
                    result.Permissions.Add(new SharedMailboxPermission
                    {
                        PermissionType = "MailboxPermission",
                        User = entry.Properties["User"]?.Value?.ToString() ?? "Unknown",
                        Rights = entry.Properties["AccessRights"]?.Value?.ToString() ?? "",
                        AutoMap = entry.Properties["AutoMapping"]?.Value?.ToString()
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            foreach (var entry in sendAs.Take(MaxResults))
            {
                try
                {
                    result.Permissions.Add(new SharedMailboxPermission
                    {
                        PermissionType = "Send As",
                        User = entry.Properties["Trustee"]?.Value?.ToString() ?? "Unknown",
                        Rights = entry.Properties["AccessRights"]?.Value?.ToString() ?? "",
                        AutoMap = null
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some permission entries failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Shared mailbox details failed.", ex);
            return new SharedMailboxDetailsResult
            {
                Email = mailboxEmail ?? string.Empty,
                Error = "Shared mailbox details failed. See log for details."
            };
        }
    }
}

/// <summary>
/// Audits group metadata and membership.
/// </summary>
public sealed class GroupExplorerService : IGroupExplorerService
{
    private const int MaxSearchLength = 256;
    private const int MaxResults = 500;

    private readonly PowerShellHelper _powerShell;

    public string DisplayName => "Group Membership Explorer";
    public string ServiceId => "group-explorer";
    public string Category => "Exchange";
    public string[] RequiredScopes => new[] { "Group.Read.All" };
    public string[] RequiredPowerShellModules => ["ExchangeOnlineManagement"];

    public GroupExplorerService(PowerShellHelper powerShell)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
    }

    /// <summary>
    /// Check whether the PowerShell runspace is ready for use.
    /// </summary>
    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    /// <summary>
    /// Returns a short description used in the UI.
    /// </summary>
    public string GetDescription() => "Analyze Microsoft 365 group membership and owners.";

    /// <summary>
    /// Search for distribution and unified groups by display name.
    /// </summary>
    public async Task<GroupExplorerResult> SearchGroupsAsync(
        string searchTerm,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = new List<string>();
        if (string.IsNullOrWhiteSpace(searchTerm))
        {
            validationErrors.Add("Search term is required.");
        }
        else if (searchTerm.Length > MaxSearchLength)
        {
            validationErrors.Add($"Search term exceeds max length {MaxSearchLength}.");
        }

        if (validationErrors.Count > 0)
        {
            return new GroupExplorerResult
            {
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Searching groups..." });

            var filterValue = AuditValidation.EscapeFilterValue(searchTerm);
            var filter = $"DisplayName -like '*{filterValue}*'";
            var result = new GroupExplorerResult();
            var errors = new List<string>();

            var unifiedGroups = await _powerShell.ExecuteCommandAsync(
                "Get-UnifiedGroup",
                new Dictionary<string, object?>
                {
                    ["Filter"] = filter,
                    ["ResultSize"] = MaxResults
                },
                cancellationToken).ConfigureAwait(false);

            foreach (var entry in unifiedGroups.Take(MaxResults))
            {
                try
                {
                    result.Groups.Add(new GroupInfo
                    {
                        DisplayName = entry.Properties["DisplayName"]?.Value?.ToString() ?? "Unknown",
                        Email = entry.Properties["PrimarySmtpAddress"]?.Value?.ToString() ?? "",
                        GroupType = "Unified",
                        MemberCount = 0
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            var distributionGroups = await _powerShell.ExecuteCommandAsync(
                "Get-DistributionGroup",
                new Dictionary<string, object?>
                {
                    ["Filter"] = filter,
                    ["ResultSize"] = MaxResults
                },
                cancellationToken).ConfigureAwait(false);

            foreach (var entry in distributionGroups.Take(MaxResults))
            {
                try
                {
                    result.Groups.Add(new GroupInfo
                    {
                        DisplayName = entry.Properties["DisplayName"]?.Value?.ToString() ?? "Unknown",
                        Email = entry.Properties["PrimarySmtpAddress"]?.Value?.ToString() ?? "",
                        GroupType = "Distribution",
                        MemberCount = 0
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some groups failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Group search failed.", ex);
            return new GroupExplorerResult
            {
                Error = "Group search failed. See log for details."
            };
        }
    }

    /// <summary>
    /// Retrieve members for a distribution or unified group.
    /// </summary>
    public async Task<GroupMembersResult> GetGroupMembersAsync(
        string groupIdentity,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var validationErrors = new List<string>();
        if (string.IsNullOrWhiteSpace(groupIdentity))
        {
            validationErrors.Add("Group identity is required.");
        }
        else if (groupIdentity.Length > MaxSearchLength)
        {
            validationErrors.Add($"Group identity exceeds max length {MaxSearchLength}.");
        }

        if (validationErrors.Count > 0)
        {
            return new GroupMembersResult
            {
                GroupName = groupIdentity ?? string.Empty,
                Error = AuditValidation.BuildValidationError(validationErrors)
            };
        }

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Exchange Online..." });
            await _powerShell.EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Fetching group members..." });

            var recipient = await _powerShell.ExecuteCommandAsync(
                "Get-Recipient",
                new Dictionary<string, object?>
                {
                    ["Identity"] = groupIdentity
                },
                cancellationToken).ConfigureAwait(false);

            var recipientType = recipient.FirstOrDefault()?.Properties["RecipientTypeDetails"]?.Value?.ToString() ?? string.Empty;
            var isUnified = recipientType.Contains("GroupMailbox", StringComparison.OrdinalIgnoreCase);

            IReadOnlyList<PSObject> members;
            if (isUnified)
            {
                members = await _powerShell.ExecuteCommandAsync(
                    "Get-UnifiedGroupLinks",
                    new Dictionary<string, object?>
                    {
                        ["Identity"] = groupIdentity,
                        ["LinkType"] = "Members"
                    },
                    cancellationToken).ConfigureAwait(false);
            }
            else
            {
                members = await _powerShell.ExecuteCommandAsync(
                    "Get-DistributionGroupMember",
                    new Dictionary<string, object?>
                    {
                        ["Identity"] = groupIdentity
                    },
                    cancellationToken).ConfigureAwait(false);
            }

            var result = new GroupMembersResult { GroupName = groupIdentity };
            var errors = new List<string>();
            foreach (var entry in members.Take(MaxResults))
            {
                try
                {
                    result.Members.Add(new GroupMember
                    {
                        DisplayName = entry.Properties["DisplayName"]?.Value?.ToString() ?? "Unknown",
                        Email = entry.Properties["PrimarySmtpAddress"]?.Value?.ToString() ?? "",
                        RecipientType = entry.Properties["RecipientTypeDetails"]?.Value?.ToString()
                            ?? entry.Properties["RecipientType"]?.Value?.ToString()
                            ?? "Unknown",
                        IsNestedGroup = string.Equals(entry.Properties["RecipientTypeDetails"]?.Value?.ToString(), "GroupMailbox", StringComparison.OrdinalIgnoreCase)
                    });
                }
                catch (Exception ex)
                {
                    AuditValidation.TryAddError(errors, ex.Message);
                }
            }

            if (errors.Count > 0)
            {
                result.Error = AuditValidation.BuildBatchError("Some group members failed to parse", errors);
                AppLogger.WriteError(result.Error);
            }

            progress?.Report(new AuditProgress { Current = 2, Total = 2, Status = "Complete" });
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Group member lookup failed.", ex);
            return new GroupMembersResult
            {
                GroupName = groupIdentity ?? string.Empty,
                Error = "Group member lookup failed. See log for details."
            };
        }
    }
}

/// <summary>
/// Shared validation and parsing helpers for audits.
/// </summary>
internal static class AuditValidation
{
    public const int MaxErrors = 50;
    public static string EscapeFilterValue(string value)
    {
        return value.Replace("'", "''", StringComparison.Ordinal);
    }

    public static async Task<bool> TryInitializeAsync(PowerShellHelper powerShell, CancellationToken cancellationToken)
    {
        try
        {
            await powerShell.InitializeAsync(cancellationToken).ConfigureAwait(false);
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("PowerShell availability check failed.", ex);
            return false;
        }
    }

    public static List<string> ValidateEmail(string? value, int maxLength, string fieldName)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(value))
        {
            errors.Add($"{fieldName} is required.");
            return errors;
        }

        if (value.Length > maxLength)
        {
            errors.Add($"{fieldName} exceeds max length {maxLength}.");
        }

        return errors;
    }

    public static List<string> ValidateDateRange(DateTime? startDate, DateTime? endDate)
    {
        var errors = new List<string>();
        if (startDate.HasValue && endDate.HasValue && startDate.Value > endDate.Value)
        {
            errors.Add("Start date must be earlier than end date.");
        }

        return errors;
    }

    public static string BuildValidationError(List<string> errors)
    {
        return "Validation failed: " + string.Join(" ", errors);
    }

    public static void TryAddError(List<string> errors, string message)
    {
        if (errors.Count >= MaxErrors)
        {
            return;
        }

        errors.Add(message);
    }

    public static string BuildBatchError(string summary, List<string> errors)
    {
        if (errors.Count == 0)
        {
            return summary;
        }

        var detail = string.Join(" | ", errors);
        return $"{summary} ({errors.Count} errors). Details: {detail}";
    }

    public static JsonElement? ParseAuditData(string? auditData)
    {
        if (string.IsNullOrWhiteSpace(auditData))
        {
            return null;
        }

        try
        {
            var doc = JsonDocument.Parse(auditData);
            return doc.RootElement.Clone();
        }
        catch
        {
            return null;
        }
    }

    public static string? GetJsonString(JsonElement? element, string propertyName)
    {
        if (element == null)
        {
            return null;
        }

        if (element.Value.TryGetProperty(propertyName, out var value))
        {
            return value.GetString();
        }

        return null;
    }

    public static DateTime? GetJsonDateTime(JsonElement? element, string propertyName)
    {
        var text = GetJsonString(element, propertyName);
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var value))
        {
            return value.ToUniversalTime();
        }

        return null;
    }
}

/// <summary>
/// Azure AD/Entra ID user management via Microsoft Graph PowerShell.
/// Uses the existing Microsoft Graph authentication session.
/// </summary>
public sealed class AzureADService : IAzureADService
{
    private const int MaxFilterLength = 200;
    private const int MaxUserResults  = 5000;
    private const string GraphScope   = "https://graph.microsoft.com/.default";

    private readonly PowerShellHelper _powerShell;
    private readonly IAuthService _authService;
    private readonly SemaphoreSlim _connectMutex = new(1, 1);
    private bool _graphConnected;

    public string DisplayName   => "Azure AD Users";
    public string ServiceId     => "azuread-users";
    public string Category      => "Azure AD";
    public string[] RequiredScopes => ["User.Read.All", "Directory.Read.All"];
    public string[] RequiredPowerShellModules => ["Microsoft.Graph"];

    public AzureADService(PowerShellHelper powerShell, IAuthService authService)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
        _authService = authService ?? throw new ArgumentNullException(nameof(authService));
    }

    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    public string GetDescription()
        => "Browse Azure AD users: list all users with account status,job title, and user type. Filterable by display name.";

    public async Task<AzureADUsersResult> GetUsersAsync(
        string? filterByName = null,
        bool includeGuestUsers = true,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (filterByName != null && filterByName.Length > MaxFilterLength)
            return new AzureADUsersResult { Error = $"Filter exceeds maximum length ({MaxFilterLength})." };

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Microsoft Graph..." });
            await EnsureGraphConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Fetching users from Azure AD..." });

            var filter = string.Empty;
            if (!string.IsNullOrWhiteSpace(filterByName))
            {
                filter = $"startsWith(displayName,'{AuditValidation.EscapeFilterValue(filterByName)}')";
                if (!includeGuestUsers)
                    filter += " and userType eq 'Member'";
            }
            else if (!includeGuestUsers)
            {
                filter = "userType eq 'Member'";
            }

            var parameters = new Dictionary<string, object?>
            {
                ["All"] = null,
                ["Property"] = new[] { "Id", "DisplayName", "UserPrincipalName", "UserType", "AccountEnabled", "JobTitle", "Department", "Mail" }
            };
            if (!string.IsNullOrEmpty(filter))
                parameters["Filter"] = filter;

            var results = await _powerShell.ExecuteRawCommandAsync(
                "Get-MgUser",
                parameters,
                cancellationToken).ConfigureAwait(false);

            var users = results
                .Take(MaxUserResults)
                .Select(PSObjectToAzureADUser)
                .Where(u => u != null)
                .Cast<AzureADUser>()
                .OrderBy(u => u.DisplayName)
                .ToList();

            AppLogger.WriteInfo($"AzureADService: found {users.Count} user(s).");
            return new AzureADUsersResult { Users = users };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("AzureADService.GetUsersAsync failed.", ex);
            return new AzureADUsersResult { Error = ex.Message };
        }
    }

    private async Task EnsureGraphConnectedAsync(CancellationToken cancellationToken)
    {
        if (_graphConnected)
            return;

        await _connectMutex.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_graphConnected)
                return;

            AppLogger.WriteInfo("Connecting to Microsoft Graph using existing authentication...");
            
            // Get access token from existing Azure auth (reuses MSAL session)
            var accessToken = await _authService.GetAccessTokenAsync(GraphScope, cancellationToken)
                .ConfigureAwait(false);

            // Connect using the access token (no interactive prompt)
            await _powerShell.ExecuteRawCommandAsync(
                "Connect-MgGraph",
                new Dictionary<string, object?> { ["AccessToken"] = accessToken },
                cancellationToken).ConfigureAwait(false);

            _graphConnected = true;
            AppLogger.WriteInfo("Microsoft Graph connection established.");
        }
        finally
        {
            _connectMutex.Release();
        }
    }

    private static AzureADUser? PSObjectToAzureADUser(PSObject psObject)
    {
        try
        {
            return new AzureADUser
            {
                DisplayName       = GetPSProperty(psObject, "DisplayName") ?? "(No Name)",
                UserPrincipalName = GetPSProperty(psObject, "UserPrincipalName") ?? string.Empty,
                Id                = GetPSProperty(psObject, "Id") ?? string.Empty,
                UserType          = GetPSProperty(psObject, "UserType") ?? "Member",
                AccountEnabled    = string.Equals(GetPSProperty(psObject, "AccountEnabled"), "True", StringComparison.OrdinalIgnoreCase),
                JobTitle          = GetPSProperty(psObject, "JobTitle"),
                Department        = GetPSProperty(psObject, "Department"),
                Mail              = GetPSProperty(psObject, "Mail")
            };
        }
        catch
        {
            return null;
        }
    }

    private static string? GetPSProperty(PSObject psObject, string propertyName)
    {
        if (psObject?.Members[propertyName]?.Value is var value && value != null)
            return value.ToString();
        return null;
    }
}

/// <summary>
/// Microsoft Intune device management via Microsoft Graph PowerShell.
/// Uses the existing Microsoft Graph authentication session.
/// </summary>
public sealed class IntuneService : IIntuneService
{
    private const int MaxFilterLength = 200;
    private const int MaxDeviceResults = 5000;
    private const string GraphScope   = "https://graph.microsoft.com/.default";

    private readonly PowerShellHelper _powerShell;
    private readonly IAuthService _authService;
    private readonly SemaphoreSlim _connectMutex = new(1, 1);
    private bool _graphConnected;

    public string DisplayName   => "Intune Devices";
    public string ServiceId     => "intune-devices";
    public string Category      => "Intune";
    public string[] RequiredScopes => ["DeviceManagementManagedDevices.Read.All"];
    public string[] RequiredPowerShellModules => ["Microsoft.Graph"];

    public IntuneService(PowerShellHelper powerShell, IAuthService authService)
    {
        _powerShell = powerShell ?? throw new ArgumentNullException(nameof(powerShell));
        _authService = authService ?? throw new ArgumentNullException(nameof(authService));
    }

    public Task<bool> IsAvailableAsync(CancellationToken cancellationToken = default)
        => AuditValidation.TryInitializeAsync(_powerShell, cancellationToken);

    public string GetDescription()
        => "Browse Intune managed devices: list all devices with compliance status, OS version, and last sync time. Filterable by device name.";

    public async Task<IntuneDevicesResult> GetManagedDevicesAsync(
        string? filterByName = null,
        IProgress<AuditProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (filterByName != null && filterByName.Length > MaxFilterLength)
            return new IntuneDevicesResult { Error = $"Filter exceeds maximum length ({MaxFilterLength})." };

        try
        {
            progress?.Report(new AuditProgress { Current = 0, Total = 2, Status = "Connecting to Microsoft Graph..." });
            await EnsureGraphConnectedAsync(cancellationToken).ConfigureAwait(false);

            progress?.Report(new AuditProgress { Current = 1, Total = 2, Status = "Fetching managed devices from Intune..." });

            var parameters = new Dictionary<string, object?>
            {
                ["All"] = null
            };
            if (!string.IsNullOrWhiteSpace(filterByName))
            {
                parameters["Filter"] = $"startsWith(deviceName,'{AuditValidation.EscapeFilterValue(filterByName)}')";
            }

            var results = await _powerShell.ExecuteRawCommandAsync(
                "Get-MgDeviceManagementManagedDevice",
                parameters,
                cancellationToken).ConfigureAwait(false);

            var devices = results
                .Take(MaxDeviceResults)
                .Select(PSObjectToIntuneDevice)
                .Where(d => d != null)
                .Cast<IntuneDevice>()
                .OrderBy(d => d.DeviceName)
                .ToList();

            AppLogger.WriteInfo($"IntuneService: found {devices.Count} managed device(s).");
            return new IntuneDevicesResult { Devices = devices };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("IntuneService.GetManagedDevicesAsync failed.", ex);
            return new IntuneDevicesResult { Error = ex.Message };
        }
    }

    private async Task EnsureGraphConnectedAsync(CancellationToken cancellationToken)
    {
        if (_graphConnected)
            return;

        await _connectMutex.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_graphConnected)
                return;

            AppLogger.WriteInfo("Connecting to Microsoft Graph (Intune) using existing authentication...");
            
            // Get access token from existing Azure auth (reuses MSAL session)
            var accessToken = await _authService.GetAccessTokenAsync(GraphScope, cancellationToken)
                .ConfigureAwait(false);

            // Connect using the access token (no interactive prompt)
            await _powerShell.ExecuteRawCommandAsync(
                "Connect-MgGraph",
                new Dictionary<string, object?> { ["AccessToken"] = accessToken },
                cancellationToken).ConfigureAwait(false);

            _graphConnected = true;
            AppLogger.WriteInfo("Microsoft Graph connection (Intune) established.");
        }
        finally
        {
            _connectMutex.Release();
        }
    }

    private static IntuneDevice? PSObjectToIntuneDevice(PSObject psObject)
    {
        try
        {
            var lastSync = GetPSProperty(psObject, "LastSyncDateTime");
            DateTime? lastSyncDate = null;
            if (lastSync != null && DateTime.TryParse(lastSync, out var parsed))
                lastSyncDate = parsed;

            return new IntuneDevice
            {
                DeviceName       = GetPSProperty(psObject, "DeviceName") ?? "(Unknown)",
                Id               = GetPSProperty(psObject, "Id") ?? string.Empty,
                OperatingSystem  = GetPSProperty(psObject, "OperatingSystem") ?? "Unknown",
                OSVersion        = GetPSProperty(psObject, "OSVersion") ?? "Unknown",
                ComplianceState  = GetPSProperty(psObject, "ComplianceState") ?? "Unknown",
                ManagementState  = GetPSProperty(psObject, "ManagementState") ?? "Unknown",
                LastSyncDateTime = lastSyncDate,
                UserPrincipalName = GetPSProperty(psObject, "UserPrincipalName"),
                Model            = GetPSProperty(psObject, "Model")
            };
        }
        catch
        {
            return null;
        }
    }

    private static string? GetPSProperty(PSObject psObject, string propertyName)
    {
        if (psObject?.Members[propertyName]?.Value is var value && value != null)
            return value.ToString();
        return null;
    }
}
