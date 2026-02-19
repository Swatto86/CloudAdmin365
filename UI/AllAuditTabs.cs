namespace CloudAdmin365.UI;

using CloudAdmin365.Services;
using CloudAdmin365.Utilities;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

/// <summary>
/// Tab for room booking audit visualization.
/// </summary>
public class RoomBookingTab : AuditTabBase
{
    private readonly IRoomBookingAuditService _service;
    private TextBox? _roomEmailInput;

    public RoomBookingTab(IRoomBookingAuditService service)
        : base("Room Booking Audit", new[] { "Subject", "Organizer", "Start Time", "Status" },
               new[] { 30, 30, 20, 20 })
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        var lblEmail = new Label { Text = "Room Email:", Location = new Point(5, 10), AutoSize = true, Font = AppTheme.DefaultFont };
        _roomEmailInput = new TextBox { Location = new Point(100, 10), Size = new Size(320, 22), Font = AppTheme.DefaultFont };
        InputPanel.Controls.AddRange(new Control[] { lblEmail, _roomEmailInput });

        var lblDesc = new Label { Text = _service.GetDescription(), Location = new Point(5, 38), MaximumSize = new Size(1000, 0), AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8) };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_roomEmailInput?.Text))
        {
            UpdateStatus("Please enter a room email address.", AppTheme.ThemeRed);
            return;
        }

        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
            });

            var result = await _service.AuditRoomBookingsAsync(_roomEmailInput.Text, null, null, progress, cancellationToken);

            if (string.IsNullOrEmpty(result.Error))
            {
                RenderResults(result);
                UpdateStatus($"Complete: {result.Bookings.Count} booking(s) found", Color.FromArgb(0, 100, 0));
                EnableExport();
            }
            else
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not RoomBookingAuditResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var booking in result.Bookings)
        {
            ResultsGrid.Rows.Add(booking.Subject, booking.Organizer, booking.StartTime, booking.Status);
        }
    }
}

/// <summary>
/// Tab for calendar diagnostic logs visualization.
/// </summary>
public class CalendarDiagnosticTab : AuditTabBase
{
    private readonly ICalendarDiagnosticService _service;
    private TextBox? _mailboxEmailInput;
    private NumericUpDown? _hoursBackInput;

    public CalendarDiagnosticTab(ICalendarDiagnosticService service)
        : base("Calendar Diagnostic Logs", new[] { "Timestamp", "Operation", "Subject", "Details" },
               new[] { 20, 20, 30, 30 })
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        var lblEmail = new Label { Text = "Mailbox Email:", Location = new Point(5, 10), AutoSize = true, Font = AppTheme.DefaultFont };
        _mailboxEmailInput = new TextBox { Location = new Point(100, 10), Size = new Size(320, 22), Font = AppTheme.DefaultFont };

        var lblHours = new Label { Text = "Hours Back:", Location = new Point(430, 10), AutoSize = true, Font = AppTheme.DefaultFont };
        _hoursBackInput = new NumericUpDown { Location = new Point(515, 10), Size = new Size(80, 22), Value = 24, Minimum = 1, Maximum = 168 };

        InputPanel.Controls.AddRange(new Control[] { lblEmail, _mailboxEmailInput, lblHours, _hoursBackInput });

        var lblDesc = new Label { Text = _service.GetDescription(), Location = new Point(5, 38), MaximumSize = new Size(1000, 0), AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8) };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_mailboxEmailInput?.Text))
        {
            UpdateStatus("Please enter a mailbox email address.", AppTheme.ThemeRed);
            return;
        }

        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
            });

            var result = await _service.GetDiagnosticLogsAsync(_mailboxEmailInput.Text, (int)(_hoursBackInput?.Value ?? 24), progress, cancellationToken);

            if (string.IsNullOrEmpty(result.Error))
            {
                RenderResults(result);
                UpdateStatus($"Complete: {result.Logs.Count} log(s) found", Color.FromArgb(0, 100, 0));
                EnableExport();
            }
            else
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not CalendarDiagnosticResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var log in result.Logs)
        {
            ResultsGrid.Rows.Add(log.Timestamp, log.Operation, log.ItemSubject, log.Details);
        }
    }
}

/// <summary>
/// Tab for mailbox permissions visualization.
/// </summary>
public class MailboxPermissionsTab : AuditTabBase
{
    private readonly IMailboxPermissionsService _service;
    private TextBox? _mailboxEmailInput;

    public MailboxPermissionsTab(IMailboxPermissionsService service)
        : base("Mailbox Permissions", new[] { "Permission Type", "User", "Rights", "AutoMap" },
               new[] { 25, 35, 20, 20 })
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        var lblEmail = new Label { Text = "Mailbox Email:", Location = new Point(5, 10), AutoSize = true, Font = AppTheme.DefaultFont };
        _mailboxEmailInput = new TextBox { Location = new Point(100, 10), Size = new Size(320, 22), Font = AppTheme.DefaultFont };
        InputPanel.Controls.AddRange(new Control[] { lblEmail, _mailboxEmailInput });

        var lblDesc = new Label { Text = _service.GetDescription(), Location = new Point(5, 38), MaximumSize = new Size(1000, 0), AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8) };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_mailboxEmailInput?.Text))
        {
            UpdateStatus("Please enter a mailbox email address.", AppTheme.ThemeRed);
            return;
        }

        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
            });

            var result = await _service.AuditMailboxAsync(_mailboxEmailInput.Text, progress, cancellationToken);

            if (string.IsNullOrEmpty(result.Error))
            {
                RenderResults(result);
                UpdateStatus($"Complete: {result.Permissions.Count} permission(s) found", Color.FromArgb(0, 100, 0));
                EnableExport();
            }
            else
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not MailboxPermissionsResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var perm in result.Permissions)
        {
            ResultsGrid.Rows.Add(perm.PermissionType, perm.User, perm.Rights, perm.AutoMap);
        }
    }
}

/// <summary>
/// Tab for mail forwarding audit visualization.
/// </summary>
public class MailForwardingTab : AuditTabBase
{
    private readonly IMailForwardingAuditService _service;
    private TextBox? _mailboxEmailInput;
    private CheckBox? _allMailboxesCheckbox;

    public MailForwardingTab(IMailForwardingAuditService service)
        : base("Mail Forwarding Audit", new[] { "Mailbox", "Forwarding Address", "Deliver to Mailbox", "Created Date" },
               new[] { 25, 35, 20, 20 })
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        _allMailboxesCheckbox = new CheckBox { Text = "Audit all mailboxes", Location = new Point(5, 10), AutoSize = true, Checked = true };

        var lblEmail = new Label { Text = "Mailbox Email:", Location = new Point(5, 32), AutoSize = true, Font = AppTheme.DefaultFont };
        _mailboxEmailInput = new TextBox { Location = new Point(100, 32), Size = new Size(300, 22), Font = AppTheme.DefaultFont, Enabled = false };

        _allMailboxesCheckbox.CheckedChanged += (s, e) => { if (_mailboxEmailInput != null) _mailboxEmailInput.Enabled = !_allMailboxesCheckbox.Checked; };

        InputPanel.Controls.AddRange(new Control[] { _allMailboxesCheckbox, lblEmail, _mailboxEmailInput });

        var lblDesc = new Label { Text = _service.GetDescription(), Location = new Point(5, 55), MaximumSize = new Size(1000, 0), AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8) };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        if (!(_allMailboxesCheckbox?.Checked ?? false) && string.IsNullOrWhiteSpace(_mailboxEmailInput?.Text))
        {
            UpdateStatus("Please enter a mailbox email or check 'Audit all mailboxes'.", AppTheme.ThemeRed);
            return;
        }

        UpdateProgress(0, 1);

        try
        {
            var mailboxEmail = (_allMailboxesCheckbox?.Checked ?? false) ? null : _mailboxEmailInput?.Text;
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
            });

            var result = await _service.AuditForwardingRulesAsync(mailboxEmail, progress, cancellationToken);

            if (string.IsNullOrEmpty(result.Error))
            {
                RenderResults(result);
                UpdateStatus($"Complete: {result.ForwardingRules.Count} rule(s) found", Color.FromArgb(0, 100, 0));
                EnableExport();
            }
            else
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not MailForwardingAuditResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var rule in result.ForwardingRules)
        {
            ResultsGrid.Rows.Add(rule.Mailbox, rule.ForwardingAddress, rule.DeliverToMailbox, rule.CreatedDate);
        }
    }
}

/// <summary>
/// Tab for shared mailbox audit visualization.
/// </summary>
public class SharedMailboxTab : AuditTabBase
{
    private readonly ISharedMailboxService _service;
    private TextBox? _searchTermInput;

    public SharedMailboxTab(ISharedMailboxService service)
        : base("Shared Mailboxes", new[] { "Display Name", "Email", "Size (MB)", "Item Count" },
               new[] { 30, 30, 20, 20 })
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        var lblSearch = new Label { Text = "Search (optional):", Location = new Point(5, 10), AutoSize = true, Font = AppTheme.DefaultFont };
        _searchTermInput = new TextBox { Location = new Point(130, 10), Size = new Size(290, 22), Font = AppTheme.DefaultFont };
        InputPanel.Controls.AddRange(new Control[] { lblSearch, _searchTermInput });

        var lblDesc = new Label { Text = _service.GetDescription(), Location = new Point(5, 38), MaximumSize = new Size(1000, 0), AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8) };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
            });

            var result = await _service.GetSharedMailboxesAsync(string.IsNullOrWhiteSpace(_searchTermInput?.Text) ? null : _searchTermInput.Text, progress, cancellationToken);

            if (string.IsNullOrEmpty(result.Error))
            {
                RenderResults(result);
                UpdateStatus($"Complete: {result.Mailboxes.Count} mailbox(es) found", Color.FromArgb(0, 100, 0));
                EnableExport();
            }
            else
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not SharedMailboxAuditResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var mailbox in result.Mailboxes)
        {
            ResultsGrid.Rows.Add(mailbox.DisplayName, mailbox.Email, mailbox.SizeMB, mailbox.ItemCount);
        }
    }
}

/// <summary>
/// Tab for group explorer visualization.
/// </summary>
public class GroupExplorerTab : AuditTabBase
{
    private readonly IGroupExplorerService _service;
    private TextBox? _searchTermInput;

    public GroupExplorerTab(IGroupExplorerService service)
        : base("Group Explorer", new[] { "Display Name", "Email", "Group Type", "Member Count" },
               new[] { 30, 30, 20, 20 })
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        var lblSearch = new Label { Text = "Search Term:", Location = new Point(5, 10), AutoSize = true, Font = AppTheme.DefaultFont };
        _searchTermInput = new TextBox { Location = new Point(100, 10), Size = new Size(320, 22), Font = AppTheme.DefaultFont };
        InputPanel.Controls.AddRange(new Control[] { lblSearch, _searchTermInput });

        var lblDesc = new Label { Text = _service.GetDescription(), Location = new Point(5, 38), MaximumSize = new Size(1000, 0), AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8) };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_searchTermInput?.Text))
        {
            UpdateStatus("Please enter a search term.", AppTheme.ThemeRed);
            return;
        }

        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
            });

            var result = await _service.SearchGroupsAsync(_searchTermInput.Text, progress, cancellationToken);

            if (string.IsNullOrEmpty(result.Error))
            {
                RenderResults(result);
                UpdateStatus($"Complete: {result.Groups.Count} group(s) found", Color.FromArgb(0, 100, 0));
                EnableExport();
            }
            else
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not GroupExplorerResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var group in result.Groups)
        {
            ResultsGrid.Rows.Add(group.DisplayName, group.Email, group.GroupType, group.MemberCount);
        }
    }
}

/// <summary>
/// Tab for Azure AD user management: browse and filter users by name.
/// </summary>
public class AzureADUsersTab : AuditTabBase
{
    private readonly IAzureADService _service;
    private TextBox? _filterInput;
    private CheckBox? _includeGuestsCheckbox;

    public AzureADUsersTab(IAzureADService service)
        : base("Azure AD — Users",
               ["Display Name", "UPN", "User Type", "Enabled", "Job Title", "Department"],
               [25, 30, 10, 8, 15, 12])
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        var lblFilter = new Label
        {
            Text     = "Filter by Name:",
            Location = new Point(5, 10),
            AutoSize = true,
            Font     = AppTheme.DefaultFont
        };
        _filterInput = new TextBox
        {
            Location    = new Point(115, 7),
            Size        = new Size(280, 22),
            Font        = AppTheme.DefaultFont,
            PlaceholderText = "(leave blank for all users)"
        };

        _includeGuestsCheckbox = new CheckBox
        {
            Text     = "Include Guest Users",
            Location = new Point(410, 8),
            AutoSize = true,
            Checked  = true,
            Font     = AppTheme.DefaultFont
        };

        InputPanel.Controls.AddRange(new Control[] { lblFilter, _filterInput, _includeGuestsCheckbox });

        var lblDesc = new Label
        {
            Text         = _service.GetDescription(),
            Location     = new Point(5, 36),
            MaximumSize  = new Size(1000, 0),
            AutoSize     = true,
            ForeColor    = Color.Gray,
            Font         = new Font("Segoe UI", 8)
        };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
                UpdateProgress(p.Current, p.Total);
            });

            var result = await _service.GetUsersAsync(
                string.IsNullOrWhiteSpace(_filterInput?.Text) ? null : _filterInput.Text,
                _includeGuestsCheckbox?.Checked ?? true,
                progress,
                cancellationToken);

            if (!string.IsNullOrEmpty(result.Error))
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
                return;
            }

            RenderResults(result);
            UpdateStatus($"Complete: {result.Users.Count} user(s) found.", Color.FromArgb(0, 100, 0));
            UpdateProgress(1, 1);
            EnableExport();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not AzureADUsersResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var user in result.Users)
        {
            ResultsGrid.Rows.Add(
                user.DisplayName,
                user.UserPrincipalName,
                user.UserType,
                user.AccountEnabled ? "Yes" : "No",
                user.JobTitle ?? string.Empty,
                user.Department ?? string.Empty);
        }
    }
}

/// <summary>
/// Tab for Intune device management: browse and filter managed devices.
/// </summary>
public class IntuneDevicesTab : AuditTabBase
{
    private readonly IIntuneService _service;
    private TextBox? _filterInput;

    public IntuneDevicesTab(IIntuneService service)
        : base("Intune — Devices",
               ["Device Name", "OS", "OS Version", "Compliance", "Last Sync", "User"],
               [25, 12, 12, 12, 20, 19])
    {
        ArgumentNullException.ThrowIfNull(service);
        _service = service;
        SetupInputPanel();
    }

    protected override void SetupInputPanel()
    {
        var lblFilter = new Label
        {
            Text     = "Filter by Device Name:",
            Location = new Point(5, 10),
            AutoSize = true,
            Font     = AppTheme.DefaultFont
        };
        _filterInput = new TextBox
        {
            Location    = new Point(160, 7),
            Size        = new Size(280, 22),
            Font        = AppTheme.DefaultFont,
            PlaceholderText = "(leave blank for all devices)"
        };

        InputPanel.Controls.AddRange(new Control[] { lblFilter, _filterInput });

        var lblDesc = new Label
        {
            Text         = _service.GetDescription(),
            Location     = new Point(5, 36),
            MaximumSize  = new Size(1000, 0),
            AutoSize     = true,
            ForeColor    = Color.Gray,
            Font         = new Font("Segoe UI", 8)
        };
        InputPanel.Controls.Add(lblDesc);
    }

    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
                UpdateProgress(p.Current, p.Total);
            });

            var result = await _service.GetManagedDevicesAsync(
                string.IsNullOrWhiteSpace(_filterInput?.Text) ? null : _filterInput.Text,
                progress,
                cancellationToken);

            if (!string.IsNullOrEmpty(result.Error))
            {
                UpdateStatus($"Error: {result.Error}", AppTheme.ThemeRed);
                return;
            }

            RenderResults(result);
            UpdateStatus($"Complete: {result.Devices.Count} managed device(s) found.", Color.FromArgb(0, 100, 0));
            UpdateProgress(1, 1);
            EnableExport();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    protected override void RenderResults(object? data)
    {
        if (data is not IntuneDevicesResult result)
            return;

        ResultsGrid.Rows.Clear();
        foreach (var device in result.Devices)
        {
            ResultsGrid.Rows.Add(
                device.DeviceName,
                device.OperatingSystem,
                device.OSVersion,
                device.ComplianceState,
                device.LastSyncDateTime?.ToString("g") ?? "Never",
                device.UserPrincipalName ?? "(Unassigned)");
        }
    }
}
