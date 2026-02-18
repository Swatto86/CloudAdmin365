namespace CloudAdmin365.UI;

using System.Windows.Forms;
using CloudAdmin365.Services;
using CloudAdmin365.Utilities;

/// <summary>
/// Audit tab for room calendar permissions.
/// Inherits from AuditTabBase to reuse standard layout and logic.
/// </summary>
public class RoomPermissionsTab : AuditTabBase
{
    private readonly IRoomPermissionsService _auditService;
    private TextBox? _txtRoomEmail;

    public RoomPermissionsTab(IRoomPermissionsService auditService)
        : base("Room Calendar Permissions", new[] { "Permission Type", "Delegate", "Access Level", "Granted Via" },
               new[] { 20, 30, 25, 25 })
    {
        ArgumentNullException.ThrowIfNull(auditService);
        _auditService = auditService;
        SetupInputPanel();
    }

    /// <summary>
    /// Setup the input panel with room email textbox.
    /// </summary>
    protected override void SetupInputPanel()
    {
        var lblEmail = new Label
        {
            Text = "Room Email:",
            Location = new Point(5, 10),
            AutoSize = true,
            Font = AppTheme.DefaultFont,
            TextAlign = ContentAlignment.MiddleLeft
        };
        InputPanel.Controls.Add(lblEmail);

        _txtRoomEmail = new TextBox
        {
            Location = new Point(100, 10),
            Size = new Size(320, 22),
            Font = AppTheme.DefaultFont
        };
        InputPanel.Controls.Add(_txtRoomEmail);

        var lblDescription = new Label
        {
            Text = _auditService.GetDescription(),
            Location = new Point(5, 38),
            Size = new Size(500, 20),
            AutoSize = false,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        InputPanel.Controls.Add(lblDescription);
    }

    /// <summary>
    /// Run the audit for the specified room.
    /// </summary>
    protected override async Task RunAuditAsync(CancellationToken cancellationToken)
    {
        if (_txtRoomEmail?.Text == null || string.IsNullOrWhiteSpace(_txtRoomEmail.Text))
        {
            UpdateStatus("Please enter a room email address.", AppTheme.ThemeRed);
            return;
        }

        var roomEmail = _txtRoomEmail.Text.Trim();
        UpdateProgress(0, 1);

        try
        {
            var progress = new Progress<AuditProgress>(p =>
            {
                UpdateProgress(p.Current, p.Total);
                if (p.Status != null)
                    UpdateStatus(p.Status, AppTheme.ThemeDimGray);
            });

            var result = await _auditService.AuditRoomAsync(roomEmail, progress, cancellationToken);

            if (result.Success)
            {
                RenderResults(result);
                UpdateStatus($"Audit complete: {result.Permissions.Count} permission(s) found", Color.FromArgb(0, 100, 0));
                EnableExport();
            }
            else
            {
                UpdateStatus($"Audit failed: {result.Error}", AppTheme.ThemeRed);
            }
        }
        catch (OperationCanceledException)
        {
            throw;  // Let base class handle
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error: {ex.Message}", AppTheme.ThemeRed);
        }
    }

    /// <summary>
    /// Render results to the grid.
    /// </summary>
    protected override void RenderResults(object? data)
    {
        if (data is not RoomPermissionResult result)
            return;

        ResultsGrid.Rows.Clear();

        foreach (var perm in result.Permissions)
        {
            ResultsGrid.Rows.Add(
                perm.PermissionType,
                perm.Delegate,
                perm.AccessLevel,
                perm.GrantedVia ?? "-");
        }
    }
}
