namespace CloudAdmin365.UI;

using System.Windows.Forms;
using CloudAdmin365.Utilities;

/// <summary>
/// Base class for all module panels.
/// Provides standard layout: Input → Buttons → Results Grid → Status Bar.
/// Change from TabPage to UserControl allows hosting in any container (nav panel, tabs, etc.)
/// </summary>
public abstract class AuditTabBase : UserControl
{
    protected Panel InputPanel { get; private set; } = null!;
    protected Button RunButton { get; private set; } = null!;
    protected Button CancelButton { get; private set; } = null!;
    protected Button ExportButton { get; private set; } = null!;
    protected DataGridView ResultsGrid { get; private set; } = null!;
    protected ProgressBar ProgressBar { get; private set; } = null!;
    protected Label StatusLabel { get; private set; } = null!;

    protected CancellationTokenSource? _cancellationTokenSource;
    private bool _isRunning;

    /// <summary>
    /// Initialize the tab with standard layout.
    /// </summary>
    protected AuditTabBase(string title, string[] gridColumns, int[]? columnWeights = null)
    {
        Text = title;   // Module name — shown in the left nav panel
        Padding = new Padding(8);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4
        };

        // Row 0: Input panel (configurable by subclass)
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
        InputPanel = new Panel { Dock = DockStyle.Fill };
        layout.Controls.Add(InputPanel, 0, 0);

        // Row 1: Action buttons
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        var buttonPanel = CreateButtonPanel();
        layout.Controls.Add(buttonPanel, 0, 1);

        // Row 2: Results grid
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        ResultsGrid = UiHelpers.CreateThemedDataGrid(gridColumns, columnWeights);
        layout.Controls.Add(ResultsGrid, 0, 2);

        // Row 3: Status bar
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
        var statusPanel = CreateStatusPanel();
        layout.Controls.Add(statusPanel, 0, 3);

        Controls.Add(layout);
    }

    /// <summary>
    /// Create button panel: Run, Cancel, Export.
    /// </summary>
    private Panel CreateButtonPanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill };

        RunButton = new Button
        {
            Text = "Run Audit",
            Location = new Point(0, 5),
            Size = new Size(120, 28),
            FlatStyle = FlatStyle.Flat,
            BackColor = AppTheme.ThemeGreen,
            ForeColor = Color.White,
            Font = AppTheme.BoldFont,
            Cursor = Cursors.Hand
        };
        RunButton.Click += RunButton_Click;
        panel.Controls.Add(RunButton);

        CancelButton = new Button
        {
            Text = "Cancel",
            Location = new Point(128, 5),
            Size = new Size(65, 28),
            FlatStyle = FlatStyle.Flat,
            Enabled = false
        };
        CancelButton.Click += CancelButton_Click;
        panel.Controls.Add(CancelButton);

        ExportButton = new Button
        {
            Text = "Export CSV",
            Location = new Point(203, 5),
            Size = new Size(90, 28),
            FlatStyle = FlatStyle.Flat,
            Enabled = false
        };
        ExportButton.Click += ExportButton_Click;
        panel.Controls.Add(ExportButton);

        return panel;
    }

    /// <summary>
    /// Create status bar with progress and label.
    /// </summary>
    private Panel CreateStatusPanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill };

        ProgressBar = new ProgressBar
        {
            Location = new Point(0, 5),
            Size = new Size(280, 20),
            Minimum = 0,
            Maximum = 100
        };
        panel.Controls.Add(ProgressBar);

        StatusLabel = new Label
        {
            Text = "Ready.",
            Location = new Point(290, 7),
            AutoSize = true,
            ForeColor = AppTheme.ThemeDimGray,
            MaximumSize = new Size(700, 0)
        };
        panel.Controls.Add(StatusLabel);

        return panel;
    }

    /// <summary>
    /// Setup input controls (abstract for each audit).
    /// Called by subclass constructor.
    /// </summary>
    protected abstract void SetupInputPanel();

    /// <summary>
    /// Run the audit (abstract for each audit).
    /// </summary>
    protected abstract Task RunAuditAsync(CancellationToken cancellationToken);

    /// <summary>
    /// Render results to grid (abstract for each audit).
    /// </summary>
    protected abstract void RenderResults(object? data);

    /// <summary>
    /// Handle Run button click.
    /// </summary>
    private async void RunButton_Click(object? sender, EventArgs e)
    {
        if (_isRunning)
            return;

        _isRunning = true;
        RunButton.Enabled = false;
        CancelButton.Enabled = true;
        ExportButton.Enabled = false;
        ResultsGrid.Rows.Clear();
        ProgressBar.Value = 0;
        StatusLabel.Text = "Running...";
        StatusLabel.ForeColor = AppTheme.ThemeOrange;

        _cancellationTokenSource = new CancellationTokenSource();

        try
        {
            await RunAuditAsync(_cancellationTokenSource.Token);
        }
        catch (OperationCanceledException)
        {
            StatusLabel.Text = "Audit cancelled.";
            StatusLabel.ForeColor = AppTheme.ThemeOrange;
        }
        catch (Exception ex)
        {
            StatusLabel.Text = $"Error: {ex.Message}";
            StatusLabel.ForeColor = AppTheme.ThemeRed;
        }
        finally
        {
            _isRunning = false;
            RunButton.Enabled = true;
            CancelButton.Enabled = false;
            _cancellationTokenSource?.Dispose();
        }
    }

    /// <summary>
    /// Handle Cancel button click.
    /// </summary>
    private void CancelButton_Click(object? sender, EventArgs e)
    {
        _cancellationTokenSource?.Cancel();
        CancelButton.Enabled = false;
    }

    /// <summary>
    /// Handle Export button click.
    /// </summary>
    private void ExportButton_Click(object? sender, EventArgs e)
    {
        UiHelpers.ExportDataGridToCsv(ResultsGrid, GetType().Name);
    }

    /// <summary>
    /// Update status label safely from background thread.
    /// </summary>
    protected void UpdateStatus(string text, Color? color = null)
    {
        if (InvokeRequired)
        {
            Invoke(() => UpdateStatus(text, color));
            return;
        }

        StatusLabel.Text = text;
        StatusLabel.ForeColor = color ?? AppTheme.ThemeDimGray;
    }

    /// <summary>
    /// Update progress bar safely from background thread.
    /// </summary>
    protected void UpdateProgress(int current, int total)
    {
        if (InvokeRequired)
        {
            Invoke(() => UpdateProgress(current, total));
            return;
        }

        ProgressBar.Maximum = Math.Max(total, 1);
        ProgressBar.Value = Math.Min(current, ProgressBar.Maximum);
    }

    /// <summary>
    /// Enable export button once results are loaded.
    /// </summary>
    protected void EnableExport()
    {
        if (InvokeRequired)
        {
            Invoke(EnableExport);
            return;
        }
        ExportButton.Enabled = ResultsGrid.Rows.Count > 0;
    }
}
