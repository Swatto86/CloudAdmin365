namespace CloudAdmin365.UI;

using System.Windows.Forms;
using CloudAdmin365;
using CloudAdmin365.Services;
using CloudAdmin365.Utilities;

/// <summary>
/// Main application form.
/// Layout: header | (left nav panel + right content panel) | footer.
/// Scales to any number of modules by grouping them in the navigation tree.
/// </summary>
public class MainForm : Form
{
    // ── Nav palette ───────────────────────────────────────────────────────
    private static readonly Color NavBg         = Color.FromArgb(27,  43,  58);
    private static readonly Color NavCategoryBg = Color.FromArgb(20,  33,  46);
    private static readonly Color NavItemHover  = Color.FromArgb(42,  68, 100);
    private static readonly Color NavItemActive = Color.FromArgb(0,  120, 212);
    private static readonly Color NavFg         = Color.FromArgb(205, 220, 232);
    private static readonly Color NavCategoryFg = Color.FromArgb(110, 145, 170);
    private static readonly Font  NavCategoryFont = new("Segoe UI", 7.5f, FontStyle.Bold);
    private static readonly Font  NavItemFont     = new("Segoe UI", 9f);

    // ── Module category display order ─────────────────────────────────────
    // Categories align with the PowerShell module that owns them.
    // Future modules (SharePoint, Azure, Security) get their own section
    // when the corresponding service and PS module are registered.
    private static readonly string[] CategoryOrder =
        ["Exchange", "Teams", "SharePoint", "Azure", "Security"];

    private readonly IAuthService _authService;
    private readonly IAuditServiceProvider _auditProvider;
    private readonly IReadOnlyDictionary<string, bool> _moduleAvailability;

    private Panel _contentPanel = null!;
    private Label _connectionStatus = null!;
    private Panel? _activeNavItem;
    private SplitContainer _splitContainer = null!;
    private readonly Dictionary<string, UserControl> _moduleCache = new();

    public MainForm(IAuthService authService, IAuditServiceProvider auditProvider,
                    IReadOnlyDictionary<string, bool> moduleAvailability)
    {
        ArgumentNullException.ThrowIfNull(authService);
        ArgumentNullException.ThrowIfNull(auditProvider);
        ArgumentNullException.ThrowIfNull(moduleAvailability);
        _authService        = authService;
        _auditProvider      = auditProvider;
        _moduleAvailability = moduleAvailability;

        InitializeComponent();
        SetupLayout();
    }

    private void InitializeComponent()
    {
        Text            = "CloudAdmin365 — M365 Administration";
        Size            = new Size(1400, 900);
        MinimumSize     = new Size(1024, 680);
        StartPosition   = FormStartPosition.CenterScreen;
        AutoScaleMode   = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96, 96);
        Font            = AppTheme.DefaultFont;
        Icon            = IconGenerator.GetAppIcon();
        FormClosing    += MainForm_FormClosing;
        Shown          += MainForm_Shown;
    }

    // ──────────────────────────────────────────────────────────────────────
    // Layout construction
    // ──────────────────────────────────────────────────────────────────────

    private void SetupLayout()
    {
        var outer = new TableLayoutPanel
        {
            Dock      = DockStyle.Fill,
            ColumnCount = 1,
            RowCount  = 3,
            Padding   = new Padding(0)
        };
        outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));   // header
        outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // body
        outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));   // footer

        outer.Controls.Add(BuildHeader(),  0, 0);
        outer.Controls.Add(BuildBody(),    0, 1);
        outer.Controls.Add(BuildFooter(), 0, 2);

        Controls.Add(outer);
    }

    // ── Header ────────────────────────────────────────────────────────────

    private Panel BuildHeader()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = AppTheme.ThemeHeaderBg
        };

        // App title (left)
        var lblTitle = new Label
        {
            Text      = "  CloudAdmin365",
            Font      = AppTheme.TitleFont,
            ForeColor = AppTheme.ThemeHeaderFg,
            Dock      = DockStyle.Left,
            AutoSize  = true,
            TextAlign = ContentAlignment.MiddleLeft
        };
        panel.Controls.Add(lblTitle);

        // Switch account button (right)
        var btnSwitch = new Button
        {
            Text      = "  Switch Account  ",
            Dock      = DockStyle.Right,
            FlatStyle = FlatStyle.Flat,
            BackColor = AppTheme.ThemeHeaderBg,
            ForeColor = AppTheme.ThemeHeaderFg,
            Font      = AppTheme.DefaultFont,
            Width     = 130,
            Cursor    = Cursors.Hand,
            Margin    = new Padding(0, 0, 10, 0)
        };
        btnSwitch.FlatAppearance.BorderColor          = AppTheme.ThemeHeaderFg;
        btnSwitch.FlatAppearance.BorderSize           = 1;
        btnSwitch.FlatAppearance.MouseOverBackColor   = Color.FromArgb(30, 108, 212);
        btnSwitch.Click += async (_, _) => await SwitchAccountAsync();
        panel.Controls.Add(btnSwitch);

        // Connection indicator (right of title, left of button)
        _connectionStatus = new Label
        {
            Text      = "● Connected",
            Dock      = DockStyle.Right,
            AutoSize  = true,
            TextAlign = ContentAlignment.MiddleRight,
            ForeColor = Color.FromArgb(100, 220, 130),
            Padding   = new Padding(0, 0, 20, 0)
        };
        panel.Controls.Add(_connectionStatus);

        return panel;
    }

    // ── Body: nav | content ───────────────────────────────────────────────

    private Control BuildBody()
    {
        // NOTE: Panel1MinSize, Panel2MinSize, and SplitterDistance are NOT set
        // here because the control has zero width during construction. WinForms
        // validates: Panel1MinSize <= SplitterDistance <= Width - Panel2MinSize.
        // When Width == 0 that constraint cannot be satisfied, causing an
        // InvalidOperationException. All three properties are deferred to the
        // Shown event (MainForm_Shown) when the form has its real dimensions.
        _splitContainer = new SplitContainer
        {
            Dock            = DockStyle.Fill,
            FixedPanel      = FixedPanel.Panel1,
            IsSplitterFixed = true
        };

        // Left: navigation panel
        var navPanel = BuildNavPanel();
        _splitContainer.Panel1.Controls.Add(navPanel);
        _splitContainer.Panel1.BackColor = NavBg;

        // Right: content area (swapped when nav item is selected)
        _contentPanel = new Panel
        {
            Dock      = DockStyle.Fill,
            BackColor = Color.FromArgb(248, 250, 252),
            Padding   = new Padding(0)
        };
        _splitContainer.Panel2.Controls.Add(_contentPanel);

        return _splitContainer;
    }

    // ── Left navigation panel ─────────────────────────────────────────────

    private Panel BuildNavPanel()
    {
        var outer = new Panel
        {
            Dock      = DockStyle.Fill,
            BackColor = NavBg
        };

        var scroll = new FlowLayoutPanel
        {
            Dock            = DockStyle.Fill,
            FlowDirection   = FlowDirection.TopDown,
            WrapContents    = false,
            AutoScroll      = true,
            BackColor       = NavBg,
            Padding         = new Padding(0)
        };

        // Group by category, in defined order
        var allAudits = _auditProvider.GetAllAudits();
        var grouped   = allAudits
            .GroupBy(a => a.Category)
            .OrderBy(g =>
            {
                var idx = Array.IndexOf(CategoryOrder, g.Key);
                return idx >= 0 ? idx : 999;
            })
            .ThenBy(g => g.Key);

        Panel? firstItem = null;
        IAuditService? firstService = null;

        foreach (var group in grouped)
        {
            scroll.Controls.Add(BuildNavCategoryHeader(group.Key));

            foreach (var audit in group.OrderBy(a => a.DisplayName))
            {
                var item = BuildNavItem(audit);
                scroll.Controls.Add(item);

                // Only auto-activate the first *enabled* module.
                bool isDisabled = audit.RequiredPowerShellModules
                    .Any(m => !(_moduleAvailability.TryGetValue(m, out bool avail) && avail));

                if (firstItem == null && !isDisabled)
                {
                    firstItem    = item;
                    firstService = audit;
                }
            }
        }

        outer.Controls.Add(scroll);

        // Select the first module by default once the form loads
        if (firstItem != null && firstService != null)
        {
            var capturedItem    = firstItem;
            var capturedService = firstService;
            Load += (_, _) => ActivateNavItem(capturedItem, capturedService);
        }

        return outer;
    }

    private Panel BuildNavCategoryHeader(string category)
    {
        var header = new Panel
        {
            Height    = 30,
            Width     = 210,
            BackColor = NavCategoryBg,
            Padding   = new Padding(0)
        };

        var lbl = new Label
        {
            Text      = category.ToUpperInvariant(),
            Font      = NavCategoryFont,
            ForeColor = NavCategoryFg,
            Dock      = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding   = new Padding(12, 0, 0, 0)
        };

        header.Controls.Add(lbl);
        return header;
    }

    private Panel BuildNavItem(IAuditService audit)
    {
        // Determine whether the module's required PS modules are all available.
        bool isDisabled = audit.RequiredPowerShellModules
            .Any(m => !(_moduleAvailability.TryGetValue(m, out bool avail) && avail));

        var item = new Panel
        {
            Height    = 38,
            Width     = 210,
            BackColor = NavBg,
            Tag       = audit,
            Cursor    = isDisabled ? Cursors.Default : Cursors.Hand
        };

        var lbl = new Label
        {
            Text      = "  " + audit.DisplayName,
            Font      = NavItemFont,
            ForeColor = isDisabled ? NavCategoryFg : NavFg,
            Dock      = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding   = new Padding(8, 0, 0, 0)
        };

        item.Controls.Add(lbl);

        if (isDisabled)
        {
            // Show tooltip listing the missing modules.
            var missingModules = audit.RequiredPowerShellModules
                .Where(m => !(_moduleAvailability.TryGetValue(m, out bool avail) && avail))
                .ToList();
            var tooltipText = "Requires: " + string.Join(", ", missingModules) + " (not installed)";

            var tip = new ToolTip();
            tip.SetToolTip(item, tooltipText);
            tip.SetToolTip(lbl,  tooltipText);

            // Clicking shows a quick install reminder.
            void OnDisabledClick(object? s, EventArgs e)
            {
                MessageBox.Show(
                    $"This module requires the following PowerShell module(s) to be installed:\n\n" +
                    string.Join("\n", missingModules.Select(m => $"  • {m}")) +
                    "\n\nInstall the module(s) with:\n" +
                    string.Join("\n", missingModules.Select(m =>
                        $"  Install-Module -Name {m} -Scope CurrentUser")) +
                    "\n\nThen restart CloudAdmin365.",
                    $"{audit.DisplayName} — Module Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            item.Click += OnDisabledClick;
            lbl.Click  += OnDisabledClick;
        }
        else
        {
            void OnEnter(object? s, EventArgs e) { if (item != _activeNavItem) item.BackColor = NavItemHover; }
            void OnLeave(object? s, EventArgs e) { if (item != _activeNavItem) item.BackColor = NavBg; }
            void OnClick(object? s, EventArgs e) { ActivateNavItem(item, audit); }

            item.MouseEnter += OnEnter; item.MouseLeave += OnLeave; item.Click += OnClick;
            lbl.MouseEnter  += OnEnter; lbl.MouseLeave  += OnLeave; lbl.Click  += OnClick;
        }

        return item;
    }

    private void ActivateNavItem(Panel navItem, IAuditService audit)
    {
        if (_activeNavItem != null)
            _activeNavItem.BackColor = NavBg;

        _activeNavItem = navItem;
        navItem.BackColor = NavItemActive;

        LoadModule(audit);
    }

    // ── Content area ──────────────────────────────────────────────────────

    private void LoadModule(IAuditService audit)
    {
        if (!_moduleCache.TryGetValue(audit.ServiceId, out var control))
        {
            control = ModuleRegistry.CreateTab(audit);
            if (control != null)
                _moduleCache[audit.ServiceId] = control;
        }

        _contentPanel.SuspendLayout();
        _contentPanel.Controls.Clear();

        if (control != null)
        {
            control.Dock = DockStyle.Fill;
            _contentPanel.Controls.Add(control);
        }
        else
        {
            // Fallback: show a placeholder
            _contentPanel.Controls.Add(BuildPlaceholder(audit.DisplayName));
        }

        _contentPanel.ResumeLayout(true);
    }

    // CreateModuleControl has been replaced by ModuleRegistry.CreateTab(audit).
    // To map a new module to its tab, add an entry in ModuleRegistry.cs.

    private static Panel BuildPlaceholder(string moduleName)
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(248, 250, 252) };
        var lbl = new Label
        {
            Text      = $"Module not yet implemented:\n{moduleName}",
            Dock      = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Color.Gray,
            Font      = new Font("Segoe UI", 11f)
        };
        panel.Controls.Add(lbl);
        return panel;
    }

    // ── Footer ────────────────────────────────────────────────────────────

    private Panel BuildFooter()
    {
        var panel = new Panel
        {
            Dock      = DockStyle.Fill,
            BackColor = Color.FromArgb(240, 243, 248)
        };

        var line = new Panel
        {
            Dock      = DockStyle.Top,
            Height    = 1,
            BackColor = Color.FromArgb(210, 215, 225)
        };
        panel.Controls.Add(line);

        var lbl = new Label
        {
            Text      = $"  CloudAdmin365 v1.0  |  Microsoft 365 Administration Platform  |  {_authService.CurrentUser?.Email ?? "Not authenticated"}",
            Dock      = DockStyle.Fill,
            ForeColor = Color.Gray,
            Font      = new Font("Segoe UI", 8),
            TextAlign = ContentAlignment.MiddleLeft
        };
        panel.Controls.Add(lbl);

        return panel;
    }

    // ──────────────────────────────────────────────────────────────────────
    // Actions
    // ──────────────────────────────────────────────────────────────────────

    private async Task SwitchAccountAsync()
    {
        var confirm = MessageBox.Show(
            "Disconnect current session and switch to a different account?\n\nCached module state will be cleared.",
            "Switch Account",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes)
            return;

        await _authService.LogoutAsync();
        Close();
    }

    /// <summary>
    /// Configures the SplitContainer once the form is fully rendered and sized.
    /// This avoids the WinForms constraint violation that occurs when setting
    /// Panel1MinSize / Panel2MinSize / SplitterDistance on a zero-width control.
    /// </summary>
    private void MainForm_Shown(object? sender, EventArgs e)
    {
        try
        {
            _splitContainer.Panel1MinSize   = 180;
            _splitContainer.Panel2MinSize   = 400;
            _splitContainer.SplitterDistance = 210;
        }
        catch (InvalidOperationException ex)
        {
            AppLogger.WriteDebug(
                $"SplitContainer (width={_splitContainer.Width}): could not configure splitter — {ex.Message}");
        }
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        _ = _authService.LogoutAsync();
    }
}
