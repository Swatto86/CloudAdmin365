namespace CloudAdmin365.Services.Implementations;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CloudAdmin365.Services;
using CloudAdmin365.Utilities;

/// <summary>
/// Manages runtime dependencies: checks for required PowerShell modules and facilitates installation.
/// Derives the required module list from the registered services, so no hardcoded lists are needed.
/// Returns a per-module availability map that MainForm uses to enable/disable nav items.
/// </summary>
public sealed class DependencyManager
{
    private const int CommandTimeoutSeconds      = 30;
    private const int ModuleInstallTimeoutSeconds = 180;

    // ── Public API ────────────────────────────────────────────────────────

    /// <summary>
    /// Checks all required PowerShell modules (derived from <paramref name="services"/>),
    /// ALWAYS shows the dependency dialog so users see which modules are available,
    /// and offers to install any missing ones. Returns a map of module → installed status.
    /// Modules that are not installed will have their nav sections disabled in the main UI.
    /// Also verifies the .NET runtime version.
    /// </summary>
    /// <param name="services">All registered services; their RequiredPowerShellModules are unioned.</param>
    /// <returns>
    /// Dictionary keyed by module name; <c>true</c> = installed, <c>false</c> = not installed.
    /// </returns>
    public static async Task<IReadOnlyDictionary<string, bool>> CheckAndInstallDependenciesAsync(
        IEnumerable<IAuditService> services,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(services);
        AppLogger.WriteInfo("Checking runtime dependencies...");

        try
        {
            CheckDotNetRuntime();

            // Derive the unique set of required modules from all registered services.
            var requiredModules = services
                .SelectMany(s => s.RequiredPowerShellModules)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(m => m)
                .ToList();

            AppLogger.WriteDebug($"Required PS modules derived from services: [{string.Join(", ", requiredModules)}]");

            if (requiredModules.Count == 0)
            {
                AppLogger.WriteInfo("No PowerShell modules required.");
                return new Dictionary<string, bool>();
            }

            // Check which modules are installed.
            var availability = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            foreach (var module in requiredModules)
            {
                bool installed = await IsModuleInstalledAsync(module, cancellationToken).ConfigureAwait(false);
                availability[module] = installed;
                AppLogger.WriteInfo($"Module {module}: {(installed ? "installed" : "NOT FOUND")}");
            }

            var missing = availability.Where(kv => !kv.Value).Select(kv => kv.Key).ToList();

            // Always show the dependency dialog so the user can see the status of all
            // modules and install any missing ones. When everything is installed the
            // dialog still shows — with all green ticks — for transparency.
            AppLogger.WriteInfo(missing.Count == 0
                ? "All required PowerShell modules are satisfied. Showing status dialog."
                : $"Missing modules: [{string.Join(", ", missing)}]. Showing dependency dialog.");

            var installed2 = await ShowDependencyDialogAsync(availability, missing, cancellationToken)
                .ConfigureAwait(false);

            // Merge any newly-installed modules back into availability.
            foreach (var m in installed2)
                availability[m] = true;

            AppLogger.WriteInfo("Final module availability: " +
                string.Join(", ", availability.Select(kv => $"{kv.Key}={kv.Value}")));

            return availability;
        }
        catch (Exception ex)
        {
            AppLogger.WriteError("Fatal error during dependency check.", ex);
            throw;
        }
    }

    // ── .NET Runtime check ────────────────────────────────────────────────

    private static void CheckDotNetRuntime()
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName               = "dotnet",
                Arguments              = "--version",
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                CreateNoWindow         = true
            };

            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException(".NET Runtime check failed: unable to start process.");

            if (!process.WaitForExit(CommandTimeoutSeconds * 1000))
            {
                process.Kill();
                throw new TimeoutException($".NET Runtime check timed out after {CommandTimeoutSeconds} seconds.");
            }

            var version = process.StandardOutput.ReadToEnd().Trim();
            AppLogger.WriteInfo($".NET Runtime detected: {version}");

            // Accept .NET 8, 9, or 10+
            if (!System.Text.RegularExpressions.Regex.IsMatch(version, @"^([89]|[1-9]\d)\.\d"))
            {
                throw new InvalidOperationException(
                    $"CloudAdmin365 requires .NET Runtime 8.0 or later. Found: {version}. " +
                    "Download from: https://dotnet.microsoft.com/en-us/download/dotnet");
            }
        }
        catch (System.ComponentModel.Win32Exception)
        {
            throw new InvalidOperationException(
                ".NET Runtime is not installed or not in PATH. " +
                "Download .NET Runtime 8.0 from: https://dotnet.microsoft.com/en-us/download/dotnet/8.0");
        }
    }

    // ── Module checks ─────────────────────────────────────────────────────

    private static async Task<bool> IsModuleInstalledAsync(string moduleName, CancellationToken cancellationToken)
    {
        try
        {
            var script = $"Get-Module -ListAvailable -Name '{moduleName}' | Select-Object -First 1 | Out-String";

            AppLogger.WriteDebug($"Checking PS module via Get-Module -ListAvailable: {moduleName}");

            var startInfo = new ProcessStartInfo
            {
                FileName               = "powershell.exe",
                Arguments              = $"-NoProfile -Command \"{script}\"",
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true
            };

            using var process = Process.Start(startInfo);
            if (process == null) return false;

            var completed = process.WaitForExit(CommandTimeoutSeconds * 1000);
            if (!completed) { process.Kill(); return false; }

            var output = process.StandardOutput.ReadToEnd().Trim();
            return output.Length > 0 && !output.Contains("No match found");
        }
        catch (Exception ex)
        {
            AppLogger.WriteInfo($"Error checking module {moduleName}: {ex.Message}");
            return false;
        }
    }

    private static async Task InstallModuleAsync(string moduleName, CancellationToken cancellationToken)
    {
        AppLogger.WriteInfo($"Installing PowerShell module: {moduleName}");
        AppLogger.WriteDebug($"Install command: Install-Module -Name '{moduleName}' -Scope CurrentUser -Force -AllowClobber");

        var script = $@"
Install-Module -Name '{moduleName}' `
    -Scope CurrentUser `
    -Force `
    -AllowClobber `
    -ErrorAction Stop
Write-Host 'Installation completed successfully.'
";

        var startInfo = new ProcessStartInfo
        {
            FileName               = "powershell.exe",
            Arguments              = $"-NoProfile -Command \"{script}\"",
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = false
        };

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start PowerShell installer process.");

        var completed = process.WaitForExit(ModuleInstallTimeoutSeconds * 1000);
        var output    = await process.StandardOutput.ReadToEndAsync(cancellationToken);
        var errors    = await process.StandardError.ReadToEndAsync(cancellationToken);

        if (!completed)
        {
            process.Kill();
            throw new TimeoutException($"Module installation timed out after {ModuleInstallTimeoutSeconds} seconds.");
        }

        AppLogger.WriteInfo($"Module install output for {moduleName}:\n{output}");

        if (process.ExitCode != 0)
        {
            AppLogger.WriteError($"Module install errors for {moduleName}:\n{errors}");
            throw new InvalidOperationException(
                $"Failed to install {moduleName}.\n\nError:\n{errors}\n\n" +
                $"To install manually:\n  Install-Module -Name {moduleName} -Scope CurrentUser");
        }

        AppLogger.WriteInfo($"Successfully installed module: {moduleName}");
    }

    // ── Dependency dialog ─────────────────────────────────────────────────

    /// <summary>
    /// Shows the dependency dialog and installs any modules the user requests.
    /// Returns the list of module names successfully installed during this call.
    /// </summary>
    private static async Task<List<string>> ShowDependencyDialogAsync(
        Dictionary<string, bool> availability,
        List<string> missing,
        CancellationToken cancellationToken)
    {
        var newlyInstalled = new List<string>();

        // Marshal to UI thread — DependencyManager may be called before the message loop is running,
        // so we run via a TaskCompletionSource on the calling (STA) thread.
        await Task.Run(() =>
        {
            // Must run on an STA thread. Program.Main is [STAThread], so Application-level
            // code (before Application.Run) can call ShowDialog directly.
            using var dlg = new DependencyDialog(availability);
            var result = dlg.ShowDialog() == DialogResult.OK ? dlg.ModulesToInstall : [];

            foreach (var module in result)
            {
                try
                {
                    InstallModuleAsync(module, cancellationToken).GetAwaiter().GetResult();
                    newlyInstalled.Add(module);
                }
                catch (Exception ex)
                {
                    AppLogger.WriteError($"Failed to install {module}.", ex);
                    MessageBox.Show(
                        $"Failed to install {module}:\n\n{ex.Message}",
                        "Installation Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }, cancellationToken).ConfigureAwait(false);

        return newlyInstalled;
    }

    // ── Inner form ────────────────────────────────────────────────────────

    /// <summary>
    /// Dependency status dialog: shows each required module with a ✓/✗ indicator,
    /// lets the user install missing modules or proceed without them.
    /// Modules that are skipped will have their nav items disabled in the main window.
    /// </summary>
    private sealed class DependencyDialog : Form
    {
        private static readonly Color InstalledColor   = Color.FromArgb(0, 180, 90);
        private static readonly Color MissingColor     = Color.FromArgb(220, 60, 60);
        private static readonly Color DialogBackground = Color.FromArgb(248, 250, 252);
        private static readonly Color AccentBlue       = Color.FromArgb(0, 120, 212);

        public IReadOnlyList<string> ModulesToInstall { get; private set; } = [];

        private readonly Dictionary<string, bool> _availability;
        private readonly List<CheckBox>           _installChecks = [];
        private readonly bool                     _allInstalled;

        public DependencyDialog(Dictionary<string, bool> availability)
        {
            _availability = availability;
            _allInstalled = availability.Values.All(v => v);
            BuildUI();
        }

        private void BuildUI()
        {
            // ── Form properties ──────────────────────────────────────────
            Text            = "CloudAdmin365 \u2014 Module Dependencies";
            StartPosition   = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            BackColor       = DialogBackground;
            Font            = new Font("Segoe UI", 9f);
            AutoScaleMode   = AutoScaleMode.Dpi;
            Icon            = IconGenerator.GetAppIcon();
            ShowInTaskbar   = true;

            // We build everything inside a non-docked panel so we can
            // measure its preferred size and set the form's ClientSize once.
            var outer = new TableLayoutPanel
            {
                AutoSize     = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount  = 1,
                RowCount     = 3,
                Padding      = new Padding(20, 16, 20, 16)
            };
            outer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // header
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // module rows
            outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // buttons

            // ── Header ───────────────────────────────────────────────────
            var lblTitle = new Label
            {
                Text      = _allInstalled
                    ? "All Modules Installed"
                    : "Required PowerShell Modules",
                Font      = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(20, 60, 100),
                AutoSize  = true,
                Margin    = new Padding(0, 0, 0, 4)
            };

            var introText = _allInstalled
                ? "All required PowerShell modules are installed and ready."
                : "CloudAdmin365 requires the following PowerShell modules.\n" +
                  "Tick any missing modules to install them automatically,\n" +
                  "or click Continue to proceed (disabled modules\u2019 features won\u2019t be available).";

            var lblIntro = new Label
            {
                Text      = introText,
                AutoSize  = true,
                ForeColor = Color.FromArgb(60, 80, 100),
                Margin    = new Padding(0, 0, 0, 14)
            };

            var headerPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize      = true,
                AutoSizeMode  = AutoSizeMode.GrowAndShrink,
                WrapContents  = false,
                Margin        = new Padding(0)
            };
            headerPanel.Controls.Add(lblTitle);
            headerPanel.Controls.Add(lblIntro);
            outer.Controls.Add(headerPanel, 0, 0);

            // ── Module rows ──────────────────────────────────────────────
            var modulePanel = new TableLayoutPanel
            {
                ColumnCount  = 3,
                AutoSize     = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin       = new Padding(0, 0, 0, 18)
            };
            modulePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 28));   // status icon
            modulePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230));  // module name
            modulePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));  // status / checkbox

            foreach (var (module, isInstalled) in _availability.OrderBy(kv => kv.Key))
            {
                modulePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

                var lblStatus = new Label
                {
                    Text      = isInstalled ? "\u2713" : "\u2717",
                    ForeColor = isInstalled ? InstalledColor : MissingColor,
                    Font      = new Font("Segoe UI", 12f, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Dock      = DockStyle.Fill,
                    Margin    = new Padding(0)
                };

                var lblName = new Label
                {
                    Text      = module,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Dock      = DockStyle.Fill,
                    ForeColor = Color.FromArgb(30, 50, 80),
                    Margin    = new Padding(4, 0, 0, 0)
                };

                modulePanel.Controls.Add(lblStatus);
                modulePanel.Controls.Add(lblName);

                if (!isInstalled)
                {
                    var chk = new CheckBox
                    {
                        Text    = "Install automatically",
                        Checked = true,
                        Dock    = DockStyle.Fill,
                        Tag     = module,
                        Margin  = new Padding(0)
                    };
                    _installChecks.Add(chk);
                    modulePanel.Controls.Add(chk);
                }
                else
                {
                    var lblOk = new Label
                    {
                        Text      = "Installed",
                        ForeColor = InstalledColor,
                        Font      = new Font("Segoe UI", 9f, FontStyle.Regular),
                        TextAlign = ContentAlignment.MiddleLeft,
                        Dock      = DockStyle.Fill,
                        Margin    = new Padding(0)
                    };
                    modulePanel.Controls.Add(lblOk);
                }
            }

            outer.Controls.Add(modulePanel, 0, 1);

            // ── Buttons ──────────────────────────────────────────────────
            var btnPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize      = true,
                AutoSizeMode  = AutoSizeMode.GrowAndShrink,
                WrapContents  = false,
                Margin        = new Padding(0)
            };

            if (_allInstalled)
            {
                // All modules present — single "Continue" button.
                var btnContinue = CreateButton("Continue", AccentBlue, Color.White, 160);
                btnContinue.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };
                btnPanel.Controls.Add(btnContinue);
            }
            else
            {
                // Some missing — "Install & Continue" + "Continue" (skip).
                var btnInstall = CreateButton("Install && Continue", AccentBlue, Color.White, 180);
                btnInstall.Click += BtnInstall_Click;
                btnPanel.Controls.Add(btnInstall);

                btnPanel.Controls.Add(new Panel { Width = 10, Height = 1 }); // spacer

                var btnSkip = CreateButton("Continue", Color.FromArgb(230, 235, 242),
                    Color.FromArgb(50, 60, 80), 120);
                btnSkip.FlatAppearance.BorderColor = Color.FromArgb(190, 200, 210);
                btnSkip.FlatAppearance.BorderSize  = 1;
                btnSkip.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };
                btnPanel.Controls.Add(btnSkip);
            }

            outer.Controls.Add(btnPanel, 0, 2);
            Controls.Add(outer);

            // Size the form to fit the content.
            var preferred = outer.GetPreferredSize(new Size(540, 0));
            ClientSize = new Size(
                Math.Max(preferred.Width, 460),
                preferred.Height);
            MinimumSize = Size; // lock to the computed size
        }

        /// <summary>Creates a styled flat button with consistent appearance.</summary>
        private static Button CreateButton(string text, Color backColor, Color foreColor, int width)
        {
            var btn = new Button
            {
                Text      = text,
                Width     = width,
                Height    = 34,
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = foreColor,
                Font      = new Font("Segoe UI", 9f),
                Cursor    = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void BtnInstall_Click(object? sender, EventArgs e)
        {
            ModulesToInstall = _installChecks
                .Where(c => c.Checked)
                .Select(c => (string)c.Tag!)
                .ToList();

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
