using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CloudAdmin365.Services;
using CloudAdmin365.Services.Implementations;
using CloudAdmin365.UI;
using CloudAdmin365.Utilities;

namespace CloudAdmin365;

static class Program
{
    /// <summary>
    /// Enable DPI awareness for high-resolution displays.
    /// </summary>
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetProcessDPIAware();

    [DllImport("shcore.dll", SetLastError = true)]
    private static extern int SetProcessDpiAwareness(int awareness);

    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // Enable DPI awareness
        try { SetProcessDpiAwareness(2); }
        catch { try { SetProcessDPIAware(); } catch { } }

        AppLogger.Initialize("CloudAdmin365");
        AppLogger.WriteInfo("Application starting.");
        AppLogger.WriteInfo($"Executable: {Application.ExecutablePath}");
        AppLogger.WriteInfo($"Version: {Assembly.GetExecutingAssembly().GetName().Version}");

        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, args) => HandleFatalException("UI thread", args.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            var ex = args.ExceptionObject as Exception ?? new Exception("Unknown fatal error.");
            HandleFatalException("AppDomain", ex);
        };

        try
        {
            ApplicationConfiguration.Initialize();
            Application.EnableVisualStyles();

            // Setup dependency injection (manually for simplicity; could use DI container)
            // Use real Azure AD authentication with MSAL interactive browser login
            IAuthService authService = new AzureIdentityAuthService();

            // Show login dialog
            using (var loginDlg = new LoginDialog(authService))
            {
                if (loginDlg.ShowDialog() != DialogResult.OK)
                {
                    AppLogger.WriteInfo("Login canceled.");
                    Application.Exit();
                    return;
                }
            }

            using var powerShellHelper = new PowerShellHelper(authService);
            powerShellHelper.EnableExecutionPolicyBypass();
            powerShellHelper.InitializeAsync().GetAwaiter().GetResult();

            // Register all modules via the central ModuleRegistry.
            // To add a new module, only ModuleRegistry.cs needs to change.
            var auditProvider = new AuditServiceProvider();
            ModuleRegistry.RegisterAll(auditProvider, powerShellHelper, authService);

            // Check / install required PowerShell modules, derived from the registered services.
            var moduleAvailability = DependencyManager
                .CheckAndInstallDependenciesAsync(auditProvider.GetAllAudits())
                .GetAwaiter()
                .GetResult();

            // Run main form — passes availability map so nav items are enabled/disabled accordingly.
            var mainForm = new MainForm(authService, auditProvider, moduleAvailability);
            Application.Run(mainForm);
        }
        catch (Exception ex)
        {
            HandleFatalException("Startup", ex);
        }
    }

    private static void HandleFatalException(string source, Exception exception)
    {
        AppLogger.WriteError($"Fatal error in {source}.", exception);

        try
        {
            var message =
                "CloudAdmin365 encountered a fatal error and must close." + Environment.NewLine +
                Environment.NewLine +
                exception.Message + Environment.NewLine +
                Environment.NewLine +
                "Log: " + AppLogger.LogFilePath;

            MessageBox.Show(
                message,
                "CloudAdmin365 Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        catch
        {
        }

        Environment.Exit(1);
    }

}