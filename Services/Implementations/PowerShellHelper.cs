namespace CloudAdmin365.Services.Implementations;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CloudAdmin365.Utilities;

/// <summary>
/// Executes Exchange Online PowerShell commands in a dedicated runspace.
/// Handles module checks, connection management, and retry for transient errors.
/// </summary>
public sealed class PowerShellHelper : IDisposable
{
    private const int MaxCommandNameLength = 200;
    private const int MaxParameterCount = 50;
    private const int MaxRetryAttempts = 3;
    private const int RetryBaseDelayMs = 50;

    private readonly IAuthService _authService;
    private readonly SemaphoreSlim _mutex = new(1, 1);
    private Runspace? _runspace;
    private bool _initialized;
    private bool _connected;
    private bool _disposed;
    private bool _allowExecutionPolicyBypass;

    /// <summary>
    /// Create a PowerShell helper bound to the authenticated user context.
    /// </summary>
    public PowerShellHelper(IAuthService authService)
    {
        _authService = authService ?? throw new ArgumentNullException(nameof(authService));
    }

    /// <summary>
    /// Indicates whether the runspace has been initialized.
    /// </summary>
    public bool IsInitialized => _initialized;

    /// <summary>
    /// Indicates whether Exchange Online is currently connected.
    /// </summary>
    public bool IsConnected => _connected;

    /// <summary>
    /// Allow a process-scoped execution policy bypass for this runspace.
    /// </summary>
    public void EnableExecutionPolicyBypass()
    {
        _allowExecutionPolicyBypass = true;
    }

    /// <summary>
    /// Initialize the runspace and verify the ExchangeOnlineManagement module is installed.
    /// </summary>
    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();

        await _mutex.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_initialized)
            {
                return;
            }

            // CreateDefault2() is used instead of CreateDefault() because CreateDefault()
            // calls PSSnapInReader.ReadEnginePSSnapIns() which reads an ApplicationBase registry
            // path that resolves to null when running as a PublishSingleFile bundle.
            // CreateDefault2() is the correct choice for SDK-hosted/embedded runspaces — it
            // loads built-in cmdlets without touching the snap-in registry.
            var initialState = InitialSessionState.CreateDefault2();
            if (_allowExecutionPolicyBypass)
            {
                initialState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
                AppLogger.WriteInfo("PowerShell execution policy bypass enabled for this session.");
            }

            _runspace = RunspaceFactory.CreateRunspace(initialState);
            _runspace.Open();

            AppLogger.WriteInfo("PowerShell runspace created.");

            var modules = await ExecuteInternalAsync(
                ps => ps.AddCommand("Get-Module")
                    .AddParameter("ListAvailable")
                    .AddParameter("Name", "ExchangeOnlineManagement"),
                cancellationToken,
                allowReconnect: false,
                skipConnect: true).ConfigureAwait(false);

            if (modules.Count == 0)
            {
                throw new PowerShellModuleMissingException(
                    "ExchangeOnlineManagement module is not installed. " +
                    "Install it with: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser");
            }

            _initialized = true;
        }
        finally
        {
            _mutex.Release();
        }
    }

    /// <summary>
    /// Ensure a connected Exchange Online session is available.
    /// </summary>
    public async Task EnsureConnectedAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();

        await _mutex.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_connected)
            {
                return;
            }

            await ConnectExchangeOnlineAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _mutex.Release();
        }
    }

    /// <summary>
    /// Execute a PowerShell command with parameters inside the shared runspace.
    /// </summary>
    public async Task<IReadOnlyList<PSObject>> ExecuteCommandAsync(
        string commandName,
        IDictionary<string, object?>? parameters = null,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();

        if (string.IsNullOrWhiteSpace(commandName))
        {
            throw new ArgumentException("Command name is required.", nameof(commandName));
        }

        if (commandName.Length > MaxCommandNameLength)
        {
            throw new ArgumentException($"Command name exceeds max length {MaxCommandNameLength}.", nameof(commandName));
        }

        parameters ??= new Dictionary<string, object?>();
        if (parameters.Count > MaxParameterCount)
        {
            throw new ArgumentException($"Too many parameters. Max is {MaxParameterCount}.", nameof(parameters));
        }

        await EnsureConnectedAsync(cancellationToken).ConfigureAwait(false);

        AppLogger.WriteDebug($"Executing PowerShell command: {commandName}");

        return await ExecuteInternalAsync(ps =>
        {
            ps.AddCommand(commandName);
            foreach (var kvp in parameters)
            {
                ps.AddParameter(kvp.Key, kvp.Value);
            }
            ps.AddParameter("ErrorAction", "Stop");
        }, cancellationToken, allowReconnect: true, skipConnect: false).ConfigureAwait(false);
    }

    private async Task ConnectExchangeOnlineAsync(CancellationToken cancellationToken)
    {
        if (!_initialized || _runspace == null)
        {
            throw new InvalidOperationException("PowerShell runspace not initialized.");
        }

        var userPrincipalName = _authService.CurrentUser?.UserPrincipalName;
        AppLogger.WriteInfo("Connecting to Exchange Online.");

        // Get the main window handle to enable interactive authentication
        // This is needed for browser-based authentication to work from a background runspace
        var mainWindowHandle = IntPtr.Zero;
        try
        {
            var activeForm = Form.ActiveForm;
            if (activeForm != null)
            {
                mainWindowHandle = activeForm.Handle;
                AppLogger.WriteDebug($"Using window handle {mainWindowHandle} for interactive authentication.");
            }
        }
        catch { /* If we can't get the handle, proceed anyway */ }

        // Try simple method first: just use the UPN with the currently authenticated session
        // This works if the user already has exchange permissions
        try
        {
            await ExecuteInternalAsync(ps =>
            {
                ps.AddCommand("Connect-ExchangeOnline");
                ps.AddParameter("ShowBanner", false);
                ps.AddParameter("SkipLoadingCmdletHelp", true);
                ps.AddParameter("ShowProgress", false);
                if (!string.IsNullOrWhiteSpace(userPrincipalName))
                {
                    ps.AddParameter("UserPrincipalName", userPrincipalName);
                }
                ps.AddParameter("ErrorAction", "Stop");
            }, cancellationToken, allowReconnect: false, skipConnect: true).ConfigureAwait(false);

            _connected = true;
            AppLogger.WriteInfo("Exchange Online connection established via user context.");
            return;
        }
        catch (Exception ex)
        {
            AppLogger.WriteInfo($"Simple connection method failed: {ex.Message}. Connecting with interactive authentication...");
        }

        // Fall back to interactive authentication
        // The browser should open automatically when Connect-ExchangeOnline is called
        try
        {
            AppLogger.WriteInfo("Opening browser for Exchange Online authentication. Please complete the authentication process.");
            
            // Show a brief message to let user know what's happening
            MessageBox.Show(
                "A browser window will open for Exchange Online authentication.\n\n" +
                "Please sign in with your Exchange Online admin account and authorize the connection.",
                "Exchange Online Authentication",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            // Call Connect-ExchangeOnline without parameters - will trigger browser auth
            await ExecuteInternalAsync(ps =>
            {
                ps.AddCommand("Connect-ExchangeOnline");
                ps.AddParameter("ShowBanner", false);
                ps.AddParameter("SkipLoadingCmdletHelp", true);
                ps.AddParameter("ShowProgress", false);
                ps.AddParameter("ErrorAction", "Stop");
            }, cancellationToken, allowReconnect: false, skipConnect: true).ConfigureAwait(false);

            _connected = true;
            AppLogger.WriteInfo("Exchange Online connection established after browser authentication.");
        }
        catch (Exception ex)
        {
            throw new PowerShellCommandException("Failed to establish Exchange Online connection.", ex);
        }
    }

    private async Task<IReadOnlyList<PSObject>> ExecuteInternalAsync(
        Action<PowerShell> configure,
        CancellationToken cancellationToken,
        bool allowReconnect,
        bool skipConnect)
    {
        if (_runspace == null)
        {
            throw new InvalidOperationException("PowerShell runspace not initialized.");
        }

        var attempt = 0;
        Exception? lastException = null;

        while (attempt < MaxRetryAttempts)
        {
            attempt++;

            using var ps = PowerShell.Create();
            ps.Runspace = _runspace;
            configure(ps);

            using var registration = cancellationToken.Register(() => ps.Stop());

            try
            {
                var results = await Task.Run(() => ps.Invoke(), cancellationToken).ConfigureAwait(false);
                var errors = ps.Streams.Error?.ToList() ?? new List<ErrorRecord>();
                
                // Capture all output streams for device code messages
                var infoMessages = ps.Streams.Information?.ToList() ?? new List<InformationRecord>();
                var warningMessages = ps.Streams.Warning?.ToList() ?? new List<WarningRecord>();
                var verboseMessages = ps.Streams.Verbose?.ToList() ?? new List<VerboseRecord>();
                var debugMessages = ps.Streams.Debug?.ToList() ?? new List<DebugRecord>();
                
                // Log and check all streams for device code information
                var allMessages = new List<string>();
                allMessages.AddRange(infoMessages.Select(i => i.ToString()));
                allMessages.AddRange(warningMessages.Select(w => w.ToString()));
                allMessages.AddRange(verboseMessages.Select(v => v.ToString()));
                allMessages.AddRange(debugMessages.Select(d => d.ToString()));
                
                foreach (var msg in allMessages)
                {
                    if (string.IsNullOrWhiteSpace(msg)) continue;
                    
                    AppLogger.WriteInfo($"PowerShell Output: {msg}");
                    
                    // If this looks like a device code message, show it to the user
                    if (msg.Contains("devicelogin", StringComparison.OrdinalIgnoreCase) ||
                        msg.Contains("code:", StringComparison.OrdinalIgnoreCase) ||
                        (msg.Contains("https://", StringComparison.OrdinalIgnoreCase) && msg.Contains("code", StringComparison.OrdinalIgnoreCase)))
                    {
                        MessageBox.Show(
                            msg,
                            "Exchange Online Device Code",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }

                if (errors.Count > 0)
                {
                    var message = string.Join(Environment.NewLine, errors.Select(e => e.ToString()));
                    var ex = new PowerShellCommandException("PowerShell command failed.", message);

                    if (allowReconnect && IsDisconnectError(errors))
                    {
                        _connected = false;
                    }

                    if (IsTransientError(errors) && attempt < MaxRetryAttempts)
                    {
                        await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(false);
                        lastException = ex;
                        continue;
                    }

                    throw ex;
                }

                return results;
            }
            catch (PowerShellCommandException)
            {
                throw;
            }
            catch (Exception ex) when (IsTransientException(ex) && attempt < MaxRetryAttempts)
            {
                lastException = ex;
                await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                throw new PowerShellCommandException("PowerShell command execution failed.", ex);
            }
        }

        if (lastException != null)
        {
            throw new PowerShellCommandException("PowerShell command failed after retries.", lastException);
        }

        throw new PowerShellCommandException("PowerShell command failed after retries.", "No additional error details.");
    }

    private static Task DelayForRetryAsync(int attempt, CancellationToken cancellationToken)
    {
        var delay = RetryBaseDelayMs * (int)Math.Pow(2, Math.Max(0, attempt - 1));
        return Task.Delay(delay, cancellationToken);
    }

    private static bool IsTransientError(IEnumerable<ErrorRecord> errors)
    {
        foreach (var error in errors)
        {
            if (IsTransientException(error.Exception))
            {
                return true;
            }

            var message = error.ToString();
            if (ContainsTransientMessage(message))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsTransientException(Exception? ex)
    {
        if (ex == null)
        {
            return false;
        }

        return ex is TimeoutException
            || ex is System.Net.Http.HttpRequestException
            || ex is System.Management.Automation.Remoting.PSRemotingTransportException;
    }

    private static bool ContainsTransientMessage(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        var text = message.ToLowerInvariant();
        return text.Contains("temporarily")
            || text.Contains("timeout")
            || text.Contains("throttle")
            || text.Contains("429")
            || text.Contains("server busy");
    }

    private static bool IsDisconnectError(IEnumerable<ErrorRecord> errors)
    {
        return errors.Any(e => e.ToString().Contains("Connect-ExchangeOnline", StringComparison.OrdinalIgnoreCase)
            || e.ToString().Contains("not connected", StringComparison.OrdinalIgnoreCase));
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(PowerShellHelper));
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _connected = false;

        try
        {
            _runspace?.Dispose();
        }
        catch
        {
        }

        _mutex.Dispose();
    }
}

/// <summary>
/// PowerShell error for missing ExchangeOnlineManagement module.
/// </summary>
public sealed class PowerShellModuleMissingException : Exception
{
    public PowerShellModuleMissingException(string message) : base(message)
    {
    }
}

/// <summary>
/// PowerShell error for command failures.
/// </summary>
public sealed class PowerShellCommandException : Exception
{
    public string? CommandOutput { get; }

    public PowerShellCommandException(string message, string? output) : base(message)
    {
        CommandOutput = output;
    }

    public PowerShellCommandException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
