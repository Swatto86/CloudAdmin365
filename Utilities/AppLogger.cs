namespace CloudAdmin365.Utilities;

using System;
using System.Globalization;
using System.IO;

/// <summary>
/// Minimal file logger for startup and fatal error diagnostics.
/// </summary>
public static class AppLogger
{
    private const long MaxLogBytes = 2 * 1024 * 1024;
    private const string LogFileName = "app.log";
    private const string LogBackupName = "app.log.old";
    private static readonly object Sync = new();

    private static bool _initialized;
    private static bool _debugEnabled;
    private static string _logDirectory = string.Empty;
    private static string _logFilePath = string.Empty;

    /// <summary>
    /// Full path to the active log file.
    /// </summary>
    public static string LogFilePath => _logFilePath;

    /// <summary>
    /// Initialize logging to %LocalAppData%\{appName}\logs.
    /// </summary>
    public static void Initialize(string appName)
    {
        if (_initialized)
        {
            return;
        }

        _debugEnabled = IsDebugEnabled();

        var baseDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            appName,
            "logs");

        try
        {
            Directory.CreateDirectory(baseDir);
            _logDirectory = baseDir;
            _logFilePath = Path.Combine(baseDir, LogFileName);
            _initialized = true;
        }
        catch
        {
            _initialized = false;
            return;
        }

        WriteInfo("Logger initialized.");
        WriteDebug("Debug logging enabled.");
    }

    /// <summary>
    /// Write an informational message.
    /// </summary>
    public static void WriteInfo(string message)
    {
        Write("INFO", message, null, allowDebug: true);
    }

    /// <summary>
    /// Write a debug message when debug is enabled.
    /// </summary>
    public static void WriteDebug(string message)
    {
        if (!_debugEnabled)
        {
            return;
        }

        Write("DEBUG", message, null, allowDebug: true);
    }

    /// <summary>
    /// Write an error message and optional exception details.
    /// </summary>
    public static void WriteError(string message, Exception? exception = null)
    {
        Write("ERROR", message, exception, allowDebug: true);
    }

    private static void Write(string level, string message, Exception? exception, bool allowDebug)
    {
        if (!_initialized)
        {
            return;
        }

        try
        {
            lock (Sync)
            {
                RotateIfNeeded();

                var timestamp = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
                var line = $"{timestamp} [{level}] {message}";
                File.AppendAllText(_logFilePath, line + Environment.NewLine);

                if (exception != null)
                {
                    File.AppendAllText(_logFilePath, exception + Environment.NewLine);
                }
            }
        }
        catch
        {
        }
    }

    private static void RotateIfNeeded()
    {
        if (string.IsNullOrWhiteSpace(_logFilePath) || !File.Exists(_logFilePath))
        {
            return;
        }

        try
        {
            var fileInfo = new FileInfo(_logFilePath);
            if (fileInfo.Length <= MaxLogBytes)
            {
                return;
            }

            var backupPath = Path.Combine(_logDirectory, LogBackupName);
            if (File.Exists(backupPath))
            {
                File.Delete(backupPath);
            }

            File.Move(_logFilePath, backupPath);
        }
        catch
        {
        }
    }

    private static bool IsDebugEnabled()
    {
        var value = Environment.GetEnvironmentVariable("CloudAdmin365_DEBUG");
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return value.Equals("1", StringComparison.OrdinalIgnoreCase)
            || value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }
}
