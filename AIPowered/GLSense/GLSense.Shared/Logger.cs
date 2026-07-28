// Logger.cs in GLSense.Shared 
using GLSense.Contracts;
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace GLSense.Shared
{
    public class Logger : MarshalByRefObject, Contracts.ILogger
    {
        private readonly NLog.Logger _logger;
        private readonly bool _isInitialized = false;
        private readonly object _initLock = new();
        private readonly IGLSenseContext _context;  // Reference to context

        // Buffer for debug logs (used when DebugMode is true)
        private static readonly List<string> _debugBuffer = new List<string>();
        private static readonly object _bufferLock = new object();

        // Thread-local scope depth for nested indentation
        [ThreadStatic]
        private static int _scopeDepth;

        // DebugMode comes from context
        private bool DebugMode => _context?.DebugMode ?? false;

        public Logger(IGLSenseContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));

            lock (_initLock)
            {
                if (_isInitialized) return;

                try
                {
                    // Create header layout with dynamic date evaluation
                    string HdrText = BuildLogHeader();

                    // Use PathProvider for log path
                    string logPath = PathProvider.Instance.Logs;
                    string fileNamePattern = logPath + @"\GLSense_Logs_${date:format=dd-MMM-yyyy}.log";
                    var fileNameLayout = NLog.Layouts.Layout.FromString(fileNamePattern);
                    var archiveFileNamePattern = logPath + @"\GLSense_Logs_{#}.log";

                    var logfile = new FileTarget("logfile")
                    {
                        FileName = fileNameLayout,
                        Header = HdrText,
                        AutoFlush = true,
                        Layout = "${longdate}|${message:withException=true:exceptionSeparator=|}",
                        KeepFileOpen = false,
                        DeleteOldFileOnStartup = false,
                        ArchiveAboveSize = 20 * 1024 * 1024,  // 20MB archive size
                        MaxArchiveFiles = 30,
                        ArchiveFileName = archiveFileNamePattern
                    };

                    var GLSenseLoggerConfiguration = new LoggingConfiguration();
                    GLSenseLoggerConfiguration.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);

                    LogManager.Configuration = GLSenseLoggerConfiguration;

                    _logger = LogManager.GetCurrentClassLogger();

                    _isInitialized = true;
                }
                catch (Exception ex)
                {
                    // Simple direct file write
                    try
                    {
                        string logPath = PathProvider.Instance.Logs;
                        string logFile = Path.Combine(logPath, $"GLSense_Error_{DateTime.Now:dd-MMM-yyyy}.log");

                        Directory.CreateDirectory(logPath);

                        string logContent = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}|ERROR|Logger initialization failed: {ex.Message}|{ex.StackTrace}";
                        File.WriteAllText(logFile, logContent);
                    }
                    catch
                    {
                        // Absolute fallback - write to temp directory
                        try
                        {
                            string tempLog = Path.Combine(Path.GetTempPath(), $"GLSense_Error_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                            File.WriteAllText(tempLog, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}|ERROR|{ex.Message}|{ex.StackTrace}");
                        }
                        catch { /* Give up */ }
                    }
                }
            }
        }

        private string BuildLogHeader()
        {
            string defaultVersion = PathProvider.Instance.LatestVersion;
            string defaultCommitDate = PathProvider.Instance.LatestReleaseDate;
            var sb = new StringBuilder();

            string header = $"Orbit GLSense(version : {defaultVersion} Released on : {defaultCommitDate}). Logs As On {DateTime.Now:dddd, dd MMMM yyyy}. Time Zone: {TimeZoneInfo.Local.DisplayName}";
            sb.AppendLine(header);
            sb.AppendLine(new string('-', header.Length));

            return sb.ToString();
        }

        private void IncrementScope()
        {
            _scopeDepth = Math.Max(0, _scopeDepth) + 1;
        }

        private void DecrementScope()
        {
            _scopeDepth = Math.Max(0, _scopeDepth - 1);
        }

        private string Indent()
        {
            int safeDepth = Math.Max(0, _scopeDepth);
            return new string(' ', safeDepth * 2);
        }

        #region Logging Methods

        public void LogInfo(string msg)
        {
            var logMessage = $"{Indent()}INFO  | {DateTime.Now:HH:mm:ss} | {msg}";
            _logger?.Info(logMessage);
        }

        public void LogWarn(string msg)
        {
            var logMessage = $"{Indent()}WARN  | {DateTime.Now:HH:mm:ss} | {msg}";
            _logger?.Warn(logMessage);
        }

        public void LogError(string msg, Exception ex = null)
        {
            var logMessage = $"{Indent()}ERROR | {DateTime.Now:HH:mm:ss} | {msg}";
            if (ex != null)
                _logger?.Error(ex, logMessage);
            else
                _logger?.Error(logMessage);
        }

        public void LogDebug(string msg)
        {
            // ✅ DebugMode comes from context
            if (!DebugMode)
                return;

            var logMessage = $"{Indent()}DEBUG | {DateTime.Now:HH:mm:ss} | {msg}";

            if (_logger != null)
            {
                _logger.Debug(logMessage);
            }
            else
            {
                lock (_bufferLock)
                {
                    _debugBuffer.Add(logMessage);
                }
            }
        }

        public void LogException(Exception ex, string context = "")
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{Indent()}========== Exception ==========");
            if (!string.IsNullOrEmpty(context))
                sb.AppendLine($"{Indent()}Context: {context}");
            sb.AppendLine($"{Indent()}Type: {ex.GetType().FullName}");
            sb.AppendLine($"{Indent()}Message: {ex.Message}");
            sb.AppendLine($"{Indent()}Source: {ex.Source}");
            sb.AppendLine($"{Indent()}TargetSite: {ex.TargetSite}");
            sb.AppendLine($"{Indent()}StackTrace:");
            foreach (var line in ex.StackTrace?.Split(new[] { Environment.NewLine }, StringSplitOptions.None) ?? Array.Empty<string>())
                sb.AppendLine($"{Indent()}{line}");
            if (ex.InnerException != null)
            {
                sb.AppendLine($"{Indent()}----- Inner Exception -----");
                sb.AppendLine($"{Indent()}Type: {ex.InnerException.GetType().FullName}");
                sb.AppendLine($"{Indent()}Message: {ex.InnerException.Message}");
                sb.AppendLine($"{Indent()}{ex.InnerException.StackTrace}");
            }
            sb.AppendLine($"{Indent()}============================");

            var exceptionMessage = sb.ToString();
            _logger?.Error(exceptionMessage);
        }

        public void LogRawJson(string context, string rawJson)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{Indent()}----- Raw JSON {(string.IsNullOrWhiteSpace(context) ? string.Empty : "(" + context + ")")} -----");
            sb.AppendLine(string.IsNullOrEmpty(rawJson) ? "<empty>" : rawJson);
            sb.AppendLine($"{Indent()}----- End Raw JSON -----");
            _logger?.Error(sb.ToString());
        }

        #endregion

        #region Scope Methods

        public void LogMethodEntry([CallerMemberName] string methodName = "")
        {
            LogDebug($"Entering {methodName}");
            IncrementScope();
        }

        public void LogMethodExit([CallerMemberName] string methodName = "")
        {
            DecrementScope();
            LogDebug($"Exiting {methodName}");
        }

        public class LogScope : IDisposable
        {
            private readonly string _scopeName;
            private readonly Logger _logger;
            private bool _disposed = false;

            public LogScope(Logger logger, string scopeName)
            {
                _logger = logger;
                _scopeName = scopeName;
                _logger.LogDebug($"BEGIN: {_scopeName}");
                _logger.IncrementScope();
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!_disposed)
                {
                    if (disposing)
                    {
                        _logger.DecrementScope();
                        _logger.LogDebug($"END: {_scopeName}");
                    }
                    _disposed = true;
                }
            }
        }

        #endregion

        #region Flush

        public void FlushDebugLogs(string section = "Buffered Logs")
        {
            lock (_bufferLock)
            {
                if (_debugBuffer.Count == 0) return;

                var header = $"===== {section} | {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====";
                var underline = new string('-', header.Length);
                var sb = new StringBuilder();
                sb.AppendLine(header);
                sb.AppendLine(underline);
                foreach (var line in _debugBuffer)
                    sb.AppendLine(line);
                sb.AppendLine(new string('-', underline.Length));
                sb.AppendLine();

                _logger?.Debug(sb.ToString());
                _debugBuffer.Clear();
            }
        }

        #endregion

        public override object InitializeLifetimeService() => null;
    }
}
