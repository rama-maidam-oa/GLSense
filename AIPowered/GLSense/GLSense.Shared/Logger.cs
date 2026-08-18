// Logger.cs in GLSense.Shared
using GLSense.Contracts;
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;

namespace GLSense.Shared
{
    public class Logger : MarshalByRefObject, Contracts.ILogger
    {
        private readonly NLog.Logger _logger;
        private readonly bool _isInitialized = false;
        private readonly object _initLock = new();
        private readonly IGLSenseContext _context;  // Reference to context

        // DebugMode comes from context
        private bool DebugMode => _context?.DebugMode ?? false;

        // ================= Per-action debug-log buffering =================
        //
        // Ported from FinalWorkingCode's identical LogUtility.cs overhaul - see that
        // file's header comment for the full design rationale. Debug-mode log lines are
        // buffered per logical action (one top-level LogScope - one ribbon click, one
        // API call, one window's lifecycle) and flushed to disk as one batched write
        // when the outermost scope closes, instead of one file open+write+flush+close
        // cycle per line (NLog's FileTarget here uses AutoFlush=true).
        //
        // This Logger instance lives in the host's own AppDomain (created once by
        // GLSenseContext, see that file) - every call made from GLSense.Addin.Core code
        // crosses the AppDomain boundary via this MarshalByRefObject proxy. AsyncLocal
        // still applies correctly here: a plain synchronous cross-domain call does not
        // disrupt ExecutionContext flow any differently than a same-domain nested call
        // would, and the case that actually matters - Addin.Core's own async
        // continuations resuming on a different threadpool thread after
        // ConfigureAwait(false) - is exactly what AsyncLocal is designed to follow
        // regardless of which domain the method body executes in.
        private sealed class ActionBuffer
        {
            public readonly Guid Id = Guid.NewGuid();
            public readonly List<string> Lines = new List<string>();
            public readonly object Lock = new object();
            public string RootScopeName;
            public int Depth;
            public DateTime OldestUnflushedAtUtc = DateTime.UtcNow;
        }

        private static readonly AsyncLocal<ActionBuffer> _currentBuffer = new AsyncLocal<ActionBuffer>();

        // Every action buffer currently open, keyed by its own id - lets the time-based
        // safety net and FlushDebugLogs (called from the shutdown/unhandled-exception
        // hooks in both AppDomains) reach buffers that live on a different async flow
        // than whichever thread happens to run them.
        private static readonly ConcurrentDictionary<Guid, ActionBuffer> _openBuffers = new ConcurrentDictionary<Guid, ActionBuffer>();

        // If a single action's buffer has had unflushed lines sitting in it longer than
        // this, the safety net flushes what's there so far (and keeps buffering under
        // the same scope) - caps how much a stuck/very-long-running action can lose.
        private static readonly TimeSpan SafetyNetMaxAge = TimeSpan.FromSeconds(30);
        private static Timer _safetyNetTimer;
        private static readonly object _safetyNetInitLock = new object();

        // Very early startup only: LogDebug/etc. calls that happen before this Logger's
        // own NLog setup has completed (_logger still null) have nowhere to write yet.
        // Held here and flushed once the logger becomes ready - unrelated to the
        // per-action buffering above, which requires the logger to already exist.
        private static readonly List<string> _startupFallbackBuffer = new List<string>();
        private static readonly object _startupFallbackLock = new object();

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
                        // Was false: with per-line writes that meant every single log
                        // call opened, wrote, flushed, and closed the file handle -
                        // expensive under verbose Debug-mode tracing. Now that logging is
                        // batched per action (see the buffering section above) and writes
                        // are infrequent-but-large instead of one-line-per-call, keeping
                        // the handle open for the session avoids repeating that open/close
                        // overhead on every flush too. AutoFlush stays true - nothing here
                        // trades durability for speed.
                        KeepFileOpen = true,
                        DeleteOldFileOnStartup = false,
                        ArchiveAboveSize = 20 * 1024 * 1024,  // 20MB archive size
                        MaxArchiveFiles = 30,
                        ArchiveFileName = archiveFileNamePattern
                    };

                    var GLSenseLoggerConfiguration = new LoggingConfiguration();
                    GLSenseLoggerConfiguration.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);

                    LogManager.Configuration = GLSenseLoggerConfiguration;

                    _logger = LogManager.GetCurrentClassLogger();
                    _currentLoggerForStaticFlush = this;

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

            AppendEnvironmentSnapshot(sb, defaultVersion, defaultCommitDate);

            return sb.ToString();
        }

        // Moved here from AddinEntry.LogEnvironmentSnapshot, which used to log this as a
        // regular LogInfo call from AddinEntry.Initialize - that method fires once per
        // Excel session, but the log file itself is per-day (GLSense_Logs_{date}.log), so
        // every subsequent Excel open on the same day re-appended an identical snapshot
        // into that day's file instead of writing it once. NLog's FileTarget.Header is
        // only written when the target actually creates a new file, so folding this into
        // the header instead makes it genuinely once-per-file (once-per-day) for free,
        // with no new file-existence tracking needed here.
        //
        // Runs from this Logger's own constructor (see above), BEFORE LogManager.Configuration
        // is assigned and before ServiceLocator is initialized - LogUtility/ServiceLocator.Logger
        // calls are not usable yet here, so failures are swallowed silently with a safe
        // fallback value instead of being logged, unlike the original method's LogWarn calls.
        // _context.ExcelApp is already assigned by GLSenseContext's constructor before it
        // creates this Logger instance, so it's available here - read via reflection (not a
        // static Excel.Application cast) since GLSense.Shared deliberately has no reference
        // to the Excel interop assembly (ExcelApp is typed `object` on IGLSenseContext for
        // exactly this reason).
        private void AppendEnvironmentSnapshot(StringBuilder sb, string version, string releaseDate)
        {
            string excelVersion = "unknown";
            try
            {
                var excelApp = _context?.ExcelApp;
                if (excelApp != null)
                {
                    excelVersion = excelApp.GetType().InvokeMember(
                        "Version", BindingFlags.GetProperty, null, excelApp, null) as string ?? "unknown";
                }
            }
            catch
            {
                // ExcelApp may not be assigned yet depending on call order - "unknown" is
                // an acceptable fallback for a one-time header line.
            }

            double dpi = 96d;
            try
            {
                using (var g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    dpi = g.DpiX;
                }
            }
            catch
            {
                // Fall back to the 96 DPI (100% scale) default set above.
            }

            sb.AppendLine("===== Environment Snapshot =====");
            sb.AppendLine($"GLSense version: {version} (released {releaseDate})");
            sb.AppendLine($"Excel version: {excelVersion}, process bitness: {(Environment.Is64BitProcess ? "64-bit" : "32-bit")}");
            sb.AppendLine($"OS: {Environment.OSVersion.VersionString}, {(Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit")} OS");
            sb.AppendLine($".NET runtime: {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}");
            sb.AppendLine($"Screen DPI: {dpi:F0} ({dpi / 96d * 100:F0}% scale)");
            sb.AppendLine($"Culture: {CultureInfo.CurrentCulture.Name} (UI: {CultureInfo.CurrentUICulture.Name})");
            sb.AppendLine($"Machine: {Environment.MachineName}, User: {Environment.UserName}");
            sb.AppendLine("=================================");
        }

        private static void EnsureSafetyNetTimerStarted()
        {
            if (_safetyNetTimer != null) return;
            lock (_safetyNetInitLock)
            {
                if (_safetyNetTimer != null) return;
                _safetyNetTimer = new Timer(_ => RunSafetyNetSweep(), null, TimeSpan.FromSeconds(15), TimeSpan.FromSeconds(15));
            }
        }

        private static void RunSafetyNetSweep()
        {
            try
            {
                var cutoff = DateTime.UtcNow - SafetyNetMaxAge;
                foreach (var kvp in _openBuffers)
                {
                    var buffer = kvp.Value;
                    bool stale;
                    lock (buffer.Lock)
                    {
                        stale = buffer.Lines.Count > 0 && buffer.OldestUnflushedAtUtc <= cutoff;
                    }
                    if (stale)
                    {
                        FlushBufferStatic(buffer, $"{buffer.RootScopeName} - safety-net flush, action still running");
                    }
                }
            }
            catch
            {
                // The safety-net timer must never itself take anything down.
            }
        }

        // Called by LogScope's constructor, BEFORE it logs its own "BEGIN:" line -
        // deliberately does not touch indentation depth, so a scope's own BEGIN/END
        // markers are logged at the SAME depth as the code that opened it; only content
        // logged *inside* the scope (between IncrementDepth/DecrementDepth) is indented
        // one level deeper. Returns the buffer this LogScope instance OWNS (and must
        // flush on Dispose), or null if this is a nested scope reusing an already-open
        // buffer.
        private object BeginScope(string scopeName)
        {
            if (!DebugMode)
                return null;

            if (_currentBuffer.Value != null)
                return null;

            var buffer = new ActionBuffer { RootScopeName = scopeName };
            _currentBuffer.Value = buffer;
            _openBuffers[buffer.Id] = buffer;
            EnsureSafetyNetTimerStarted();
            return buffer;
        }

        private void IncrementDepth()
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null)
                buffer.Depth++;
        }

        private void DecrementDepth()
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null && buffer.Depth > 0)
                buffer.Depth--;
        }

        // Called by LogScope.Dispose(), AFTER it logs its own "END:" line. "owned" is
        // whatever BeginScope returned for this same instance - only non-null for the
        // scope that actually created the buffer, which is the only one that flushes it.
        private void EndScope(object owned)
        {
            if (owned is ActionBuffer buffer)
            {
                FlushBufferStatic(buffer, buffer.RootScopeName);
                _openBuffers.TryRemove(buffer.Id, out _);
                _currentBuffer.Value = null;
            }
        }

        private string Indent()
        {
            int depth = _currentBuffer.Value?.Depth ?? 0;
            return new string(' ', Math.Max(0, depth) * 2);
        }

        private static void FlushBufferStatic(ActionBuffer buffer, string label)
        {
            List<string> lines;
            lock (buffer.Lock)
            {
                if (buffer.Lines.Count == 0) return;
                lines = new List<string>(buffer.Lines);
                buffer.Lines.Clear();
                buffer.OldestUnflushedAtUtc = DateTime.UtcNow;
            }

            var header = $"===== {label} | {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====";
            var sb = new StringBuilder();
            sb.AppendLine(header);
            sb.AppendLine(new string('-', header.Length));
            foreach (var line in lines)
                sb.AppendLine(line);
            sb.AppendLine(new string('-', header.Length));

            _currentLoggerForStaticFlush?._logger?.Debug(sb.ToString());
        }

        // FlushBufferStatic/RunSafetyNetSweep are static (the buffer registry itself is
        // static, shared across the one Logger instance that actually exists), but still
        // need an NLog.Logger to write through - captured here once at construction time.
        private static Logger _currentLoggerForStaticFlush;

        private void FlushCurrentBuffer(string reason)
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null)
                FlushBufferStatic(buffer, $"{buffer.RootScopeName} - {reason}");
        }

        #region Logging Methods

        public void LogInfo(string msg)
        {
            var logMessage = $"{Indent()}INFO  | {DateTime.Now:HH:mm:ss} | {msg}";
            WriteImmediate(logMessage);
        }

        public void LogWarn(string msg)
        {
            var logMessage = $"{Indent()}WARN  | {DateTime.Now:HH:mm:ss} | {msg}";
            WriteImmediate(logMessage);
            FlushCurrentBuffer("warning logged");
        }

        public void LogError(string msg, Exception ex = null)
        {
            var logMessage = $"{Indent()}ERROR | {DateTime.Now:HH:mm:ss} | {msg}";
            if (ex != null)
                _logger?.Error(ex, logMessage);
            else
                WriteImmediate(logMessage);
            FlushCurrentBuffer("error logged");
        }

        public void LogDebug(string msg)
        {
            if (!DebugMode)
                return;

            var logMessage = $"{Indent()}DEBUG | {DateTime.Now:HH:mm:ss} | {msg}";

            var buffer = _currentBuffer.Value;
            if (buffer == null)
            {
                // No action scope is open - nothing to group this line with, so it must
                // still reach disk on its own rather than being silently dropped.
                WriteImmediate(logMessage);
                return;
            }

            lock (buffer.Lock)
            {
                if (buffer.Lines.Count == 0)
                    buffer.OldestUnflushedAtUtc = DateTime.UtcNow;
                buffer.Lines.Add(logMessage);
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
            WriteImmediate(exceptionMessage);
            FlushCurrentBuffer("exception logged");
        }

        public void LogRawJson(string context, string rawJson)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{Indent()}----- Raw JSON {(string.IsNullOrWhiteSpace(context) ? string.Empty : "(" + context + ")")} -----");
            sb.AppendLine(string.IsNullOrEmpty(rawJson) ? "<empty>" : rawJson);
            sb.AppendLine($"{Indent()}----- End Raw JSON -----");
            WriteImmediate(sb.ToString());
        }

        #endregion

        private void WriteImmediate(string logMessage)
        {
            if (_logger != null)
            {
                if (_startupFallbackBuffer.Count > 0)
                    FlushStartupFallbackBuffer();

                _logger.Debug(logMessage);
            }
            else
            {
                lock (_startupFallbackLock)
                {
                    _startupFallbackBuffer.Add(logMessage);
                }
            }
        }

        private void FlushStartupFallbackBuffer()
        {
            List<string> pending;
            lock (_startupFallbackLock)
            {
                if (_startupFallbackBuffer.Count == 0) return;
                pending = new List<string>(_startupFallbackBuffer);
                _startupFallbackBuffer.Clear();
            }

            foreach (var line in pending)
                _logger.Debug(line);
        }

        #region Scope Methods

        public void LogMethodEntry([CallerMemberName] string methodName = "")
        {
            LogDebug($"Entering {methodName}");
        }

        public void LogMethodExit([CallerMemberName] string methodName = "")
        {
            LogDebug($"Exiting {methodName}");
        }

        // MarshalByRefObject so a LogScope created here (inside the host AppDomain, by
        // BeginLogScope below) crosses back to GLSense.Addin.Core's AppDomain as a
        // remoting proxy rather than being marshaled by value - Dispose() on the
        // Addin.Core side needs to call back into THIS instance's fields (_logger,
        // _ownedBuffer) to correctly close the same buffer it opened, which a
        // by-value copy in the other AppDomain could never do.
        public class LogScope : MarshalByRefObject, IDisposable
        {
            private readonly string _scopeName;
            private readonly Logger _logger;
            private readonly object _ownedBuffer;
            private bool _disposed = false;

            public LogScope(Logger logger, string scopeName)
            {
                _logger = logger;
                _scopeName = scopeName;
                _ownedBuffer = _logger.BeginScope(scopeName);
                _logger.LogDebug($"BEGIN: {_scopeName}");
                _logger.IncrementDepth();
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
                        _logger.DecrementDepth();
                        _logger.LogDebug($"END: {_scopeName}");
                        _logger.EndScope(_ownedBuffer);
                    }
                    _disposed = true;
                }
            }

            public override object InitializeLifetimeService() => null;
        }

        /// <summary>
        /// ILogger-facing entry point for opening a buffered action scope from
        /// GLSense.Addin.Core code, which only ever holds an ILogger reference (never a
        /// concrete Logger) - see ILogger.BeginLogScope's doc comment.
        /// </summary>
        public IDisposable BeginLogScope(string scopeName) => new LogScope(this, scopeName);

        #endregion

        #region Flush

        /// <summary>
        /// Flushes every action buffer currently open, whatever state it's in. Used when
        /// the Debug checkbox is switched off mid-action, and by the shutdown and
        /// unhandled-exception hooks (in both AppDomains - this Logger instance is
        /// reached identically from either side via the ILogger contract), so nothing
        /// buffered is ever silently lost.
        /// </summary>
        public void FlushDebugLogs(string section = "Buffered Logs")
        {
            foreach (var kvp in _openBuffers)
            {
                FlushBufferStatic(kvp.Value, $"{kvp.Value.RootScopeName} - {section}");
            }
        }

        #endregion

        public override object InitializeLifetimeService() => null;
    }
}
