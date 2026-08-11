using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace GLSense.Utilities
{
    public static class LogUtility
    {
        // Toggle debug buffering - the ribbon's Debug checkbox (RibDebug_OnClick).
        public static bool DebugMode => AppState.Instance.DebugLogs;

        // Duplicate exception prevention
        private static readonly object _lock = new object();
        private static readonly TimeSpan _dedupeInterval = TimeSpan.FromSeconds(5);
        private static readonly Dictionary<string, DateTime> _exceptionTimestamps = new Dictionary<string, DateTime>();

        // ================= Per-action debug-log buffering =================
        //
        // Debug-mode logging used to write straight to NLog on every single call - and
        // NLog's FileTarget here is configured with AutoFlush=true/KeepFileOpen=false
        // (see LogHelper.cs), meaning every LogDebug call opened, wrote, flushed, and
        // closed the log file. Fine for the occasional warning, expensive for verbose
        // Debug-mode tracing (full request/response payloads, per-cell operations, etc).
        //
        // Instead: log lines produced while Debug mode is on are buffered per logical
        // "action" - one top-level LogScope (one ribbon click, one API call, one
        // window's lifecycle) - and only actually written to disk once, as a single
        // batched block, when that outermost scope closes. Nested LogScopes (very
        // common - e.g. "WebView2 Initialization" inside a ribbon click) share the same
        // buffer as their parent; only the scope that CREATED the buffer flushes it.
        //
        // AsyncLocal<T>, not [ThreadStatic] (which is what the old scope-depth counter
        // used): this codebase awaits everywhere with ConfigureAwait(false), so a
        // logical action's continuations routinely resume on a different threadpool
        // thread than the one that started them - [ThreadStatic] does not reliably
        // follow that. AsyncLocal does follow the logical async/await flow regardless of
        // which thread each continuation actually lands on, and as a side effect, two
        // unrelated concurrent actions (e.g. a background refresh running while the user
        // clicks another ribbon button) each get their own isolated buffer for free,
        // since forking an async flow (Task.Run, parallel awaits, etc.) copies the
        // *pointer* to the current buffer, never lets a child flow's writes bleed back
        // into an unrelated sibling/parent flow's buffer.
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
        // safety net and the global shutdown/unhandled-exception flush reach buffers
        // that live on a different async flow than whichever thread happens to run them.
        private static readonly ConcurrentDictionary<Guid, ActionBuffer> _openBuffers = new ConcurrentDictionary<Guid, ActionBuffer>();

        // If a single action's buffer has had unflushed lines sitting in it longer than
        // this, the safety net flushes what's there so far (and keeps buffering under
        // the same scope) - caps how much a stuck/very-long-running action can lose if
        // it never reaches a clean Dispose.
        private static readonly TimeSpan SafetyNetMaxAge = TimeSpan.FromSeconds(30);
        private static Timer _safetyNetTimer;
        private static readonly object _safetyNetInitLock = new object();

        // Very early startup only: LogDebug/LogInfo/etc. calls that happen before
        // LogHelper.InitializeLogger() has run (AddinModule.Logger is still null) have
        // nowhere to write to yet. Held here and flushed once the logger becomes ready -
        // unrelated to the per-action buffering above, which requires the logger to
        // already exist.
        private static readonly List<string> _startupFallbackBuffer = new List<string>();

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
                        FlushBuffer(buffer, $"{buffer.RootScopeName} - safety-net flush, action still running");
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
        // logged *inside* the scope (between IncrementDepth/DecrementDepth below) is
        // indented one level deeper. Returns the buffer this LogScope instance OWNS (and
        // must flush on Dispose), or null if this is a nested scope reusing an
        // already-open buffer.
        internal static object BeginScope(string scopeName)
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

        internal static void IncrementDepth()
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null)
                buffer.Depth++;
        }

        internal static void DecrementDepth()
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null && buffer.Depth > 0)
                buffer.Depth--;
        }

        // Called by LogScope.Dispose(), AFTER it logs its own "END:" line. "owned" is
        // whatever BeginScope returned for this same instance - only non-null for the
        // scope that actually created the buffer, which is the only one that flushes it.
        internal static void EndScope(object owned)
        {
            if (owned is ActionBuffer buffer)
            {
                FlushBuffer(buffer, buffer.RootScopeName);
                _openBuffers.TryRemove(buffer.Id, out _);
                _currentBuffer.Value = null;
            }
        }

        private static string Indent()
        {
            int depth = _currentBuffer.Value?.Depth ?? 0;
            return new string(' ', Math.Max(0, depth) * 2);
        }

        private static void FlushBuffer(ActionBuffer buffer, string label)
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

            AddinModule.Logger?.Debug(sb.ToString());
        }

        /// <summary>
        /// Flushes every action buffer currently open, whatever state it's in. Used when
        /// the Debug checkbox is switched off mid-action, and by the global shutdown and
        /// unhandled-exception hooks, so nothing buffered is ever silently lost.
        /// </summary>
        public static void FlushAllOpenBuffers(string reason)
        {
            foreach (var kvp in _openBuffers)
            {
                FlushBuffer(kvp.Value, $"{kvp.Value.RootScopeName} - {reason}");
            }
        }

        #region Logging Methods
        public static void LogInfo(string message)
        {
            var logMessage = $"{Indent()}INFO  | {DateTime.Now:HH:mm:ss} | {message}";
            WriteImmediate(logMessage);
        }

        public static void LogWarn(string message)
        {
            var logMessage = $"{Indent()}WARN  | {DateTime.Now:HH:mm:ss} | {message}";
            WriteImmediate(logMessage);
            FlushCurrentBuffer("warning logged");
        }

        public static void LogError(string message)
        {
            var logMessage = $"{Indent()}ERROR | {DateTime.Now:HH:mm:ss} | {message}";
            WriteImmediate(logMessage);
            FlushCurrentBuffer("error logged");
        }

        public static void LogDebug(string message)
        {
            if (!DebugMode)
                return;

            var logMessage = $"{Indent()}DEBUG | {DateTime.Now:HH:mm:ss} | {message}";

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

        public static void LogException(Exception ex, string context = "", bool forceLog = false)
        {
            // If forceLog is true, skip deduplication
            if (!forceLog)
            {
                // Generate unique key for this exception
                string exceptionKey = GenerateExceptionKey(ex, context);

                lock (_lock)
                {
                    // Check if this exception was already logged recently
                    if (_exceptionTimestamps.TryGetValue(exceptionKey, out DateTime lastLogged))
                    {
                        // If logged within the dedupe interval, skip
                        if (DateTime.UtcNow - lastLogged < _dedupeInterval)
                        {
                            LogDebug($"Duplicate exception suppressed: {ex.Message}");
                            return;
                        }
                    }

                    // Update timestamp
                    _exceptionTimestamps[exceptionKey] = DateTime.UtcNow;

                    // Keep dictionary from growing too large
                    if (_exceptionTimestamps.Count > 1000)
                    {
                        CleanupOldEntries();
                    }
                }
            }

            // Build and log the exception message
            var sb = new StringBuilder();
            sb.AppendLine($"{Indent()}========== Exception ==========");
            if (!string.IsNullOrEmpty(context))
                sb.AppendLine($"{Indent()}Context: {context}");
            sb.AppendLine($"{Indent()}Type: {ex.GetType().FullName}");
            sb.AppendLine($"{Indent()}Message: {ex.Message}");
            sb.AppendLine($"{Indent()}Source: {ex.Source}");
            sb.AppendLine($"{Indent()}TargetSite: {ex.TargetSite}");
            sb.AppendLine($"{Indent()}StackTrace:");
            foreach (var line in ex.StackTrace?.Split(new[] { Environment.NewLine }, StringSplitOptions.None) ?? [])
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

        private static void FlushCurrentBuffer(string reason)
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null)
                FlushBuffer(buffer, $"{buffer.RootScopeName} - {reason}");
        }

        private static string GenerateExceptionKey(Exception ex, string context)
        {
            // Create a unique key based on exception type and message only (ignoring stack trace)
            var keyBuilder = new StringBuilder();
            keyBuilder.Append(context ?? "");
            keyBuilder.Append("|");
            keyBuilder.Append(ex.GetType().FullName);
            keyBuilder.Append("|");

            // Use the first 200 characters of the message as key
            string messageKey = ex.Message.Length > 200 ? ex.Message.Substring(0, 200) : ex.Message;
            keyBuilder.Append(messageKey);

            // Include inner exception if present
            if (ex.InnerException != null)
            {
                keyBuilder.Append("|Inner:");
                keyBuilder.Append(ex.InnerException.GetType().FullName);
                keyBuilder.Append("|");
                string innerMsg = ex.InnerException.Message.Length > 200 ?
                    ex.InnerException.Message.Substring(0, 200) :
                    ex.InnerException.Message;
                keyBuilder.Append(innerMsg);
            }

            return keyBuilder.ToString();
        }

        private static void CleanupOldEntries()
        {
            var cutoff = DateTime.UtcNow - TimeSpan.FromMinutes(10);
            var keysToRemove = new List<string>();

            foreach (var kvp in _exceptionTimestamps)
            {
                if (kvp.Value < cutoff)
                {
                    keysToRemove.Add(kvp.Key);
                }
            }

            foreach (var key in keysToRemove)
            {
                _exceptionTimestamps.Remove(key);
            }
        }

        public static void ClearExceptionCache()
        {
            lock (_lock)
            {
                _exceptionTimestamps.Clear();
            }
        }

        public static void LogRawJson(string context, string rawJson)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{Indent()}----- Raw JSON {(string.IsNullOrWhiteSpace(context) ? string.Empty : "(" + context + ")")} -----");
            sb.AppendLine(string.IsNullOrEmpty(rawJson) ? "<empty>" : rawJson);
            sb.AppendLine($"{Indent()}----- End Raw JSON -----");
            WriteImmediate(sb.ToString());
        }
        #endregion

        private static void WriteImmediate(string logMessage)
        {
            var logger = AddinModule.Logger;
            if (logger != null)
            {
                if (_startupFallbackBuffer.Count > 0)
                    FlushStartupFallbackBuffer(logger);

                logger.Debug(logMessage);
            }
            else
            {
                lock (_lock)
                {
                    _startupFallbackBuffer.Add(logMessage);
                }
            }
        }

        private static void FlushStartupFallbackBuffer(NLog.Logger logger)
        {
            List<string> pending;
            lock (_lock)
            {
                if (_startupFallbackBuffer.Count == 0) return;
                pending = new List<string>(_startupFallbackBuffer);
                _startupFallbackBuffer.Clear();
            }

            foreach (var line in pending)
                logger.Debug(line);
        }

        #region Additional Helper Methods (Optional)
        public static void LogMethodEntry([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            LogDebug($"Entering {methodName}");
        }

        public static void LogMethodExit([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            LogDebug($"Exiting {methodName}");
        }

        public sealed class LogScope : IDisposable
        {
            private readonly string _scopeName;
            private readonly object _ownedBuffer;
            private bool _disposed;

            public LogScope(string scopeName)
            {
                _scopeName = scopeName;
                _ownedBuffer = LogUtility.BeginScope(scopeName);
                LogUtility.LogDebug($"BEGIN: {_scopeName}");
                LogUtility.IncrementDepth();
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private void Dispose(bool disposing)
            {
                if (_disposed) return;
                if (disposing)
                {
                    LogUtility.DecrementDepth();
                    LogUtility.LogDebug($"END: {_scopeName}");
                    LogUtility.EndScope(_ownedBuffer);
                }
                _disposed = true;
            }
        }
        #endregion
    }
}
