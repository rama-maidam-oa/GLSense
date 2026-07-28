using System;
using System.Collections.Generic;
using System.Text;

namespace GLSense.Utilities
{
    public static class LogUtility
    {
        // Buffer for debug logs (used when DebugMode is true)
        private static readonly List<string> _debugBuffer = new List<string>();

        // Toggle debug buffering
        public static bool DebugMode => AppState.Instance.DebugLogs;

        // Thread-local scope depth for nested indentation
        [ThreadStatic]
        private static int _scopeDepth;

        // Duplicate exception prevention
        private static readonly object _lock = new object();
        private static readonly TimeSpan _dedupeInterval = TimeSpan.FromSeconds(5);
        private static readonly Dictionary<string, DateTime> _exceptionTimestamps = new Dictionary<string, DateTime>();

        internal static void IncrementScope()
        {
            _scopeDepth = Math.Max(0, _scopeDepth) + 1;
        }

        internal static void DecrementScope()
        {
            _scopeDepth = Math.Max(0, _scopeDepth - 1);
        }

        private static string Indent()
        {
            int safeDepth = Math.Max(0, _scopeDepth);
            return new string(' ', safeDepth * 2);
        }

        #region Logging Methods
        public static void LogInfo(string message)
        {
            var logMessage = $"{Indent()}INFO  | {DateTime.Now:HH:mm:ss} | {message}";
            AddinModule.Logger?.Info(logMessage);
        }

        public static void LogWarn(string message)
        {
            var logMessage = $"{Indent()}WARN  | {DateTime.Now:HH:mm:ss} | {message}";
            AddinModule.Logger?.Warn(logMessage);
        }

        public static void LogError(string message)
        {
            var logMessage = $"{Indent()}ERROR | {DateTime.Now:HH:mm:ss} | {message}";
            AddinModule.Logger?.Error(logMessage);
        }

        public static void LogDebug(string message)
        {
            if (!DebugMode)
                return;

            var logMessage = $"{Indent()}DEBUG | {DateTime.Now:HH:mm:ss} | {message}";

            var logger = AddinModule.Logger;
            if (logger != null)
            {
                logger.Debug(logMessage);
            }
            else
            {
                _debugBuffer.Add(logMessage);
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
            AddinModule.Logger?.Error(exceptionMessage);
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
            AddinModule.Logger?.Error(sb.ToString());
        }
        #endregion

        #region Flush
        public static void FlushDebugLogs(string section = "Buffered Logs")
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

            AddinModule.Logger?.Debug(sb.ToString());
            _debugBuffer.Clear();
        }
        #endregion

        #region Additional Helper Methods (Optional)
        public static void LogMethodEntry([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            LogDebug($"Entering {methodName}");
            IncrementScope();
        }

        public static void LogMethodExit([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            DecrementScope();
            LogDebug($"Exiting {methodName}");
        }

        public class LogScope : IDisposable
        {
            private readonly string _scopeName;
            private bool _disposed = false;
            public LogScope(string scopeName)
            {
                _scopeName = scopeName;
                LogUtility.LogDebug($"BEGIN: {_scopeName}");
                LogUtility.IncrementScope();
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
                        LogUtility.DecrementScope();
                        LogUtility.LogDebug($"END: {_scopeName}");
                    }
                    _disposed = true;
                }
            }
        }
        #endregion
    }
}