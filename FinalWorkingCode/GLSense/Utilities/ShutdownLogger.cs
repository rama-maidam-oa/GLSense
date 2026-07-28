using System;
using System.Diagnostics;
using System.IO;

namespace GLSense.Utilities
{
    public static class ShutdownLogger
    {
        private static readonly string LogPath;
        static ShutdownLogger()
        {
            try
            {
                string logFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ORBIT", "Excel_Logs", "GLSense_Logs", "Logs");

                Directory.CreateDirectory(logFolder);
                LogPath = Path.Combine(logFolder, $"shutdown_errors_{DateTime.Now:yyyyMMdd}.log");
            }
            catch (Exception ex)
            {
                // Non-fatal: fallback to temp path. Log to debug so we have visibility during development.
                LogPath = Path.Combine(Path.GetTempPath(), $"ExcelShutdownErrors_{DateTime.Now:yyyyMMdd}.log");
                Debug.WriteLine($"ShutdownLogger: could not create log folder, using temp path. Error: {ex.Message}");
            }
        }

        public static void LogError(string message, Exception ex = null)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string errorMsg = ex != null
                    ? $"{message} - Exception: {ex.Message}\nStack Trace: {ex.StackTrace}"
                    : message;

                string logEntry = $"{timestamp} [ERROR] {errorMsg}";
                File.AppendAllText(LogPath, logEntry + Environment.NewLine);
                Debug.WriteLine(logEntry);
            }
            catch (Exception logEx)
            {
                // During shutdown avoid throwing; record to debug so we can diagnose if necessary
                Debug.WriteLine($"ShutdownLogger.LogError failed: {logEx.Message} (non-critical during shutdown)");
            }
        }

        public static void LogWarn(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string logEntry = $"{timestamp} [WARN] {message}";
                File.AppendAllText(LogPath, logEntry + Environment.NewLine);
                Debug.WriteLine(logEntry);
            }
            catch (Exception logEx)
            {
                Debug.WriteLine($"ShutdownLogger.LogWarn failed: {logEx.Message} (non-critical during shutdown)");
            }
        }
    }
}
