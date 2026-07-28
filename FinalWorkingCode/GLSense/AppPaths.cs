using System;
using System.Diagnostics;
using System.IO;
using GLSense.Utilities;

namespace GLSense
{
    public static class AppPaths
    {
        private static readonly Lazy<string> _baseFolder = new(() =>
        {
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"ORBIT\Excel_Logs");

            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _glSenseLogsFolder = new(() =>
        {
            string path = Path.Combine(BaseFolder, "GLSense_Logs");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _logFolder = new(() =>
        {
            string path = Path.Combine(GLSenseLogsFolder, "Logs");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _databasePath = new(() =>
            Path.Combine(GLSenseLogsFolder, "Database", "GLSense.sqlite"));

        private static readonly Lazy<string> _tempFilesPath = new(() =>
        {
            string path = Path.Combine(GLSenseLogsFolder, "TempFiles");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _tempUrlsPath = new(() =>
        {
            string path = Path.Combine(BaseFolder, "ORBIT_URLS.xml");
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                EnsureDirectoryExists(directory);
            }
            return path;
        });

        private static readonly Lazy<string> _browserLogsFolder = new(() =>
        {
            string path = Path.Combine(GLSenseLogsFolder, "BrowserLogs");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _drilldownBrowserLogs = new(() =>
        {
            string path = Path.Combine(BrowserLogsFolder, "Drilldown");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _loginBrowserLogs = new(() =>
        {
            string path = Path.Combine(BrowserLogsFolder, "Login");
            EnsureDirectoryExists(path);
            return path;
        });

        // Public properties
        public static string BaseFolder => _baseFolder.Value;
        public static string LogFolder => _logFolder.Value;
        public static string GLSenseLogsFolder => _glSenseLogsFolder.Value;
        public static string DatabasePath => _databasePath.Value;
        public static string TempFilesPath => _tempFilesPath.Value;
        public static string TempUrlsPath => _tempUrlsPath.Value;
        public static string BrowserLogsFolder => _browserLogsFolder.Value;
        public static string DrilldownBrowserLogsPath => _drilldownBrowserLogs.Value;
        public static string LoginBrowserLogsPath => _loginBrowserLogs.Value;

        private static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                try { Directory.CreateDirectory(path); }
                catch (Exception ex)
                {
                    // Best-effort: LogUtility may not be initialized yet this early (e.g. when
                    // resolving the log folder itself, before LogHelper.InitializeLogger runs),
                    // so also emit to Debug output as a fallback for that chicken-and-egg case.
                    Debug.WriteLine($"AppPaths.EnsureDirectoryExists: failed to create '{path}': {ex.Message}");
                    LogUtility.LogException(ex, $"AppPaths.EnsureDirectoryExists: failed to create '{path}'");
                }
            }
        }
    }
}
