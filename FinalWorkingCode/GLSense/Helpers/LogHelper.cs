using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.IO;
using System.Net.NetworkInformation;
using System.Text;
using System.Windows.Forms;

namespace GLSense.Helpers
{
    public static class LogHelper
    {
        private static bool _isInitialized = false;
        private static readonly object _initLock = new();

        public static void InitializeLogger()
        {
            lock (_initLock)
            {
                if (_isInitialized) return;

                try
                {
                    // Create header layout with dynamic date evaluation
                    string HdrText = BuildLogHeader();

                    // IMPORTANT: Use NLog's date pattern, NOT a pre-evaluated date
                    string fileNamePattern = AppPaths.LogFolder + @"\GLSense_Logs_${date:format=dd-MMM-yyyy}.log";
                    var fileNameLayout = NLog.Layouts.Layout.FromString(fileNamePattern);

                    var logfile = new FileTarget("logfile")
                    {
                        FileName = fileNameLayout,  // This will be evaluated at runtime
                        Header = HdrText,  // NLog will write this header when creating new files
                        AutoFlush = true,
                        Layout = "${longdate}|${level:uppercase=true}|${message:withException=true:exceptionSeparator=|}",
                        // Was false: with per-line writes that meant every single log call
                        // opened, wrote, flushed, and closed the file handle - expensive
                        // under verbose Debug-mode tracing. Now that LogUtility batches
                        // Debug-mode logging per action (see its own header comment) and
                        // writes are infrequent-but-large instead of one-line-per-call,
                        // keeping the handle open for the session avoids repeating that
                        // open/close overhead on every flush too. AutoFlush=true still
                        // guarantees each write reaches disk immediately - nothing here
                        // trades durability for speed, only removes redundant open/close
                        // cycles.
                        KeepFileOpen = true,
                        DeleteOldFileOnStartup = false,
                        ArchiveAboveSize = AppConstants.LogMaxFileSizeBytes,  // 20MB archive size
                        MaxArchiveFiles = AppConstants.LogMaxArchiveFiles,
                        ArchiveFileName = AppPaths.LogFolder + @"\GLSense_Logs_{#}.log"
                    };

                    var GLSenseLoggerConfiguration = new LoggingConfiguration();

                    // Single rule covering the required level range to avoid duplicate routing
                    // Multiple overlapping rules that target the same target will cause messages
                    // to be written multiple times. Use a single rule instead.
                    GLSenseLoggerConfiguration.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);

                    LogManager.Configuration = GLSenseLoggerConfiguration;

                    // Store the logger instance
                    AddinModule.LoggerConfiguration = GLSenseLoggerConfiguration;
                    AddinModule.Logger = LogManager.GetCurrentClassLogger();

                    _isInitialized = true;
                }
                catch (Exception ex)
                {
                    // LogError equivalent in C#
                    MessageBox.Show($"Error initializing logger: {ex.Message}", "Orbit GLSense", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // You might want to implement proper error logging here
                }
            }
        }

        private static string BuildLogHeader()
        {

            var sb = new StringBuilder();

            string header = $"Orbit GLSense logs generated on : {AppConstants.DefaultCommitDate}). Logs As On {DateTime.Now:dddd, dd MMMM yyyy}. Time Zone: {TimeZoneInfo.Local.DisplayName}";
            sb.AppendLine(header);

            // Add underline that exactly matches the header length in characters
            sb.AppendLine(new string('-', header.Length));

            return sb.ToString();
        }
    }
}
