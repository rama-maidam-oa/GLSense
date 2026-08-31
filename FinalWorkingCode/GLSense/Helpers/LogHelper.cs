using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Globalization;
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
                        // Deliberately no ArchiveFileName: FileName already uses ${date} (dynamic
                        // per-day naming), and NLog's own guidance is to never combine that with an
                        // explicit ArchiveFileName - doing so forces the "Legacy/unstable" file-move
                        // archive handler (FileTarget.cs's CreateFileArchiveHandler), which fights
                        // KeepFileOpen's exclusive lock on the active file. Leaving ArchiveFileName
                        // unset routes size-based rollover through NLog 6's RollingArchiveFileHandler
                        // instead - it opens a new, already-numbered file rather than renaming the
                        // full one, so there's no lock contention. ArchiveSuffixFormat only gets
                        // appended once sequenceNumber > 0 (BuildFullFilePath), so today's first/
                        // active chunk stays plain GLSense_Logs_{date}.log, and each subsequent
                        // 20MB rollover produces GLSense_Logs_{date}(1).log, (2).log, etc.
                        ArchiveSuffixFormat = "({0})"
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

            AppendEnvironmentSnapshot(sb);

            return sb.ToString();
        }

        // Moved here from AddinModule.LogEnvironmentSnapshot, which used to log this as a
        // regular LogInfo call from AddinModule_OnRibbonLoaded - that method fires once per
        // Excel session, but the log file itself is per-day
        // (GLSense_Logs_{date}.log), so every subsequent Excel open on the same day
        // re-appended an identical snapshot into that day's file instead of writing it
        // once. NLog's FileTarget.Header is only written when the target actually creates
        // a new file, so folding this into the header instead makes it genuinely
        // once-per-file (once-per-day) for free, with no new file-existence tracking
        // needed here.
        //
        // Runs BEFORE LogManager.Configuration is assigned (see InitializeLogger above),
        // so LogUtility/AddinModule.Logger calls are not usable yet here - failures are
        // swallowed silently with a safe fallback value instead of being logged, unlike
        // the original method's LogWarn calls.
        private static void AppendEnvironmentSnapshot(StringBuilder sb)
        {
            string excelVersion = "unknown";
            try
            {
                excelVersion = AppState.Instance.ExcelApp?.Version ?? "unknown";
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
            sb.AppendLine($"GLSense version: {AppConstants.DefaultVersion} (released {AppConstants.DefaultCommitDate})");
            sb.AppendLine($"Excel version: {excelVersion}, process bitness: {(Environment.Is64BitProcess ? "64-bit" : "32-bit")}");
            sb.AppendLine($"OS: {Environment.OSVersion.VersionString}, {(Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit")} OS");
            sb.AppendLine($".NET runtime: {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}");
            sb.AppendLine($"Screen DPI: {dpi:F0} ({dpi / 96d * 100:F0}% scale)");
            sb.AppendLine($"Culture: {CultureInfo.CurrentCulture.Name} (UI: {CultureInfo.CurrentUICulture.Name})");
            sb.AppendLine($"Machine: {Environment.MachineName}, User: {Environment.UserName}");
            sb.AppendLine("=================================");
        }
    }
}
