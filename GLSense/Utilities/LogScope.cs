using System;
using System.Diagnostics;

namespace GLSense.Utilities
{
    public sealed class LogScope : IDisposable
    {
        private readonly string _sectionName;
        private readonly Stopwatch _stopwatch;

        public LogScope(string sectionName)
        {
            _sectionName = sectionName;
            _stopwatch = Stopwatch.StartNew();

            LogUtility.IncrementScope();
            if (LogUtility.DebugMode) LogUtility.LogDebug($"▶ Entering: {_sectionName}");
        }

        public void Dispose()
        {
            _stopwatch.Stop();
            if (LogUtility.DebugMode) LogUtility.LogDebug($"◀ Exiting: {_sectionName} (Duration: {_stopwatch.Elapsed.TotalSeconds:F3}s)");

            // Flush buffered debug logs automatically
            if (LogUtility.DebugMode)
                LogUtility.FlushDebugLogs(_sectionName);

            LogUtility.DecrementScope();
        }
    }
}
