using GLSense.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Helpers
{
#nullable enable
    /// <summary>
    /// Manages a singleton-like guard for the session so you can call Apply/Restore from anywhere.
    /// </summary>
    public static class ExcelErrorCheckingService
    {
        private static ExcelErrorCheckingGuard? _guard;

        /// <summary>
        /// Capture current settings and apply your session values (no-op if already applied).
        /// </summary>
        public static void Apply(Excel.Application app)
        {
            if (_guard != null)
            {
                LogUtility.LogDebug("ExcelErrorCheckingService.Apply: guard already applied, no-op.");
                return;        // already applied
            }

            LogUtility.LogDebug("ExcelErrorCheckingService.Apply: creating and applying guard.");
            _guard = new ExcelErrorCheckingGuard(app);
            _guard.Apply();
        }

        /// <summary>
        /// Restore original settings if previously applied.
        /// </summary>
        public static void Restore()
        {
            LogUtility.LogDebug("ExcelErrorCheckingService.Restore: disposing guard (if any).");
            _guard?.Dispose();
            _guard = null;
        }
    }
#nullable disable
}
