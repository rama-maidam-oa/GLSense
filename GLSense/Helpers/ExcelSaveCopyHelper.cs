using GLSense.Utilities;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Helpers
{
    public static class ExcelSaveCopyHelper
    {

        /// <summary>
        /// Saves a copy of the active workbook to the specified <paramref name="destinationPath"/>.
        /// All COM calls happen on the caller's thread (must be the Excel UI/STA thread).
        /// Then waits until the file is readable/unlocked. Returns the path on success.
        /// </summary>
        /// <param name="excelApp">The running Excel Application (e.g., Globals.ThisAddIn.Application).</param>
        /// <param name="destinationPath">Target .xlsx path (directory must exist).</param>
        /// <param name="timeout">Max time to wait for file readiness.</param>
        /// <param name="ct">Cancellation token.</param>
        public static async Task<string> SaveActiveWorkbookCopyAsync(
            Excel.Application excelApp,
            string destinationPath,
            TimeSpan timeout,
            CancellationToken ct = default)
        {
            if (excelApp == null) throw new ArgumentNullException(nameof(excelApp));
            if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentException("Destination path is required.", nameof(destinationPath));
            if (!Path.GetExtension(destinationPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Destination must be a .xlsx path.", nameof(destinationPath));

            // Ensure directory exists
            var destDir = Path.GetDirectoryName(destinationPath);
            if (string.IsNullOrWhiteSpace(destDir))
                throw new ArgumentException("Invalid destination directory.", nameof(destinationPath));
            Directory.CreateDirectory(destDir);

            LogUtility.LogDebug($"ExcelSaveCopyHelper.SaveActiveWorkbookCopyAsync: saving copy to '{destinationPath}'");

            // COM boundary: get active workbook and save copy synchronously (must be on Excel thread)
            Excel.Workbook wb = null;
            try
            {
                wb = excelApp.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook.");
                wb.SaveCopyAs(destinationPath);
            }
            finally
            {
                // Release the COM reference we created
                SafeFinalReleaseCom(wb);
            }

            // Wait until the saved copy is fully written & unlocked (do NOT touch COM during await)
            await WaitForFileReadyAsync(destinationPath, timeout, ct).ConfigureAwait(false);

            LogUtility.LogDebug($"ExcelSaveCopyHelper.SaveActiveWorkbookCopyAsync: copy saved and ready at '{destinationPath}'");

            return destinationPath;
        }

        /// <summary>
        /// Generates a unique temp file path for saving a copy (e.g., in %TEMP%\YourApp\yyyyMMdd\GUID.xlsx).
        /// </summary>
        public static string CreateUniqueTempCopyPath(string prefix = "wbcopy_", string subfolder = "YourApp")
        {
            var root = Path.Combine(Path.GetTempPath(), subfolder, DateTime.UtcNow.ToString("yyyyMMdd"));
            Directory.CreateDirectory(root);
            var file = $"{prefix}{Guid.NewGuid():N}.xlsx";
            return Path.Combine(root, file);
        }

        // -----------------------------
        // Helpers (file readiness & COM)
        // -----------------------------

        private static async Task WaitForFileReadyAsync(string path, TimeSpan timeout, CancellationToken ct)
        {
            var deadline = DateTime.UtcNow + timeout;
            Exception last = null;

            while (DateTime.UtcNow < deadline)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    // Attempt exclusive open → ensures writer/locks are released
                    using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
                    // Optionally touch the stream to ensure read is OK
                    if (fs.Length >= 0) { /* noop */ }
                    return;
                }
                catch (Exception ex)
                {
                    last = ex;
                    LogUtility.LogDebug($"ExcelSaveCopyHelper.WaitForFileReadyAsync: file '{path}' not yet ready ({ex.GetType().Name}: {ex.Message}), retrying...");
                    await Task.Delay(100, ct).ConfigureAwait(false);
                }
            }

            LogUtility.LogWarn($"ExcelSaveCopyHelper.WaitForFileReadyAsync: file '{path}' not ready after {timeout}.");
            throw new IOException($"File '{path}' not ready after {timeout}.", last);
        }

        private static void SafeFinalReleaseCom(object com)
        {
            if (com != null && Marshal.IsComObject(com))
            {
                try { Marshal.FinalReleaseComObject(com); } catch { /* ignore */ }
            }
        }

    }
}
