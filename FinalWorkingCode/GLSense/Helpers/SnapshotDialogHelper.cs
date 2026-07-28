using GLSense.Utilities;
using GLSense.Views;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Helpers
{
    public static class SnapshotDialogHelper
    {

        /// <summary>
        /// End-to-end: hide progress, prompt for snapshot file path, then resume progress.
        /// Returns the selected file path or null if cancelled.
        /// </summary>

        public static async Task<string> PromptSnapshotAsync(
            GLWaitWindow progressWindow,           // WPF progress window; pass null if not using WPF
            CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            Excel.Application excelApp = AppState.Instance.ExcelApp;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;

            // Hide the progress UI (equivalent to CloseProgressNew)
            await HideProgressAsync(progressWindow);

            // Resolve a safe initial directory
            string initialDir = ResolveInitialDirectory(workbook);

            // Build dialog title from active workbook
            string title = BuildDialogTitle(excelApp);

            // Show the SaveFileDialog on an STA thread and await the result
            var (success, filePath) = await ShowSaveFileDialogOnStaAsync(initialDir, title, ct);

            // If the user cancelled, just return
            if (!success || string.IsNullOrWhiteSpace(filePath))
            {
                LogUtility.LogDebug("SnapshotDialogHelper.PromptSnapshotAsync: user cancelled the save dialog.");
                return null;
            }

            LogUtility.LogDebug($"SnapshotDialogHelper.PromptSnapshotAsync: user selected snapshot path '{filePath}'");

            // Resume progress UI and give the UI a small breathing space
            await ShowProgressAsync(progressWindow);
            await Task.Delay(100, ct);

            return filePath;

        }

        // -------------------------
        // WPF progress helpers
        // -------------------------

        private static Task HideProgressAsync(GLWaitWindow win)
        {
            if (win == null) return Task.CompletedTask;

            // Ensure this runs on the window's Dispatcher (UI thread)
            return win.Dispatcher.InvokeAsync(
                () => win.Visibility = System.Windows.Visibility.Hidden,
                DispatcherPriority.Normal
            ).Task;
        }

        private static Task ShowProgressAsync(GLWaitWindow win)
        {
            if (win == null) return Task.CompletedTask;

            return win.Dispatcher.InvokeAsync(
                () => win.Visibility = System.Windows.Visibility.Visible,
                DispatcherPriority.Normal
            ).Task;
        }

        // -------------------------
        // Initial directory logic
        // -------------------------

        private static string ResolveInitialDirectory(Excel.Workbook workbook)
        {
            string defaultFallbackDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (workbook == null) return defaultFallbackDir;

            try
            {
                var workbookDir = Path.GetDirectoryName(workbook.FullName);
                if (string.IsNullOrWhiteSpace(workbookDir)) return defaultFallbackDir;

                if (!Directory.Exists(workbookDir))
                {
                    LogUtility.LogWarn($"Workbook directory does not exist: {workbookDir}");
                    return defaultFallbackDir;
                }

                // Test read access explicitly (handles permission issues)
                try
                {
                    // Minimal read test
                    var _ = Directory.EnumerateFileSystemEntries(workbookDir).Take(1).ToList();
                    return workbookDir;
                }
                catch (UnauthorizedAccessException uaEx)
                {
                    LogUtility.LogWarn($"No permission to access workbook directory: {uaEx.Message}");
                    return defaultFallbackDir;
                }
                catch (Exception exInner)
                {
                    LogUtility.LogWarn($"Error accessing workbook directory: {exInner.Message}");
                    return defaultFallbackDir;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"Exception getting workbook directory: {ex.Message}");
                return defaultFallbackDir;
            }
        }

        private static string BuildDialogTitle(Excel.Application excelApp)
        {
            try
            {
                var name = excelApp?.ActiveWorkbook?.Name ?? "Workbook";
                return $"Snapshot of {name}";
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SnapshotDialogHelper.BuildDialogTitle");
                return "Snapshot";
            }
        }

        // -------------------------
        // SaveFileDialog on STA thread
        // -------------------------

        private static Task<(bool Success, string FilePath)> ShowSaveFileDialogOnStaAsync(
            string initialDirectory,
            string title,
            CancellationToken ct)
        {
            // If we're already on an STA thread (Excel add-in typically is), show directly
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                return Task.FromResult(ShowDialogOnce(initialDirectory, title));
            }

            // Otherwise, spin up a dedicated STA thread and marshal the dialog there
            var tcs = new TaskCompletionSource<(bool Success, string FilePath)>();

            var thread = new Thread(() =>
            {
                try
                {
                    var result = ShowDialogOnce(initialDirectory, title);
                    tcs.TrySetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            })
            {
                IsBackground = true
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            // Honor cancellation (optional)
            if (ct.CanBeCanceled)
            {
                ct.Register(() =>
                {
                    // If the dialog is open, we cannot forcibly close it from here.
                    // We can just mark as cancelled; caller should handle.
                    tcs.TrySetCanceled(ct);
                });
            }

            return tcs.Task;
        }

        private static (bool Success, string FilePath) ShowDialogOnce(string initialDirectory, string title)
        {
            using var dlg = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = title,
                InitialDirectory = SafeInitialDir(initialDirectory),
                OverwritePrompt = true,
                AddExtension = true,
                RestoreDirectory = true
            };
            var res = dlg.ShowDialog();
            if (res == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.FileName))
                return (true, dlg.FileName);

            return (false, string.Empty);
        }

        private static string SafeInitialDir(string initialDirectory)
        {
            if (string.IsNullOrWhiteSpace(initialDirectory))
                return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            try
            {
                return Directory.Exists(initialDirectory)
                    ? initialDirectory
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"SnapshotDialogHelper.SafeInitialDir: '{initialDirectory}'");
                return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
        }

    }
}
