// ExcelWindowHelper.cs in GLSense.Addin.Core
// Port of GLSense\Helpers\ExcelWindowPositioning.cs (FinalWorkingCode) lines ~38-59
// (ActivateExcelMainWindow only - the rest of that file's DPI/positioning helpers are
// not needed by Group D and are not ported here).
// Re-pointed vs. the original: AppState.Instance.ExcelApp default -> ServiceLocator.ExcelApp.
// The IsIconic/ShowWindow/SetForegroundWindow p/invoke trio is declared locally rather
// than reusing Utilities.CommonFunctions.NativeMethods, which only declares
// SetForegroundWindow (not IsIconic/ShowWindow) and is a private nested class scoped to
// that file's own usage - not meant to be a shared p/invoke surface.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Helpers
{
    public static class ExcelWindowHelper
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private const int SW_RESTORE = 9;

        /// <summary>
        /// Attempts to bring the Excel main window to the foreground, restoring it if minimized.
        /// </summary>
        public static void ActivateExcelMainWindow(Excel.Application excelApp = null)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("ExcelWindowHelper.ActivateExcelMainWindow: attempting to restore/foreground Excel's main window");

                excelApp ??= ServiceLocator.ExcelApp;
                if (excelApp == null)
                {
                    ServiceLocator.Logger?.LogWarn("ExcelWindowHelper.ActivateExcelMainWindow: no Excel Application available (neither passed in nor ServiceLocator.ExcelApp).");
                    return;
                }

                IntPtr excelHandle = new IntPtr(excelApp.Hwnd);
                if (excelHandle == IntPtr.Zero)
                {
                    ServiceLocator.Logger?.LogWarn("ExcelWindowHelper.ActivateExcelMainWindow: Excel Hwnd is zero.");
                    return;
                }

                if (IsIconic(excelHandle))
                {
                    ServiceLocator.Logger?.LogDebug("ExcelWindowHelper.ActivateExcelMainWindow: Excel window is minimized, restoring.");
                    ShowWindow(excelHandle, SW_RESTORE);
                }

                ForceSetForegroundWindow(excelHandle);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ActivateExcelMainWindow error");
            }
        }

        /// <summary>
        /// Reliably brings <paramref name="hWnd"/> to the foreground.
        /// A plain SetForegroundWindow call is not enough here: Windows silently
        /// refuses it (it just flashes the taskbar icon instead) whenever some other
        /// process currently holds the "foreground activation" right - which is exactly
        /// what happens if the user was interacting with another app (e.g. reading the
        /// debug log in Notepad) while a long-running GLSense operation (balance
        /// refresh, snapshot, drilldown, etc.) was still in progress. This was reported
        /// as: after a Balance Refresh finished, focus stayed on Notepad instead of
        /// returning to Excel.
        /// The standard workaround is to temporarily attach this thread's input queue
        /// to whichever thread currently owns the foreground window - that grants us
        /// the right to steal foreground focus - call SetForegroundWindow, then detach
        /// again immediately.
        /// </summary>
        private static void ForceSetForegroundWindow(IntPtr hWnd)
        {
            if (SetForegroundWindow(hWnd))
                return;

            ServiceLocator.Logger?.LogDebug("ExcelWindowHelper.ForceSetForegroundWindow: plain SetForegroundWindow did not take effect, retrying via AttachThreadInput.");

            IntPtr foregroundWindow = GetForegroundWindow();
            uint foregroundThreadId = GetWindowThreadProcessId(foregroundWindow, IntPtr.Zero);
            uint currentThreadId = GetCurrentThreadId();

            if (foregroundThreadId == 0 || foregroundThreadId == currentThreadId)
            {
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
                return;
            }

            bool attached = AttachThreadInput(currentThreadId, foregroundThreadId, true);
            try
            {
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
            }
            finally
            {
                if (attached)
                    AttachThreadInput(currentThreadId, foregroundThreadId, false);
            }
        }
    }
}
