using GLSense;
using GLSense.Utilities;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Helpers
{
    public class ExcelWindowHelper
    {
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, ref RECT rectangle);

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

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }


        /// <summary>
        /// Attempts to bring the Excel main window to the foreground, restoring it if minimized.
        /// </summary>
        public static void ActivateExcelMainWindow(Excel.Application excelApp = null)
        {
            try
            {
                excelApp ??= GLSense.AppState.Instance.ExcelApp;
                if (excelApp == null)
                {
                    LogUtility.LogDebug("ExcelWindowHelper.ActivateExcelMainWindow: no Excel application available.");
                    return;
                }

                IntPtr excelHandle = new IntPtr(excelApp.Hwnd);
                if (excelHandle == IntPtr.Zero)
                {
                    LogUtility.LogDebug("ExcelWindowHelper.ActivateExcelMainWindow: Excel window handle is zero.");
                    return;
                }

                if (IsIconic(excelHandle))
                {
                    LogUtility.LogDebug("ExcelWindowHelper.ActivateExcelMainWindow: Excel window minimized, restoring.");
                    ShowWindow(excelHandle, SW_RESTORE);
                }

                ForceSetForegroundWindow(excelHandle);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ExcelWindowHelper.ActivateExcelMainWindow");
            }
        }

        /// <summary>
        /// Reliably brings <paramref name="hWnd"/> to the foreground. A plain
        /// SetForegroundWindow call is not enough on its own: Windows silently refuses it
        /// (it just flashes the taskbar icon instead) whenever some other process currently
        /// holds the "foreground activation" right - attaching this thread's input queue to
        /// whichever thread currently owns the foreground window temporarily grants that
        /// right. Generic (not Excel-specific) despite living on this class - reused by
        /// ExcelRefEditControl's Window-hosted reactivation path after Excel's native
        /// cell-picker InputBox closes.
        /// </summary>
        public static void ForceSetForegroundWindow(IntPtr hWnd)
        {
            if (SetForegroundWindow(hWnd))
                return;

            LogUtility.LogDebug("ExcelWindowHelper.ForceSetForegroundWindow: plain SetForegroundWindow did not take effect, retrying via AttachThreadInput.");

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