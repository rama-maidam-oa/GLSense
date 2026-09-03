// WindowChromeHelper.cs in GLSense\Views
// Shared by GLReloadSourcePicker/GLReleaseHistoryBrowser - both are plain host-side WPF
// windows (see CLAUDE.md section 40.4: they can't inherit GLSense.Addin.Core.Views.BaseWindow,
// since they must keep working even when Addin.Core isn't loaded). Pure WPF has no
// Window property to hide the title-bar minimize/maximize buttons while keeping a normal
// title bar and Close button - ResizeMode="NoResize" only disables (grays out) them, it
// doesn't remove them. The standard, well-documented workaround is stripping the
// WS_MINIMIZEBOX/WS_MAXIMIZEBOX bits from the native window style once the HWND exists.
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace GLSense
{
    internal static class WindowChromeHelper
    {
        private const int GWL_STYLE = -16;
        private const long WS_MINIMIZEBOX = 0x00020000L;
        private const long WS_MAXIMIZEBOX = 0x00010000L;

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
        {
            return IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : new IntPtr(GetWindowLong32(hWnd, nIndex));
        }

        private static void SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 8) SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            else SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32());
        }

        /// <summary>Call from a Window's SourceInitialized handler (HWND must already exist).
        /// Removes the minimize/maximize buttons from the native title bar while leaving the
        /// title, icon, Close button, and (if enabled) the resize border untouched.</summary>
        public static void RemoveMinimizeMaximizeButtons(Window window)
        {
            try
            {
                IntPtr hwnd = new WindowInteropHelper(window).Handle;
                if (hwnd == IntPtr.Zero) return;

                long style = GetWindowLongPtr(hwnd, GWL_STYLE).ToInt64();
                style &= ~(WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
                SetWindowLongPtr(hwnd, GWL_STYLE, new IntPtr(style));
            }
            catch
            {
                // Purely cosmetic - never let a chrome tweak block the window from opening.
            }
        }
    }
}
