using GLSense.Addin.Core.Infrastructure;
using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace GLSense.Addin.Core.Utilities
{
    // Fix for OISR: "scrollbar not working properly unless clicked and dragged" - mouse
    // wheel/touchpad scroll gestures did nothing on hover, only an explicit
    // click-and-drag of the scrollbar thumb worked.
    //
    // Root cause: Win32 delivers WM_MOUSEWHEEL to whichever window currently holds
    // keyboard focus, NOT to whichever window is under the mouse cursor. Every normal
    // top-level Window shown via Show()/ShowDialog() is given OS focus automatically the
    // moment it opens, so wheel scrolling "just works" there without any help - but the
    // Balance Configurator's WPF content is its own top-level Window, HWND-reparented as
    // a WS_CHILD into the host's task pane (see ConfiguratorPaneHost/GLConfiguratorPane),
    // and reparenting alone does NOT give it Win32 keyboard focus just because the mouse
    // hovers over it. Only an explicit click (which WPF/Win32 resolve to a focus change
    // as a side effect) ever gave it focus, so wheel gestures were silently routed to
    // whatever else currently had focus (e.g. Excel's own grid/ribbon) and never reached
    // the ScrollViewer at all - explaining exactly why drag-to-scroll (which only needs
    // mouse capture, not focus) worked while the wheel did not.
    //
    // Fix: grab native Win32 focus for this element's own hwnd as soon as the mouse
    // enters it, so any immediately-following wheel notch is already routed correctly.
    // Applied generically (via BaseWindow's constructor) so every window gets this safety
    // net, not just the one pane that was actually broken; it is a no-op for ordinary
    // top-level windows that already own focus (SetFocus on a hwnd that already has focus
    // is harmless).
    public static class MouseWheelFocusHelper
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        public static void EnableHoverToScroll(FrameworkElement element)
        {
            if (element == null) return;

            element.MouseEnter -= OnElementMouseEnter;
            element.MouseEnter += OnElementMouseEnter;
        }

        private static void OnElementMouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                if (!(sender is FrameworkElement element)) return;

                var hwndSource = PresentationSource.FromVisual(element) as HwndSource;
                if (hwndSource == null || hwndSource.IsDisposed) return;

                SetFocus(hwndSource.Handle);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"MouseWheelFocusHelper.OnElementMouseEnter failed: {ex.Message}");
            }
        }
    }
}
