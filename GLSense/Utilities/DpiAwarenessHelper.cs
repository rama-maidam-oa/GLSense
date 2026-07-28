using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace GLSense.Utilities
{
    public static class DpiAwarenessHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
        private const int MDT_EFFECTIVE_DPI = 0;


        public static IDisposable SetPerMonitorAware()
        {
            IntPtr oldContext = IntPtr.Zero;

            try
            {
                oldContext = SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
            }
            catch
            {
                // Fallback for older Windows versions
            }

            return new DpiContextDisposer(oldContext);
        }

        public static double GetWindowDpi(Window window)
        {
            try
            {
                var presentationSource = PresentationSource.FromVisual(window);
                if (presentationSource?.CompositionTarget != null)
                {
                    return 96.0 * presentationSource.CompositionTarget.TransformToDevice.M11;
                }
            }
            catch
            {
                // Ignore
            }

            try
            {
                var handle = new System.Windows.Interop.WindowInteropHelper(window).Handle;
                if (handle != IntPtr.Zero)
                {
                    // Try to get monitor DPI first for better accuracy
                    var monitor = MonitorFromWindow(handle, MONITOR_DEFAULTTONEAREST);
                    if (monitor != IntPtr.Zero)
                    {
                        if (GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, out uint dpiX, out uint dpiY) == 0)
                        {
                            return dpiX;
                        }
                    }

                    // Fallback to window DPI
                    return GetDpiForWindow(handle);
                }
            }
            catch
            {
                // Ignore
            }


            return 96.0;
        }

        private class DpiContextDisposer : IDisposable
        {
            private readonly IntPtr _oldContext;
            private bool _disposed = false;

            public DpiContextDisposer(IntPtr oldContext)
            {
                _oldContext = oldContext;
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (_disposed)
                    return;

                if (disposing)
                {
                    // Dispose managed resources here (none in this case)
                }

                // Dispose unmanaged resources
                if (_oldContext != IntPtr.Zero)
                {
                    try
                    {
                        SetThreadDpiAwarenessContext(_oldContext);
                    }
                    catch
                    {
                        // Exception can be ignored - this is cleanup code
                    }
                }

                _disposed = true;
            }

            // Note: No finalizer needed since we have no unmanaged resources to clean up
            // The IntPtr is just a handle we're restoring, not something we own
        }
    }
}