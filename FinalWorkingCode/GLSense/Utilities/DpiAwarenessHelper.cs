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

        // OISR-23017: GetDpiForMonitor lives in shcore.dll - declared against user32.dll it threw
        // EntryPointNotFoundException on every call (caught and logged), so it never returned a DPI.
        [DllImport("shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
        private const int MDT_EFFECTIVE_DPI = 0;


        // OISR-23017: WPF only re-renders a Per-Monitor-V2 window at its own monitor's DPI (and on
        // WM_DPICHANGED) when this compatibility switch is off. Hosted in Excel there is no app.config
        // of ours to set it, and its default leaves WPF drawing every window at the DPI Excel started
        // with - so after a display-scale change with Excel open, windows came out too small/too big.
        // AppContext switches are read once and cached, so this must run before WPF reads it, i.e.
        // before the first WPF window/HwndSource is created.
        public static void EnableWpfPerMonitorScaling()
        {
            try
            {
                AppContext.SetSwitch("Switch.System.Windows.DoNotScaleForDpiChanges", false);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwarenessHelper.EnableWpfPerMonitorScaling");
            }
        }

        // OISR-23017: the effective DPI of the monitor a window is on, independent of the calling
        // thread's DPI awareness (GetDpiForMonitor reports the system DPI to a non-per-monitor caller).
        // Returns 0 if it can't be read.
        public static uint GetMonitorDpi(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero)
                    return 0;

                using (SetPerMonitorAware())
                {
                    var monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
                    if (monitor != IntPtr.Zero && GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, out uint dpiX, out _) == 0)
                        return dpiX;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwarenessHelper.GetMonitorDpi");
            }

            return 0;
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForSystem();

        private static int _lastLoggedDpi;

        // OISR-23017: the log's Environment Snapshot records the scale Excel started at; a scale change
        // after that (docking, Settings > Display) was invisible in the log. Logs one WARN line per change
        // (not Debug-gated). Baseline is the system DPI, the same value the snapshot shows.
        public static void LogIfDpiChanged(uint dpi, string source)
        {
            try
            {
                if (dpi == 0)
                    return;

                int previous = System.Threading.Interlocked.Exchange(ref _lastLoggedDpi, (int)dpi);
                if (previous == 0)
                    previous = (int)GetDpiForSystem();

                if (previous != 0 && previous != (int)dpi)
                    LogUtility.LogWarn($"Display scale changed: {previous} DPI ({previous * 100 / 96}%) -> {dpi} DPI ({dpi * 100 / 96}%), detected by {source}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DpiAwarenessHelper.LogIfDpiChanged");
            }
        }

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