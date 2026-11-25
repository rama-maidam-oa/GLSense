// GLSense.Addin.Core/Helpers/DpiAwarenessHelper.cs
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace GLSense.Addin.Core.Helpers
{
    public static class DpiAwarenessHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MonitorInfo lpmi);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForSystem();

        [StructLayout(LayoutKind.Sequential)]
        private struct MonitorInfo
        {
            public int cbSize;
            public NativeRect rcMonitor;
            public NativeRect rcWork;
            public uint dwFlags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left, Top, Right, Bottom;
        }

        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
        private static int _lastLoggedDpi;

        // OISR-23017: WPF only re-renders a Per-Monitor-V2 window at its monitor's DPI (and on WM_DPICHANGED)
        // when this compatibility switch is off; hosted in Excel its default left every window drawn at the DPI
        // Excel started with. AppContext switches are read once, so this must run before the first WPF window.
        public static void EnableWpfPerMonitorScaling()
        {
            try
            {
                AppContext.SetSwitch("Switch.System.Windows.DoNotScaleForDpiChanges", false);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DpiAwarenessHelper.EnableWpfPerMonitorScaling");
            }
        }

        // OISR-23017: effective DPI of the monitor a window is on, whatever the calling thread's awareness. 0 if unknown.
        public static uint GetMonitorDpi(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero)
                    return 0;

                using (SetPerMonitorAware())
                {
                    var monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
                    if (monitor != IntPtr.Zero && GetDpiForMonitor(monitor, 0, out uint dpiX, out _) == 0)
                        return dpiX;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DpiAwarenessHelper.GetMonitorDpi");
            }

            return 0;
        }

        // OISR-23017: the work area of the window's own monitor in that window's DIPs. SystemParameters.WorkArea
        // is in system-DPI units, wrong after a display-scale change with Excel open. Null if unavailable.
        public static Rect? GetMonitorWorkAreaDip(IntPtr hwnd, double scaleFactor)
        {
            try
            {
                if (hwnd == IntPtr.Zero || scaleFactor <= 0)
                    return null;

                using (SetPerMonitorAware())
                {
                    var monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
                    var info = new MonitorInfo { cbSize = Marshal.SizeOf(typeof(MonitorInfo)) };
                    if (monitor == IntPtr.Zero || !GetMonitorInfo(monitor, ref info))
                        return null;

                    var w = info.rcWork;
                    return new Rect(w.Left / scaleFactor, w.Top / scaleFactor, (w.Right - w.Left) / scaleFactor, (w.Bottom - w.Top) / scaleFactor);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DpiAwarenessHelper.GetMonitorWorkAreaDip");
                return null;
            }
        }

        // OISR-23017: the log header records the scale at start-up only; log each later change once (WARN, not
        // Debug-gated). Baseline is the system DPI, the value the start-up snapshot shows.
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
                    ServiceLocator.Logger?.LogWarn($"Display scale changed: {previous} DPI ({previous * 100 / 96}%) -> {dpi} DPI ({dpi * 100 / 96}%), detected by {source}");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DpiAwarenessHelper.LogIfDpiChanged");
            }
        }

        public static IDisposable SetPerMonitorAware()
        {
            IntPtr oldContext = IntPtr.Zero;

            try
            {
                oldContext = SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
            }
            catch (Exception ex)
            {
                // Fallback for older Windows versions
                ServiceLocator.Logger?.LogDebug($"DpiAwarenessHelper.SetPerMonitorAware: SetThreadDpiAwarenessContext not available (older Windows?) - {ex.Message}");
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
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogDebug($"DpiAwarenessHelper.GetWindowDpi: PresentationSource lookup failed, falling back to GetDpiForWindow - {ex.Message}");
                var handle = new System.Windows.Interop.WindowInteropHelper(window).Handle;
                if (handle != IntPtr.Zero)
                {
                    return GetDpiForWindow(handle);
                }
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
                    catch (Exception ex)
                    {
                        // Exception can be ignored - this is cleanup code
                        ServiceLocator.Logger?.LogWarn($"DpiAwarenessHelper.DpiContextDisposer: failed to restore previous DPI awareness context - {ex.Message}");
                    }
                }

                _disposed = true;
            }

            // Note: No finalizer needed since we have no unmanaged resources to clean up
            // The IntPtr is just a handle we're restoring, not something we own
        }
    }
}
