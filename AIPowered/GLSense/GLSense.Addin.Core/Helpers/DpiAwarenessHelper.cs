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

        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);

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
