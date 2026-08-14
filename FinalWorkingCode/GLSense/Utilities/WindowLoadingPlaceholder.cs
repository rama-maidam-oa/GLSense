using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace GLSense.Utilities
{
    // Targets the confirmed root cause behind the "blank window before loading" flash:
    // showing a *new* real window (GLCubeDetails, GLLOVs, etc.) always pays a genuine
    // per-instance cost - a fresh native HWND, a fresh visual tree, a fresh first layout/
    // paint pass - between the moment it's shown and ContentRendered. This is per-instance,
    // not per-type, so no amount of pre-warming a window *type* (WpfWarmup's generic
    // control warm-up, or the since-removed per-window-type warm-up) can pre-pay it; direct
    // proof from an earlier investigation showed a warmed-up GLLOVs opened for real 40s
    // later still hit the identical gap as a cold process.
    //
    // Every earlier attempt to hide this gap by manipulating the *real* window itself
    // (Opacity=0 until ready, parking it off-screen until ContentRendered) caused a worse,
    // independently-reproduced DWM black-flash regression - twice in this investigation,
    // and once before it in this repo's history (commit cdd1168/caa346c). This takes a
    // different approach entirely: never touch the real window's Opacity/Position/
    // Visibility. Instead, show one small, deliberately trivial, reused loading indicator
    // window near where the real window will appear, and hide it once the real window's
    // own ContentRendered fires. Reusing a single instance (Hide(), never Close()) means
    // its own one-time per-instance cost is paid once, off-screen, at ribbon load - every
    // later ShowNear() just toggles visibility of an HWND that already exists and has
    // already painted, which is fast. This also avoids the ghost Task-View-tile mechanism
    // that hit the old per-window-type warm-up: that came from constructing+Show()ing+
    // Close()ing 16 *different* window types back-to-back in a tight startup loop, not
    // from a single window being shown/hidden sparingly across a session.
    public static class WindowLoadingPlaceholder
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private static readonly object _lock = new object();
        private static Window _window;
        private static int _generation;
        private static DispatcherTimer _safetyTimer;

        public static void WarmUpInBackground()
        {
            var app = Application.Current;
            if (app == null)
            {
                LogUtility.LogWarn("WindowLoadingPlaceholder.WarmUpInBackground: no Application.Current yet, skipping.");
                return;
            }

            try
            {
                app.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        EnsureCreated();
                        LogUtility.LogDebug("WindowLoadingPlaceholder: warmed up (paid its own first-paint cost off-screen).");
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "WindowLoadingPlaceholder.WarmUpInBackground (inner)");
                    }
                }), DispatcherPriority.ApplicationIdle);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WindowLoadingPlaceholder.WarmUpInBackground");
            }
        }

        /// <summary>
        /// Shows the reused loading indicator near excelOwnerHwnd (falls back to the
        /// primary screen's work area if unavailable). Returns a generation token - pass
        /// it to Hide(generation) so a stale dismissal (e.g. from a window that's still
        /// loading after a newer one already took over the placeholder) can't hide a
        /// placeholder that a different, later ShowNear() call is now responsible for.
        /// </summary>
        public static int ShowNear(IntPtr excelOwnerHwnd)
        {
            int myGeneration;
            try
            {
                EnsureCreated();
                PositionNear(excelOwnerHwnd);

                lock (_lock)
                {
                    myGeneration = ++_generation;
                }

                _window.Show();

                _safetyTimer?.Stop();
                _safetyTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(3000) };
                _safetyTimer.Tick += (s, e) => Hide(myGeneration);
                _safetyTimer.Start();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WindowLoadingPlaceholder.ShowNear");
                myGeneration = -1;
            }

            return myGeneration;
        }

        public static void Hide(int generation)
        {
            try
            {
                lock (_lock)
                {
                    if (generation != _generation)
                        return; // superseded by a newer ShowNear() - not ours to hide
                }

                _safetyTimer?.Stop();
                _safetyTimer = null;
                _window?.Hide();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WindowLoadingPlaceholder.Hide");
            }
        }

        private static void EnsureCreated()
        {
            if (_window != null)
                return;

            var ring = new Ellipse
            {
                Width = 48,
                Height = 48,
                Stroke = new SolidColorBrush((Color)ColorConverter.ConvertFromString(AppConstants.GLAccentHex)),
                StrokeThickness = 5,
                StrokeDashArray = new DoubleCollection { 2.4, 1.6 },
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform = new RotateTransform(0)
            };

            var spin = new DoubleAnimation(0, 360, TimeSpan.FromSeconds(1))
            {
                RepeatBehavior = RepeatBehavior.Forever
            };
            ring.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, spin);

            var text = new TextBlock
            {
                Text = "Loading...",
                Margin = new Thickness(0, 16, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
                FontSize = 18,
                Foreground = Brushes.White
            };

            var panel = new StackPanel
            {
                Orientation = Orientation.Vertical,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            panel.Children.Add(ring);
            panel.Children.Add(text);

            var border = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(235, 45, 45, 48)),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(28),
                Child = panel
            };

            _window = new Window
            {
                Content = border,
                // Sized to whatever the ring+text actually need plus the Border's Padding,
                // instead of a hardcoded box that has to be kept in sync by hand if either
                // one ever changes.
                SizeToContent = SizeToContent.WidthAndHeight,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = false,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                ShowActivated = false,
                Topmost = true,
                WindowStartupLocation = WindowStartupLocation.Manual,
                Left = -32000,
                Top = -32000
            };

            // Pay the one-time per-instance HWND-creation/first-paint cost right here,
            // off-screen, instead of on whichever real window first triggers ShowNear().
            _window.Show();
            _window.Hide();
        }

        private static void PositionNear(IntPtr excelOwnerHwnd)
        {
            try
            {
                double centerX, centerY;

                if (excelOwnerHwnd != IntPtr.Zero && GetWindowRect(excelOwnerHwnd, out RECT rect) &&
                    rect.Right > rect.Left && rect.Bottom > rect.Top)
                {
                    // GetWindowRect returns raw screen pixels; Window.Left/Top are
                    // device-independent units (96 DPI), so convert using this monitor's
                    // scale factor - otherwise this lands wrong on anything above 100%.
                    double scale = 1.0;
                    try
                    {
                        uint dpi = GetDpiForWindow(excelOwnerHwnd);
                        if (dpi > 0)
                            scale = dpi / 96.0;
                    }
                    catch
                    {
                        // fall back to 1.0
                    }

                    centerX = ((rect.Left + rect.Right) / 2.0) / scale;
                    centerY = ((rect.Top + rect.Bottom) / 2.0) / scale;
                }
                else
                {
                    var workArea = SystemParameters.WorkArea;
                    centerX = workArea.Left + (workArea.Width / 2.0);
                    centerY = workArea.Top + (workArea.Height / 2.0);
                }

                // Width/Height are NaN under SizeToContent - ActualWidth/ActualHeight hold
                // the real size instead, and are already resolved by the warm-up Show()+
                // Hide() cycle in EnsureCreated() by the time this runs.
                _window.Left = centerX - (_window.ActualWidth / 2.0);
                _window.Top = centerY - (_window.ActualHeight / 2.0);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WindowLoadingPlaceholder.PositionNear");
            }
        }
    }
}
