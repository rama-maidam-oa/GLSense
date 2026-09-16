using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace GLSense.Utilities
{
    // Fixes the reported "every window shows a completely blank frame for a moment
    // when first opened" bug: the first time WPF ever shows a Window/DataGrid/custom
    // control of a given type in this process, it pays a one-time cost to parse XAML,
    // apply styles/ControlTemplates, and JIT-compile the generated code behind them -
    // and the native HWND becomes visible (Show()/ShowDialog() returns control to
    // Windows) before that first frame is actually composited, so the user sees a
    // blank/white rectangle until it catches up. This is a well-known WPF cold-start
    // effect, not specific to any one window - it hit every DpiAwareWindow-derived
    // window in the reported video, since they all share the same DataGrid/
    // ExcelRefEditControl/AppOverlay controls this class exists to pre-warm.
    //
    // Runs once, off-screen (zero opacity, positioned far outside any monitor,
    // ShowActivated=false so it never steals focus from Excel), dispatched at
    // ApplicationIdle priority right after ribbon load finishes - so this entire cost
    // gets paid silently in the background before the user ever opens a real window,
    // instead of being visible on whichever window they happen to open first.
    public static class WpfWarmup
    {
        private static bool _warmupStarted;
        private static readonly object _lock = new object();

        public static void WarmUpInBackground()
        {
            if (_warmupStarted)
                return;

            lock (_lock)
            {
                if (_warmupStarted)
                    return;
                _warmupStarted = true;
            }

            var app = Application.Current;
            if (app == null)
            {
                LogUtility.LogWarn("WpfWarmup.WarmUpInBackground: no Application.Current yet, skipping.");
                return;
            }

            try
            {
                app.Dispatcher.BeginInvoke(new Action(RunWarmup), DispatcherPriority.ApplicationIdle);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WpfWarmup.WarmUpInBackground");
            }
        }

        private static void RunWarmup()
        {
            try
            {
                LogUtility.LogDebug("WpfWarmup: starting off-screen warm-up of DataGrid/ExcelRefEditControl/AppOverlay.");

                var grid = new DataGrid
                {
                    Width = 200,
                    Height = 100,
                    ItemsSource = new[]
                    {
                        new { A = "warm", B = "up" },
                        new { A = "warm", B = "up" }
                    }
                };

                var refEdit = new GLSense.Views.ExcelRefEditControl();
                var overlay = new GLSense.Views.AppOverlay();

                var panel = new StackPanel();
                panel.Children.Add(grid);
                panel.Children.Add(refEdit);
                panel.Children.Add(overlay);

                var warmupWindow = new Window
                {
                    Content = panel,
                    Width = 50,
                    Height = 50,
                    Left = -32000,
                    Top = -32000,
                    Opacity = 0,
                    ShowActivated = false,
                    ShowInTaskbar = false,
                    WindowStyle = WindowStyle.None,
                    ResizeMode = ResizeMode.NoResize
                };

                warmupWindow.ContentRendered += (s, e) =>
                {
                    try
                    {
                        warmupWindow.Close();
                        LogUtility.LogDebug("WpfWarmup: warm-up window closed - JIT/style/template cost already paid ahead of the first real window open.");
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "WpfWarmup: closing warm-up window");
                    }
                };

                warmupWindow.Show();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WpfWarmup.RunWarmup");
            }
        }
    }
}
