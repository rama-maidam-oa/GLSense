// GLReloadProgressWindow.xaml.cs in GLSense\Views
//
// Host-side busy indicator shown for the duration of AddinModule.ReloadAddinCore()'s
// blocking AppDomain unload/resolve/load work - previously the only feedback during that
// entire operation was a WinForms wait cursor, with no visible indication of what was
// happening or how long it might take, and nothing stopping a stray click/Alt+F4/Escape.
//
// Plain host-side WPF window, same reasoning as GLReloadSourcePicker/GLReleaseHistoryBrowser
// (see CLAUDE.md section 40.4) - can't inherit GLSense.Addin.Core.Views.BaseWindow, since it
// must keep working even while Addin.Core is between instances (that's the whole reason it
// exists). The caller must set Owner via WindowInteropHelper before calling ShowAndRender(),
// same pattern as those two windows.
//
// Deliberately NOT closeable while the reload is in progress: Closing is always cancelled
// unless AllowCloseAndClose() (called by the caller once ReloadAddinCore actually finishes,
// success or failure) has set _allowClose first - covers the title bar's X button, Alt+F4,
// and Escape all in one place, since all three ultimately raise the same Closing event.
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace GLSense
{
    public partial class GLReloadProgressWindow : Window
    {
        private bool _allowClose;

        public GLReloadProgressWindow(string message, string currentVersion, string currentReleaseDate)
        {
            InitializeComponent();
            SetStatus(message, 0);
            SetCurrentRelease(currentVersion, currentReleaseDate);

            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
            Closing += GLReloadProgressWindow_Closing;
        }

        private void GLReloadProgressWindow_Closing(object sender, CancelEventArgs e)
        {
            if (!_allowClose) e.Cancel = true;
        }

        /// <summary>
        /// Show()'s WPF equivalent of BaseWindow.PumpDispatcherFrame() (see CLAUDE.md
        /// section 1.4d) - Show() alone returns immediately without painting anything, and
        /// the caller is about to run genuinely blocking, non-yielding work (AppDomain
        /// unload/load) on this same thread right afterward, with no further opportunity for
        /// this window to render until that work finishes. Forcing two dispatcher passes here
        /// (Render, then Background) guarantees the window's first frame is actually on
        /// screen before the freeze starts, instead of a blank/unpainted window appearing and
        /// disappearing around a block it never got to render into.
        /// </summary>
        public void ShowAndRender()
        {
            Show();
            Dispatcher.Invoke(DispatcherPriority.Render, new Action(() => { }));
            Dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { }));
        }

        public void SetStatus(string message, int completedSteps)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => SetStatus(message, completedSteps));
                return;
            }

            TxtMessage.Text = string.IsNullOrWhiteSpace(message)
                ? "Reloading GLSense Add-in..."
                : message;

            int clampedSteps = Math.Max(0, Math.Min(4, completedSteps));
            ReloadProgress.Value = clampedSteps;

            SetStepState(TxtStep1, clampedSteps >= 1);
            SetStepState(TxtStep2, clampedSteps >= 2);
            SetStepState(TxtStep3, clampedSteps >= 3);
            SetStepState(TxtStep4, clampedSteps >= 4);

            if (IsVisible)
            {
                Dispatcher.Invoke(DispatcherPriority.Render, new Action(() => { }));
            }
        }

        public void ShowStage(string message, int completedSteps, int minimumDisplayMilliseconds)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => ShowStage(message, completedSteps, minimumDisplayMilliseconds));
                return;
            }

            SetStatus(message, completedSteps);
            if (!IsVisible || minimumDisplayMilliseconds <= 0) return;

            var frame = new DispatcherFrame();
            var timer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher)
            {
                Interval = TimeSpan.FromMilliseconds(minimumDisplayMilliseconds)
            };

            timer.Tick += (sender, args) =>
            {
                timer.Stop();
                frame.Continue = false;
            };
            timer.Start();
            Dispatcher.PushFrame(frame);
        }

        public void SetRelease(string version, string releaseDate)
        {
            SetReleaseDetails(TxtTargetVersion, TxtTargetReleaseDate, version, releaseDate);
        }

        private void SetCurrentRelease(string version, string releaseDate)
        {
            SetReleaseDetails(TxtCurrentVersion, TxtCurrentReleaseDate, version, releaseDate);
        }

        private void SetReleaseDetails(TextBlock versionText, TextBlock releaseDateText, string version, string releaseDate)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => SetReleaseDetails(versionText, releaseDateText, version, releaseDate));
                return;
            }

            versionText.Text = string.IsNullOrWhiteSpace(version) ? "Not available" : version;
            releaseDateText.Text = FormatReleaseDate(releaseDate);
        }

        public void AllowCloseAndClose()
        {
            _allowClose = true;
            Close();
        }

        private static void SetStepState(TextBlock step, bool completed)
        {
            step.Foreground = completed
                ? new SolidColorBrush(Color.FromRgb(46, 134, 171))
                : new SolidColorBrush(Color.FromRgb(107, 120, 131));
            step.FontWeight = completed ? FontWeights.SemiBold : FontWeights.Normal;
        }

        private static string FormatReleaseDate(string releaseDate)
        {
            if (DateTime.TryParse(releaseDate, out DateTime parsedReleaseDate))
            {
                return parsedReleaseDate.ToString("dd MMM yyyy, hh:mm tt");
            }

            return string.IsNullOrWhiteSpace(releaseDate) ? "Not available" : releaseDate;
        }
    }
}
