// GLWaitWindow.xaml.cs in GLSense.Addin.Core
// Ported from GLSense\Views\GLWaitWindow.xaml.cs (FinalWorkingCode), including the
// SizeToContent="Height" + Auto-rows height fix already made there.
//
// Adjustments made when porting into this project's architecture:
//   - Base class is now BaseWindow (a plain System.Windows.Window - WPF-UI's FluentWindow
//     base class was removed from this project) instead of DpiAwareWindow. BaseWindow
//     already sets the Excel owner automatically (ServiceLocator.ExcelHandle
//     + ModalToExcel), so there is no separate ShowWithOwner()/SetExcelOwner() call here.
//   - LogUtility.* (static) -> ServiceLocator.Logger.*.
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> a dedicated title-bar drag
//     handler (TitleBar_MouseLeftButtonDown), matching the pattern already used by
//     GLLogin.xaml/.cs in this project.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLWaitWindow.xaml
    /// </summary>
    public partial class GLWaitWindow : BaseWindow, IDisposable
    {
        // ---- Fields kept for minimal functionality ----
        private readonly CancellationHelper _helper;
        private readonly Stopwatch _stopwatch;
        private readonly DispatcherTimer _timer;

        // Close gate & idempotent close
        private volatile bool _allowClose = false; // only allow closing when you call RequestClose()
        private volatile bool _isClosing = false; // prevent re-entry/double-close
        private readonly object _closeLock = new object();

        private bool _disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeManagedResources();

                _helper?.Dispose();

                // Clean up event handlers
                BtnCancel.Click -= BtnCancel_Click;
                this.Closing -= OnClosingGate;
                this.Closed -= OnClosedCleanup;
            }

            _disposed = true;
        }
        private void DisposeManagedResources()
        {
            CleanupWindow();
            StopTimerAndStopwatch();
        }
        private void CleanupWindow()
        {
            try
            {
                if (IsLoaded) RequestClose();
            }
            catch (Exception ex)
            {
                // Ignore window cleanup errors, but still log for root-cause analysis
                ServiceLocator.Logger?.LogException(ex, "GLWaitWindow.CleanupWindow");
            }
        }
        private void StopTimerAndStopwatch()
        {
            try
            {
                _timer?.Stop();
                _stopwatch?.Reset();
            }
            catch (Exception ex)
            {
                // Ignore timer errors, but still log for root-cause analysis
                ServiceLocator.Logger?.LogException(ex, "GLWaitWindow.StopTimerAndStopwatch");
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public GLWaitWindow(CancellationHelper helper = null)
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLWaitWindow constructor invoked");

            EnableEscapeToClose = false;

            _helper = helper ?? new CancellationHelper();

            // Wire up Cancel
            BtnCancel.Click += BtnCancel_Click;

            // Elapsed time display
            _stopwatch = new Stopwatch();
            _timer = new DispatcherTimer(DispatcherPriority.Normal)
            {
                Interval = TimeSpan.FromSeconds(0.5)
            };
            _timer.Tick += (_, __) =>
            {
                ElapsedTimeLabel.Text = $"Time Elapsed: {_stopwatch.Elapsed:hh\\:mm\\:ss}";
            };

            // Prevent user/system close unless allowed
            this.Closing += OnClosingGate;

            // Cleanup
            this.Closed += OnClosedCleanup;

            BtnCancel.IsEnabled = true;
        }

        // ---------- Lifecycle & UI ----------

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "TitleBar_MouseLeftButtonDown error");
            }
        }

        private void OnClosingGate(object sender, CancelEventArgs e)
        {
            // Block close unless gate is open
            if (!_allowClose)
            {
                e.Cancel = true;
                this.Activate();
            }
        }

        private void OnClosedCleanup(object sender, EventArgs e)
        {
            lock (_closeLock)
            {
                if (_isClosing) return;
                _isClosing = true;
            }

            try
            {
                _timer?.Stop();
                _stopwatch?.Reset();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLWaitWindow.OnClosedCleanup");
            }
            finally
            {
                try { _helper?.Dispose(); }
                catch (Exception ex) { ServiceLocator.Logger?.LogException(ex, "GLWaitWindow.OnClosedCleanup: helper dispose"); }
            }
        }

        /// <summary>
        /// Start progress monitoring (marquee + timer + initial labels).
        /// </summary>
        public void StartMonitoring()
        {
            VerifyAccess(); // Ensure we're on UI thread
            ServiceLocator.Logger?.LogDebug("GLWaitWindow.StartMonitoring invoked");

            _stopwatch.Restart();
            ElapsedTimeLabel.Text = "Time Elapsed: 00:00:00";
            _timer.Start();

            ProgressBarControl.IsIndeterminate = true;
            BtnCancel.IsEnabled = true;
        }

        /// <summary>Update the title text.</summary>
        public void SetProcessTitle(string title)
        {
            VerifyAccess();

            if (!string.IsNullOrWhiteSpace(title))
            {
                ServiceLocator.Logger?.LogDebug($"GLWaitWindow.SetProcessTitle: {title}");
                // Set the visible title-bar text directly (txtTitle.Text is not
                // data-bound - see the note in GLWaitWindow.xaml).
                txtTitle.Text = title;
                // Also keep the real Window.Title DP in sync (harmless; not shown
                // anywhere since ShowInTaskbar="False", but keeps things consistent
                // for anything that inspects Title/WindowCaption later).
                WindowCaption = title;
            }
        }

        /// <summary>Update the status/message text.</summary>
        public void SetProcessMessage(string message)
        {
            VerifyAccess();
            ServiceLocator.Logger?.LogDebug($"GLWaitWindow.SetProcessMessage: {message}");

            ProcessNameLabel.Text = message;
        }

        /// <summary>
        /// Allow the window to close now and close safely (idempotent).
        /// </summary>
        public void RequestClose()
        {
            ServiceLocator.Logger?.LogDebug("GLWaitWindow.RequestClose invoked");
            lock (_closeLock)
            {
                if (_isClosing) return;
                _allowClose = true;
            }

            // Use Dispatcher.BeginInvoke to ensure we're on UI thread
            if (Dispatcher.CheckAccess())
            {
                SafeClose();
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(SafeClose), DispatcherPriority.Background);
            }
        }
        private void SafeClose()
        {
            lock (_closeLock)
            {
                if (_isClosing) return;
                _isClosing = true;
            }

            try
            {
                Close();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error closing wait window: {ex.Message}");
            }
        }

        // ---------- Cancel flow (uses confirm overlay only) ----------

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLWaitWindow.BtnCancel_Click invoked");
            BtnCancel.IsEnabled = false;

            try
            {
                _helper.Cancel();  // cancels the ribbon action's Task.Run + APIs
                ProcessNameLabel.Text = "Cancelling...";
                RequestClose();
            }
            catch (Exception ex)
            {
                // Log and re-enable cancel
                ServiceLocator.Logger?.LogException(ex, "GLWaitWindow.BtnCancel_Click");
                BtnCancel.IsEnabled = true;
            }
        }

        /// <summary>
        /// Shows the confirm overlay on the window's dispatcher and returns user choice.
        /// true = Yes, false = No, null = Cancel/dismissed.
        /// </summary>
        public async Task<bool?> ShowConfirmToastAsync(string message)
        {
            ServiceLocator.Logger?.LogDebug($"GLWaitWindow.ShowConfirmToastAsync invoked: {message}");
            // Run entirely on the UI thread and unwrap inner Task<bool?>
            return await Dispatcher
                .InvokeAsync(() => AppOverlayControl.ShowConfirmAsync(message), DispatcherPriority.Normal)
                .Task.Unwrap()
                .ConfigureAwait(false);
        }
    }
}
