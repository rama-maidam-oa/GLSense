using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLWaitWindow.xaml
    /// </summary>
    public partial class GLWaitWindow : DpiAwareWindow, IDisposable
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
            LogUtility.LogDebug($"GLWaitWindow.Dispose invoked - disposing={disposing}");
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
                // Ignore window cleanup errors, non-fatal during dispose.
                LogUtility.LogWarn($"GLWaitWindow.CleanupWindow: exception during window cleanup (ignored): {ex.Message}");
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
                // Ignore timer errors, non-fatal during dispose.
                LogUtility.LogWarn($"GLWaitWindow.StopTimerAndStopwatch: exception stopping timer (ignored): {ex.Message}");
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public GLWaitWindow(CancellationHelper helper = null)
        {
            LogUtility.LogDebug("GLWaitWindow.ctor invoked");
            InitializeComponent();

            EnableEscapeToClose = false;

            _helper = helper ?? new CancellationHelper();

            EnhancedDragDropHelper.EnableWindowDrag(this);


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

        private void OnClosingGate(object sender, CancelEventArgs e)
        {
            // Block close unless gate is open
            if (!_allowClose)
            {
                LogUtility.LogDebug("GLWaitWindow.OnClosingGate: close blocked - gate not open (RequestClose not yet called)");
                e.Cancel = true;
                this.Activate();
            }
        }

        private void OnClosedCleanup(object sender, EventArgs e)
        {
            LogUtility.LogDebug("GLWaitWindow.OnClosedCleanup invoked");
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
                LogUtility.LogWarn($"GLWaitWindow.OnClosedCleanup: exception stopping timer (ignored): {ex.Message}");
            }
            finally
            {
                try { _helper?.Dispose(); }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"GLWaitWindow.OnClosedCleanup: exception disposing helper (ignored): {ex.Message}");
                }
                ExcelWindowHelper.ActivateExcelMainWindow(GLSense.AppState.Instance.ExcelApp);
            }
        }
        /// <summary>
        /// Start progress monitoring (marquee + timer + initial labels).
        /// </summary>
        public void StartMonitoring()
        {
            LogUtility.LogDebug("GLWaitWindow.StartMonitoring invoked");
            VerifyAccess(); // Ensure we're on UI thread

            _stopwatch.Restart();
            ElapsedTimeLabel.Text = "Time Elapsed: 00:00:00";
            _timer.Start();

            ProgressBarControl.IsIndeterminate = true;
            ProgressBarControl.Visibility = Visibility.Visible;
            BtnCancel.IsEnabled = true;
        }

        /// <summary>Update the title text.</summary>
        public void SetProcessTitle(string title)
        {
            VerifyAccess();
            LogUtility.LogDebug($"GLWaitWindow.SetProcessTitle invoked - title={title}");

            if (!string.IsNullOrWhiteSpace(title))
                txtTitle.Text = title;
        }

        /// <summary>Update the status/message text.</summary>
        public void SetProcessMessage(string message)
        {
            VerifyAccess();
            LogUtility.LogDebug($"GLWaitWindow.SetProcessMessage invoked - message={message}");

            ProcessNameLabel.Text = message;
        }
        /// <summary>
        /// Allow the window to close now and close safely (idempotent).
        /// </summary>
        public void RequestClose()
        {
            LogUtility.LogDebug("GLWaitWindow.RequestClose invoked");
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
            LogUtility.LogDebug("GLWaitWindow.SafeClose invoked");
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
                LogUtility.LogError($"Error closing wait window: {ex.Message}");
            }
        }

        // ---------- Cancel flow (uses confirm overlay only) ----------

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLWaitWindow.BtnCancel_Click invoked");
            BtnCancel.IsEnabled = false;

            try
            {
                _helper.Cancel();  // ⚠️ THIS cancels the ribbon's Task.Run + APIs!
                ProcessNameLabel.Text = "Cancelling...";
                RequestClose();
            }
            catch (Exception ex)
            {
                // Log and re-enable cancel
                LogUtility.LogException(ex, "GLWaitWindow.BtnCancel_Click");
                BtnCancel.IsEnabled = true;
            }
        }
        /// <summary>
        /// Shows the confirm overlay on the window's dispatcher and returns user choice.
        /// true = Yes, false = No, null = Cancel/dismissed.
        /// </summary>
        public async Task<bool?> ShowConfirmToastAsync(string message)
        {
            LogUtility.LogDebug($"GLWaitWindow.ShowConfirmToastAsync invoked - message={message}");
            // Run entirely on the UI thread and unwrap inner Task<bool?>
            var result = await Dispatcher
                .InvokeAsync(() => AppOverlayControl.ShowConfirmAsync(message), DispatcherPriority.Normal)
                .Task.Unwrap()
                .ConfigureAwait(false);
            LogUtility.LogDebug($"GLWaitWindow.ShowConfirmToastAsync: user response={result}");
            return result;
        }
    }
}

