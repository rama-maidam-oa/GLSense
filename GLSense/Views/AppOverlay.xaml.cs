using MahApps.Metro.IconPacks;
using GLSense.Utilities;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Windows.Media.Effects;
using System.Collections.Generic;

namespace GLSense.Views
{
    public partial class AppOverlay : UserControl
    {
        // Attached property to mark elements blurred by this overlay (across instances)
        public static readonly DependencyProperty BlurAppliedProperty =
            DependencyProperty.RegisterAttached("BlurApplied", typeof(bool), typeof(AppOverlay), new PropertyMetadata(false));

        public static void SetBlurApplied(UIElement element, bool value)
        {
            try
            {
                element?.SetValue(BlurAppliedProperty, value);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"SetBlurApplied failed: {ex.Message}");
            }
        }

        public static bool GetBlurApplied(UIElement element)
        {
            try
            {
                return (bool)(element?.GetValue(BlurAppliedProperty) ?? false);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"GetBlurApplied failed: {ex.Message}");
                return false;
            }
        }
        // Keep track of elements that were blurred so we can restore them
        private readonly System.Collections.Generic.List<(UIElement Element, Effect OriginalEffect, bool OriginalHitTest)> _blurredElements = new();

        private DispatcherTimer _busyTimer;
        private DateTime? _busyStart;
        private DispatcherTimer _toastTimer;
        private TaskCompletionSource<bool> _activeToastTcs;
        private RoutedEventHandler YesHandler, NoHandler, CancelHandler, BusyCancelHandler;
        private EventHandler _hideBusyHandler;

        // 2026-07-15 fix: confirm popup clipping (see ConfirmPopup's XAML comment).
        // The owning window (e.g. GLWaitWindow, SizeToContent="Height" + a small
        // MaxHeight) never grows to accommodate a longer confirm message, since
        // ConfirmPopup's own MaxHeight is only a fraction of that window's *current*,
        // usually-small height. Temporarily switching the owner to a fixed size sized to
        // what THIS message actually needs - and restoring it afterward - fixes that
        // without touching every window that happens to host a confirm, and without
        // over-growing the window for short messages (which left dead space behind before
        // this file also got a Height="*" spacer row - see GLWaitWindow.xaml).
        private Window _confirmOwnerWindow;
        private double _ownerOriginalHeight;
        private double _ownerOriginalMinHeight;
        private double _ownerOriginalMaxHeight;
        private SizeToContent _ownerOriginalSizeToContent;
        private bool _ownerSizeAdjusted;

        public AppOverlay()
        {
            InitializeComponent();
            // Not logged: fires once per AppOverlay instance (one per window that hosts
            // one), carries no data, and adds nothing the action-level logs below
            // (ShowConfirm, ShowBusyasyn, ShowToast, etc.) don't already say better.
        }

        public bool IsBusyVisible => BusyOverlay.Visibility == Visibility.Visible;
        public bool IsConfirmVisible => ConfirmOverlay.Visibility == Visibility.Visible;

        // === Toast ===
        // blurBackground/blockInput default to true to preserve existing behavior for every
        // pre-existing caller. Set both false for benign, non-error notifications (e.g. "no
        // jobs found") where blurring the whole window behind the toast would misleadingly
        // make a routine "nothing to show" state look like something went wrong
        // (OISR-21811: "window becomes blurred even if there is no error message").
        public void ShowToast(string message, PackIconFontAwesomeKind icon, Brush color, int durationSeconds = 60, bool blurBackground = true, bool blockInput = true)
        {
            LogUtility.LogDebug($"AppOverlay.ShowToast invoked - message={message}, durationSeconds={durationSeconds}, blurBackground={blurBackground}, blockInput={blockInput}");
            if (Toast == null)
            {
                LogUtility.LogDebug("AppOverlay.ShowToast: Toast control is null, aborting");
                return;
            }

            this.Visibility = Visibility.Visible;
            Toast.Visibility = Visibility.Visible;
            Toast.Opacity = 1;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Panel.SetZIndex(Toast, 10001);

            // Block input to underlying UI while toast is visible
            if (blockInput && ToastInputBlocker != null)
            {
                ToastInputBlocker.Visibility = Visibility.Visible;
                try
                {
                    ToastInputBlocker.Focus();
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"Could not focus ToastInputBlocker: {ex.Message}");
                }
            }

            // Apply blur to underlying sibling elements (mirror/blur effect)
            if (blurBackground)
                ApplyBlurToSiblings();

            ToastMessage.Text = message;
            ToastIcon.Kind = icon;
            ToastIcon.Foreground = color;

            Panel.SetZIndex(this, 9999);

            _toastTimer?.Stop();
            _toastTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(durationSeconds) };
            _toastTimer.Tick += OnToastTimerTick;
            _toastTimer.Start();
        }

        public void ShowSuccess(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleCheckSolid, Brushes.LimeGreen, 60);
        public void ShowError(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleXmarkSolid, Brushes.Red, 60);
        public void ShowWarning(string message) => ShowToast(message, PackIconFontAwesomeKind.TriangleExclamationSolid, Brushes.Orange, 60);
        public void ShowInfo(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleInfoSolid, Brushes.DodgerBlue, 60);

        public Task ShowToastAsync(string message, PackIconFontAwesomeKind icon, Brush color, int durationSeconds = 60, bool blurBackground = true, bool blockInput = true)
        {
            LogUtility.LogDebug($"AppOverlay.ShowToastAsync invoked - message={message}, durationSeconds={durationSeconds}, blurBackground={blurBackground}, blockInput={blockInput}");
            _activeToastTcs?.TrySetResult(true);
            var tcs = new TaskCompletionSource<bool>();
            _activeToastTcs = tcs;

            this.Visibility = Visibility.Visible;
            Toast.Visibility = Visibility.Visible;
            Toast.Opacity = 1;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Panel.SetZIndex(Toast, 10001);

            // Block input to underlying UI while toast is visible
            if (blockInput && ToastInputBlocker != null)
            {
                ToastInputBlocker.Visibility = Visibility.Visible;
                try
                {
                    ToastInputBlocker.Focus();
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"Could not focus ToastInputBlocker: {ex.Message}");
                }
            }

            // Apply blur to underlying sibling elements (mirror/blur effect)
            if (blurBackground)
                ApplyBlurToSiblings();

            ToastMessage.Text = message;
            ToastIcon.Kind = icon;
            ToastIcon.Foreground = color;

            Panel.SetZIndex(this, 9999);

            _toastTimer?.Stop();
            _toastTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(durationSeconds) };
            _toastTimer.Tick += (s, e) =>
            {
                _toastTimer.Stop();
                var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.5));
                fade.Completed += (s2, e2) =>
                {
                    Toast.Opacity = 0;
                    // Hide the input blocker when toast fades out
                    if (ToastInputBlocker != null)
                        ToastInputBlocker.Visibility = Visibility.Collapsed;

                    // Remove blur from siblings
                    RemoveBlurFromSiblings();
                    if (BusyOverlay.Visibility != Visibility.Visible && ConfirmOverlay.Visibility != Visibility.Visible)
                        this.Visibility = Visibility.Collapsed;
                    tcs.TrySetResult(true);
                    if (_activeToastTcs == tcs)
                        _activeToastTcs = null;
                };
                Toast.BeginAnimation(Border.OpacityProperty, fade);
            };
            _toastTimer.Start();

            return tcs.Task;
        }

        // Lightweight, non-blocking notification for benign/expected states that are NOT
        // errors (e.g. "no jobs found", "no drilldown jobs exist"). Deliberately skips the
        // background blur and input-blocking that ShowToast/ShowToastAsync apply for real
        // errors/warnings - see OISR-21811.
        public Task ShowStatusAsync(string message, int durationSeconds = 6)
            => ShowToastAsync(message, PackIconFontAwesomeKind.CircleInfoSolid, Brushes.DodgerBlue, durationSeconds, blurBackground: false, blockInput: false);

        public async Task ShowSuccessAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.CircleCheckSolid, Brushes.LimeGreen, 60);
        }

        public async Task ShowErrorAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.CircleXmarkSolid, Brushes.Red, 60);
        }

        public async Task ShowWarningAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.TriangleExclamationSolid, Brushes.Orange, 60);
        }

        public async Task ShowInfoAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.CircleInfoSolid, Brushes.DodgerBlue, 60);
        }

        public void DismissToast()
        {
            LogUtility.LogDebug("AppOverlay.DismissToast invoked");
            if (Toast == null || Toast.Visibility != Visibility.Visible)
            {
                LogUtility.LogDebug("AppOverlay.DismissToast: Toast is null or not visible, nothing to dismiss");
                return;
            }

            _toastTimer?.Stop();
            _activeToastTcs?.TrySetResult(true);
            _activeToastTcs = null;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Toast.Opacity = 0;
            Toast.Visibility = Visibility.Collapsed;

            if (ToastInputBlocker != null)
                ToastInputBlocker.Visibility = Visibility.Collapsed;

            // Remove blur from siblings
            RemoveBlurFromSiblings();

            if (BusyOverlay.Visibility != Visibility.Visible && ConfirmOverlay.Visibility != Visibility.Visible)
                this.Visibility = Visibility.Collapsed;
        }
        private void OnToastTimerTick(object sender, EventArgs e)
        {
            LogUtility.LogDebug("AppOverlay.OnToastTimerTick invoked - toast duration elapsed");
            _toastTimer.Stop();
            // Immediately dismiss toast and remove blur when timer elapses
            DismissToast();
        }

        private void BtnCloseToast_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("AppOverlay.BtnCloseToast_Click invoked");
            // Immediately dismiss toast and remove blur when user clicks close
            _toastTimer?.Stop();
            _activeToastTcs?.TrySetResult(true);
            _activeToastTcs = null;
            DismissToast();
        }

        // === Busy Overlay ===
        public async Task ShowBusyasynTask(string message = "Please wait...", Func<Task> cancelAction = null)
        {
            ShowBusyasyn(message, cancelAction);
            await Task.CompletedTask;
        }
        public void ShowBusyasyn(string message = "Please wait...", Func<Task> cancelAction = null)
        {
            LogUtility.LogDebug($"AppOverlay.ShowBusyasyn invoked - message={message}, cancelActionProvided={cancelAction != null}");
            BusyMessage.Text = message ?? "Please wait...";
            this.Visibility = Visibility.Visible;
            this.Opacity = 1.0;
            Panel.SetZIndex(this, 9999);

            // 🔥 Bring to front of parent container
            if (this.Parent is UIElement parent)
                Panel.SetZIndex(parent, 0);

            BusyOverlay.Visibility = Visibility.Visible;
            BusyOverlay.Opacity = 1;
            BusyOverlay.IsHitTestVisible = true;

            if (BusyCancelHandler != null)
            {
                BtnCancelBusy.Click -= BusyCancelHandler;
                BusyCancelHandler = null;
            }

            BusyCancelHandler = async (sender, e) =>
            {
                LogUtility.LogDebug("AppOverlay.ShowBusyasyn: busy cancel button clicked");
                BtnCancelBusy.IsEnabled = false;
                try
                {
                    // 🟢 Run async cancel if provided
                    if (cancelAction != null)
                        await cancelAction();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Busy cancel handler");
                }
                finally
                {
                    await HideBusyAsync();
                }
            };

            BtnCancelBusy.Click += BusyCancelHandler;
            BtnCancelBusy.Visibility = Visibility.Visible;
            BtnCancelBusy.IsEnabled = true;

            // Start elapsed timer
            try
            {
                _busyStart = DateTime.UtcNow;
                BusyElapsed.Text = "Time Elapsed: 00:00:00";

                _busyTimer?.Stop();
                _busyTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
                _busyTimer.Tick += (s, e) =>
                {
                    if (_busyStart == null) return;
                    var span = DateTime.UtcNow - _busyStart.Value;
                    BusyElapsed.Text = $"Time Elapsed: {FormatTimeSpan(span)}";
                };
                _busyTimer.Start();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Start busy timer");
            }

            if (this.Resources["ShowBusy"] is Storyboard sb)
                sb.Begin(this);
        }
        private async void BtnHideBusy_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("AppOverlay.BtnHideBusy_Click invoked");
            await HideBusyAsync();
        }
        public async Task HideBusyAsync()
        {
            LogUtility.LogDebug("AppOverlay.HideBusyAsync invoked");
            var tcs = new TaskCompletionSource<bool>();

            await Dispatcher.InvokeAsync(() =>
            {
                if (this.Resources["HideBusy"] is not Storyboard sb)
                {
                    // Fallback: if storyboard missing, hide immediately
                    LogUtility.LogDebug("AppOverlay.HideBusyAsync: HideBusy storyboard missing, hiding immediately");
                    StopBusyTimer();
                    BusyOverlay.Visibility = Visibility.Collapsed;
                    this.Visibility = Visibility.Collapsed;
                    tcs.TrySetResult(true);
                    return;
                }

                // Remove previous handler to prevent multiple triggers
                if (_hideBusyHandler != null)
                    sb.Completed -= _hideBusyHandler;

                _hideBusyHandler = (s, e) =>
                {
                    StopBusyTimer();

                    BusyOverlay.Visibility = Visibility.Collapsed;

                    // Hide root only if no other overlay is visible
                    if (Math.Abs(Toast.Opacity - 0) < 0.0001 && ConfirmOverlay.Visibility != Visibility.Visible)
                        this.Visibility = Visibility.Collapsed;

                    sb.Completed -= _hideBusyHandler; // Clean up handler
                    _hideBusyHandler = null;

                    tcs.TrySetResult(true);
                };

                sb.Completed += _hideBusyHandler;

                // Ensure it stays visible during fade
                BusyOverlay.Visibility = Visibility.Visible;
                BusyOverlay.IsHitTestVisible = false; // allow clicks to pass during fade
                sb.Begin(this, true); // true = controllable animation
            });

            await tcs.Task;
        }

        private void StopBusyTimer()
        {
            try
            {
                _busyTimer?.Stop();
                _busyTimer = null;
                _busyStart = null;
                if (BusyElapsed != null)
                    BusyElapsed.Text = string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Stop busy timer");
            }
        }

        private static string FormatTimeSpan(TimeSpan span)
        {
            // Format as HH:MM:SS
            return string.Format("{0:D2}:{1:D2}:{2:D2}", (int)span.TotalHours, span.Minutes, span.Seconds);
        }


        // === Confirmation ===
        public void ShowConfirm(string message, Action yesAction, Action noAction = null, Action cancelAction = null)
        {
            LogUtility.LogDebug($"AppOverlay.ShowConfirm invoked - message={message}");
            ConfirmText.Text = message;
            EnsureOwnerWindowRoomForConfirm();
            ConfirmOverlay.Visibility = Visibility.Visible;
            this.Visibility = Visibility.Visible;
            Panel.SetZIndex(this, 9999);

            ConfirmPopup.RenderTransform = new ScaleTransform(0.8, 0.8);

            if (YesHandler != null) BtnYes.Click -= YesHandler;
            if (NoHandler != null) BtnNo.Click -= NoHandler;
            if (CancelHandler != null) BtnCancel.Click -= CancelHandler;

            YesHandler = (s, e) => { LogUtility.LogDebug("AppOverlay.ShowConfirm: user clicked Yes"); HideConfirm(); yesAction?.Invoke(); };
            NoHandler = (s, e) => { LogUtility.LogDebug("AppOverlay.ShowConfirm: user clicked No"); HideConfirm(); noAction?.Invoke(); };
            CancelHandler = (s, e) => { LogUtility.LogDebug("AppOverlay.ShowConfirm: user clicked Cancel"); HideConfirm(); cancelAction?.Invoke(); };

            BtnYes.Click += YesHandler;
            BtnNo.Click += NoHandler;
            BtnCancel.Click += CancelHandler;

            if (this.Resources["ShowConfirm"] is Storyboard sb)
                sb.Begin(this);
        }

        public Task<bool?> ShowConfirmAsync(string message)
        {
            LogUtility.LogDebug($"AppOverlay.ShowConfirmAsync invoked - message={message}");
            var tcs = new TaskCompletionSource<bool?>();

            ConfirmText.Text = message;
            EnsureOwnerWindowRoomForConfirm();
            ConfirmOverlay.Visibility = Visibility.Visible;
            this.Visibility = Visibility.Visible;
            Panel.SetZIndex(this, 9999);

            ConfirmPopup.RenderTransform = new ScaleTransform(0.8, 0.8);

            if (YesHandler != null) BtnYes.Click -= YesHandler;
            if (NoHandler != null) BtnNo.Click -= NoHandler;
            if (CancelHandler != null) BtnCancel.Click -= CancelHandler;

            YesHandler = (s, e) => { LogUtility.LogDebug("AppOverlay.ShowConfirmAsync: user clicked Yes"); HideConfirm(); tcs.TrySetResult(true); };
            NoHandler = (s, e) => { LogUtility.LogDebug("AppOverlay.ShowConfirmAsync: user clicked No"); HideConfirm(); tcs.TrySetResult(false); };
            CancelHandler = (s, e) => { LogUtility.LogDebug("AppOverlay.ShowConfirmAsync: user clicked Cancel"); HideConfirm(); tcs.TrySetResult(null); };

            BtnYes.Click += YesHandler;
            BtnNo.Click += NoHandler;
            BtnCancel.Click += CancelHandler;

            if (this.Resources["ShowConfirm"] is Storyboard sb)
                sb.Begin(this);

            return tcs.Task;
        }

        private void HideConfirm()
        {
            LogUtility.LogDebug("AppOverlay.HideConfirm invoked");
            if (this.Resources["HideConfirm"] is Storyboard sb)
            {
                EventHandler onComplete = null;
                onComplete = (s, e) =>
                {
                    ConfirmOverlay.Visibility = Visibility.Collapsed;
                    if (BusyOverlay.Visibility != Visibility.Visible && Math.Abs(Toast.Opacity - 0) < 0.0001)
                        this.Visibility = Visibility.Collapsed;

                    sb.Completed -= onComplete;
                };

                sb.Completed += onComplete;
                sb.Begin(this);
            }
            else
            {
                // Fallback when storyboard is missing
                ConfirmOverlay.Visibility = Visibility.Collapsed;
                if (BusyOverlay.Visibility != Visibility.Visible && Math.Abs(Toast.Opacity - 0) < 0.0001)
                    this.Visibility = Visibility.Collapsed;
            }

            RestoreOwnerWindowSize();
        }

        // 2026-07-15 fix: see field comments above and the ConfirmPopup XAML comment.
        // Temporarily gives the owning window enough height for THIS confirm popup's
        // message + Yes/No/Cancel row to render fully, instead of being squeezed by
        // whatever small size that window currently happens to be (e.g. GLWaitWindow's
        // SizeToContent="Height" + MaxHeight="350" while just showing a progress bar) -
        // and without growing it any more than the message actually needs. Call this
        // AFTER ConfirmText.Text is set, so the measurement below reflects the real
        // message.
        private void EnsureOwnerWindowRoomForConfirm()
        {
            LogUtility.LogDebug("AppOverlay.EnsureOwnerWindowRoomForConfirm invoked");
            try
            {
                var wnd = Window.GetWindow(this);
                if (wnd == null)
                {
                    LogUtility.LogDebug("AppOverlay.EnsureOwnerWindowRoomForConfirm: no owner window found, skipping resize");
                    return;
                }

                _confirmOwnerWindow = wnd;
                _ownerOriginalHeight = wnd.Height;
                _ownerOriginalMinHeight = wnd.MinHeight;
                _ownerOriginalMaxHeight = wnd.MaxHeight;
                _ownerOriginalSizeToContent = wnd.SizeToContent;
                _ownerSizeAdjusted = true;

                // Measure how tall the message text actually wants to be for this
                // confirm, instead of guessing a flat number that leaves dead space for
                // short messages. Measuring ConfirmText directly (rather than the
                // ConfirmPopup Border) matters: ConfirmPopup's own MaxHeight is bound to
                // 0.8x RootOverlay's *current* (still-small, pre-resize) ActualHeight, so
                // measuring the Border itself would just measure into that stale cap
                // instead of the text's real desired size.
                double popupMaxWidth = Math.Max(280, wnd.ActualWidth > 0 ? wnd.ActualWidth * 0.95 : 320);
                double innerTextWidth = Math.Max(100, popupMaxWidth - 40); // minus Border Padding="20" each side
                ConfirmText.Measure(new Size(innerTextWidth, double.PositiveInfinity));
                double textHeight = ConfirmText.DesiredSize.Height;

                // Known fixed chrome inside ConfirmPopup: Border Padding="20" top+bottom,
                // ScrollViewer Margin="0,0,0,20" under the text, and the Yes/No/Cancel
                // button row (Height="32" + Margin="5" top+bottom).
                const double borderPadding = 40;
                const double scrollMargin = 20;
                const double buttonRowHeight = 32 + 10;
                double popupNeeds = textHeight + borderPadding + scrollMargin + buttonRowHeight;

                // ConfirmPopup's own MaxHeight binding caps it at 80% of RootOverlay's
                // height, so ask for enough total height that 80% of it comfortably
                // covers what the popup wants, plus a little chrome above (title bar)
                // and a safety margin.
                const double chromeAboveOverlay = 40;
                const double safetyMargin = 30;
                double neededHeight = chromeAboveOverlay + (popupNeeds / 0.8) + safetyMargin;

                double currentHeight = double.IsNaN(wnd.ActualHeight) || wnd.ActualHeight <= 0
                    ? wnd.Height
                    : wnd.ActualHeight;
                double targetHeight = Math.Max(neededHeight, currentHeight);
                targetHeight = Math.Min(targetHeight, SystemParameters.WorkArea.Height * 0.9);

                // SizeToContent fights any explicit Height we set below, so suspend it
                // while the confirm popup is up and restore it in RestoreOwnerWindowSize.
                wnd.SizeToContent = SizeToContent.Manual;

                // Only MaxHeight can block our explicit Height below (MinHeight is just a
                // floor, and targetHeight is always >= the window's current height, so it
                // never needs raising) - bump it if it's currently more restrictive.
                if (double.IsNaN(wnd.MaxHeight) || wnd.MaxHeight < targetHeight)
                    wnd.MaxHeight = targetHeight;

                wnd.Height = targetHeight;
                LogUtility.LogDebug($"AppOverlay.EnsureOwnerWindowRoomForConfirm: resized owner window to targetHeight={targetHeight}");
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"EnsureOwnerWindowRoomForConfirm failed: {ex.Message}");
            }
        }

        private void RestoreOwnerWindowSize()
        {
            LogUtility.LogDebug("AppOverlay.RestoreOwnerWindowSize invoked");
            if (!_ownerSizeAdjusted || _confirmOwnerWindow == null)
            {
                LogUtility.LogDebug("AppOverlay.RestoreOwnerWindowSize: no adjustment to restore, skipping");
                return;
            }

            try
            {
                _confirmOwnerWindow.Height = _ownerOriginalHeight;
                _confirmOwnerWindow.MinHeight = _ownerOriginalMinHeight;
                _confirmOwnerWindow.MaxHeight = _ownerOriginalMaxHeight;
                _confirmOwnerWindow.SizeToContent = _ownerOriginalSizeToContent;
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"RestoreOwnerWindowSize failed: {ex.Message}");
            }
            finally
            {
                _ownerSizeAdjusted = false;
                _confirmOwnerWindow = null;
            }
        }

        // Add helper methods for blur
        private void ApplyBlurToSiblings()
        {
            LogUtility.LogDebug("AppOverlay.ApplyBlurToSiblings invoked");
            try
            {
                // find parent panel that contains this overlay
                if (this.Parent is Panel parentPanel)
                {
                    _blurredElements.Clear();
                    foreach (UIElement child in parentPanel.Children)
                    {
                        if (child == this) continue;

                        var originalEffect = child.Effect;
                        var originalHit = child.IsHitTestVisible;

                        // store original state
                        _blurredElements.Add((child, originalEffect, originalHit));

                        // apply blur and disable hit testing (overlay + blocker will handle input)
                        child.Effect = new BlurEffect { Radius = 6 };
                        child.IsHitTestVisible = false;
                        SetBlurApplied(child, true);
                    }
                    LogUtility.LogDebug($"AppOverlay.ApplyBlurToSiblings: blurred {_blurredElements.Count} sibling(s) in parent panel");
                    return;
                }

                // fallback: try window content
                var wnd = Window.GetWindow(this);
                if (wnd?.Content is Panel wndPanel)
                {
                    _blurredElements.Clear();
                    foreach (UIElement child in wndPanel.Children)
                    {
                        if (child == this) continue;

                        var originalEffect = child.Effect;
                        var originalHit = child.IsHitTestVisible;
                        _blurredElements.Add((child, originalEffect, originalHit));

                        child.Effect = new BlurEffect { Radius = 6 };
                        child.IsHitTestVisible = false;
                        SetBlurApplied(child, true);
                    }
                    LogUtility.LogDebug($"AppOverlay.ApplyBlurToSiblings: blurred {_blurredElements.Count} sibling(s) in window content (fallback path)");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"ApplyBlurToSiblings failed: {ex.Message}");
            }
        }

        private void RemoveBlurFromSiblings()
        {
            LogUtility.LogDebug("AppOverlay.RemoveBlurFromSiblings invoked");
            try
            {
                if (_blurredElements == null || _blurredElements.Count == 0)
                {
                    LogUtility.LogDebug("AppOverlay.RemoveBlurFromSiblings: no blurred elements tracked, nothing to restore");
                    return;
                }

                foreach (var entry in _blurredElements)
                {
                    try
                    {
                        if (entry.Element == null) continue;
                        entry.Element.Effect = entry.OriginalEffect;
                        entry.Element.IsHitTestVisible = entry.OriginalHitTest;
                        SetBlurApplied(entry.Element, false);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogWarn($"Restore blur failed for element: {ex.Message}");
                    }
                }

                _blurredElements.Clear();
                // Fallback: if other AppOverlay instances applied blur but didn't
                // populate our _blurredElements (different instance), clear any
                // residual blurred elements in all windows that were tagged.
                try
                {
                    if (Application.Current != null)
                    {
                        foreach (Window w in Application.Current.Windows)
                        {
                            try
                            {
                                var root = w?.Content as DependencyObject;
                                if (root == null) continue;
                                ClearBlurFromVisualTree(root);
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogWarn($"Fallback blur clear failed for window: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"Fallback blur clear failed: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"RemoveBlurFromSiblings failed: {ex.Message}");
            }
        }

        private void ClearBlurFromVisualTree(DependencyObject node)
        {
            if (node == null) return;

            int children = System.Windows.Media.VisualTreeHelper.GetChildrenCount(node);
            for (int i = 0; i < children; i++)
            {
                try
                {
                    var child = System.Windows.Media.VisualTreeHelper.GetChild(node, i);
                    if (child is UIElement ui)
                    {
                        try
                        {
                            if (GetBlurApplied(ui) || (ui.Effect is BlurEffect))
                            {
                                ui.Effect = null;
                                ui.IsHitTestVisible = true;
                                SetBlurApplied(ui, false);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogWarn($"ClearBlurFromVisualTree - UIElement: {ex.Message}");
                        }
                    }

                    // Recurse
                    ClearBlurFromVisualTree(child);
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"ClearBlurFromVisualTree - child iteration: {ex.Message}");
                }
            }
        }
    }
}