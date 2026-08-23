using GLSense.Addin.Core.Infrastructure;
using MahApps.Metro.IconPacks;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Windows.Media.Effects;
using System.Collections.Generic;

namespace GLSense.Addin.Core.Views
{
    public partial class AppOverlay : UserControl
    {
        public static readonly DependencyProperty BlurAppliedProperty =
            DependencyProperty.RegisterAttached("BlurApplied", typeof(bool), typeof(AppOverlay), new PropertyMetadata(false));

        public static void SetBlurApplied(UIElement element, bool value)
        {
            try { element?.SetValue(BlurAppliedProperty, value); }
            catch (Exception ex) { ServiceLocator.Logger?.LogWarn($"AppOverlay.SetBlurApplied failed: {ex.Message}"); }
        }

        public static bool GetBlurApplied(UIElement element)
        {
            try { return (bool)(element?.GetValue(BlurAppliedProperty) ?? false); }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"AppOverlay.GetBlurApplied failed: {ex.Message}");
                return false;
            }
        }

        private readonly List<(UIElement Element, Effect OriginalEffect, bool OriginalHitTest)> _blurredElements = new();
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
        }

        public bool IsBusyVisible => BusyOverlay.Visibility == Visibility.Visible;
        public bool IsConfirmVisible => ConfirmOverlay.Visibility == Visibility.Visible;
        public bool IsToastVisible => Toast != null && Toast.Visibility == Visibility.Visible && Toast.Opacity > 0.1;

        // dimBackground defaults to true to preserve existing behavior for every pre-existing
        // caller. Pass false for benign, non-error notifications (e.g. "no jobs found") where
        // dimming the whole window behind the toast would misleadingly make a routine
        // "nothing to show" state look like something went wrong (OISR-21811: "window becomes
        // blurred even if there is no error message").
        public void ShowToast(string message, PackIconFontAwesomeKind icon, Brush color, int durationSeconds = 60, bool dimBackground = true)
        {
            if (Toast == null)
            {
                ServiceLocator.Logger?.LogWarn("AppOverlay.ShowToast: Toast control is null, cannot show toast");
                return;
            }

            ServiceLocator.Logger?.LogDebug($"AppOverlay.ShowToast invoked: message={message}, durationSeconds={durationSeconds}, dimBackground={dimBackground}");
            if (dimBackground)
            {
                ShowToastOverlay();
                ApplyBlurToSiblings();
            }

            this.Visibility = Visibility.Visible;
            Toast.Visibility = Visibility.Visible;
            Toast.Opacity = 1;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Panel.SetZIndex(Toast, 10001);

            ToastMessage.Text = message;
            ToastIcon.Kind = icon;
            ToastIcon.Foreground = color;

            Panel.SetZIndex(this, 9999);

            if (this.Resources["ToastSlideIn"] is Storyboard sb)
                sb.Begin(Toast);

            _toastTimer?.Stop();
            _toastTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(durationSeconds) };
            _toastTimer.Tick += OnToastTimerTick;
            _toastTimer.Start();
        }

        private void ShowToastOverlay()
        {
            if (ToastInputBlocker == null)
                return;

            ToastInputBlocker.Visibility = Visibility.Visible;
            ToastInputBlocker.IsHitTestVisible = true;
        }

        private void HideToastOverlay()
        {
            if (ToastInputBlocker != null)
            {
                ToastInputBlocker.Visibility = Visibility.Collapsed;
                ToastInputBlocker.IsHitTestVisible = false;
            }
            RemoveBlurFromSiblings();
        }

        public void ShowSuccess(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleCheckSolid, Brushes.LimeGreen, 60);
        public void ShowError(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleXmarkSolid, Brushes.Red, 60);
        public void ShowWarning(string message) => ShowToast(message, PackIconFontAwesomeKind.TriangleExclamationSolid, Brushes.Orange, 60);
        public void ShowInfo(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleInfoSolid, Brushes.DodgerBlue, 60);

        public Task ShowToastAsync(string message, PackIconFontAwesomeKind icon, Brush color, int durationSeconds = 60, bool dimBackground = true)
        {
            ServiceLocator.Logger?.LogDebug($"AppOverlay.ShowToastAsync invoked: message={message}, durationSeconds={durationSeconds}, dimBackground={dimBackground}");
            _activeToastTcs?.TrySetResult(true);
            var tcs = new TaskCompletionSource<bool>();
            _activeToastTcs = tcs;

            if (dimBackground)
            {
                ShowToastOverlay();
                ApplyBlurToSiblings();
            }

            this.Visibility = Visibility.Visible;
            Toast.Visibility = Visibility.Visible;
            Toast.Opacity = 1;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Panel.SetZIndex(Toast, 10001);

            ToastMessage.Text = message;
            ToastIcon.Kind = icon;
            ToastIcon.Foreground = color;

            Panel.SetZIndex(this, 9999);

            if (this.Resources["ToastSlideIn"] is Storyboard sb)
                sb.Begin(Toast);

            _toastTimer?.Stop();
            _toastTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(durationSeconds) };
            _toastTimer.Tick += (s, e) =>
            {
                _toastTimer.Stop();
                var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.5));
                fade.Completed += (s2, e2) =>
                {
                    Toast.Opacity = 0;
                    HideToastOverlay();
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

        // Lightweight, non-blocking notification for benign/expected states that are NOT
        // errors (e.g. "no jobs found", "no drilldown jobs exist"). Deliberately skips the
        // background dim that ShowToast/ShowToastAsync apply for real errors/warnings -
        // see OISR-21811.
        public Task ShowStatusAsync(string message, int durationSeconds = 6)
            => ShowToastAsync(message, PackIconFontAwesomeKind.CircleInfoSolid, Brushes.DodgerBlue, durationSeconds, dimBackground: false);

        public void DismissToast()
        {
            if (Toast == null || Toast.Visibility != Visibility.Visible)
                return;

            _toastTimer?.Stop();
            _activeToastTcs?.TrySetResult(true);
            _activeToastTcs = null;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Toast.Opacity = 0;
            Toast.Visibility = Visibility.Collapsed;
            HideToastOverlay();
            RemoveBlurFromSiblings();

            if (BusyOverlay.Visibility != Visibility.Visible && ConfirmOverlay.Visibility != Visibility.Visible)
                this.Visibility = Visibility.Collapsed;
        }

        private void OnToastTimerTick(object sender, EventArgs e)
        {
            _toastTimer.Stop();
            HideToast();
        }

        private void BtnCloseToast_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("AppOverlay.BtnCloseToast_Click invoked");
            _toastTimer?.Stop();
            _activeToastTcs?.TrySetResult(true);
            _activeToastTcs = null;
            HideToast();
        }

        private void HideToast()
        {
            if (this.Resources["ToastSlideOut"] is Storyboard slideOut)
            {
                slideOut.Completed += (s, e) =>
                {
                    Toast.Opacity = 0;
                    HideToastOverlay();
                    RemoveBlurFromSiblings();
                    if (BusyOverlay.Visibility != Visibility.Visible && ConfirmOverlay.Visibility != Visibility.Visible)
                        this.Visibility = Visibility.Collapsed;
                };
                slideOut.Begin(Toast);
            }
            else
            {
                var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.5));
                fade.Completed += (s, e) =>
                {
                    Toast.Opacity = 0;
                    HideToastOverlay();
                    RemoveBlurFromSiblings();
                    if (BusyOverlay.Visibility != Visibility.Visible && ConfirmOverlay.Visibility != Visibility.Visible)
                        this.Visibility = Visibility.Collapsed;
                };
                Toast.BeginAnimation(Border.OpacityProperty, fade);
            }
        }

        private void ToastOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            e.Handled = true;
        }

        private void ToastOverlay_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            e.Handled = true;
        }

        private void ApplyBlurToSiblings()
        {
            try
            {
                var panel = Parent as Panel ?? Window.GetWindow(this)?.Content as Panel;
                if (panel == null)
                    return;

                _blurredElements.Clear();
                foreach (UIElement child in panel.Children)
                {
                    if (child == this)
                        continue;

                    _blurredElements.Add((child, child.Effect, child.IsHitTestVisible));
                    child.Effect = new BlurEffect { Radius = 6 };
                    child.IsHitTestVisible = false;
                    SetBlurApplied(child, true);
                }
                ServiceLocator.Logger?.LogDebug($"AppOverlay.ApplyBlurToSiblings: blurred {_blurredElements.Count} sibling(s)");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"AppOverlay.ApplyBlurToSiblings failed: {ex.Message}");
            }
        }

        private void RemoveBlurFromSiblings()
        {
            try
            {
                foreach (var entry in _blurredElements)
                {
                    try
                    {
                        entry.Element.Effect = entry.OriginalEffect;
                        entry.Element.IsHitTestVisible = entry.OriginalHitTest;
                        SetBlurApplied(entry.Element, false);
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogWarn($"AppOverlay.RemoveBlurFromSiblings failed for element: {ex.Message}");
                    }
                }
                _blurredElements.Clear();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"AppOverlay.RemoveBlurFromSiblings failed: {ex.Message}");
            }
        }

        // Busy Overlay
        public async Task ShowBusyasynTask(string message = "Please wait...", Func<Task> cancelAction = null)
        {
            ShowBusyasyn(message, cancelAction);
            await Task.CompletedTask;
        }

        public void ShowBusyasyn(string message = "Please wait...", Func<Task> cancelAction = null)
        {
            ServiceLocator.Logger?.LogDebug($"AppOverlay.ShowBusyasyn invoked: message={message}, hasCancelAction={cancelAction != null}");
            BusyMessage.Text = message ?? "Please wait...";
            this.Visibility = Visibility.Visible;
            this.Opacity = 1.0;
            Panel.SetZIndex(this, 9999);

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
                ServiceLocator.Logger?.LogDebug("AppOverlay.BusyCancelHandler invoked (user clicked Cancel on busy overlay)");
                BtnCancelBusy.IsEnabled = false;
                try
                {
                    if (cancelAction != null)
                        await cancelAction();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "AppOverlay.BusyCancelHandler");
                }
                finally
                {
                    await HideBusyAsync();
                }
            };

            BtnCancelBusy.Click += BusyCancelHandler;
            BtnCancelBusy.Visibility = Visibility.Visible;
            BtnCancelBusy.IsEnabled = true;

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
                ServiceLocator.Logger?.LogException(ex, "Start busy timer");
            }

            if (this.Resources["ShowBusy"] is Storyboard sb)
                sb.Begin(this);
        }

        private async void BtnHideBusy_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("AppOverlay.BtnHideBusy_Click invoked");
            await HideBusyAsync();
        }

        public async Task HideBusyAsync()
        {
            ServiceLocator.Logger?.LogDebug("AppOverlay.HideBusyAsync invoked");
            var tcs = new TaskCompletionSource<bool>();

            await Dispatcher.InvokeAsync(() =>
            {
                if (this.Resources["HideBusy"] is not Storyboard sb)
                {
                    StopBusyTimer();
                    BusyOverlay.Visibility = Visibility.Collapsed;
                    this.Visibility = Visibility.Collapsed;
                    tcs.TrySetResult(true);
                    return;
                }

                if (_hideBusyHandler != null)
                    sb.Completed -= _hideBusyHandler;

                _hideBusyHandler = (s, e) =>
                {
                    StopBusyTimer();
                    BusyOverlay.Visibility = Visibility.Collapsed;

                    if (Math.Abs(Toast.Opacity - 0) < 0.0001 && ConfirmOverlay.Visibility != Visibility.Visible)
                        this.Visibility = Visibility.Collapsed;

                    sb.Completed -= _hideBusyHandler;
                    _hideBusyHandler = null;
                    tcs.TrySetResult(true);
                };

                sb.Completed += _hideBusyHandler;
                BusyOverlay.Visibility = Visibility.Visible;
                BusyOverlay.IsHitTestVisible = false;
                sb.Begin(this, true);
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
                ServiceLocator.Logger?.LogException(ex, "Stop busy timer");
            }
        }

        private static string FormatTimeSpan(TimeSpan span)
        {
            return string.Format("{0:D2}:{1:D2}:{2:D2}", (int)span.TotalHours, span.Minutes, span.Seconds);
        }

        // Confirmation
        public void ShowConfirm(string message, Action yesAction, Action noAction = null, Action cancelAction = null)
        {
            ServiceLocator.Logger?.LogDebug($"AppOverlay.ShowConfirm invoked: message={message}");
            ConfirmText.Text = message;
            EnsureOwnerWindowRoomForConfirm();
            ConfirmOverlay.Visibility = Visibility.Visible;
            this.Visibility = Visibility.Visible;
            Panel.SetZIndex(this, 9999);

            ConfirmPopup.RenderTransform = new ScaleTransform(0.8, 0.8);

            if (YesHandler != null) BtnYes.Click -= YesHandler;
            if (NoHandler != null) BtnNo.Click -= NoHandler;
            if (CancelHandler != null) BtnCancel.Click -= CancelHandler;

            YesHandler = (s, e) => { ServiceLocator.Logger?.LogDebug("AppOverlay.ShowConfirm: user clicked Yes"); HideConfirm(); yesAction?.Invoke(); };
            NoHandler = (s, e) => { ServiceLocator.Logger?.LogDebug("AppOverlay.ShowConfirm: user clicked No"); HideConfirm(); noAction?.Invoke(); };
            CancelHandler = (s, e) => { ServiceLocator.Logger?.LogDebug("AppOverlay.ShowConfirm: user clicked Cancel"); HideConfirm(); cancelAction?.Invoke(); };

            BtnYes.Click += YesHandler;
            BtnNo.Click += NoHandler;
            BtnCancel.Click += CancelHandler;

            if (this.Resources["ShowConfirm"] is Storyboard sb)
                sb.Begin(this);
        }

        public Task<bool?> ShowConfirmAsync(string message)
        {
            ServiceLocator.Logger?.LogDebug($"AppOverlay.ShowConfirmAsync invoked: message={message}");
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

            YesHandler = (s, e) => { ServiceLocator.Logger?.LogDebug("AppOverlay.ShowConfirmAsync: user clicked Yes"); HideConfirm(); tcs.TrySetResult(true); };
            NoHandler = (s, e) => { ServiceLocator.Logger?.LogDebug("AppOverlay.ShowConfirmAsync: user clicked No"); HideConfirm(); tcs.TrySetResult(false); };
            CancelHandler = (s, e) => { ServiceLocator.Logger?.LogDebug("AppOverlay.ShowConfirmAsync: user clicked Cancel"); HideConfirm(); tcs.TrySetResult(null); };

            BtnYes.Click += YesHandler;
            BtnNo.Click += NoHandler;
            BtnCancel.Click += CancelHandler;

            if (this.Resources["ShowConfirm"] is Storyboard sb)
                sb.Begin(this);

            return tcs.Task;
        }

        private void HideConfirm()
        {
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
            try
            {
                var wnd = Window.GetWindow(this);
                if (wnd == null) return;

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
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"EnsureOwnerWindowRoomForConfirm failed: {ex.Message}");
            }
        }

        private void RestoreOwnerWindowSize()
        {
            if (!_ownerSizeAdjusted || _confirmOwnerWindow == null)
                return;

            try
            {
                _confirmOwnerWindow.Height = _ownerOriginalHeight;
                _confirmOwnerWindow.MinHeight = _ownerOriginalMinHeight;
                _confirmOwnerWindow.MaxHeight = _ownerOriginalMaxHeight;
                _confirmOwnerWindow.SizeToContent = _ownerOriginalSizeToContent;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"RestoreOwnerWindowSize failed: {ex.Message}");
            }
            finally
            {
                _ownerSizeAdjusted = false;
                _confirmOwnerWindow = null;
            }
        }
    }
}