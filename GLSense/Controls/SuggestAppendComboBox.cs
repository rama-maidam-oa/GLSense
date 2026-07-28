using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using GLSense.Utilities;

namespace GLSense.Controls
{
    [TemplatePart(Name = "PART_TextBox", Type = typeof(TextBox))]
    [TemplatePart(Name = "PART_Popup", Type = typeof(Popup))]
    [TemplatePart(Name = "PART_ListBox", Type = typeof(ListBox))]
    public class SuggestAppendComboBox : Control
    {
        private const string _IsSelectedPropName = "IsSelected";

        private TextBox _textBox;
        private Popup _popup;
        private ListBox _listBox;
        private ICollectionView _view;
        private bool _isInternalUpdate;
        private readonly DispatcherTimer _debounceTimer;
        private string _pendingUserInput = string.Empty;
        private Key _lastKey = Key.None;
        private TextBlock _toolTipTextBlock;

        // ADXTaskPane mouse-wheel fix (see the "Popup mouse-wheel forwarding" region near
        // the bottom of this class for the full explanation): tracks whichever instance's
        // popup is currently open, and caches that instance's internal ScrollViewer once
        // found via the visual tree.
        private static SuggestAppendComboBox _openInstance;
        private ScrollViewer _listBoxScrollViewer;

        // The open popup's on-screen rectangle, as plain primitives - read from the
        // dedicated mouse-hook thread (see below), which must never touch WPF
        // DispatcherObjects directly (thread-affinity would throw). Reference assignment
        // is atomic in .NET, and `volatile` gives the hook thread a guaranteed-fresh read
        // without needing a lock for this simple set-then-read pattern.
        private static volatile ScreenRect _openPopupScreenRect;

        // cached reflection props for performance
        private PropertyInfo _displayProp;
        private PropertyInfo _isSelectedProp;
        private IEnumerable _attachedItemsSource;
        private Button _clearButton;

        public static readonly DependencyProperty ComboClearProperty =
        DependencyProperty.Register(nameof(ComboClear), typeof(bool), typeof(SuggestAppendComboBox),
        new PropertyMetadata(false, OnComboClearChanged));

        public bool ComboClear
        {
            get => (bool)GetValue(ComboClearProperty);
            set => SetValue(ComboClearProperty, value);
        }

        private static void OnComboClearChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is SuggestAppendComboBox combo)
            {
                combo.UpdateClearButtonVisibility();
            }
        }
        private void UpdateClearButtonVisibility()
        {
            if (_clearButton == null) return;

            _clearButton.Visibility = ComboClear ? Visibility.Visible : Visibility.Collapsed;
        }
        public int DebounceIntervalMs
        {
            get => (int)GetValue(DebounceIntervalMsProperty);
            set => SetValue(DebounceIntervalMsProperty, value);
        }
        public static readonly DependencyProperty DebounceIntervalMsProperty =
            DependencyProperty.Register(nameof(DebounceIntervalMs), typeof(int), typeof(SuggestAppendComboBox),
                new PropertyMetadata(50, (d, e) =>
                {
                    if (d is SuggestAppendComboBox c)
                        c._debounceTimer.Interval = TimeSpan.FromMilliseconds((int)e.NewValue);
                }));


        public double ArrowWidth { get => (double)GetValue(ArrowWidthProperty); set => SetValue(ArrowWidthProperty, value); }
        public static readonly DependencyProperty ArrowWidthProperty =
            DependencyProperty.Register(nameof(ArrowWidth), typeof(double), typeof(SuggestAppendComboBox), new PropertyMetadata(10.0));

        public double ArrowHeight { get => (double)GetValue(ArrowHeightProperty); set => SetValue(ArrowHeightProperty, value); }
        public static readonly DependencyProperty ArrowHeightProperty =
            DependencyProperty.Register(nameof(ArrowHeight), typeof(double), typeof(SuggestAppendComboBox), new PropertyMetadata(8.0));

        public System.Windows.Media.Brush ArrowBrush { get => (System.Windows.Media.Brush)GetValue(ArrowBrushProperty); set => SetValue(ArrowBrushProperty, value); }
        public static readonly DependencyProperty ArrowBrushProperty =
             DependencyProperty.Register(nameof(ArrowBrush), typeof(System.Windows.Media.Brush), typeof(SuggestAppendComboBox),
                 new PropertyMetadata(System.Windows.Media.Brushes.SteelBlue));

        // 🟩 MultiSelect property
        public static readonly DependencyProperty IsMultiSelectProperty =
            DependencyProperty.Register(nameof(IsMultiSelect), typeof(bool), typeof(SuggestAppendComboBox),
                new PropertyMetadata(false, (d, e) => ((SuggestAppendComboBox)d).UpdateListBoxTemplate()));

        public bool IsMultiSelect
        {
            get => (bool)GetValue(IsMultiSelectProperty);
            set => SetValue(IsMultiSelectProperty, value);
        }

        private void UpdateListBoxTemplate()
        {
            if (_listBox == null) return;


            if (IsMultiSelect)
            {
                var key = new ComponentResourceKey(typeof(SuggestAppendComboBox), "SuggestMultiSelectItemTemplate");

                if (TryFindResource(key) is not DataTemplate template)
                {
                    var factory = new FrameworkElementFactory(typeof(StackPanel));
                    factory.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);

                    var chk = new FrameworkElementFactory(typeof(CheckBox));
                    chk.SetBinding(CheckBox.IsCheckedProperty, new Binding(_IsSelectedPropName) { Mode = BindingMode.TwoWay });
                    if (!string.IsNullOrEmpty(DisplayMemberPath))
                        chk.SetBinding(CheckBox.ContentProperty, new Binding(DisplayMemberPath));
                    else
                        chk.SetBinding(CheckBox.ContentProperty, new Binding("."));

                    factory.AppendChild(chk);

                    template = new DataTemplate { VisualTree = factory };
                }

                _listBox.ItemTemplate = template;
                _listBox.SelectionMode = SelectionMode.Multiple;
            }
            else
            {
                _listBox.ClearValue(ItemsControl.ItemTemplateProperty);
                _listBox.SelectionMode = SelectionMode.Single;
            }

        }

        public string Text
        {
            get => (string)GetValue(TextProperty);
            set => SetValue(TextProperty, value);
        }
        public static readonly DependencyProperty TextProperty =
            DependencyProperty.Register(nameof(Text), typeof(string), typeof(SuggestAppendComboBox),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnTextPropertyChanged));

        private static void OnTextPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var c = (SuggestAppendComboBox)d;
            c.UpdateTextBoxFromProperty((string)e.NewValue);
            c.UpdateToolTipText();
        }

        private void UpdateTextBoxFromProperty(string newValue)
        {
            if (_textBox == null) return;
            try
            {
                _isInternalUpdate = true;
                _textBox.Text = newValue ?? string.Empty;
                _textBox.Select(_textBox.Text.Length, 0);
            }
            finally { _isInternalUpdate = false; }
        }

        public event Action<object> SelectionCommitted;
        public event Action<string> InvalidSelection;

        static SuggestAppendComboBox()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(SuggestAppendComboBox),
                new FrameworkPropertyMetadata(typeof(SuggestAppendComboBox)));
        }

        public SuggestAppendComboBox()
        {
            _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(50) };
            _debounceTimer.Tick += (s, e) => { _debounceTimer.Stop(); ApplyFilterAndSuggest(_pendingUserInput); };

            SelectedItems = new ObservableCollection<object>();
        }

        #region ItemsSource / DisplayMemberPath / SelectedItem
        public IEnumerable ItemsSource
        {
            get => (IEnumerable)GetValue(ItemsSourceProperty);
            set => SetValue(ItemsSourceProperty, value);
        }
        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register(nameof(ItemsSource), typeof(IEnumerable), typeof(SuggestAppendComboBox),
                new PropertyMetadata(null, (d, e) =>
                {
                    var c = (SuggestAppendComboBox)d;
                    c._view = e.NewValue != null ? CollectionViewSource.GetDefaultView(e.NewValue) : null;
                    // Attach listeners to the underlying items so that external changes
                    // to item properties (e.g. IsSelected) are reflected in the control
                    // even when the popup/listbox has not been opened.
                    try
                    {
                        c.DetachListenersFromItemsSource(e.OldValue as IEnumerable);
                        c.AttachListenersToItemsSource(e.NewValue as IEnumerable);
                        // Update the displayed text from the new items immediately
                        c.UpdateDisplayedTextFromEnumerable(e.NewValue as IEnumerable);
                    }
                    catch (Exception ex)
                    {
                        // Non-fatal: template may not be ready. Log for diagnostics.
                        LogUtility.LogWarn($"SuggestAppendComboBox: ItemsSource change handler exception (non-fatal): {ex.Message}");
                    }
                }));

        public string DisplayMemberPath
        {
            get => (string)GetValue(DisplayMemberPathProperty);
            set => SetValue(DisplayMemberPathProperty, value);
        }
        public static readonly DependencyProperty DisplayMemberPathProperty =
            DependencyProperty.Register(nameof(DisplayMemberPath), typeof(string), typeof(SuggestAppendComboBox),
                new PropertyMetadata(string.Empty));

        public object SelectedItem
        {
            get => GetValue(SelectedItemProperty);
            set => SetValue(SelectedItemProperty, value);
        }
        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register(nameof(SelectedItem), typeof(object), typeof(SuggestAppendComboBox),
                        new PropertyMetadata(null, (d, e) =>
                        {
                            var c = (SuggestAppendComboBox)d;
                            c.UpdateTextFromSelectedItem();
                            c.UpdateToolTipText();
                        }));

        public ObservableCollection<object> SelectedItems
        {
            get => (ObservableCollection<object>)GetValue(SelectedItemsProperty);
            set => SetValue(SelectedItemsProperty, value);
        }

        public static readonly DependencyProperty SelectedItemsProperty =
            DependencyProperty.Register(
                nameof(SelectedItems),
                typeof(ObservableCollection<object>),
                typeof(SuggestAppendComboBox),
                new PropertyMetadata(null));
        #endregion
        private void ApplySelectedItemToListBox()
        {
            if (IsMultiSelect || _listBox == null || SelectedItem == null) return;

            _listBox.SelectedItem = SelectedItem;
            _listBox.ScrollIntoView(SelectedItem);
        }
        private void UpdateTextFromSelectedItem()
        {
            if (IsMultiSelect)
            {
                if (SelectedItem == null && _textBox != null)
                {
                    try
                    {
                        _isInternalUpdate = true;
                        _textBox.Text = string.Empty;
                    }
                    finally
                    {
                        _isInternalUpdate = false;
                    }

                    UpdateToolTipText();
                }

                return;
            }

            if (_textBox == null) return;

            string newText = "";

            if (SelectedItem != null)
            {
                if (!string.IsNullOrWhiteSpace(DisplayMemberPath))
                {
                    var prop = SelectedItem.GetType().GetProperty(DisplayMemberPath);
                    if (prop != null)
                        newText = prop.GetValue(SelectedItem)?.ToString() ?? "";
                    else
                        newText = SelectedItem.ToString();
                }
                else
                {
                    newText = SelectedItem.ToString();
                }
            }

            _isInternalUpdate = true;
            _textBox.Text = newText;
            _isInternalUpdate = false;
        }

        private System.Windows.Point _mouseDownPos;
        private bool _isMouseDragging = false;

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            FindAndStoreTemplateParts();
            SetupClearButton();
            SetupTextBox();
            SetupListBox();
            SetupDropDownButton();
            SetupPopup();
            InitializeCollectionView();
            FindToolTipTextBlock();

            // Make sure any pre-set Text value is pushed into the template parts
            // (important when the control is initially collapsed and templated later).
            UpdateTextBoxFromProperty(Text);
            UpdateToolTipText();

            // ensure ESC closes popup even if focus is elsewhere inside the control
            RemoveHandler(Keyboard.PreviewKeyDownEvent, new KeyEventHandler(OnControlPreviewKeyDown));
            AddHandler(Keyboard.PreviewKeyDownEvent, new KeyEventHandler(OnControlPreviewKeyDown), true);
        }

        private void FindAndStoreTemplateParts()
        {
            _clearButton = GetTemplateChild("PART_ClearButton") as Button;
            _textBox = GetTemplateChild("PART_TextBox") as TextBox;
            _popup = GetTemplateChild("PART_Popup") as Popup;
            _listBox = GetTemplateChild("PART_ListBox") as ListBox;
        }

        private void FindToolTipTextBlock()
        {
            if (_textBox != null && _textBox.ToolTip is ToolTip toolTip && toolTip.Content is StackPanel stackPanel)
            {
                foreach (var child in stackPanel.Children)
                {
                    if (child is TextBlock textBlock && textBlock.Name == "toolTipTextBlock")
                    {
                        _toolTipTextBlock = textBlock;
                        UpdateToolTipText();
                        break;
                    }
                }
            }
        }

        private void SetupClearButton()
        {
            if (_clearButton == null) return;

            UpdateClearButtonVisibility();
            _clearButton.Click += OnClearButtonClick;
        }

        private void SetupTextBox()
        {
            if (_textBox == null) return;

            _textBox.TextChanged -= OnTextChanged;
            _textBox.PreviewKeyDown -= OnPreviewKeyDown;
            _textBox.LostFocus -= OnTextBoxLostFocus;

            _textBox.TextChanged += OnTextChanged;
            _textBox.PreviewKeyDown += OnPreviewKeyDown;
            _textBox.LostFocus += OnTextBoxLostFocus;
        }

        private void SetupListBox()
        {
            if (_listBox == null) return;

            SetupMouseDragScrollDetection();
            SetupSelectionBehavior();
            SetupDisplayMemberPathForSingleSelect();
            _listBox.PreviewKeyDown += OnListBoxPreviewKeyDown;
        }

        private void SetupMouseDragScrollDetection()
        {
            _listBox.PreviewMouseDown += (s, e) =>
            {
                _mouseDownPos = e.GetPosition(_listBox);
                _isMouseDragging = false;
            };

            _listBox.PreviewMouseMove += (s, e) =>
            {
                if (_isMouseDragging) return;

                var currentPos = e.GetPosition(_listBox);
                if (Math.Abs(currentPos.Y - _mouseDownPos.Y) > 4)
                {
                    _isMouseDragging = true;
                }
            };

            _listBox.PreviewMouseUp += (s, e) =>
            {
                if (_isMouseDragging)
                {
                    e.Handled = true;
                    _isMouseDragging = false;
                }
            };
        }

        private void SetupSelectionBehavior()
        {
            _listBox.PreviewMouseUp += (s, e) =>
            {
                if (_isMouseDragging) return;

                if (IsMultiSelect)
                {
                    e.Handled = false;
                }
                else
                {
                    CommitSelectionFromList();
                    e.Handled = true;
                }
            };
        }

        private void SetupDisplayMemberPathForSingleSelect()
        {
            if (!IsMultiSelect)
            {
                _listBox.DisplayMemberPath = DisplayMemberPath;
            }
        }

        private void SetupDropDownButton()
        {
            if (GetTemplateChild("PART_DropDownButton") is not Button dropDownButton)
                return;

            dropDownButton.Click += (s, e) =>
            {
                if (_popup == null) return;

                bool shouldOpen = ShouldOpenPopup();

                if (shouldOpen)
                {
                    OpenPopup();
                }
                else
                {
                    _popup.IsOpen = false;
                    VisualStateManager.GoToState(this, "PopupClosed", true);
                }
            };
        }

        private bool ShouldOpenPopup()
        {
            if (GetTemplateChild("ArrowRotate") is RotateTransform arrowRotate)
            {
                return Math.Abs(arrowRotate.Angle - 0) < 0.001;
            }

            return _popup?.IsOpen != true;
        }

        private void SetupPopup()
        {
            if (_popup == null) return;

            _popup.StaysOpen = false;
            _popup.Opened += (s, e) =>
            {
                // ADXTaskPane mouse-wheel fix: track this instance for every open path
                // (dropdown-button click via OpenPopup(), and the type-to-filter
                // autosuggest path in ApplyFilterAndSuggest, which sets _popup.IsOpen =
                // true directly). Popup itself always raises Opened regardless of which
                // call site flipped IsOpen, so this is the one place that reliably catches
                // all of them. See the "Popup mouse-wheel forwarding" region below for the
                // full mechanism (a dedicated background thread hosting a low-level mouse
                // hook, decoupled from this - the WPF UI - thread).
                _openInstance = this;
                UpdateOpenPopupScreenRect();
                EnsureMouseHookThreadRunning();

                VisualStateManager.GoToState(this, "PopupOpen", true);
                if (GetTemplateChild("ArrowRotate") is RotateTransform arrow)
                {
                    arrow.Angle = 180;
                }
            };

            _popup.Closed += Popup_Closed;
        }

        private void OnControlPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape && _popup is { IsOpen: true })
            {
                _popup.IsOpen = false;
                e.Handled = true;
            }
        }

        #region Popup mouse-wheel forwarding (ADXTaskPane fix)
        //
        // Fix for: mouse wheel not scrolling this control's popup when hosted inside the
        // Balance Configurator's ADXExcelTaskPane (confirmed by the user NOT to reproduce
        // on regular windows like GLSegmentValues, where the popup's own internal
        // ScrollViewer already handles the wheel natively).
        //
        // ATTEMPT 1 (did not work): removing a nested ScrollViewer in Generic.xaml - a real
        // bug, but not this one's cause.
        //
        // ATTEMPT 2 (did not work): a WPF class handler for PreviewMouseWheel registered on
        // every Window. The wheel message never reaches any WPF Window's routed-event
        // tunnel in this hosting context, so this could never have worked.
        //
        // ATTEMPT 3 (fixed scrolling, but caused a new regression): a low-level, system-
        // wide mouse hook (WH_MOUSE_LL) installed on this class's own (shared, busy) WPF
        // Dispatcher thread. WH_MOUSE_LL hooks always invoke their callback ON THE
        // INSTALLING THREAD, and since the hook is inherently global/system-wide (thread-
        // scoped installation isn't supported for the _LL hook types), EVERY mouse-move
        // message on the entire desktop had to round-trip through that one callback on our
        // shared thread before Windows could continue. Whenever that thread was even
        // briefly busy (closing GLSegmentRef, populating GLAccountsRef), mouse/redraw
        // processing backed up system-wide - the reported "window doesn't close until you
        // click elsewhere" / general sluggishness regression.
        //
        // ATTEMPT 4 (did not work either): moved the interception to GLConfiguratorPane's
        // own WndProc (the WinForms ADXExcelTaskPane host), on the theory that
        // WM_MOUSEWHEEL is delivered there natively. Disabling the attempt-3 hook to test
        // this proved wrong: with no hook installed, WndProc never saw WM_MOUSEWHEEL for
        // the popup either - the message genuinely isn't delivered through any normal
        // Win32 routing path in this hosting context; it's ONLY visible via a global
        // low-level hook, which intercepts raw input before Windows even decides which
        // window to route it to.
        //
        // REAL FIX: attempt 3 was the only mechanism that actually saw the message - the
        // defect was WHICH THREAD it ran on, not the hook itself. This installs the exact
        // same WH_MOUSE_LL hook, but on its own small, dedicated, otherwise-idle background
        // thread with its own native message loop (required for any low-level hook),
        // completely decoupled from the busy shared WPF Dispatcher thread. That dedicated
        // thread does nothing except this hook, so Windows always finds it responsive
        // regardless of what the main UI thread is doing - eliminating attempt 3's
        // regression while keeping its (only proven-effective) interception mechanism.
        //
        // Since the hook thread is NOT the WPF Dispatcher thread, it must never touch WPF
        // DispatcherObjects directly (FrameworkElement.PointToScreen,
        // ScrollViewer.ScrollToVerticalOffset, etc. all enforce thread affinity and would
        // throw). So:
        //   1. The open popup's on-screen rectangle is computed on the WPF thread itself,
        //      whenever the popup opens (UpdateOpenPopupScreenRect, called from
        //      Popup.Opened), and cached as a plain, thread-safe-to-read ScreenRect
        //      (primitives only).
        //   2. The hook callback (running on the dedicated thread) does a pure numeric
        //      point-in-rect test against that cached rect - no WPF calls at all.
        //   3. Only on an actual hit does it touch WPF, and even then only by marshaling
        //      onto the WPF thread asynchronously via Dispatcher.BeginInvoke - never
        //      blocking the hook thread waiting for the (possibly busy) UI thread to catch
        //      up.
        private static readonly object _hookLock = new object();
        private static Thread _hookThread;
        private static uint _hookThreadId;
        private static readonly ManualResetEventSlim _hookThreadReady = new ManualResetEventSlim(false);
        private static NativeMethods.LowLevelMouseProc _mouseHookProc;
        private static IntPtr _mouseHookHandle = IntPtr.Zero;

        private static void EnsureMouseHookThreadRunning()
        {
            lock (_hookLock)
            {
                if (_hookThread != null) return;

                _hookThreadReady.Reset();
                _hookThread = new Thread(MouseHookThreadProc)
                {
                    IsBackground = true,
                    Name = "SuggestAppendComboBox.MouseWheelHook"
                };
                _hookThread.Start();
            }

            // Best-effort: the thread sets this within microseconds of starting. If it
            // somehow takes longer than 2s (should never happen), proceed anyway rather
            // than hang the UI thread indefinitely.
            _hookThreadReady.Wait(2000);
        }

        private static void MouseHookThreadProc()
        {
            try
            {
                _hookThreadId = NativeMethods.GetCurrentThreadId();
                _mouseHookProc = LowLevelMouseHookCallback;
                _mouseHookHandle = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL, _mouseHookProc, IntPtr.Zero, 0);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SuggestAppendComboBox.MouseHookThreadProc: hook install failed");
            }
            finally
            {
                _hookThreadReady.Set();
            }

            try
            {
                // Low-level hooks require a running message loop on the installing
                // thread. This thread does nothing else, ever - it stays free/responsive
                // indefinitely no matter how busy the main WPF Dispatcher thread gets.
                while (NativeMethods.GetMessage(out NativeMethods.MSG msg, IntPtr.Zero, 0, 0) > 0)
                {
                    NativeMethods.TranslateMessage(ref msg);
                    NativeMethods.DispatchMessage(ref msg);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SuggestAppendComboBox.MouseHookThreadProc: message loop failed");
            }
        }

        private static IntPtr LowLevelMouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && wParam.ToInt32() == NativeMethods.WM_MOUSEWHEEL)
                {
                    var rect = _openPopupScreenRect; // volatile read - thread-safe, no WPF call
                    var openCombo = _openInstance;
                    if (rect != null && openCombo != null)
                    {
                        var hookStruct = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                        if (rect.Contains(hookStruct.pt.x, hookStruct.pt.y))
                        {
                            short wheelDelta = (short)((hookStruct.mouseData >> 16) & 0xFFFF);

                            // Fire-and-forget onto the WPF thread - this hook thread must
                            // never block waiting for the (possibly busy) UI thread; that
                            // decoupling is the entire point of the dedicated thread.
                            try
                            {
                                openCombo.Dispatcher.BeginInvoke(new Action(() => ScrollOpenPopup(openCombo, wheelDelta)));
                            }
                            catch (Exception dispatchEx)
                            {
                                LogUtility.LogException(dispatchEx, "SuggestAppendComboBox.LowLevelMouseHookCallback: BeginInvoke failed");
                            }

                            return new IntPtr(1); // swallow - already scheduled above
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SuggestAppendComboBox.LowLevelMouseHookCallback");
            }

            return NativeMethods.CallNextHookEx(_mouseHookHandle, nCode, wParam, lParam);
        }

        private static void ScrollOpenPopup(SuggestAppendComboBox combo, short wheelDelta)
        {
            // Runs on the WPF thread (marshaled via Dispatcher.BeginInvoke above) - safe to
            // touch WPF objects here. Re-check everything since this runs asynchronously
            // and the popup may have closed in the meantime.
            if (!ReferenceEquals(_openInstance, combo) || combo._popup is not { IsOpen: true } || combo._listBox == null)
                return;

            var scrollViewer = combo._listBoxScrollViewer ??= FindVisualChild<ScrollViewer>(combo._listBox);
            if (scrollViewer == null) return;

            double newOffset = scrollViewer.VerticalOffset - (wheelDelta / 3.0);
            newOffset = Math.Max(0, Math.Min(newOffset, scrollViewer.ScrollableHeight));
            scrollViewer.ScrollToVerticalOffset(newOffset);
        }

        // Computed on the WPF thread (Popup.Opened) - never on the hook thread.
        private void UpdateOpenPopupScreenRect()
        {
            _openPopupScreenRect = null;

            if (_popup?.Child is not FrameworkElement child)
                return;

            try
            {
                if (!child.IsLoaded)
                    return;

                // Fully-qualified System.Windows.Point: this file also imports
                // System.Drawing (for other, unrelated members), and a bare "Point" would
                // be ambiguous between the two namespaces.
                var topLeft = child.PointToScreen(new System.Windows.Point(0, 0));
                var bottomRight = child.PointToScreen(new System.Windows.Point(child.ActualWidth, child.ActualHeight));
                _openPopupScreenRect = new ScreenRect(topLeft.X, topLeft.Y, bottomRight.X, bottomRight.Y);
            }
            catch (Exception ex)
            {
                // PointToScreen can throw if the popup is mid-teardown; leave the rect
                // null (the hook then simply has nothing to hit-test against).
                LogUtility.LogWarn($"SuggestAppendComboBox.UpdateOpenPopupScreenRect: {ex.Message}");
            }
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;

            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typed) return typed;

                var descendant = FindVisualChild<T>(child);
                if (descendant != null) return descendant;
            }

            return null;
        }

        /// <summary>
        /// Tears down the dedicated hook thread and unhooks WH_MOUSE_LL. This codebase is
        /// a monolith (no hot-reload/AppDomain teardown concept like AIPowered), so this
        /// is mainly relevant if the add-in is ever disabled/unloaded without closing
        /// Excel entirely; harmless to leave running until process exit otherwise.
        /// </summary>
        public static void ShutdownMouseHook()
        {
            lock (_hookLock)
            {
                try
                {
                    if (_mouseHookHandle != IntPtr.Zero)
                    {
                        NativeMethods.UnhookWindowsHookEx(_mouseHookHandle);
                        _mouseHookHandle = IntPtr.Zero;
                    }

                    if (_hookThread != null && _hookThreadId != 0)
                    {
                        NativeMethods.PostThreadMessage(_hookThreadId, NativeMethods.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                        _hookThread.Join(1000);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "SuggestAppendComboBox.ShutdownMouseHook");
                }
                finally
                {
                    _hookThread = null;
                    _mouseHookProc = null;
                    _hookThreadId = 0;
                    _openPopupScreenRect = null;
                }
            }
        }

        /// <summary>Plain-data screen rectangle - deliberately NOT a WPF type, so the
        /// dedicated hook thread can read it without any DispatcherObject thread-affinity
        /// concerns.</summary>
        private sealed class ScreenRect
        {
            private readonly double _left, _top, _right, _bottom;

            public ScreenRect(double left, double top, double right, double bottom)
            {
                _left = left;
                _top = top;
                _right = right;
                _bottom = bottom;
            }

            public bool Contains(int x, int y) => x >= _left && x <= _right && y >= _top && y <= _bottom;
        }

        /// <summary>Win32 interop for the low-level mouse hook above, kept in its own
        /// nested class so the P/Invoke surface is easy to find/remove if this control is
        /// ever revisited.</summary>
        private static class NativeMethods
        {
            public const int WH_MOUSE_LL = 14;
            public const int WM_MOUSEWHEEL = 0x020A;
            public const int WM_QUIT = 0x0012;

            public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

            [StructLayout(LayoutKind.Sequential)]
            public struct POINT
            {
                public int x;
                public int y;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct MSLLHOOKSTRUCT
            {
                public POINT pt;
                public uint mouseData;
                public uint flags;
                public uint time;
                public IntPtr dwExtraInfo;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct MSG
            {
                public IntPtr hwnd;
                public uint message;
                public IntPtr wParam;
                public IntPtr lParam;
                public uint time;
                public POINT pt;
            }

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll")]
            public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll")]
            public static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

            [DllImport("user32.dll")]
            public static extern bool TranslateMessage(ref MSG lpMsg);

            [DllImport("user32.dll")]
            public static extern IntPtr DispatchMessage(ref MSG lpMsg);

            [DllImport("user32.dll")]
            public static extern bool PostThreadMessage(uint idThread, uint msg, IntPtr wParam, IntPtr lParam);

            [DllImport("kernel32.dll")]
            public static extern uint GetCurrentThreadId();
        }
        #endregion

        private void InitializeCollectionView()
        {
            if (ItemsSource != null)
            {
                _view = CollectionViewSource.GetDefaultView(ItemsSource);
            }
        }

        private void UpdateToolTipText()
        {
            if (_toolTipTextBlock == null) return;

            if (string.IsNullOrEmpty(_textBox?.Text))
            {
                _toolTipTextBlock.Text = "No selection";
            }
            else if (IsMultiSelect)
            {
                var items = _textBox.Text.Split(';')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToList();

                if (items.Count == 0)
                {
                    _toolTipTextBlock.Text = "No selection";
                }
                else if (items.Count == 1)
                {
                    _toolTipTextBlock.Text = items[0];
                }
                else
                {
                    _toolTipTextBlock.Text = $"{items.Count} items selected";
                }
            }
            else
            {
                _toolTipTextBlock.Text = _textBox.Text;
            }
        }

        private void OnClearButtonClick(object sender, RoutedEventArgs e)
        {
            _textBox?.Clear();

            SelectedItem = null;
            UpdateToolTipText();

            if (IsMultiSelect && _listBox?.ItemsSource != null)
            {
                foreach (var item in _listBox.ItemsSource.Cast<object>())
                {
                    _isSelectedProp?.SetValue(item, false);
                }
                UpdateDisplayedText();
            }
        }
        private void OpenPopup()
        {
            if (_popup == null) return;

            PrepareView();
            PopulateListBox();
            _textBox?.Focus();
            _popup.IsOpen = true;
            _openInstance = this;
            VisualStateManager.GoToState(this, "PopupOpen", true);

            CacheItemReflectionProperties();
            RestoreMultiSelectionIfNeeded();
            SubscribeToItemPropertyChangesIfMultiSelect();

            if (!IsMultiSelect)
                ApplySelectedItemToListBox();

        }
        private void PrepareView()
        {
            if (_view == null) return;

            _view.Filter = null;
            _view.Refresh();
        }

        private void PopulateListBox()
        {
            if (_listBox == null) return;

            UpdateListBoxTemplate();
            _listBox.ItemsSource = _view?.Cast<object>().ToList();
            AttachItemListeners();
        }

        private void CacheItemReflectionProperties()
        {
            if (_listBox?.ItemsSource is not IEnumerable<object> items)
                return;

            var firstItem = items.FirstOrDefault();
            if (firstItem == null) return;

            var itemType = firstItem.GetType();

            _displayProp = !string.IsNullOrEmpty(DisplayMemberPath)
                ? itemType.GetProperty(DisplayMemberPath)
                : null;

            _isSelectedProp = itemType.GetProperty(_IsSelectedPropName);
        }

        private void RestoreMultiSelectionIfNeeded()
        {
            if (!IsMultiSelect || _listBox?.ItemsSource == null)
                return;

            var selectedTexts = _textBox.Text
                .Split(';', (char)StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var item in _listBox.ItemsSource)
            {
                var displayText = GetItemDisplayText(item);
                bool shouldBeSelected = selectedTexts.Contains(displayText);

                _isSelectedProp?.SetValue(item, shouldBeSelected);
            }
        }

        private string GetItemDisplayText(object item)
        {
            if (item == null) return string.Empty;

            var text = _displayProp?.GetValue(item)?.ToString();
            return string.IsNullOrEmpty(text) ? item.ToString() ?? string.Empty : text;
        }

        private void SubscribeToItemPropertyChangesIfMultiSelect()
        {
            if (!IsMultiSelect || _listBox?.ItemsSource == null)
                return;

            foreach (var item in _listBox.ItemsSource.Cast<object>())
            {
                if (item is INotifyPropertyChanged npc)
                {
                    npc.PropertyChanged -= OnItemPropertyChanged;
                    npc.PropertyChanged += OnItemPropertyChanged;
                }
            }
        }
        private void AttachItemListeners()
        {
            if (_listBox?.ItemsSource == null)
                return;

            foreach (var item in _listBox.ItemsSource)
            {
                if (item is INotifyPropertyChanged npc)
                {
                    npc.PropertyChanged -= OnItemPropertyChanged;
                    npc.PropertyChanged += OnItemPropertyChanged;
                }
            }
        }

        private void AttachListenersToItemsSource(IEnumerable items)
        {
            if (items == null) return;

            _attachedItemsSource = items;

            foreach (var item in items)
            {
                if (item is INotifyPropertyChanged npc)
                {
                    npc.PropertyChanged -= OnItemPropertyChanged;
                    npc.PropertyChanged += OnItemPropertyChanged;
                }
            }
            CacheItemReflectionPropertiesFromEnumerable(items);
        }

        private void DetachListenersFromItemsSource(IEnumerable items)
        {
            if (items == null) return;
                foreach (var item in items)
                {
                    if (item is INotifyPropertyChanged npc)
                    {
                        try { npc.PropertyChanged -= OnItemPropertyChanged; }
                        catch (Exception ex)
                        {
                            // Detach should not throw; log and continue
                            LogUtility.LogWarn($"SuggestAppendComboBox: failed to detach PropertyChanged handler: {ex.Message}");
                        }
                    }
                }
            _attachedItemsSource = null;
        }

        private void CacheItemReflectionPropertiesFromEnumerable(IEnumerable items)
        {
            var first = items.Cast<object>().FirstOrDefault();
            if (first == null) return;
            var itemType = first.GetType();
            _displayProp = !string.IsNullOrEmpty(DisplayMemberPath) ? itemType.GetProperty(DisplayMemberPath) : null;
            _isSelectedProp = itemType.GetProperty(_IsSelectedPropName);
        }

        private void UpdateDisplayedTextFromEnumerable(IEnumerable items)
        {
            if (!IsMultiSelect || items == null || _isSelectedProp == null) return;

            var selectedNames = items.Cast<object>()
                .Where(i => (bool?)_isSelectedProp.GetValue(i) == true)
                .Select(i => _displayProp?.GetValue(i)?.ToString() ?? i.ToString())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            if (_textBox == null) return;

            try
            {
                _isInternalUpdate = true;
                _textBox.Text = string.Join(";", selectedNames);
                _textBox.Select(_textBox.Text.Length, 0);
            }
            finally
            {
                _isInternalUpdate = false;
            }

            UpdateToolTipText();
        }
        private void UpdateDisplayedText()
        {
            if (_listBox?.ItemsSource == null || _textBox == null)
                return;

            var selectedNames = _listBox.ItemsSource
                .Cast<object>()
                .Where(i => (bool?)_isSelectedProp?.GetValue(i) == true)
                .Select(i => _displayProp?.GetValue(i)?.ToString())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            try
            {
                _isInternalUpdate = true;
                _textBox.Text = string.Join(";", selectedNames);
                _textBox.Select(_textBox.Text.Length, 0);
            }
            finally
            {
                _isInternalUpdate = false;
            }

            UpdateToolTipText();
        }
        private void OnItemPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (!string.Equals(e.PropertyName, _IsSelectedPropName, StringComparison.Ordinal))
                return;

            if (_isSelectedProp == null || sender == null)
                return;

            if (!_isSelectedProp.DeclaringType.IsInstanceOfType(sender))
                return;

            Dispatcher.InvokeAsync(() =>
            {
                if (!IsMultiSelect || _listBox?.ItemsSource == null)
                    return;

                UpdateDisplayedText();
            }, DispatcherPriority.Background);

        }

        private void Popup_Closed(object sender, EventArgs e)
        {
            if (ReferenceEquals(_openInstance, this))
            {
                _openInstance = null;
                _openPopupScreenRect = null;
            }

            VisualStateManager.GoToState(this, "PopupClosed", true);
            if (GetTemplateChild("ArrowRotate") is RotateTransform arrowRotate)
            {
                arrowRotate.Angle = 0;
            }

            if (!IsMultiSelect || _listBox?.ItemsSource == null) return;

            var selectedNames = _listBox.ItemsSource
                .Cast<object>()
                .Where(i => (bool?)_isSelectedProp?.GetValue(i) == true)
                .Select(i => _displayProp?.GetValue(i)?.ToString())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            try
            {
                _isInternalUpdate = true;
                _textBox.Text = string.Join(";", selectedNames);
            }
            finally { _isInternalUpdate = false; }

            UpdateToolTipText();

            if (IsMultiSelect && _listBox?.ItemsSource != null)
            {
                foreach (var item in _listBox.ItemsSource.Cast<object>())
                {
                    if (item is INotifyPropertyChanged npc)
                        npc.PropertyChanged -= OnItemPropertyChanged;
                }
            }
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            _lastKey = e.Key;

            if (e.Key is Key.Back or Key.Delete)
            {
                e.Handled = false;
                return;
            }

            switch (e.Key)
            {
                case Key.Enter:
                    HandleEnterKey(e);
                    break;

                case Key.Tab:
                    HandleTabKey();
                    break;

                case Key.Escape:
                    HandleEscapeKey();
                    break;

                case Key.Down:
                case Key.Up:
                    HandleArrowKeys(e);
                    break;

                default:
                    break;
            }
        }

        private void HandleEnterKey(KeyEventArgs e)
        {
            _debounceTimer.Stop();
            CommitSelectionFromTextOrList();
            e.Handled = true;
        }

        private void HandleTabKey()
        {
            _debounceTimer.Stop();

            if (string.IsNullOrWhiteSpace(_textBox?.Text))
            {
                SelectedItem = null;
                return;
            }

            CommitSelectionFromTextOrList();
        }

        private void HandleEscapeKey()
        {
            _debounceTimer.Stop();
            if (_popup is { IsOpen: true })
            {
                _popup.IsOpen = false;
            }
        }

        private void HandleArrowKeys(KeyEventArgs e)
        {
            if (_popup is not { IsOpen: true } || _listBox == null)
                return;

            e.Handled = true;

            _listBox.Focus();

            if (_listBox.Items.Count == 0)
                return;

            if (_listBox.SelectedIndex < 0)
            {
                _listBox.SelectedIndex = e.Key == Key.Down ? 0 : _listBox.Items.Count - 1;
            }

            if (_listBox.ItemContainerGenerator.ContainerFromIndex(_listBox.SelectedIndex) is ListBoxItem item)
            {
                item.Focus();
            }
        }

        private void OnListBoxPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitSelectionFromList();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                _textBox?.Focus();
                _popup.IsOpen = false;
                e.Handled = true;
            }
        }

        private void OnTextBoxLostFocus(object sender, RoutedEventArgs e)
        {
            if (_listBox != null && _listBox.IsKeyboardFocusWithin) return;

            _debounceTimer.Stop();
            ValidateInputImmediate();
        }

        private void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isInternalUpdate) return;

            if (_textBox == null) return;

            SetCurrentValue(TextProperty, _textBox.Text);
            UpdateToolTipText();

            var text = _textBox.Text ?? string.Empty;
            int selStart = _textBox.SelectionStart;
            int selLen = _textBox.SelectionLength;

            string userInput;
            if (selLen > 0 && selStart <= text.Length)
                userInput = text.Substring(0, selStart);
            else
                userInput = text;


            _pendingUserInput = userInput;
            _debounceTimer.Stop();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(DebounceIntervalMs);
            _debounceTimer.Start();

            // If the user cleared the text (via Backspace/Delete) we need to immediately
            // clear multi-select item flags and notify selection change so dependent
            // controls (e.g., refedit) can update their enabled state without waiting
            // for lost-focus.
            if (string.IsNullOrWhiteSpace(text))
            {
                try
                {
                    var items = _attachedItemsSource ?? ItemsSource ?? (_listBox?.ItemsSource as IEnumerable);
                    if (IsMultiSelect && items != null)
                    {
                        if (_isSelectedProp == null)
                            CacheItemReflectionPropertiesFromEnumerable(items);

                        foreach (var item in items)
                        {
                            if (_isSelectedProp != null)
                            {
                                try { _isSelectedProp.SetValue(item, false); }
                                catch (Exception ex)
                                {
                                    LogUtility.LogWarn($"SuggestAppendComboBox: failed to clear IsSelected on item (non-fatal): {ex.Message}");
                                }
                            }
                        }

                        try { SelectedItems?.Clear(); } catch (Exception ex) { LogUtility.LogWarn($"SuggestAppendComboBox: failed to clear SelectedItems (non-fatal): {ex.Message}"); }
                        SelectedItem = null;
                        UpdateDisplayedTextFromEnumerable(items);
                        SelectionCommitted?.Invoke(new List<object>());
                    }
                    else if (!IsMultiSelect)
                    {
                        SelectedItem = null;
                        UpdateToolTipText();
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"SuggestAppendComboBox.OnTextChanged handling clear (non-fatal): {ex.Message}");
                }
            }
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            ApplyFilterAndSuggest(_pendingUserInput);
        }

        private void ApplyFilterAndSuggest(string userInput)
        {
            if (_view == null || _textBox == null || _listBox == null || _popup == null)
                return;

            if (string.IsNullOrEmpty(userInput))
            {
                _view.Filter = null;
                _view.Refresh();
                _listBox.ItemsSource = null;
                _popup.IsOpen = false;
                return;
            }

            _view.Filter = (obj) =>
            {
                var txt = GetDisplayText(obj);
                return txt.StartsWith(userInput, StringComparison.OrdinalIgnoreCase);
            };
            _view.Refresh();

            var matches = _view.Cast<object>().ToList();

            if (matches.Count == 0)
            {
                _listBox.ItemsSource = null;
                _popup.IsOpen = false;
                return;
            }

            _listBox.ItemsSource = matches;
            _popup.IsOpen = true;

            if (_lastKey != Key.Back && _lastKey != Key.Delete)
            {
                string first = GetDisplayText(matches[0]);
                if (!string.IsNullOrEmpty(first) && first.Length > userInput.Length)
                {
                    int appendSelStart = userInput.Length;

                    try
                    {
                        _isInternalUpdate = true;
                        _textBox.Text = first;
                        _textBox.Select(appendSelStart, first.Length - appendSelStart);
                    }
                    finally
                    {
                        _isInternalUpdate = false;
                    }
                }
            }
        }

        private void CommitSelectionFromList()
        {
            if (IsMultiSelect)
            {
                if (_listBox == null) return;

                var selectedItems = _listBox.Items
                    .Cast<object>()
                    .Where(x =>
                    {
                        var prop = x.GetType().GetProperty(_IsSelectedPropName);
                        return prop != null && prop.GetValue(x) is bool b && b;
                    })
                    .ToList();

                SelectedItem = selectedItems;

                var multitxt = string.Join(";",
                    selectedItems.Select(i => GetDisplayText(i)));

                try
                {
                    _isInternalUpdate = true;
                    _textBox.Text = multitxt;
                    _textBox.Select(_textBox.Text.Length, 0);
                }
                finally { _isInternalUpdate = false; }

                UpdateToolTipText();
                SelectionCommitted?.Invoke(selectedItems);
                return;
            }

            if (_listBox == null || _listBox.SelectedItem == null)
            {
                if (_listBox != null && _listBox.Items.Count > 0)
                    _listBox.SelectedIndex = 0;
                else
                    return;
            }

            var sel = _listBox.SelectedItem;
            SelectedItem = sel;

            var txt = GetDisplayText(sel ?? SelectedItem ?? string.Empty);

            try
            {
                _isInternalUpdate = true;
                _textBox.Text = txt;
                _textBox.Select(_textBox.Text.Length, 0);
            }
            finally
            {
                _isInternalUpdate = false;
            }

            UpdateToolTipText();
            _popup.IsOpen = false;
            SelectionCommitted?.Invoke(sel);
        }

        private void CommitSelectionFromTextOrList()
        {
            if (_listBox != null && (
                    (!IsMultiSelect && _listBox.SelectedItem != null) ||
                    IsMultiSelect))
            {
                CommitSelectionFromList();
                return;
            }

            if (_view != null)
            {
                var first = _view.Cast<object>().FirstOrDefault();
                if (first != null)
                {
                    SelectedItem = first;
                    var txt = GetDisplayText(first);
                    try
                    {
                        _isInternalUpdate = true;
                        _textBox.Text = txt;
                        _textBox.Select(_textBox.Text.Length, 0);
                    }
                    finally
                    {
                        _isInternalUpdate = false;
                    }

                    UpdateToolTipText();
                    _popup.IsOpen = false;
                    SelectionCommitted?.Invoke(first);
                    return;
                }
            }

            var current = _textBox?.Text?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(current))
            {
                InvalidSelection?.Invoke(current);
                try
                {
                    _isInternalUpdate = true;
                    _textBox.Text = string.Empty;
                }
                finally
                {
                    _isInternalUpdate = false;
                }

                UpdateToolTipText();
            }
        }
        private void ValidateInputImmediate()
        {
            if (_textBox == null) return;

            string text = _textBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(text))
            {
                SelectedItem = null;
                UpdateToolTipText();

                // If the user cleared the textbox manually (Delete/Backspace), ensure multi-select
                // item flags are cleared as well so dependent refedit controls become enabled.
                try
                {
                    var items = _attachedItemsSource ?? ItemsSource ?? _listBox?.ItemsSource as IEnumerable;
                    if (IsMultiSelect && items != null)
                    {
                        // Ensure we have reflection info
                        if (_isSelectedProp == null)
                            CacheItemReflectionPropertiesFromEnumerable(items);

                        foreach (var item in items)
                        {
                            if (_isSelectedProp != null)
                            {
                                try
                                {
                                    _isSelectedProp.SetValue(item, false);

                                }
                                catch (Exception ex) 
                                { 
                                    LogUtility.LogWarn($"SuggestAppendComboBox: failed to clear IsSelected on item (non-fatal): {ex.Message}"); 
                                }
                            }
                        }

                        // Clear SelectedItems collection if present
                        try 
                        { 
                            SelectedItems?.Clear(); 
                        } 
                        catch (Exception ex)
                        {  
                          LogUtility.LogWarn($"SuggestAppendComboBox: failed to clear SelectedItems (non-fatal): {ex.Message}");
                        }

                        // Update visual text
                        UpdateDisplayedTextFromEnumerable(items);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"SuggestAppendComboBox: failed to update visual text (non-fatal): {ex.Message}");
                }

                return;
            }

            var match = ItemsSource?.Cast<object>()
                .FirstOrDefault(x => string.Equals(GetDisplayText(x), text, StringComparison.OrdinalIgnoreCase));

            if (match == null)
            {
                SelectedItem = null;
                try
                {
                    _isInternalUpdate = true;
                    _textBox.Text = string.Empty;
                }
                finally
                {
                    _isInternalUpdate = false;
                }

                UpdateToolTipText();
                InvalidSelection?.Invoke(text);
            }
            else
            {
                SelectedItem = match;
                UpdateToolTipText();
            }

            if (_popup != null)
            {
                _popup.IsOpen = false;
            }
        }

        private string GetDisplayText(object item)
        {
            if (item == null) return string.Empty;

            if (!string.IsNullOrEmpty(DisplayMemberPath))
            {
                var prop = item.GetType().GetProperty(DisplayMemberPath);
                if (prop != null)
                    return Convert.ToString(prop.GetValue(item)) ?? string.Empty;
            }

            return item.ToString();
        }
        public void SetSelectedItemWithText(object item)
        {
            SelectedItem = item;
            CommitSelectionFromList();
        }
    }
}