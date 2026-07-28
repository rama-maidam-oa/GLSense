using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace GLSense.Controls
{
    /// <summary>
    /// A professional numeric up/down spinner control with min/max validation.
    /// </summary>
    public partial class NumericSpinner : UserControl
    {
        private Brush _defaultBorderBrush;
        private bool _isWarningActive;
        private bool _suppressTextUpdate;

        #region Dependency Properties

        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(nameof(Value), typeof(int), typeof(NumericSpinner),
                new FrameworkPropertyMetadata(0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnValueChanged, CoerceValue));

        public static readonly DependencyProperty MinimumProperty =
            DependencyProperty.Register(nameof(Minimum), typeof(int), typeof(NumericSpinner),
                new PropertyMetadata(0, OnMinMaxChanged));

        public static readonly DependencyProperty MaximumProperty =
            DependencyProperty.Register(nameof(Maximum), typeof(int), typeof(NumericSpinner),
                new PropertyMetadata(100, OnMinMaxChanged));

        public static readonly DependencyProperty StepProperty =
            DependencyProperty.Register(nameof(Step), typeof(int), typeof(NumericSpinner),
                new PropertyMetadata(1));

        public static readonly DependencyProperty HintTextProperty =
            DependencyProperty.Register(nameof(HintText), typeof(string), typeof(NumericSpinner),
                new PropertyMetadata(string.Empty, OnHintTextChanged));

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the current value.
        /// </summary>
        public int Value
        {
            get => (int)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        /// <summary>
        /// Gets or sets the minimum allowed value.
        /// </summary>
        public int Minimum
        {
            get => (int)GetValue(MinimumProperty);
            set => SetValue(MinimumProperty, value);
        }

        /// <summary>
        /// Gets or sets the maximum allowed value.
        /// </summary>
        public int Maximum
        {
            get => (int)GetValue(MaximumProperty);
            set => SetValue(MaximumProperty, value);
        }

        /// <summary>
        /// Gets or sets the increment/decrement step value.
        /// </summary>
        public int Step
        {
            get => (int)GetValue(StepProperty);
            set => SetValue(StepProperty, value);
        }

        /// <summary>
        /// Gets or sets the hint/tooltip text displayed when hovering over the control.
        /// </summary>
        public string HintText
        {
            get => (string)GetValue(HintTextProperty);
            set => SetValue(HintTextProperty, value);
        }

        #endregion

        #region Events

        public static readonly RoutedEvent ValueChangedEvent =
            EventManager.RegisterRoutedEvent(nameof(ValueChanged), RoutingStrategy.Bubble,
                typeof(RoutedPropertyChangedEventHandler<int>), typeof(NumericSpinner));

        /// <summary>
        /// Occurs when the Value property changes.
        /// </summary>
        public event RoutedPropertyChangedEventHandler<int> ValueChanged
        {
            add => AddHandler(ValueChangedEvent, value);
            remove => RemoveHandler(ValueChangedEvent, value);
        }

        #endregion

        #region Constructor

        public NumericSpinner()
        {
            InitializeComponent();
            // Use Loaded event to ensure visual tree is ready
            Loaded += OnNumericSpinnerLoaded;
        }

        private void OnNumericSpinnerLoaded(object sender, RoutedEventArgs e)
        {
            // Ensure template is applied
            ApplyTemplate();
            _defaultBorderBrush = MainBorder?.BorderBrush;
            SyncTextToValue();
            UpdateToolTip();
            UpdateValueToolTip();
        }

        #endregion

        #region Property Changed Callbacks

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var spinner = (NumericSpinner)d;
            var args = new RoutedPropertyChangedEventArgs<int>((int)e.OldValue, (int)e.NewValue, ValueChangedEvent);
            spinner.RaiseEvent(args);
            spinner.ClearWarning();
            spinner.SyncTextToValue();
            spinner.UpdateValueToolTip();
        }

        private static object CoerceValue(DependencyObject d, object baseValue)
        {
            var spinner = (NumericSpinner)d;
            var value = (int)baseValue;
            return Math.Max(spinner.Minimum, Math.Min(spinner.Maximum, value));
        }

        private static void OnMinMaxChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var spinner = (NumericSpinner)d;
            spinner.CoerceValue(ValueProperty);
        }

        private static void OnHintTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var spinner = (NumericSpinner)d;
            spinner.UpdateToolTip();
        }

        private void UpdateToolTip()
        {
            if (MainBorder == null) return;

            if (string.IsNullOrEmpty(HintText))
            {
                MainBorder.ToolTip = null;
            }
            else
            {
                var toolTipStyle = TryFindResource("SimpleBrowserToolTip") as Style;
                var toolTip = new ToolTip
                {
                    Content = HintText
                };

                // Apply style if found
                if (toolTipStyle != null)
                {
                    toolTip.Style = toolTipStyle;
                }

                MainBorder.ToolTip = toolTip;
            }
        }

        private void UpdateValueToolTip()
        {
            if (_isWarningActive || PART_TextBox == null) return;

            var toolTipStyle = TryFindResource("SimpleBrowserToolTip") as Style;
            var toolTip = new ToolTip
            {
                Content = Value.ToString(CultureInfo.InvariantCulture)
            };

            if (toolTipStyle != null)
            {
                toolTip.Style = toolTipStyle;
            }

            PART_TextBox.ToolTip = toolTip;
        }

        #endregion

        #region Event Handlers

        private void UpButton_Click(object sender, RoutedEventArgs e)
        {
            if (Value + Step <= Maximum)
                Value += Step;
            else
                Value = Maximum;
        }

        private void DownButton_Click(object sender, RoutedEventArgs e)
        {
            if (Value - Step >= Minimum)
                Value -= Step;
            else
                Value = Minimum;
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Allow only numeric input
            e.Handled = !IsNumericInput(e.Text);
        }

        private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up)
            {
                UpButton_Click(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.Down)
            {
                DownButton_Click(sender, e);
                e.Handled = true;
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressTextUpdate)
            {
                return;
            }

            ValidateText(false);
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateText(true);
        }

        private void TextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string));
                if (!IsNumericInput(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        #endregion

        #region Helper Methods

        private static bool IsNumericInput(string text)
        {
            return Regex.IsMatch(text, @"^[0-9]+$");
        }

        public bool ValidateText(bool commitValue)
        {
            if (PART_TextBox != null && int.TryParse(PART_TextBox.Text, out int parsedValue))
            {
                if (parsedValue < Minimum || parsedValue > Maximum)
                {
                    ShowWarning($"Value must be between {Minimum} and {Maximum}.");

                    if (commitValue)
                    {
                        var clampedValue = Math.Max(Minimum, Math.Min(Maximum, parsedValue));
                        if (Value != clampedValue)
                        {
                            Value = clampedValue;
                        }

                        SyncTextToValue();
                        ClearWarning();
                    }

                    return false;
                }

                if (Value != parsedValue)
                {
                    Value = parsedValue;
                }

                ClearWarning();
                return true;
            }
            else
            {
                if (PART_TextBox != null)
                {
                    if (string.IsNullOrWhiteSpace(PART_TextBox.Text))
                    {
                        ShowWarning("Value is required.");
                    }
                    else
                    {
                        ShowWarning("Enter a numeric value.");
                    }
                }

                return false;
            }
        }

        private void ShowWarning(string message)
        {
            _isWarningActive = true;

            if (MainBorder != null)
            {
                var dangerBrush = TryFindResource("DangerBrush") as Brush;
                MainBorder.BorderBrush = dangerBrush ?? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC3545"));
            }

            var warningStyle = TryFindResource("WarningToolTipStyle") as Style;
            var toolTip = new ToolTip
            {
                Content = message
            };

            if (warningStyle != null)
            {
                toolTip.Style = warningStyle;
            }

            if (PART_TextBox != null)
            {
                PART_TextBox.ToolTip = toolTip;
                toolTip.PlacementTarget = PART_TextBox;
                toolTip.StaysOpen = false;
                toolTip.IsOpen = true;
            }
        }

        private void ClearWarning()
        {
            if (!_isWarningActive)
            {
                UpdateValueToolTip();
                return;
            }

            _isWarningActive = false;

            if (MainBorder != null)
            {
                MainBorder.BorderBrush = _defaultBorderBrush ?? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CED4DA"));
            }

            if (PART_TextBox?.ToolTip is ToolTip openToolTip)
            {
                openToolTip.IsOpen = false;
            }

            UpdateValueToolTip();
        }

        private void SyncTextToValue()
        {
            if (PART_TextBox == null)
            {
                return;
            }

            _suppressTextUpdate = true;
            PART_TextBox.Text = Value.ToString(CultureInfo.InvariantCulture);
            PART_TextBox.CaretIndex = PART_TextBox.Text.Length;
            _suppressTextUpdate = false;
        }

        #endregion
    }

    public class NumericSpinnerValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return string.Empty;
            }

            return value.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = value as string;
            if (string.IsNullOrWhiteSpace(text))
            {
                return Binding.DoNothing;
            }

            if (int.TryParse(text, out int result))
            {
                return result;
            }

            return Binding.DoNothing;
        }
    }
}
