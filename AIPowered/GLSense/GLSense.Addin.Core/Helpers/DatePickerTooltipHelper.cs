// DatePickerTooltipHelper.cs in GLSense.Addin.Core
// Verbatim port of GLSense\Helpers\DatePickerTooltipHelper.cs (FinalWorkingCode) for
// Group C (Segment/Period pickers) - used by GLGetPeriodByDate and GLDailyRates to show
// a live tooltip of the selected date on their DatePicker controls. WPF-only, no
// old-project statics referenced beyond MahApps.Metro.IconPacks (already referenced by
// this project's csproj) - nothing to re-point.
using GLSense.Addin.Core.Infrastructure;
using MahApps.Metro.IconPacks;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Helpers
{
    public static class DatePickerTooltipHelper
    {
        /// <summary>
        /// Initializes and updates tooltip for a DatePicker with all events
        /// </summary>
        /// <param name="datePicker">The DatePicker control</param>
        /// <param name="title">Title to display (e.g., "Start Date", "End Date")</param>
        /// <param name="dispatcher">Dispatcher for UI thread operations</param>
        /// <param name="dateFormat">Date format string (default: "yyyy-MM-dd")</param>
        /// <param name="instructionText">Instruction text (default: "Click calendar icon to change date")</param>
        /// <param name="onDateChangedAction">Optional action to execute when date changes (for ViewModel)</param>
        public static void InitializeTooltip(
            DatePicker datePicker,
            string title,
            Dispatcher dispatcher = null,
            string dateFormat = "yyyy-MM-dd",
            string instructionText = "Click calendar icon to change date",
            Action<DatePicker> onDateChangedAction = null)
        {
            if (datePicker == null) return;

            ServiceLocator.Logger?.LogDebug($"DatePickerTooltipHelper.InitializeTooltip: title='{title}', dateFormat='{dateFormat}'");

            // Use provided dispatcher or fallback to current dispatcher
            var targetDispatcher = dispatcher ?? Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;

            // Store custom data in Tag to use in event handler
            datePicker.Tag = new TooltipData
            {
                Title = title,
                DateFormat = dateFormat,
                InstructionText = instructionText,
                OnDateChangedAction = onDateChangedAction
            };

            // Subscribe to SelectedDateChanged event
            datePicker.SelectedDateChanged -= DatePicker_SelectedDateChanged;
            datePicker.SelectedDateChanged += DatePicker_SelectedDateChanged;

            // Hook into the DatePickerTextBox for KeyDown events
            datePicker.Loaded += (s, e) =>
            {
                AttachTextBoxKeyDownHandler(datePicker);
            };

            // Initial tooltip update
            targetDispatcher.BeginInvoke(new Action(() =>
            {
                UpdateTooltip(datePicker);
            }), DispatcherPriority.Loaded);
        }

        /// <summary>
        /// Attaches KeyDown handler to the DatePickerTextBox
        /// </summary>
        private static void AttachTextBoxKeyDownHandler(DatePicker datePicker)
        {
            if (datePicker == null) return;

            // Find the DatePickerTextBox in the visual tree
            var textBox = FindVisualChild<DatePickerTextBox>(datePicker);
            if (textBox != null)
            {
                // Find the actual TextBox within DatePickerTextBox
                var innerTextBox = FindVisualChild<TextBox>(textBox);
                if (innerTextBox != null)
                {
                    // Subscribe to KeyDown event
                    innerTextBox.KeyDown -= DatePickerTextBox_KeyDown;
                    innerTextBox.KeyDown += DatePickerTextBox_KeyDown;
                }
            }
        }

        /// <summary>
        /// Updates the tooltip for a single DatePicker
        /// </summary>
        public static void UpdateTooltip(DatePicker datePicker)
        {
            if (datePicker == null) return;

            var tooltipData = datePicker.Tag as TooltipData;
            if (tooltipData == null)
            {
                // Fallback defaults if Tag is not set
                tooltipData = new TooltipData
                {
                    Title = "Selected Date",
                    DateFormat = "yyyy-MM-dd",
                    InstructionText = "Click calendar icon to change date"
                };
            }

            try
            {
                // Build the tooltip content
                var tooltipContent = CreateTooltipContent(datePicker, tooltipData);

                // Find the SimpleBrowserToolTip style - looked up on the DatePicker itself
                // (walks its actual logical tree up to whichever window hosts it, where
                // GlobalStyles.xaml is merged), not Application.Current (GlobalStyles.xaml
                // is merged per-window, not at the Application level, so that lookup could
                // never find it in the first place). TryFindResource also degrades to null
                // instead of throwing - NumericSpinner.xaml.cs already uses this same safe
                // pattern for the identical resource; plain FindResource here could
                // surface as an uncaught WPF-dispatcher-level ResourceReferenceKeyNotFoundException
                // even with the try/catch below, since some of WPF's own internal
                // style-application work happens outside it.
                var style = datePicker.TryFindResource("SimpleBrowserToolTip") as Style;

                var newTooltip = new ToolTip
                {
                    Style = style,
                    Content = tooltipContent
                };

                datePicker.ToolTip = newTooltip;
            }
            catch (Exception ex)
            {
                // Fallback to simple tooltip if there's an error
                datePicker.ToolTip = CreateFallbackTooltip(datePicker, tooltipData);
                ServiceLocator.Logger?.LogException(ex, "DatePickerTooltipHelper.UpdateTooltip");
                System.Diagnostics.Debug.WriteLine($"Error creating tooltip: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates the tooltip content with proper styling
        /// </summary>
        private static StackPanel CreateTooltipContent(DatePicker datePicker, TooltipData tooltipData)
        {
            var stackPanel = new StackPanel();

            // Header with icon
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };

            var icon = new PackIconFontAwesome
            {
                Kind = PackIconFontAwesomeKind.CalendarDaySolid,
                Width = 12,
                Height = 12,
                Foreground = GetResourceBrush("PrimaryBrush") ?? Brushes.Blue,
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            headerPanel.Children.Add(icon);

            var headerText = new TextBlock
            {
                Text = tooltipData.Title,
                FontWeight = FontWeights.SemiBold
            };
            headerPanel.Children.Add(headerText);
            stackPanel.Children.Add(headerPanel);

            // Date display
            var dateText = new TextBlock
            {
                FontFamily = new FontFamily("Segoe UI"),
                Background = GetResourceBrush("#F5F5F5") ?? new SolidColorBrush(Color.FromRgb(245, 245, 245)),
                Padding = new Thickness(4),
                Margin = new Thickness(0, 6, 0, 0),
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 300
            };

            if (datePicker.SelectedDate.HasValue)
            {
                dateText.Text = $"Selected Date: {datePicker.SelectedDate.Value.ToString(tooltipData.DateFormat)}";
            }
            else
            {
                dateText.Text = "Selected Date: No date selected";
            }

            stackPanel.Children.Add(dateText);

            // If the DatePicker exposes a display range (DisplayDateStart/End), show
            // the available selection info. For Start vs End date pickers, present
            // the single relevant bound so the tooltip reads naturally.
            try
            {
                if (datePicker.DisplayDateStart.HasValue || datePicker.DisplayDateEnd.HasValue)
                {
                    var start = datePicker.DisplayDateStart.HasValue ? datePicker.DisplayDateStart.Value.ToString(tooltipData.DateFormat) : null;
                    var end = datePicker.DisplayDateEnd.HasValue ? datePicker.DisplayDateEnd.Value.ToString(tooltipData.DateFormat) : null;

                    bool isStartTooltip = !string.IsNullOrWhiteSpace(tooltipData.Title) && tooltipData.Title.IndexOf("start", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool isEndTooltip = !string.IsNullOrWhiteSpace(tooltipData.Title) && tooltipData.Title.IndexOf("end", StringComparison.OrdinalIgnoreCase) >= 0;

                    var rangePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 0) };

                    if (isStartTooltip && start != null)
                    {
                        var rangeLabel = new TextBlock
                        {
                            Text = "Available From:",
                            FontWeight = FontWeights.SemiBold,
                            Margin = new Thickness(0, 0, 6, 0),
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        var rangeText = new TextBlock
                        {
                            Text = start,
                            Foreground = GetResourceBrush("TextSecondaryBrush") ?? Brushes.Gray,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        rangePanel.Children.Add(rangeLabel);
                        rangePanel.Children.Add(rangeText);
                        stackPanel.Children.Add(rangePanel);
                    }
                    else if (isEndTooltip && end != null)
                    {
                        var rangeLabel = new TextBlock
                        {
                            Text = "Available Until:",
                            FontWeight = FontWeights.SemiBold,
                            Margin = new Thickness(0, 0, 6, 0),
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        var rangeText = new TextBlock
                        {
                            Text = end,
                            Foreground = GetResourceBrush("TextSecondaryBrush") ?? Brushes.Gray,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        rangePanel.Children.Add(rangeLabel);
                        rangePanel.Children.Add(rangeText);
                        stackPanel.Children.Add(rangePanel);
                    }
                    else
                    {
                        // Generic fallback: show both bounds if available
                        var rangeLabel = new TextBlock
                        {
                            Text = "Available:",
                            FontWeight = FontWeights.SemiBold,
                            Margin = new Thickness(0, 0, 6, 0),
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        var rangeText = new TextBlock
                        {
                            Text = $"{start ?? "-"}  →  {end ?? "-"}",
                            Foreground = GetResourceBrush("TextSecondaryBrush") ?? Brushes.Gray,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        rangePanel.Children.Add(rangeLabel);
                        rangePanel.Children.Add(rangeText);
                        stackPanel.Children.Add(rangePanel);
                    }
                }
            }
            catch
            {
                // Non-fatal if range cannot be read
            }

            // Footer with instruction
            var footerBorder = new Border
            {
                Background = GetResourceBrush("TooltipBorderBrush") ?? new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                CornerRadius = new CornerRadius(3),
                Margin = new Thickness(0, 8, 0, 0)
            };

            var footerPanel = new StackPanel { Orientation = Orientation.Horizontal };

            var lightbulb = new PackIconFontAwesome
            {
                Kind = PackIconFontAwesomeKind.LightbulbRegular,
                Width = 12,
                Height = 12,
                Foreground = Brushes.Black,
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            footerPanel.Children.Add(lightbulb);

            var footerText = new TextBlock
            {
                Text = tooltipData.InstructionText,
                FontSize = 11,
                FontStyle = FontStyles.Italic,
                Foreground = Brushes.Black
            };
            footerPanel.Children.Add(footerText);
            footerBorder.Child = footerPanel;
            stackPanel.Children.Add(footerBorder);

            return stackPanel;
        }

        /// <summary>
        /// Creates a fallback tooltip in case of errors
        /// </summary>
        private static ToolTip CreateFallbackTooltip(DatePicker datePicker, TooltipData tooltipData)
        {
            var text = datePicker.SelectedDate.HasValue
                ? $"{tooltipData.Title}: {datePicker.SelectedDate.Value.ToString(tooltipData.DateFormat)}"
                : $"{tooltipData.Title}: No date selected";

            return new ToolTip
            {
                Content = text
            };
        }

        /// <summary>
        /// Event handler for SelectedDateChanged
        /// </summary>
        private static void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var datePicker = sender as DatePicker;
            if (datePicker == null) return;

            // Update binding source
            datePicker.GetBindingExpression(DatePicker.SelectedDateProperty)?.UpdateSource();

            // Execute custom action if provided
            var tooltipData = datePicker.Tag as TooltipData;
            if (tooltipData?.OnDateChangedAction != null)
            {
                try
                {
                    tooltipData.OnDateChangedAction.Invoke(datePicker);
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "DatePickerTooltipHelper.DatePicker_SelectedDateChanged (OnDateChangedAction callback)");
                    System.Diagnostics.Debug.WriteLine($"Error in OnDateChangedAction: {ex.Message}");
                }
            }

            // Update tooltip on UI thread
            Application.Current?.Dispatcher?.BeginInvoke(new Action(() =>
            {
                UpdateTooltip(datePicker);
            }), DispatcherPriority.Normal);
        }

        /// <summary>
        /// Event handler for KeyDown on the inner TextBox
        /// </summary>
        private static void DatePickerTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Tab)
            {
                var tb = sender as TextBox;
                tb?.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            }
        }

        /// <summary>
        /// Helper to find child elements in visual tree
        /// </summary>
        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                    return typedChild;

                var childOfChild = FindVisualChild<T>(child);
                if (childOfChild != null)
                    return childOfChild;
            }
            return null;
        }

        /// <summary>
        /// Helper to get a brush from resources
        /// </summary>
        private static Brush GetResourceBrush(string key)
        {
            try
            {
                if (Application.Current?.Resources.Contains(key) == true)
                {
                    return Application.Current.Resources[key] as Brush;
                }

                // Try to find in merged dictionaries
                foreach (var dict in Application.Current?.Resources?.MergedDictionaries ?? new System.Collections.ObjectModel.Collection<ResourceDictionary>())
                {
                    if (dict.Contains(key))
                    {
                        return dict[key] as Brush;
                    }
                }

                // If key is a color string, create a brush
                if (key.StartsWith("#") && key.Length >= 6)
                {
                    var converter = new System.Windows.Media.BrushConverter();
                    return converter.ConvertFromString(key) as Brush;
                }
            }
            catch (Exception ex)
            {
                // Return default brush on error
                ServiceLocator.Logger?.LogDebug($"DatePickerTooltipHelper.GetResourceBrush: failed to resolve resource '{key}', using default - {ex.Message}");
                return Brushes.LightGray;
            }

            return null;
        }

        /// <summary>
        /// Clean up event handlers to prevent memory leaks
        /// </summary>
        public static void Cleanup(DatePicker datePicker)
        {
            if (datePicker == null) return;

            ServiceLocator.Logger?.LogDebug("DatePickerTooltipHelper.Cleanup: detaching tooltip event handlers.");

            datePicker.SelectedDateChanged -= DatePicker_SelectedDateChanged;
            datePicker.Loaded -= (s, e) => { };

            // Find and remove KeyDown handler from inner TextBox
            var textBox = FindVisualChild<DatePickerTextBox>(datePicker);
            if (textBox != null)
            {
                var innerTextBox = FindVisualChild<TextBox>(textBox);
                if (innerTextBox != null)
                {
                    innerTextBox.KeyDown -= DatePickerTextBox_KeyDown;
                }
            }
        }

        /// <summary>
        /// Data class for tooltip configuration
        /// </summary>
        public class TooltipData
        {
            public string Title { get; set; }
            public string DateFormat { get; set; }
            public string InstructionText { get; set; }
            public Action<DatePicker> OnDateChangedAction { get; set; }
        }
    }
}
