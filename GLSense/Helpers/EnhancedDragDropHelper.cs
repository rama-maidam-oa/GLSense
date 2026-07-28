using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace GLSense.Helpers
{
    public static class EnhancedDragDropHelper
    {
        private static readonly HashSet<Type> _interactiveTypes =
        [
            // Standard WPF controls
            typeof(Button), typeof(TextBox), typeof(ComboBox), typeof(CheckBox),
            typeof(RadioButton), typeof(Slider), typeof(ScrollBar), typeof(ScrollViewer),
            typeof(DataGrid), typeof(ListBox), typeof(ListView), typeof(TreeView),
            typeof(DatePicker), typeof(Calendar), typeof(PasswordBox), typeof(RichTextBox),
            typeof(ButtonBase), typeof(Selector), typeof(RangeBase), typeof(Thumb),
            typeof(Hyperlink),

            // Custom controls
            typeof(Controls.SuggestAppendComboBox),
            typeof(Microsoft.Web.WebView2.Wpf.WebView2)
        ];

        private static readonly HashSet<Window> _registeredWindows = new HashSet<Window>();

        public static bool IsInteractiveControl(object source)
        {
            try
            {
                if (source == null) return false;

                var current = source as DependencyObject;
                while (current != null && current is not Window)
                {
                    if (_interactiveTypes.Any(t => t.IsInstanceOfType(current)))
                        return true;

                    current = VisualTreeHelper.GetParent(current);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "EnhancedDragDropHelper.IsInteractiveControl - non-fatal, ignored");
            }

            return false;
        }

        public static void EnableWindowDrag(Window window)
        {
            try
            {
                if (_registeredWindows.Contains(window)) return;

                window.MouseLeftButtonDown += Window_MouseLeftButtonDown;
                window.Closed += Window_Closed;
                _registeredWindows.Add(window);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "EnhancedDragDropHelper.EnableWindowDrag - non-fatal, ignored");
            }
        }

        public static void DisableWindowDrag(Window window)
        {
            try
            {
                if (!_registeredWindows.Contains(window)) return;

                window.MouseLeftButtonDown -= Window_MouseLeftButtonDown;
                window.Closed -= Window_Closed;
                _registeredWindows.Remove(window);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "EnhancedDragDropHelper.DisableWindowDrag - non-fatal, ignored");
            }
        }

        public static void RegisterInteractiveType(Type type)
        {
            try { _interactiveTypes.Add(type); }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "EnhancedDragDropHelper.RegisterInteractiveType - non-fatal, ignored");
            }
        }

        public static void UnregisterInteractiveType(Type type)
        {
            try { _interactiveTypes.Remove(type); }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "EnhancedDragDropHelper.UnregisterInteractiveType - non-fatal, ignored");
            }
        }

        private static void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender is not Window window)
                    return;

                if (IsInteractiveControl(e.OriginalSource))
                    return;

                // Avoid double-click races
                if (e.ClickCount > 1)
                    return;

                // Simple, safe behavior: don’t drag from maximized
                if (window.WindowState == WindowState.Maximized)
                    return;

                // Guard against the race condition
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    e.Handled = true;
                    window.DragMove();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exception in Window_MouseLeftButtonDown, this can be ignored");
            }
        }

        private static void Window_Closed(object sender, EventArgs e)
        {
            try
            {
                if (sender is Window window)
                {
                    DisableWindowDrag(window);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "EnhancedDragDropHelper.Window_Closed - non-fatal, ignored");
            }
        }
    }
}