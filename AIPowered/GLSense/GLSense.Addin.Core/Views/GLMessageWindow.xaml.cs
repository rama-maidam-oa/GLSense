// GLMessageWindow.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLMessageWindow.xaml.cs (FinalWorkingCode) - custom-styled
// message-box replacement window. Resolves the TODO in Utilities\CommonFunctions.cs's
// GLSenseMessage (see that method - now constructs and shows this window instead of a
// plain WPF MessageBox.Show fallback).
//
// Re-pointed vs. the original:
//   - Base class DpiAwareWindow -> BaseWindow (see GLDailyRates.xaml.cs/GLWaitWindow.xaml.cs
//     for the same idiom). BaseWindow sets the Excel owner automatically via
//     ServiceLocator.ExcelHandle, so there is no explicit ShowWithOwner()/SetExcelOwner()
//     call anywhere in this file or at its call site (CommonFunctions.GLSenseMessage).
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> dedicated title-bar drag handler
//     (TitleBar_MouseLeftButtonDown), matching GLDailyRates.xaml.cs/GLWaitWindow.xaml.cs
//     (the old HeaderPanel_MouseLeftButtonDown handler is dropped along with the old
//     bespoke HeaderPanel markup - see GLMessageWindow.xaml's header comment).
//   - Constructor ADAPTED (per this pass's explicit requirement) to take WPF's own
//     System.Windows.MessageBoxImage/MessageBoxButton enums instead of the old WinForms
//     System.Windows.Forms.MessageBoxIcon/MessageBoxButtons (this project has no WinForms
//     reference). The icon-kind/button-set switch logic is otherwise unchanged - just the
//     enum types:
//       MessageBoxImage.Error   (also covers Hand/Stop - same underlying value 16)
//       MessageBoxImage.Warning (also covers Exclamation - same underlying value 32)
//       MessageBoxImage.Information (also covers Asterisk - same underlying value 64)
//       MessageBoxImage.Question
//       (default/None -> generic message icon, same as the old Icon.None fallback)
//     WPF's MessageBoxButton enum (OK/OKCancel/YesNo/YesNoCancel) already matches the old
//     WinForms MessageBoxButtons 1:1 by name, so SetupButtons' switch is otherwise
//     unchanged.
//   - No more `using System.Windows.Forms;` (was only present in the old file to
//     disambiguate System.Windows.Controls.Button from System.Windows.Forms.Button; this
//     project has no WinForms reference, so plain `Button`/`new Button` is unambiguous).
using GLSense.Addin.Core.Infrastructure;
using MahApps.Metro.IconPacks;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace GLSense.Addin.Core.Views
{
    public partial class GLMessageWindow : BaseWindow
    {
        public MessageBoxResult Result { get; private set; } = MessageBoxResult.None;

        public GLMessageWindow(string message,
                               MessageBoxImage icon,
                               MessageBoxButton buttons = MessageBoxButton.OK)
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug($"GLMessageWindow constructor invoked: icon={icon}, buttons={buttons}, message={message}");

            // Set message text
            MsgText.Text = message;

            SetMessageIcon(icon);
            SetupButtons(buttons);
        }

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

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLMessageWindow.BtnClose_Click invoked - closing with Result=None");
            Result = MessageBoxResult.None;
            DialogResult = false;
            Close();
        }

        private void SetMessageIcon(MessageBoxImage icon)
        {
            switch (icon)
            {
                case MessageBoxImage.Error: // covers Hand/Stop - same underlying enum value
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleXmarkSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(231, 76, 60));
                    break;

                case MessageBoxImage.Warning: // covers Exclamation - same underlying enum value
                    MsgIcon.Kind = PackIconFontAwesomeKind.TriangleExclamationSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(241, 196, 15));
                    break;

                case MessageBoxImage.Information: // covers Asterisk - same underlying enum value
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleInfoSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(46, 134, 171));
                    break;

                case MessageBoxImage.Question:
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleQuestionSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(41, 128, 185));
                    break;

                default:
                    MsgIcon.Kind = PackIconFontAwesomeKind.MessageSolid;
                    MsgIcon.Foreground = Brushes.Gray;
                    break;
            }
        }

        private void SetupButtons(MessageBoxButton buttons)
        {
            ButtonPanel.Children.Clear();

            switch (buttons)
            {
                case MessageBoxButton.OK:
                    AddDialogButton("OK", MessageBoxResult.OK, PackIconFontAwesomeKind.CircleCheckSolid);
                    break;

                case MessageBoxButton.OKCancel:
                    AddDialogButton("OK", MessageBoxResult.OK, PackIconFontAwesomeKind.CircleCheckSolid);
                    AddDialogButton("Cancel", MessageBoxResult.Cancel, PackIconFontAwesomeKind.CircleXmarkSolid);
                    break;

                case MessageBoxButton.YesNo:
                    AddDialogButton("Yes", MessageBoxResult.Yes, PackIconFontAwesomeKind.CircleCheckSolid);
                    AddDialogButton("No", MessageBoxResult.No, PackIconFontAwesomeKind.CircleXmarkSolid);
                    break;

                case MessageBoxButton.YesNoCancel:
                    AddDialogButton("Yes", MessageBoxResult.Yes, PackIconFontAwesomeKind.CircleCheckSolid);
                    AddDialogButton("No", MessageBoxResult.No, PackIconFontAwesomeKind.CircleCheckSolid);
                    AddDialogButton("Cancel", MessageBoxResult.Cancel, PackIconFontAwesomeKind.CircleCheckSolid);
                    break;
            }
        }

        private void AddDialogButton(string text, MessageBoxResult resultValue, PackIconFontAwesomeKind iconKind)
        {
            Brush buttonForeground = Brushes.White;
            Color baseColor;

            switch (text)
            {
                case "OK":
                case "Yes":
                    baseColor = (Color)ColorConverter.ConvertFromString("#2E86AB");
                    break;
                case "No":
                    baseColor = (Color)ColorConverter.ConvertFromString("#EC616E");
                    break;
                case "Cancel":
                    baseColor = Colors.Gray;
                    break;
                default:
                    baseColor = Colors.LightGray;
                    buttonForeground = Brushes.Black;
                    break;
            }

            var baseBrush = new SolidColorBrush(baseColor);
            var hoverBrush = new SolidColorBrush(LightenColor(baseColor, 0.15));

            var btn = new Button
            {
                Width = 88,
                MinHeight = 36,
                Height = 36,
                Padding = new Thickness(12, 4, 12, 4),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                Background = baseBrush,
                Foreground = buttonForeground,
                Margin = new Thickness(ButtonPanel.Children.Count == 0 ? 0 : 8, 0, 0, 0),
                Cursor = Cursors.Hand
            };

            // Without an explicit Template, this Button renders through WPF's default
            // chrome, which layers its own theme hover/pressed highlight on top of
            // Background - independent of (and washing out) the baseBrush/hoverBrush
            // swap done manually below, which is why the button looked like it faded
            // to near-transparent/white on hover instead of the intended lighter
            // shade. PlainButtonTemplate is a bare Border bound to Background with no
            // such overlay, matching how CloseButtonStyle/ModernButton already render
            // correctly elsewhere in this app.
            btn.Template = TryFindResource("PlainButtonTemplate") as System.Windows.Controls.ControlTemplate;

            btn.MouseEnter += (s, e) => btn.Background = hoverBrush;
            btn.MouseLeave += (s, e) => btn.Background = baseBrush;
            btn.PreviewMouseDown += (s, e) => btn.Background = new SolidColorBrush(LightenColor(baseColor, -0.1));
            btn.PreviewMouseUp += (s, e) => btn.Background = hoverBrush;

            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            var faIcon = new PackIconFontAwesome
            {
                Kind = iconKind,
                Width = 16,
                Height = 16,
                Foreground = buttonForeground,
                Margin = new Thickness(0, 0, 5, 0)
            };

            var txt = new TextBlock
            {
                Text = text,
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            };

            panel.Children.Add(faIcon);
            panel.Children.Add(txt);
            btn.Content = panel;

            btn.Click += (sender, e) =>
            {
                ServiceLocator.Logger?.LogDebug($"GLMessageWindow dialog button clicked: Result={resultValue}");
                Result = resultValue;
                this.DialogResult = true;
                this.Close();
            };

            ButtonPanel.Children.Add(btn);
        }

        private static Color LightenColor(Color color, double amount)
        {
            double h, s, l;
            RgbToHsl(color, out h, out s, out l);
            l = Math.Max(0.0, Math.Min(1.0, l + amount));
            return HslToRgb(h, s, l);
        }

        private static void RgbToHsl(Color color, out double h, out double s, out double l)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));

            h = s = l = (max + min) / 2.0;

            if (max == min)
            {
                h = s = 0;
            }
            else
            {
                double d = max - min;
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

                if (max == r)
                    h = (g - b) / d + (g < b ? 6 : 0);
                else if (max == g)
                    h = (b - r) / d + 2;
                else
                    h = (r - g) / d + 4;

                h /= 6;
            }
        }

        private static Color HslToRgb(double h, double s, double l)
        {
            double r, g, b;

            if (s == 0)
            {
                r = g = b = l;
            }
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                double p = 2 * l - q;

                r = HueToRgb(p, q, h + 1.0 / 3.0);
                g = HueToRgb(p, q, h);
                b = HueToRgb(p, q, h - 1.0 / 3.0);
            }

            return Color.FromRgb(
                (byte)(r * 255),
                (byte)(g * 255),
                (byte)(b * 255));
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6.0) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2.0) return q;
            if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6;
            return p;
        }
    }
}
