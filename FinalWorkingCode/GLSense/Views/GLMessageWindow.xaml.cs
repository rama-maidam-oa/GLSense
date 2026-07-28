using GLSense.Helpers;
using GLSense.Utilities;
using MahApps.Metro.IconPacks;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;

namespace GLSense.Views
{
    public partial class GLMessageWindow : DpiAwareWindow
    {
        public MessageBoxResult Result { get; private set; } = MessageBoxResult.None;

        public GLMessageWindow(string message,
                               MessageBoxIcon icon,
                               MessageBoxButtons buttons = MessageBoxButtons.OK)
        {
            InitializeComponent();

            EnhancedDragDropHelper.EnableWindowDrag(this);

            LogUtility.LogDebug($"GLMessageWindow constructor invoked - icon={icon}, buttons={buttons}, message={message}");

            // Set message text
            MsgText.Text = message;

            SetMessageIcon(icon);
            SetupButtons(buttons);
        }

        private void HeaderPanel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
                DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLMessageWindow.BtnClose_Click invoked - result=Cancel");
            Result = MessageBoxResult.Cancel;
            DialogResult = false;
            Close();
        }

        private void SetMessageIcon(MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.Error:
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleXmarkSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(231, 76, 60));
                    break;

                case MessageBoxIcon.Warning:
                    MsgIcon.Kind = PackIconFontAwesomeKind.TriangleExclamationSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(241, 196, 15));
                    break;

                case MessageBoxIcon.Information:
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleInfoSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(46, 134, 171));
                    break;

                case MessageBoxIcon.Question:
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleQuestionSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(41, 128, 185));
                    break;

                default:
                    MsgIcon.Kind = PackIconFontAwesomeKind.MessageSolid;
                    MsgIcon.Foreground = Brushes.Gray;
                    break;
            }
        }

        private void SetupButtons(MessageBoxButtons buttons)
        {
            ButtonPanel.Children.Clear();

            switch (buttons)
            {
                case MessageBoxButtons.OK:
                    AddDialogButton("OK", MessageBoxResult.OK, PackIconFontAwesomeKind.CircleCheckSolid);
                    break;

                case MessageBoxButtons.OKCancel:
                    AddDialogButton("OK", MessageBoxResult.OK, PackIconFontAwesomeKind.CircleCheckSolid);
                    AddDialogButton("Cancel", MessageBoxResult.Cancel, PackIconFontAwesomeKind.CircleXmarkSolid);
                    break;

                case MessageBoxButtons.YesNo:
                    AddDialogButton("Yes", MessageBoxResult.Yes, PackIconFontAwesomeKind.CircleCheckSolid);
                    AddDialogButton("No", MessageBoxResult.No, PackIconFontAwesomeKind.CircleXmarkSolid);
                    break;

                case MessageBoxButtons.YesNoCancel:
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

            var btn = new System.Windows.Controls.Button
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
                Cursor = System.Windows.Input.Cursors.Hand
            };

            btn.MouseEnter += (s, e) => btn.Background = hoverBrush;
            btn.MouseLeave += (s, e) => btn.Background = baseBrush;
            btn.PreviewMouseDown += (s, e) => btn.Background = new SolidColorBrush(LightenColor(baseColor, -0.1));
            btn.PreviewMouseUp += (s, e) => btn.Background = hoverBrush;

            var panel = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center
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
                LogUtility.LogDebug($"GLMessageWindow: dialog button clicked - result={resultValue}");
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