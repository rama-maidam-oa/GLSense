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
            // Reuse the exact same shared style every other window's buttons use (Close/Cancel/
            // Insert/Download Logs, etc. all reference Themes\GlobalStyles.xaml::
            // DynamicContentButton) instead of a hand-rolled one. That style's idle state is a
            // plain black-on-light-gray look, and only the built-in IsMouseOver/IsPressed/
            // IsEnabled triggers change Background/BorderBrush/Foreground - so pulling it in
            // here keeps this window's buttons pixel-identical (corner radius, border, padding,
            // font weight, idle/hover/pressed/disabled colors) automatically, including staying
            // in sync if that shared style is ever updated.
            var btn = new System.Windows.Controls.Button
            {
                Style = TryFindResource("DynamicContentButton") as Style,
                Width = 88,
                MinHeight = 36,
                Height = 36,
                Margin = new Thickness(ButtonPanel.Children.Count == 0 ? 0 : 8, 0, 0, 0)
            };

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
                Margin = new Thickness(0, 0, 5, 0)
            };
            // PackIconFontAwesome doesn't participate in ordinary WPF Foreground inheritance the
            // way the plain TextBlock below does, so bind it explicitly to the button's own
            // Foreground - this way it follows the shared style's Black-idle/White-hover trigger
            // automatically instead of needing separate MouseEnter/MouseLeave handlers.
            faIcon.SetBinding(PackIconFontAwesome.ForegroundProperty, new System.Windows.Data.Binding("Foreground")
            {
                RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.FindAncestor, typeof(System.Windows.Controls.Button), 1)
            });

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
    }
}