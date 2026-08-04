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
            // Ported from GLSense\Views\GLMessageWindow.xaml.cs (FinalWorkingCode): buttons now
            // reuse the shared "DynamicContentButton" style (Themes\GlobalStyles.xaml) instead of
            // a hand-rolled inline one. That style's idle state is plain black-on-light-gray, and
            // only its built-in IsMouseOver/IsPressed/IsEnabled triggers change Background/
            // BorderBrush/Foreground - previously this window (like FinalWorkingCode's) had those
            // reversed (solid accent color idle, washed-out hover), backwards relative to every
            // other window's buttons (Close/Cancel/ModernButton, etc., all built on the same
            // DynamicContentButton). Reusing the real style keeps this pixel-identical and in sync
            // automatically instead of hand-rolling colors with the LightenColor/HSL helper
            // methods this file used to have (removed - they were only used here).
            var btn = new Button
            {
                Style = TryFindResource("DynamicContentButton") as Style,
                Width = 88,
                MinHeight = 36,
                Height = 36,
                Margin = new Thickness(ButtonPanel.Children.Count == 0 ? 0 : 8, 0, 0, 0)
            };

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
                Margin = new Thickness(0, 0, 5, 0)
            };
            // PackIconFontAwesome doesn't participate in ordinary WPF Foreground inheritance the
            // way the plain TextBlock below does, so bind it explicitly to the button's own
            // Foreground - this way it follows the shared style's Black-idle/White-hover trigger
            // automatically instead of needing separate MouseEnter/MouseLeave handlers.
            faIcon.SetBinding(PackIconFontAwesome.ForegroundProperty, new System.Windows.Data.Binding("Foreground")
            {
                RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.FindAncestor, typeof(Button), 1)
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
                ServiceLocator.Logger?.LogDebug($"GLMessageWindow dialog button clicked: Result={resultValue}");
                Result = resultValue;
                this.DialogResult = true;
                this.Close();
            };

            ButtonPanel.Children.Add(btn);
        }
    }
}
