// AttachmentsDialog.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\AttachmentsDialog.xaml(.cs) (FinalWorkingCode) - lists the journal
// attachments returned by the server (AppState.Instance.JournalDictionary) as checkboxes;
// on "Download Selected" writes the checked FILE_IDs (comma-separated) into
// AppState.Instance.AttachIDs for JournalAttachments.DownloadSelectedAttachments to pick
// up. Shown (modally, via SafeShowDialog) from JournalAttachments.RunJournalAttachmentFlow
// - see that file for the rest of the journal-attachment download flow.
//
// Re-pointed vs. the original:
//   - Base class DpiAwareWindow -> BaseWindow (a plain System.Windows.Window in this
//     project - WPF-UI's FluentWindow base class was removed - see GLDailyRates.xaml.cs
//     for the same idiom). BaseWindow sets the Excel owner automatically via
//     ServiceLocator.ExcelHandle, so there is no more explicit
//     ShowWithOwner((IntPtr)AppState.Instance.ExcelApp.Hwnd) call at the show site.
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> dedicated title-bar drag handler
//     (TitleBar_MouseLeftButtonDown), matching GLDailyRates.xaml/GLWaitWindow.xaml.
//   - LogUtility.LogException -> ServiceLocator.Logger?.LogException.
//   - The ListBox/CheckBox-find-and-collect logic and AppState.Instance.JournalDictionary/
//     AttachIDs usage are a straight port (unchanged).
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for AttachmentsDialog.xaml
    /// </summary>
    public partial class AttachmentsDialog : BaseWindow
    {
        public AttachmentsDialog()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("AttachmentsDialog constructor invoked");
            DataContext = this;
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("AttachmentsDialog.Window_Loaded invoked");
            try
            {
                AttachmentsListBox.Items.Clear();
                int count = AppState.Instance.JournalDictionary?.Keys.Count ?? 0;
                foreach (string key in AppState.Instance.JournalDictionary.Keys)
                {
                    // Add file name (value), unchecked by default
                    AttachmentsListBox.Items.Add(AppState.Instance.JournalDictionary[key]);
                }
                ServiceLocator.Logger?.LogDebug($"AttachmentsDialog.Window_Loaded: loaded {count} attachment(s)");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "AttachmentsDialog.Window_Loaded");
            }
        }

        private void CmdExecute_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("AttachmentsDialog.CmdExecute_Click invoked");
            try
            {
                StringBuilder selectedIds = new();

                foreach (object item in AttachmentsListBox.Items)
                {
                    // Find the ListBoxItem container
                    ListBoxItem lbi = (ListBoxItem)AttachmentsListBox.ItemContainerGenerator.ContainerFromItem(item);
                    if (lbi == null) continue;

                    // Get the CheckBox inside the template
                    CheckBox cb = FindVisualChild<CheckBox>(lbi);
                    if (cb != null && cb.IsChecked == true)
                    {
                        string fileName = item.ToString();
                        string fileId = KeyData(fileName);

                        if (!string.IsNullOrEmpty(fileId))
                        {
                            if (selectedIds.Length > 0)
                                selectedIds.Append(",");
                            selectedIds.Append(fileId);
                        }
                    }
                }

                AppState.Instance.AttachIDs = selectedIds.Length > 0 ? selectedIds.ToString() : string.Empty;
                ServiceLocator.Logger?.LogDebug($"AttachmentsDialog.CmdExecute_Click: selected attachment IDs = '{AppState.Instance.AttachIDs}'");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "AttachmentsDialog.CmdExecute_Click");
            }
            finally
            {
                this.Close();
            }
        }

        private static string KeyData(string itemDisplayName)
        {
            try
            {
                if (AppState.Instance.JournalDictionary == null || AppState.Instance.JournalDictionary.Count == 0)
                    return string.Empty;

                foreach (string key in AppState.Instance.JournalDictionary.Keys)
                {
                    if (AppState.Instance.JournalDictionary[key] == itemDisplayName)
                        return key;
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "AttachmentsDialog.KeyData");
                return string.Empty;
            }
        }

        // Helper to find CheckBox inside ListBoxItem template
        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T found)
                    return found;

                T childItem = FindVisualChild<T>(child);
                if (childItem != null)
                    return childItem;
            }
            return null;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("AttachmentsDialog.BtnClose_Click invoked - closing window");
            Close();
        }
    }
}
