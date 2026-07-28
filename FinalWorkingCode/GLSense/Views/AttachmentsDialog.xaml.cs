using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for AttachmentsDialog.xaml
    /// </summary>
    public partial class AttachmentsDialog : DpiAwareWindow
    {
        public AttachmentsDialog()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);


            DataContext = this;
            LogUtility.LogDebug("AttachmentsDialog.AttachmentsDialog: window constructed");
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("AttachmentsDialog.Window_Loaded invoked");
            try
            {
                AttachmentsListBox.Items.Clear();
                int count = 0;
                foreach (string key in AppState.Instance.JournalDictionary.Keys)
                {
                    // Add file name (value), unchecked by default
                    AttachmentsListBox.Items.Add(AppState.Instance.JournalDictionary[key]);
                    count++;
                }
                LogUtility.LogDebug($"AttachmentsDialog.Window_Loaded: populated {count} attachment item(s)");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "AttachmentsDialog.Window_Loaded");
            }
        }
        private void CmdExecute_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("AttachmentsDialog.CmdExecute_Click invoked");
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
                LogUtility.LogDebug($"AttachmentsDialog.CmdExecute_Click: selected attachment ids set - AttachIDs={AppState.Instance.AttachIDs}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "AttachmentsDialog.CmdExecute_Click");
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
                {
                    LogUtility.LogDebug("AttachmentsDialog.KeyData: JournalDictionary is null or empty");
                    return string.Empty;
                }

                foreach (string key in AppState.Instance.JournalDictionary.Keys)
                {
                    if (AppState.Instance.JournalDictionary[key] == itemDisplayName)
                        return key;
                }
                LogUtility.LogDebug($"AttachmentsDialog.KeyData: no matching key found for itemDisplayName={itemDisplayName}");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "AttachmentsDialog.KeyData");
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
            LogUtility.LogDebug("AttachmentsDialog.BtnClose_Click invoked");
            Close();
        }
    }
}

