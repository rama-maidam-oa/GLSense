using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLServerConfiguration.xaml
    /// </summary>
    public partial class GLServerConfiguration : DpiAwareWindow
    {
        private readonly string xmlFilePath = AppPaths.TempUrlsPath;
        private readonly ObservableCollection<UrlInstance> urlInstances;
        public GLServerConfiguration()
        {
            LogUtility.LogDebug("GLServerConfiguration.ctor invoked");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            urlInstances = new ObservableCollection<UrlInstance>();
            dgInstances.ItemsSource = urlInstances;
            // Monitor collection changes so we can attach PropertyChanged handlers
            urlInstances.CollectionChanged += UrlInstances_CollectionChanged;
            LoadConfiguration();
            // Ensure delete button reflects current selection/data
            DgInstances_SelectionChanged(this, null);
            // Hide status area initially to avoid showing empty background
            try { txtStatus.Visibility = Visibility.Collapsed; } catch (Exception ex) { LogUtility.LogException(ex, "GLServerConfiguration.ctor: hide status text"); }
        }
        private void LoadConfiguration()
        {
            LogUtility.LogDebug($"GLServerConfiguration.LoadConfiguration invoked - xmlFilePath={xmlFilePath}");
            try
            {
                EnsureConfigFilePath();
                if (!File.Exists(xmlFilePath))
                {
                    LogUtility.LogDebug("GLServerConfiguration.LoadConfiguration: config file does not exist, creating default config");
                    CreateDefaultConfig();
                }

                XDocument doc = XDocument.Load(xmlFilePath);
                urlInstances.Clear();

                foreach (var urlElement in doc.Descendants("URL"))
                {
                    var instance = new UrlInstance
                    {
                        Name = urlElement.Element("Name")?.Value ?? "",
                        Address = urlElement.Element("Address")?.Value ?? "",
                        IsDefault = bool.Parse(urlElement.Element("DefaultURL")?.Value ?? "false")
                    };
                    urlInstances.Add(instance);
                    // Attach property change handler so UI state (e.g. delete button) updates immediately
                    AttachHandlersToInstance(instance);
                }

                UpdateStatus($"Configuration loaded successfully. {urlInstances.Count} instances found.", true);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLServerConfiguration.LoadConfiguration");
                UpdateStatus($"Error loading configuration: {ex.Message}", false);
            }
        }

        private void CreateDefaultConfig()
        {
            LogUtility.LogDebug("GLServerConfiguration.CreateDefaultConfig invoked");
            try
            {
                EnsureConfigFilePath();
                // Create empty XML structure with no URL entries
                var emptyConfig = new XDocument(
                    new XElement("ORBIT")
                );

                emptyConfig.Save(xmlFilePath);
                UpdateStatus("New configuration file created.", true);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLServerConfiguration.CreateDefaultConfig");
                UpdateStatus($"Error creating configuration file: {ex.Message}", false);
            }
        }

        private void SaveConfiguration()
        {
            LogUtility.LogDebug($"GLServerConfiguration.SaveConfiguration invoked - instanceCount={urlInstances.Count}");
            try
            {
                EnsureConfigFilePath();
                // Validate data before saving
                var duplicateNames = urlInstances
                    .Where(u => !string.IsNullOrWhiteSpace(u.Name))
                    .GroupBy(u => u.Name.ToLower())
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();

                if (duplicateNames.Any())
                {
                    LogUtility.LogDebug($"GLServerConfiguration.SaveConfiguration: validation failed - duplicate URL names: {string.Join(", ", duplicateNames)}");
                    UpdateStatus($"Duplicate URL Names found: {string.Join(", ", duplicateNames)}", false);
                    return;
                }

                // Validate that entries are complete: both Name and Address must be provided together
                var invalidInstancesWithIndex = urlInstances
                    .Select((u, idx) => new { Instance = u, Index = idx })
                    .Where(x => (!string.IsNullOrWhiteSpace(x.Instance.Name) && string.IsNullOrWhiteSpace(x.Instance.Address))
                                || (string.IsNullOrWhiteSpace(x.Instance.Name) && !string.IsNullOrWhiteSpace(x.Instance.Address)))
                    .ToList();

                if (invalidInstancesWithIndex.Any())
                {
                    var parts = new System.Collections.Generic.List<string>();
                    foreach (var item in invalidInstancesWithIndex)
                    {
                        var rowNum = item.Index + 1; // 1-based for user
                        if (string.IsNullOrWhiteSpace(item.Instance.Name))
                            parts.Add($"row {rowNum}: missing URL Name");
                        else if (string.IsNullOrWhiteSpace(item.Instance.Address))
                            parts.Add($"row {rowNum}: missing URL Address");
                    }

                    var msg = "Incomplete entries found - " + string.Join(", ", parts);
                    LogUtility.LogDebug($"GLServerConfiguration.SaveConfiguration: validation failed - {msg}");
                    UpdateStatus(msg, false);
                    return;
                }

                // Only save rows that have both URL Name and Address filled
                var validInstances = urlInstances.Where(u => !string.IsNullOrWhiteSpace(u.Name) && !string.IsNullOrWhiteSpace(u.Address)).ToList();

                var doc = new XDocument(
                    new XElement("ORBIT",
                        validInstances.Select(instance =>
                            new XElement("URL",
                                new XElement("Name", instance.Name?.Trim() ?? ""),
                                new XElement("Address", instance.Address?.Trim().TrimEnd('/') ?? ""),
                                new XElement("DefaultURL", instance.IsDefault.ToString())
                            )
                        )
                    )
                );

                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(xmlFilePath));
                doc.Save(xmlFilePath);
                LogUtility.LogDebug($"GLServerConfiguration.SaveConfiguration: saved {validInstances.Count} instance(s) successfully");
                UpdateStatus("Configuration saved successfully.", true);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLServerConfiguration.SaveConfiguration");
                UpdateStatus($"Error saving configuration: {ex.Message}", false);
            }
        }
        private void AutoSaveConfiguration()
        {
            LogUtility.LogDebug($"GLServerConfiguration.AutoSaveConfiguration invoked - instanceCount={urlInstances.Count}");
            try
            {
                EnsureConfigFilePath();
                // Only save rows that have both URL Name and Address filled
                var validInstances = urlInstances.Where(u => !string.IsNullOrWhiteSpace(u.Name) && !string.IsNullOrWhiteSpace(u.Address)).ToList();

                var doc = new XDocument(
                    new XElement("ORBIT",
                        validInstances.Select(instance =>
                            new XElement("URL",
                                new XElement("Name", instance.Name?.Trim() ?? ""),
                                new XElement("Address", instance.Address?.Trim().TrimEnd('/') ?? ""),
                                new XElement("DefaultURL", instance.IsDefault.ToString())
                            )
                        )
                    )
                );

                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(xmlFilePath));
                doc.Save(xmlFilePath);
                LogUtility.LogDebug($"GLServerConfiguration.AutoSaveConfiguration: saved {validInstances.Count} instance(s) successfully");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLServerConfiguration.AutoSaveConfiguration");
                UpdateStatus($"Error auto-saving configuration: {ex.Message}", false);
            }
        }
        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLServerConfiguration.BtnUpdate_Click invoked");
            SaveConfiguration();
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLServerConfiguration.BtnDelete_Click invoked");
            if (dgInstances.SelectedItem is UrlInstance selectedInstance)
            {
                string instanceName = string.IsNullOrWhiteSpace(selectedInstance.Name)
                                        ? "this instance"
                                        : $"'{selectedInstance.Name}'";

                AppOverlayControl.ShowConfirm(
                    $"Are you sure you want to delete instance {instanceName}?",
                    yesAction: () =>
                    {
                        LogUtility.LogDebug($"GLServerConfiguration.BtnDelete_Click: user confirmed delete of {instanceName}");
                        urlInstances.Remove(selectedInstance);
                        AutoSaveConfiguration(); // Auto-save after deletion
                        UpdateStatus($"Instance {instanceName} deleted successfully.", true);
                    },
                    noAction: () =>
                    {
                        // User chose No, do nothing
                        LogUtility.LogDebug($"GLServerConfiguration.BtnDelete_Click: user cancelled delete of {instanceName}");
                        UpdateStatus("Delete cancelled.", true);
                    }
                );
            }
            else
            {
                LogUtility.LogDebug("GLServerConfiguration.BtnDelete_Click: validation failed - no instance selected");
                UpdateStatus("Please select an instance to delete.", false);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLServerConfiguration.BtnClose_Click invoked");
            this.Close();
        }
        private void DgInstances_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Check if the selected item is valid and not an empty row
            if (dgInstances.SelectedItem is UrlInstance selectedInstance)
            {
                // Enable delete if the row has any data (Name or Address) so incomplete rows can be removed
                bool hasData = !string.IsNullOrWhiteSpace(selectedInstance.Name) ||
                              !string.IsNullOrWhiteSpace(selectedInstance.Address);
                btnDelete.IsEnabled = hasData;
            }
            else
            {
                btnDelete.IsEnabled = false;
            }
        }

        private void UrlInstances_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (var newItem in e.NewItems)
                {
                    if (newItem is UrlInstance inst)
                    {
                        AttachHandlersToInstance(inst);
                    }
                }
            }

            if (e.OldItems != null)
            {
                foreach (var oldItem in e.OldItems)
                {
                    if (oldItem is UrlInstance inst)
                    {
                        inst.PropertyChanged -= Instance_PropertyChanged;
                    }
                }
            }
        }

        private void AttachHandlersToInstance(UrlInstance instance)
        {
            if (instance == null) return;
            instance.PropertyChanged -= Instance_PropertyChanged;
            instance.PropertyChanged += Instance_PropertyChanged;
        }

        private void Instance_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // If the currently selected item changed its data, update delete button state immediately
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (dgInstances.SelectedItem is UrlInstance selectedInstance && ReferenceEquals(selectedInstance, sender))
                {
                    // Allow deleting when either Name or Address is provided so users can remove incomplete rows
                    bool hasData = !string.IsNullOrWhiteSpace(selectedInstance.Name) ||
                                   !string.IsNullOrWhiteSpace(selectedInstance.Address);
                    btnDelete.IsEnabled = hasData;
                }
            }));
        }
        private void DgInstances_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            // Real-time validation when user finishes editing a cell
            if (e.EditAction == DataGridEditAction.Commit && e.Column.Header.ToString() == "URL Name" && e.EditingElement is TextBox textBox)
            {
                string newValue = textBox.Text?.Trim();
                if (!string.IsNullOrEmpty(newValue))
                {
                    // Check for duplicates
                    var currentItem = e.Row.Item as UrlInstance;
                    var duplicates = urlInstances
                        .Where(u => u != currentItem && u.Name?.ToLower() == newValue.ToLower())
                        .ToList();

                    if (duplicates.Any())
                    {
                        LogUtility.LogDebug($"GLServerConfiguration.DgInstances_CellEditEnding: validation failed - duplicate URL Name '{newValue}'");
                        UpdateStatus($"URL Name '{newValue}' already exists. Please use a unique name.", false);
                        e.Cancel = true;
                    }
                    else
                    {
                        UpdateStatus("Changes made. Click 'Add' to save.", true);
                    }
                }
            }
        }
        private void UpdateStatus(string message, bool isSuccess)
        {
            LogUtility.LogDebug($"GLServerConfiguration.UpdateStatus - isSuccess={isSuccess}, message={message}");
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    if (isSuccess)
                    {
                        // Show concise success/info messages in the small status area.
                        // If there's no message, collapse the status area to avoid showing an empty colored box.
                        if (string.IsNullOrWhiteSpace(message))
                        {
                            try { txtStatus.Text = string.Empty; txtStatus.Style = null; txtStatus.Visibility = Visibility.Collapsed; } catch { }
                        }
                        else
                        {
                            var shortMsg = message ?? string.Empty;
                            if (shortMsg.Length > 100)
                                shortMsg = shortMsg.Substring(0, 97) + "...";

                            txtStatus.Text = shortMsg;
                            txtStatus.Style = (Style)FindResource("SuccessMessage");
                            txtStatus.Visibility = Visibility.Visible;
                        }
                    }
                    else
                    {
                        // Hide inline status for warnings/errors so empty background isn't shown
                        try { txtStatus.Text = string.Empty; txtStatus.Style = null; txtStatus.Visibility = Visibility.Collapsed; } catch { }

                        try
                        {
                            // Use warning toast by default
                            AppOverlayControl?.ShowWarning(message ?? "");
                        }
                        catch
                        {
                            // Fallback to status text if overlay unavailable
                            txtStatus.Text = message;
                            txtStatus.Style = (Style)FindResource("ErrorMessage");
                            txtStatus.Visibility = Visibility.Visible;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Best-effort: ensure exceptions here don't crash UI
                    LogUtility.LogWarn($"GLServerConfiguration.UpdateStatus: exception updating status UI (ignored): {ex.Message}");
                    try { txtStatus.Text = message; } catch { }
                }
            }));
        }

        private void EnsureConfigFilePath()
        {
            try
            {
                if (Directory.Exists(xmlFilePath))
                {
                    LogUtility.LogDebug($"GLServerConfiguration.EnsureConfigFilePath: xmlFilePath is unexpectedly a directory, deleting - {xmlFilePath}");
                    Directory.Delete(xmlFilePath, true);
                }

                var directory = Path.GetDirectoryName(xmlFilePath);
                if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                {
                    LogUtility.LogDebug($"GLServerConfiguration.EnsureConfigFilePath: creating missing directory - {directory}");
                    Directory.CreateDirectory(directory);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLServerConfiguration.EnsureConfigFilePath");
                UpdateStatus($"Error preparing configuration path: {ex.Message}", false);
            }
        }
    }

    public class UrlInstance : INotifyPropertyChanged
    {
        private string _name;
        private string _address;
        private bool _isDefault;

        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public string Address
        {
            get => _address;
            set
            {
                _address = value;
                OnPropertyChanged(nameof(Address));
            }
        }

        public bool IsDefault
        {
            get => _isDefault;
            set
            {
                _isDefault = value;
                OnPropertyChanged(nameof(IsDefault));
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

