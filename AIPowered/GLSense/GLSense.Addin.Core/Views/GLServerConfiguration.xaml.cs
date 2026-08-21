// GLServerConfiguration.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLServerConfiguration.xaml.cs (FinalWorkingCode) - a simple
// XML-file-backed grid editor for named server URL instances, opened by the Riburl
// ribbon button (ribbon wiring itself is out of scope here - see PORTING_GUIDE.md /
// this group's task notes).
//
// Adjustments made when porting into this project's architecture:
//   - Base class DpiAwareWindow -> BaseWindow (same as every other window in this
//     project). BaseWindow already centers/modals against the Excel owner.
//   - EnhancedDragDropHelper.EnableWindowDrag(this) -> the dedicated
//     TitleBar_MouseLeftButtonDown handler already present on every other window here
//     (see GLLogin.xaml.cs / GLCubeDetails.xaml.cs for the identical pattern).
//   - AppPaths.TempUrlsPath -> ServiceLocator.Paths.UrlsDirectory (this project's
//     IPathProvider - already used this way by GLLogin.xaml.cs).
//   - "SuccessMessage"/"ErrorMessage" TextBlock styles now resolve from this project's
//     Themes\GlobalStyles.xaml (already defined there for other windows) instead of the
//     window-local styles the original XAML declared inline.
//   - No LogUtility usage in the original file, so nothing to re-point there.
//
// 2026-07-15 merge: re-synced against the old monolith's live-edited version to pick up
// fixes made there after this file's original port:
//   - Constructor now wires UrlInstances_CollectionChanged so PropertyChanged handlers get
//     attached/detached as rows are added/removed, primes the delete-button state right
//     after load, and collapses txtStatus initially (was showing an empty colored box).
//   - SaveConfiguration now rejects "half-filled" rows (Name without Address or vice
//     versa) with a per-row message instead of silently dropping them.
//   - SaveConfiguration/AutoSaveConfiguration's final row filter requires both Name AND
//     Address (previously Name-only), consistent with the new half-filled-row check.
//   - DgInstances_SelectionChanged / Instance_PropertyChanged enable the Delete button
//     when EITHER Name or Address has data (previously required both), so incomplete rows
//     can still be removed.
//   - UpdateStatus rewritten to collapse the status bar on empty messages, truncate long
//     messages, and route error/warning text through AppOverlayControl.ShowWarning (toast)
//     instead of an inline red status line, with a safe fallback if the overlay isn't
//     available.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml.Linq;
using GLSense.Addin.Core.Utilities;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLServerConfiguration.xaml
    /// </summary>
    public partial class GLServerConfiguration : BaseWindow
    {
        private readonly string xmlFilePath = ServiceLocator.Paths.UrlsDirectory;
        private readonly ObservableCollection<UrlInstance> urlInstances;

        // Snapshot of the last-persisted (loaded or saved) valid instances, used to detect
        // whether "Add" actually has anything to add/update (requirement: clicking Add with
        // no changes should show an info toast instead of silently re-saving).
        private List<(string Name, string Address, bool IsDefault)> _lastSavedSnapshot = new List<(string, string, bool)>();

        public GLServerConfiguration()
        {
            InitializeComponent();

            urlInstances = new ObservableCollection<UrlInstance>();
            dgInstances.ItemsSource = urlInstances;
            // Monitor collection changes so we can attach PropertyChanged handlers and keep
            // the status count in sync whenever rows are added or removed.
            urlInstances.CollectionChanged += UrlInstances_CollectionChanged;
            LoadConfiguration();
            // Ensure delete button reflects current selection/data
            DgInstances_SelectionChanged(this, null);
        }

        // ---------- Title bar (drag / close) ----------

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

        private void LoadConfiguration()
        {
            ServiceLocator.Logger?.LogDebug("GLServerConfiguration.LoadConfiguration invoked");
            try
            {
                EnsureConfigFilePath();
                if (!File.Exists(xmlFilePath))
                {
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

                _lastSavedSnapshot = BuildSnapshot();
                UpdateInstanceCountStatus();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error loading server configuration");
                UpdateStatus($"Error loading configuration: {ex.Message}");
            }
        }

        private void CreateDefaultConfig()
        {
            try
            {
                EnsureConfigFilePath();
                // Create empty XML structure with no URL entries
                var emptyConfig = new XDocument(
                    new XElement("ORBIT")
                );

                emptyConfig.Save(xmlFilePath);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error creating server configuration file");
                UpdateStatus($"Error creating configuration file: {ex.Message}");
            }
        }

        private void SaveConfiguration()
        {
            ServiceLocator.Logger?.LogDebug("GLServerConfiguration.SaveConfiguration invoked");
            try
            {
                EnsureConfigFilePath();

                // Validate that entries are complete: both Name and Address must be provided
                // together. Checked BEFORE the "no changes" check below, because a half-filled
                // row never shows up in BuildSnapshot() (it only counts fully-valid rows) - so
                // without this ordering an invalid row would silently fall through to "no
                // changes" instead of being flagged.
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

                    var msg = "Invalid entry - " + string.Join(", ", parts);
                    ServiceLocator.Logger?.LogDebug($"GLServerConfiguration.SaveConfiguration: validation failed - {msg}");
                    UpdateStatus(msg);
                    return;
                }

                if (!HasUnsavedChanges())
                {
                    ServiceLocator.Logger?.LogDebug("GLServerConfiguration.SaveConfiguration: no changes detected, nothing to add/update.");
                    AppOverlayControl?.ShowInfo("There are no changes for add or update.");
                    return;
                }

                // Validate data before saving
                var duplicateNames = urlInstances
                    .Where(u => !string.IsNullOrWhiteSpace(u.Name))
                    .GroupBy(u => u.Name.ToLower())
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();

                if (duplicateNames.Any())
                {
                    ServiceLocator.Logger?.LogDebug($"GLServerConfiguration.SaveConfiguration: validation failed - duplicate URL names: {string.Join(", ", duplicateNames)}");
                    UpdateStatus($"Duplicate URL Names found: {string.Join(", ", duplicateNames)}");
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

                Directory.CreateDirectory(Path.GetDirectoryName(xmlFilePath));
                doc.Save(xmlFilePath);
                ServiceLocator.Logger?.LogDebug($"GLServerConfiguration.SaveConfiguration: saved {validInstances.Count} instances successfully");
                _lastSavedSnapshot = BuildSnapshot();
                UpdateInstanceCountStatus();
                AppOverlayControl?.ShowSuccess("Configuration saved successfully.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error saving server configuration");
                UpdateStatus($"Error saving configuration: {ex.Message}");
            }
        }

        private void AutoSaveConfiguration()
        {
            ServiceLocator.Logger?.LogDebug("GLServerConfiguration.AutoSaveConfiguration invoked");
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

                Directory.CreateDirectory(Path.GetDirectoryName(xmlFilePath));
                doc.Save(xmlFilePath);
                _lastSavedSnapshot = BuildSnapshot();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error auto-saving server configuration");
                UpdateStatus($"Error auto-saving configuration: {ex.Message}");
            }
        }

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLServerConfiguration.BtnUpdate_Click invoked");
            SaveConfiguration();
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLServerConfiguration.BtnDelete_Click invoked");
            if (dgInstances.SelectedItem is UrlInstance selectedInstance)
            {
                string instanceName = string.IsNullOrWhiteSpace(selectedInstance.Name)
                                        ? "this instance"
                                        : $"'{selectedInstance.Name}'";

                AppOverlayControl.ShowConfirm(
                    $"Are you sure you want to delete instance {instanceName}?",
                    yesAction: () =>
                    {
                        ServiceLocator.Logger?.LogDebug($"GLServerConfiguration.BtnDelete_Click: user confirmed delete of {instanceName}");
                        urlInstances.Remove(selectedInstance);
                        AutoSaveConfiguration(); // Auto-save after deletion
                        AppOverlayControl?.ShowSuccess($"Instance {instanceName} deleted successfully.");
                    },
                    noAction: () =>
                    {
                        // User chose No, do nothing - nothing changed, so no status/toast needed.
                        ServiceLocator.Logger?.LogDebug($"GLServerConfiguration.BtnDelete_Click: user cancelled delete of {instanceName}");
                    }
                );
            }
            else
            {
                ServiceLocator.Logger?.LogDebug("GLServerConfiguration.BtnDelete_Click: validation failed - no instance selected");
                UpdateStatus("Please select an instance to delete.");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLServerConfiguration.BtnClose_Click invoked - closing window");
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

            // Keep the status count in sync with every add/remove, regardless of source
            // (initial load, Add button, Delete button, or the grid's own new-row placeholder).
            UpdateInstanceCountStatus();
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

                // A row can transition into/out of being a "counted" configuration as the
                // user types (Name/Address both need to be non-empty), so re-evaluate the
                // count on every edit, not just on add/remove.
                UpdateInstanceCountStatus();
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
                        UpdateStatus($"URL Name '{newValue}' already exists. Please use a unique name.");
                        e.Cancel = true;
                    }
                    // else: valid edit committed - the status area only ever shows the live
                    // configuration count (updated via Instance_PropertyChanged), so there's
                    // nothing further to show here.
                }
            }
        }

        /// <summary>
        /// Builds a comparable snapshot of the currently "valid" (both Name and Address
        /// filled) instances, using the same normalization SaveConfiguration itself applies
        /// before writing to disk. Used to detect whether there's actually anything for
        /// "Add" to add/update.
        /// </summary>
        private List<(string Name, string Address, bool IsDefault)> BuildSnapshot()
        {
            return urlInstances
                .Where(u => !string.IsNullOrWhiteSpace(u.Name) && !string.IsNullOrWhiteSpace(u.Address))
                .Select(u => (u.Name.Trim(), u.Address.Trim().TrimEnd('/'), u.IsDefault))
                .ToList();
        }

        private bool HasUnsavedChanges()
        {
            var current = BuildSnapshot();
            if (current.Count != _lastSavedSnapshot.Count) return true;

            for (int i = 0; i < current.Count; i++)
            {
                if (!current[i].Equals(_lastSavedSnapshot[i])) return true;
            }

            return false;
        }

        /// <summary>
        /// txtStatus's one and only job: show how many configurations currently exist. Not a
        /// generic message area - never put a sentence/confirmation in here (use a toast via
        /// AppOverlayControl instead). Counts only "valid" rows (both Name and Address
        /// filled), so an in-progress blank new row doesn't inflate the count.
        /// </summary>
        private void UpdateInstanceCountStatus()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    int count = urlInstances.Count(u => !string.IsNullOrWhiteSpace(u.Name) && !string.IsNullOrWhiteSpace(u.Address));
                    txtStatus.Text = $"Total configurations loaded : {count}";
                    txtStatus.Style = (Style)FindResource("SuccessMessage");
                    txtStatus.Visibility = Visibility.Visible;
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "GLServerConfiguration.UpdateInstanceCountStatus");
                }
            }));
        }

        /// <summary>
        /// Warning/error toast - the status area no longer shows inline messages, so every
        /// call site here now always routes to the warning toast (matches this window's
        /// existing warning-toast behavior, unchanged).
        /// </summary>
        private void UpdateStatus(string message)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    AppOverlayControl?.ShowWarning(message ?? "");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "GLServerConfiguration.UpdateStatus");
                }
            }));
        }

        private void EnsureConfigFilePath()
        {
            try
            {
                if (Directory.Exists(xmlFilePath))
                {
                    Directory.Delete(xmlFilePath, true);
                }

                var directory = Path.GetDirectoryName(xmlFilePath);
                if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error preparing server configuration path");
                UpdateStatus($"Error preparing configuration path: {ex.Message}");
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
