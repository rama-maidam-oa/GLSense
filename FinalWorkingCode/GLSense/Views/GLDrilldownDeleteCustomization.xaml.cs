using GLSense.Common;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace GLSense.Views
{
    /// <summary>
    /// Lets the user pick which locally-saved drilldown-metadata types (Balances/Journals/
    /// SubLedgers/Unified - see Common\DrilldownMetadataXmlStore.cs) to delete for the
    /// currently selected cube, instead of AddinModule's old all-or-nothing "Delete
    /// Customization" ribbon click. Opened only when at least one type is saved for the cube
    /// (AddinModule.RibDDDeleteConfiguration_OnClick checks this first).
    /// </summary>
    public partial class GLDrilldownDeleteCustomization : DpiAwareWindow
    {
        private readonly long _cubeId;
        private readonly ObservableCollection<GLDrilldownTypeModel> _types;

        // Guards against a second click starting a concurrent delete while the first is
        // still running its own AppOverlayControl warn/confirm/toast sequence - same
        // re-entry shape guarded elsewhere (e.g. GLJobsMonitor.xaml.cs's BtnDelete_Click).
        private bool _actionInProgress;

        public GLDrilldownDeleteCustomization(long cubeId, IReadOnlyList<(string DdType, int RecordCount)> savedTypes)
        {
            LogUtility.LogDebug($"GLDrilldownDeleteCustomization.ctor invoked, cubeId={cubeId}, savedTypes.Count={savedTypes?.Count ?? 0}");
            InitializeComponent();

            EnhancedDragDropHelper.EnableWindowDrag(this);

            _cubeId = cubeId;
            _types = new ObservableCollection<GLDrilldownTypeModel>(
                (savedTypes ?? Array.Empty<(string DdType, int RecordCount)>())
                    .Select(t => new GLDrilldownTypeModel
                    {
                        DdType = t.DdType,
                        DisplayName = GetDisplayName(t.DdType),
                        RecordCount = t.RecordCount
                    }));

            foreach (var type in _types)
            {
                type.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(GLDrilldownTypeModel.IsSelected))
                        UpdateHeaderCheckbox();
                };
            }

            dgTypes.ItemsSource = _types;
        }

        // Maps a saved records-key (BALANCE/JOURNAL/SUBLEDGER/UNIFIED) to the same
        // human-readable label used everywhere else in this codebase
        // (DrilldownMetadata.GetDisplay), rather than showing the raw API key.
        private static string GetDisplayName(string ddType)
        {
            DrilldownType? enumType = ddType?.ToUpperInvariant() switch
            {
                "BALANCE" => DrilldownType.BL,
                "JOURNAL" => DrilldownType.JL,
                "SUBLEDGER" => DrilldownType.SL,
                "UNIFIED" => DrilldownType.UF,
                _ => null
            };

            return enumType.HasValue ? DrilldownMetadata.GetDisplay(enumType.Value) : (ddType ?? string.Empty);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDrilldownDeleteCustomization.BtnClose_Click invoked");
            Close();
        }

        private void ChkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDrilldownDeleteCustomization.ChkSelectAll_Checked invoked");
            foreach (var type in _types)
                type.IsSelected = true;
        }

        private void ChkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLDrilldownDeleteCustomization.ChkSelectAll_Unchecked invoked");
            foreach (var type in _types)
                type.IsSelected = false;
        }

        private void UpdateHeaderCheckbox()
        {
            if (_types.Count == 0)
            {
                chkSelectAll.IsChecked = false;
                return;
            }

            bool allSelected = _types.All(t => t.IsSelected);
            bool anySelected = _types.Any(t => t.IsSelected);

            chkSelectAll.IsChecked = allSelected ? true : anySelected ? (bool?)null : false;
        }

        private void DgTypesRow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject source)
            {
                if (FindAncestor<CheckBox>(source) != null)
                    return;

                var row = FindAncestor<DataGridRow>(source);
                if (row?.DataContext is GLDrilldownTypeModel type)
                {
                    type.IsSelected = !type.IsSelected;
                    LogUtility.LogDebug($"GLDrilldownDeleteCustomization.DgTypesRow_PreviewMouseLeftButtonDown: {type.DdType} toggled, IsSelected={type.IsSelected}");
                    e.Handled = true;
                }
            }
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T typed)
                    return typed;

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        private async void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (_actionInProgress)
                return;

            LogUtility.LogDebug("GLDrilldownDeleteCustomization.BtnDelete_Click invoked");
            _actionInProgress = true;
            btnDelete.IsEnabled = false;
            try
            {
                await BtnDelete_ClickCore();
            }
            finally
            {
                _actionInProgress = false;
                btnDelete.IsEnabled = true;
            }
        }

        private async Task BtnDelete_ClickCore()
        {
            var selected = _types.Where(t => t.IsSelected).ToList();

            // Safeguard (b): nothing checked.
            if (selected.Count == 0)
            {
                LogUtility.LogDebug("GLDrilldownDeleteCustomization.BtnDelete_ClickCore: no rows selected.");
                await AppOverlayControl.ShowWarningAsync("Please select at least one drilldown type to delete.");
                return;
            }

            // Safeguard (c): confirm - this cannot be reversed.
            var userAction = await AppOverlayControl.ShowConfirmAsync(
                $"Delete the saved customization for {selected.Count} drilldown type(s)? This cannot be undone.");
            if (userAction != true)
            {
                LogUtility.LogDebug("GLDrilldownDeleteCustomization.BtnDelete_ClickCore: deletion cancelled by user.");
                return;
            }

            try
            {
                var wb = AppState.Instance.ExcelApp?.ActiveWorkbook;
                var ddTypes = selected.Select(t => t.DdType).ToList();
                bool deleted = DrilldownMetadataXmlStore.Delete(wb, _cubeId, ddTypes);

                if (deleted)
                {
                    foreach (var type in selected)
                        _types.Remove(type);

                    LogUtility.LogDebug($"GLDrilldownDeleteCustomization.BtnDelete_ClickCore: deleted {selected.Count} drilldown type(s) for cubeId={_cubeId}.");
                }

                UpdateHeaderCheckbox();
                await AppOverlayControl.ShowSuccessAsync($"{selected.Count} drilldown type(s) deleted successfully.");

                if (_types.Count == 0)
                    Close();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLDrilldownDeleteCustomization.BtnDelete_ClickCore");
                await AppOverlayControl.ShowErrorAsync("Failed to delete the selected drilldown customization(s).");
            }
        }
    }
}
