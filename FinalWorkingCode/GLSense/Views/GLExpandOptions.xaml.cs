using GLSense.Helpers;
using GLSense.Utilities;
using System;
using System.Windows;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLExpandOptions.xaml
    /// New window (ported alongside GLSense.Addin.Core's copy) replacing the old
    /// "Hierarchy" ribbon menu (RibSegmentExpand ADXRibbonMenu hosting the RibExpandAll /
    /// RibbonExpand1Level ADXRibbonButton menu items). RibSegmentExpand is now a single
    /// ADXRibbonButton whose OnClick opens this dialog instead of a menu. The user picks:
    ///   - Expand Level: Expand All vs Expand 1 Level - the same "HierarchyAll"/
    ///     "Hierarchy1Level" ActionType strings SegmentDiscoverer.SegmentAction already
    ///     recognized (previously selected by which of the two ribbon buttons was clicked).
    ///   - Fill Direction: By Rows (existing behavior - reads the selection top-to-bottom
    ///     and inserts new rows for each expanded child) vs By Columns (new - reads the
    ///     selection left-to-right and inserts new columns for each expanded child). See
    ///     SegmentDiscoverer.cs's byColumns-aware overloads (ValidateAreaValuesByColumnAsync
    ///     / ExpandSummaryAccountsAsync / InsertHierarchyExpansionByColumn /
    ///     InsertColumnsAndFillData) for the actual Excel-write logic this adds.
    /// No ViewModel - this dialog only reads its own RadioButtons on submit and hands off
    /// to SegmentDiscoverer.SegmentAction(actionType, byColumns), which already owns all
    /// progress-window/validation/Excel-write logic. This window closes immediately on
    /// Expand click since SegmentAction shows its own GLWaitWindow for the actual
    /// (potentially long-running) operation - there's no reason to keep this small options
    /// dialog on screen while that runs.
    /// </summary>
    public partial class GLExpandOptions : DpiAwareWindow
    {
        public GLExpandOptions()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);
            LogUtility.LogDebug("GLExpandOptions constructor invoked");
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLExpandOptions.BtnClose_Click invoked");
            Close();
        }

        private void BtnExpand_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string actionType = rbExpand1Level.IsChecked == true ? "Hierarchy1Level" : "HierarchyAll";
                bool byColumns = rbByColumns.IsChecked == true;

                LogUtility.LogDebug($"GLExpandOptions.BtnExpand_Click: actionType={actionType}, byColumns={byColumns}");

                // Close immediately - SegmentDiscoverer.SegmentAction shows its own
                // GLWaitWindow progress dialog (with cancel support) for the actual
                // operation, so there is no reason to keep this options dialog open while
                // it runs.
                Close();

                _ = SegmentDiscoverer.SegmentAction(actionType, byColumns);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLExpandOptions.BtnExpand_Click");
            }
        }
    }
}
