// GLSegmentDiscovery.xaml.cs in GLSense.Addin.Core
// Port of GLSense\Views\GLSegmentDiscovery.xaml.cs (FinalWorkingCode) for Group D
// (Segment/Period discoverers) - modal dialog for filling segment values in a chosen
// direction. No separate ViewModel in the original and none added here - all logic stays
// in code-behind, matching the source file exactly (only the plumbing below changed):
//   - Base class DpiAwareWindow -> BaseWindow; EnhancedDragDropHelper.EnableWindowDrag(this)
//     -> TitleBar_MouseLeftButtonDown handler (same pattern as every other Group C/D
//     BaseWindow-derived view in this project, e.g. GLSegmentFunctions.xaml.cs).
//   - GLSense.Helpers/.Models/.Repositories/.Utilities -> GLSense.Addin.Core.* equivalents.
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp.
//   - LogUtility.* -> ServiceLocator.Logger.*.
//   - AddinModule.CurrentInstance.RibAsFormula.Pressed -> ServiceLocator.RibbonController.
//     GetControlPressed("RibAsFormula") - see SegmentDiscoverer.cs's header for why this
//     mechanism was added in this pass (no existing Group A/B/C code read a ribbon
//     toggle's live state back across the AppDomain boundary).
//   - Window_Loaded wired from the constructor (Loaded += Window_Loaded) instead of via
//     Loaded="Window_Loaded" in XAML, matching GLSegmentFunctions.xaml.cs's convention.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Views
{
    /// <summary>
    /// Interaction logic for GLSegmentDiscovery.xaml
    /// </summary>
    public partial class GLSegmentDiscovery : BaseWindow
    {
        private Excel.Range OrbSelection = null;
        private Excel.Range OrbFirstCell = null;
        private Excel.Range OrbActCell = null;
        private ObservableCollection<SegmentModel> AllSegments { get; set; }
        private ObservableCollection<SegmentValueModel> AllSegmentValues { get; set; }
        private Dictionary<string, bool> _summaryBySegment;
        private string[] ValueArray = null;

        private sealed record DirectionConfig
        {
            public int ItemCount { get; set; }
        }
        private enum FilterMode
        {
            All,
            Parents,
            Children
        }
        private sealed class WriteConfig
        {
            public bool IsInsert { get; set; }
            public bool IsFormula { get; set; }
            public string SheetName { get; set; }
            public string SegmentName { get; set; }
            public bool ParentChecked { get; set; }
            public bool ChildChecked { get; set; }
        }
        private enum DirectionType
        {
            None, Down, DownAll, Right, RightAll, Up, Left
        }
        private static bool IsVertical(DirectionType direction)
        {
            return direction == DirectionType.Down || direction == DirectionType.DownAll;
        }
        public GLSegmentDiscovery()
        {
            InitializeComponent();
            ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery constructor invoked");
            Loaded += Window_Loaded;
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
            ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.BtnClose_Click invoked - closing window");
            Close();
        }
        private async void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.BtnSubmit_Click invoked");

            // btnSubmit is disabled for the whole duration of a write below, but a click
            // already queued on the message loop right before that happens (e.g. an
            // impatient double-click) would still reach this handler a second time. Bail
            // out rather than starting a second overlapping write into a range that
            // already has formulas from the first write - see the Calculation-Manual
            // comment below for why an overlapping second write is what actually crashes/
            // freezes Excel (ported from FinalWorkingCode's identical fix for this file -
            // see that repo's CLAUDE.md).
            if (!btnSubmit.IsEnabled)
            {
                ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.BtnSubmit_Click: ignored - a write is already in progress");
                return;
            }

            if (!ValidatePrerequisites())
                return;

            bool busyShown = false;
            var originalCalculation = Excel.XlCalculation.xlCalculationAutomatic;
            bool calculationSaved = false;
            try
            {
                CommonMethods.DisableExcelSettings();

                // DisableExcelSettings() only toggles ScreenUpdating/DisplayAlerts/EnableEvents
                // - it leaves Calculation on Automatic. Each formula this writes (BuildFormula,
                // below) references the PREVIOUS cell in the chain, so on a cell range that's
                // already been written once before (e.g. clicking Insert a second time), every
                // single cell.Value assignment in the write loop dirties and immediately
                // recalculates its own entire downstream suffix of the chain - an O(n^2)
                // recalculation storm that's indistinguishable from a permanent freeze/crash,
                // with DisplayAlerts=false hiding any dialog that might otherwise have hinted
                // Excel was still (uselessly) working. Go Manual for the write, then do exactly
                // one Calculate() pass at the end (O(n) instead of O(n^2)).
                originalCalculation = ServiceLocator.ExcelApp.Calculation;
                calculationSaved = true;
                ServiceLocator.ExcelApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                LoadSegmentData();

                if (!ValidateOperationSelected() || ValueArray == null || ValueArray.Length == 0)
                {
                    ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.BtnSubmit_Click: no values to write (validation failed or empty ValueArray)");
                    return;
                }

                // Testing feedback: writing/inserting can take anywhere from a few
                // milliseconds to a few seconds depending on ValueArray.Length and
                // whether cells need to be shifted (Insert mode), so show a busy toast
                // for the duration instead of leaving the window looking unresponsive.
                // The actual Excel writes stay synchronous, in-process COM calls on the
                // UI thread exactly as before (WriteCellValue/PerformInsertIfNeeded are
                // unchanged) - only the ShowBusyasyn call plus periodic Dispatcher.Yield
                // calls are new. The first Yield here is what actually lets WPF paint the
                // busy overlay before the Excel calls start; without it, the overlay's
                // Visibility change and the Excel writes would both happen within the
                // same synchronous call stack and the overlay would never actually appear
                // on screen until after the writes had already finished. Follow-up
                // feedback: for a longer ValueArray, the busy overlay's spinner animation
                // itself looked frozen once the writes started - WriteValuesToExcelAsync
                // (below) now also yields every YieldBatchSize cells inside its write
                // loops (see that method's own comment) so the animation keeps rendering
                // throughout, not just before the first cell is written.
                AppOverlayControl.ShowBusyasyn("Writing segment values...");
                busyShown = true;
                btnSubmit.IsEnabled = false;
                await Dispatcher.Yield(DispatcherPriority.Render);

                ServiceLocator.Logger?.LogDebug($"GLSegmentDiscovery.BtnSubmit_Click: writing {ValueArray.Length} values to Excel");
                await WriteValuesToExcelAsync();

                // Single, deliberate recalculation pass now that every formula is in place.
                ServiceLocator.ExcelApp.Calculate();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentDiscovery.BtnSubmit_Click");
            }
            finally
            {
                btnSubmit.IsEnabled = true;
                if (calculationSaved)
                {
                    ServiceLocator.ExcelApp.Calculation = originalCalculation;
                }
                CommonMethods.TryEnableExcelSettings("GLSegmentDiscovery.BtnSubmit_Click");
                if (busyShown)
                    await AppOverlayControl.HideBusyAsync();
            }
        }
        private bool ValidatePrerequisites()
        {
            if (AppState.Instance.SelectedCube == null || AppState.Instance.SelectedLedger == null)
            {
                ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.ValidatePrerequisites: validation failed - no cube/ledger selected");
                AppOverlayControl.ShowWarning("Please select a valid Cube and Ledger before proceeding.");
                return false;
            }

            if (AppState.Instance.DefaultSegment == null || AppState.Instance.SegmentPickedIndex < 0)
            {
                ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.ValidatePrerequisites: validation failed - no default segment set");
                AppOverlayControl.ShowWarning("Please set a Default Segment in dropdown before proceeding.");
                return false;
            }

            var repository = new DataRepository();
            AllSegments = repository.GetSegments(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);

            if (AllSegments == null || AllSegments.Count == 0)
            {
                ServiceLocator.Logger?.LogWarn($"GLSegmentDiscovery.ValidatePrerequisites: no segments found for CubeId={AppState.Instance.SelectedCube.CubeId}, LedgerId={AppState.Instance.SelectedLedger.LedgerId}");
                AppOverlayControl.ShowWarning("No Segments found for the selected Cube and Ledger.");
                return false;
            }

            var defaultSeg = AllSegments.FirstOrDefault(s => s?.SegmentName == AppState.Instance.DefaultSegment);
            if (defaultSeg != null)
            {
                AllSegmentValues = DataRepository.GetSegmentValues(defaultSeg);
            }

            if (AllSegmentValues == null || AllSegmentValues.Count == 0)
            {
                ServiceLocator.Logger?.LogWarn($"GLSegmentDiscovery.ValidatePrerequisites: no segment values found for segment={AppState.Instance.DefaultSegment}");
                AppOverlayControl.ShowWarning("No Segment Values found for the selected Segment.");
                return false;
            }

            return true;
        }

        private bool ValidateOperationSelected()
        {
            bool downChecked = (bool)rbDown.IsChecked;
            bool downAllChecked = (bool)rbDownAll.IsChecked;
            bool rightChecked = (bool)rbRight.IsChecked;
            bool rightAllChecked = (bool)rbRightAll.IsChecked;
            bool upChecked = (bool)rbUp.IsChecked;
            bool leftChecked = (bool)rbLeft.IsChecked;

            if (!downChecked && !downAllChecked && !rightChecked &&
                !rightAllChecked && !leftChecked && !upChecked)
            {
                ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.ValidateOperationSelected: validation failed - no direction operation selected");
                AppOverlayControl.ShowWarning("Choose any operation for the segment discoverer and try again.");
                return false;
            }
            return true;
        }

        private void LoadSegmentData()
        {
            FillDetailstoArray();
        }

        private async Task WriteValuesToExcelAsync()
        {
            var direction = GetSelectedDirection();
            var config = new WriteConfig
            {
                IsInsert = rbInsert.IsChecked == true,
                IsFormula = ServiceLocator.RibbonController?.GetControlPressed("RibAsFormula") ?? false,
                SheetName = OrbActCell.Worksheet.Name,
                SegmentName = AppState.Instance.DefaultSegment,
                ParentChecked = chkParent.IsChecked == true,
                ChildChecked = chkChild.IsChecked == true
            };

            PerformInsertIfNeeded(direction, config);
            await WriteValuesByDirectionAsync(direction, config);
        }

        private DirectionType GetSelectedDirection()
        {
            if ((bool)rbDown.IsChecked)
            { return DirectionType.Down; }
            if ((bool)rbDownAll.IsChecked)
            { return DirectionType.DownAll; }
            if ((bool)rbLeft.IsChecked)
            { return DirectionType.Left; }
            if ((bool)rbUp.IsChecked)
            { return DirectionType.Up; }
            if ((bool)rbRight.IsChecked)
            { return DirectionType.Right; }
            if ((bool)rbRightAll.IsChecked)
            { return DirectionType.RightAll; }

            return DirectionType.None;
        }

        private void PerformInsertIfNeeded(DirectionType direction, WriteConfig config)
        {
            if (!config.IsInsert) return;

            if (direction == DirectionType.Down || direction == DirectionType.DownAll)
            {
                OrbActCell.Offset[1, 0].Resize[ValueArray.Length, 1].Select();
                var selectedRange = ServiceLocator.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            }
            else if (direction == DirectionType.Up)
            {
                OrbActCell.Offset[-ValueArray.Length, 0].Resize[ValueArray.Length, 1].Select();
                var selectedRange = ServiceLocator.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            }
            else if (direction == DirectionType.Right || direction == DirectionType.RightAll)
            {
                OrbActCell.Offset[0, 1].Resize[1, ValueArray.Length].Select();
                var selectedRange = ServiceLocator.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            else if (direction == DirectionType.Left)
            {
                OrbActCell.Offset[0, -ValueArray.Length].Resize[1, ValueArray.Length].Select();
                var selectedRange = ServiceLocator.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
        }
        // Testing feedback: the busy overlay shown around this whole operation
        // (BtnSubmit_Click) looked "frozen" - no spinner animation - for larger
        // ValueArray lengths. The Excel writes below are genuine synchronous, blocking
        // COM calls (Excel Interop is in-process here, not cross-process RPC, so there's
        // no message pump between calls) that must stay on this UI/STA thread - they
        // can't be moved to Task.Run. But blocking this thread in a tight loop for more
        // than a frame or two also starves WPF's composition engine of the periodic
        // hand-off it needs to actually paint the AppOverlay's busy Storyboard, even
        // though that animation is otherwise composition-thread-driven - a well-known WPF
        // gotcha. Fix: yield back to the dispatcher every YieldBatchSize cells so pending
        // render/animation frames get a chance to run in between chunks of Excel calls,
        // without adding a yield (and its small overhead) on every single cell.
        private const int YieldBatchSize = 10;

        private static async Task YieldEveryBatch(int index)
        {
            if ((index + 1) % YieldBatchSize == 0)
                await Dispatcher.Yield(DispatcherPriority.Render);
        }

        private async Task WriteValuesByDirectionAsync(DirectionType direction, WriteConfig config)
        {
            switch (direction)
            {
                case DirectionType.Down or DirectionType.DownAll:
                    await WriteVerticalForward(config);
                    break;
                case DirectionType.Right or DirectionType.RightAll:
                    await WriteHorizontalForward(config);
                    break;
                case DirectionType.Up:
                    await WriteVerticalBackward(config);
                    break;
                case DirectionType.Left:
                    await WriteHorizontalBackward(config);
                    break;
            }
        }
        private async Task WriteVerticalForward(WriteConfig config)
        {
            for (int i = 0; i < ValueArray.Length; i++)
            {
                var targetCell = GetVerticalCell(i);
                WriteCellValue(targetCell.Offset[1, 0], config, ValueArray[i], targetCell, DirectionType.Down);
                await YieldEveryBatch(i);
            }
        }

        private async Task WriteHorizontalForward(WriteConfig config)
        {
            for (int i = 0; i < ValueArray.Length; i++)
            {
                var targetCell = GetHorizontalCell(i);
                WriteCellValue(targetCell.Offset[0, 1], config, ValueArray[i], targetCell, DirectionType.Right);
                await YieldEveryBatch(i);
            }
        }

        private async Task WriteVerticalBackward(WriteConfig config)
        {
            for (int i = ValueArray.Length - 1; i >= 0; i--)
            {
                var targetCell = GetVerticalCell(-(i + 1));
                var formulaRef = targetCell.Offset[1, 0];
                WriteCellValue(targetCell, config, ValueArray[i], formulaRef, DirectionType.Up);
                await YieldEveryBatch(i);
            }
        }

        private async Task WriteHorizontalBackward(WriteConfig config)
        {
            for (int i = ValueArray.Length - 1; i >= 0; i--)
            {
                var targetCell = GetHorizontalCell(-(i + 1));
                var formulaRef = targetCell.Offset[0, 1];
                WriteCellValue(targetCell, config, ValueArray[i], formulaRef, DirectionType.Left);
                await YieldEveryBatch(i);
            }
        }
        private Excel.Range GetVerticalCell(int rowOffset) => OrbActCell.Offset[rowOffset, 0];
        private Excel.Range GetHorizontalCell(int colOffset) => OrbActCell.Offset[0, colOffset];

        private static void WriteCellValue(Excel.Range cell, WriteConfig config, string value, Excel.Range formulaRef, DirectionType direction)
        {
            cell.NumberFormat = config.IsFormula ? AppConstants.General : "@";
            string address = string.Empty;

            switch (direction)
            {
                case DirectionType.Down:
                case DirectionType.DownAll:
                case DirectionType.Up:
                    address = formulaRef.Address[false, true];
                    break;
                case DirectionType.Right:
                case DirectionType.RightAll:
                case DirectionType.Left:
                    address = formulaRef.Address[true, false];
                    break;
                default:
                    address = formulaRef.Address[true, true];
                    break;
            }

            cell.Value = config.IsFormula
                ? BuildFormula(config, address, direction)
                : (object)value;
        }

        private static string BuildFormula(WriteConfig config, string address, DirectionType direction)
        {
            string function = IsBackwardDirection(direction)
                ? "GLSense_GetPreviousSegment"
                : "GLSense_GetNextSegment";

            return $"={function}('{config.SheetName}'!{address},\"{config.SegmentName}\",{config.ParentChecked},{config.ChildChecked})";
        }
        private static bool IsBackwardDirection(DirectionType direction)
        {
            return direction == DirectionType.Up || direction == DirectionType.Left;
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ServiceLocator.Logger?.LogDebug("GLSegmentDiscovery.Window_Loaded invoked");
            OrbSelection = ServiceLocator.ExcelApp.Selection as Excel.Range;
            OrbFirstCell = OrbSelection.Cells[1, 1] as Excel.Range;
            OrbActCell = ServiceLocator.ExcelApp.ActiveCell;


            txtRangeSelected.Text = "'" + (OrbSelection.Parent is Excel.Worksheet worksheet ? worksheet.Name : string.Empty) + "'!" + OrbSelection.Address[RowAbsolute: true, ColumnAbsolute: true, External: false];
            txtSegmentValue.Text = OrbActCell.Value2 != null ? OrbActCell.Value2.ToString() : string.Empty;

            ResetControls();

            if (txtSegmentValue.Text == string.Empty)
            {
                btnSubmit.IsEnabled = false;
            }
        }
        private void ResetControls()
        {
            try
            {
                rbDown.IsChecked = false;
                rbDown.IsChecked = false;
                rbLeft.IsChecked = false;
                rbRight.IsChecked = false;
                rbRightAll.IsChecked = false;
                rbUp.IsChecked = false;

                rbDown.IsEnabled = false;
                rbDownAll.IsEnabled = false;
                rbLeft.IsEnabled = false;
                rbRight.IsEnabled = false;
                rbRightAll.IsEnabled = false;
                rbUp.IsEnabled = false;

                if (OrbSelection.Rows.Count == 1 && OrbSelection.Columns.Count == 1)
                {
                    rbDownAll.IsEnabled = true;
                    rbRightAll.IsEnabled = true;
                }
                else if (OrbSelection.Rows.Count > 1 && OrbActCell.Row == OrbFirstCell.Row)
                {
                    rbDown.IsEnabled = true;
                    rbDown.IsChecked = true;
                }
                else if (OrbSelection.Rows.Count > 1 && OrbActCell.Row != OrbFirstCell.Row)
                {
                    rbUp.IsEnabled = true;
                    rbUp.IsChecked = true;
                }
                else if (OrbSelection.Columns.Count > 1 && OrbActCell.Column == OrbFirstCell.Column)
                {
                    rbRight.IsEnabled = true;
                    rbRight.IsChecked = true;
                }
                else if (OrbSelection.Columns.Count > 1 && OrbActCell.Column != OrbFirstCell.Column)
                {
                    rbLeft.IsEnabled = true;
                    rbLeft.IsChecked = true;
                }

                chkParent.IsEnabled = true;
                chkParent.IsChecked = true;
                chkChild.IsEnabled = true;
                chkChild.IsChecked = true;
                rbOverwrite.IsEnabled = true;
                rbOverwrite.IsChecked = true;
                rbInsert.IsEnabled = true;
                btnSubmit.IsEnabled = true;

                UpdateInsertAvailability();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentDiscovery.ResetControls");
            }
        }

        private void DirectionRadio_Checked(object sender, RoutedEventArgs e)
        {
            UpdateInsertAvailability();
        }

        private void UpdateInsertAvailability()
        {
            if (rbInsert == null || rbOverwrite == null || rbDown == null || rbDownAll == null || rbRight == null || rbRightAll == null)
                return;

            bool insertAllowed = rbDown.IsChecked == true || rbDownAll.IsChecked == true ||
                                 rbRight.IsChecked == true || rbRightAll.IsChecked == true;

            rbInsert.IsEnabled = insertAllowed;

            if (!insertAllowed && rbInsert.IsChecked == true)
            {
                rbInsert.IsChecked = false;
                rbOverwrite.IsChecked = true;
            }
        }
        private void InitializeSummaryDictionary()
        {
            _summaryBySegment = AllSegmentValues
                .Where(x => x.SummaryFlag == "Y")  // Only include summary accounts
                .ToDictionary(x => x.SegmentValue, x => true);
        }
        private bool IsSummaryAccount(string segmentValue)
        {
            if (_summaryBySegment == null)
                InitializeSummaryDictionary();

            return _summaryBySegment.ContainsKey(segmentValue);
        }
        private void FillDetailstoArray()
        {
            try
            {
                var direction = GetSelectedDirection();
                ValueArray = direction == DirectionType.None
                    ? Array.Empty<string>()
                    : GetFilteredSegmentValues(direction);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLSegmentDiscovery.FillDetailstoArray");
            }
        }
        private string[] GetFilteredSegmentValues(DirectionType direction)
        {
            var allValues = AllSegmentValues
                .AsEnumerable()
                .Select(x => x.SegmentValue)
                .ToArray();

            var searchValue = txtSegmentValue.Text.Trim();
            int startIndex = Array.FindIndex(allValues, x => x == searchValue);
            if (startIndex == -1)
                return Array.Empty<string>();

            var directionConfig = GetDirectionConfig(direction);
            int itemsToTake = directionConfig.ItemCount;

            var candidateValues = SliceValues(allValues, startIndex, direction);
            var filteredValues = FilterValues(candidateValues, itemsToTake);
            return filteredValues;
        }
        private DirectionConfig GetDirectionConfig(DirectionType direction)
        {
            if (direction == DirectionType.Down || direction == DirectionType.Up)
                return new DirectionConfig
                {
                    ItemCount = OrbSelection.Rows.Count - 1
                };

            if (direction == DirectionType.Left || direction == DirectionType.Right)
                return new DirectionConfig
                {
                    ItemCount = OrbSelection.Columns.Count - 1
                };

            return new DirectionConfig { ItemCount = AllSegmentValues.Count() - 1 };
        }
        private static string[] SliceValues(string[] allValues, int startIndex, DirectionType direction)
        {
            if (direction == DirectionType.Up || direction == DirectionType.Left)
            {
                // Values before the current selection, nearest first
                int count = startIndex;
                var before = new string[count];
                Array.Copy(allValues, 0, before, 0, count);
                Array.Reverse(before);
                return before;
            }

            // Values after the current selection
            int remainingLength = Math.Max(0, allValues.Length - startIndex - 1);
            var after = new string[remainingLength];
            Array.Copy(allValues, startIndex + 1, after, 0, remainingLength);
            return after;
        }
        private string[] FilterValues(string[] values, int itemsToTake)
        {
            var filterMode = GetFilterMode();

            return filterMode switch
            {
                FilterMode.All => TakeUntilCount(values, itemsToTake),
                FilterMode.Parents => TakeSummaryAccounts(values, itemsToTake),
                FilterMode.Children => TakeNonSummaryAccounts(values, itemsToTake),
                _ => Array.Empty<string>()
            };
        }
        private FilterMode GetFilterMode()
        {
            bool parentChecked = chkParent.IsChecked == true;
            bool childChecked = chkChild.IsChecked == true;

            return (parentChecked, childChecked) switch
            {
                (true, true) or (false, false) => FilterMode.All,
                (true, false) => FilterMode.Parents,
                (false, true) => FilterMode.Children
            };
        }

        private static string[] TakeUntilCount(string[] values, int maxCount)
        {
            int actualCount = Math.Min(values.Length, maxCount);
            var result = new string[actualCount];
            Array.Copy(values, result, actualCount);
            return result;
        }

        private string[] TakeSummaryAccounts(string[] values, int maxCount)
        {
            var result = new List<string>();
            for (int i = 0; i < values.Length && result.Count < maxCount; i++)
            {
                if (IsSummaryAccount(values[i]))
                    result.Add(values[i]);
            }
            return result.ToArray();
        }

        private string[] TakeNonSummaryAccounts(string[] values, int maxCount)
        {
            var result = new List<string>();
            for (int i = 0; i < values.Length && result.Count < maxCount; i++)
            {
                if (!IsSummaryAccount(values[i]))
                    result.Add(values[i]);
            }
            return result.ToArray();
        }
    }

}
