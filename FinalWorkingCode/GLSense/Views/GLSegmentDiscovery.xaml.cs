using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLSegmentDiscovery.xaml
    /// </summary>
    /// 

    public partial class GLSegmentDiscovery : DpiAwareWindow

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
            EnhancedDragDropHelper.EnableWindowDrag(this);
            LogUtility.LogDebug("GLSegmentDiscovery constructor invoked");

        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentDiscovery.BtnClose_Click invoked");
            Close();
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentDiscovery.BtnSubmit_Click invoked");

            if (!ValidatePrerequisites())
                return;

            try
            {
                CommonMethods.DisableExcelSettings();
                LoadSegmentData();

                if (!ValidateOperationSelected() || ValueArray == null || ValueArray.Length == 0)
                {
                    LogUtility.LogDebug($"GLSegmentDiscovery.BtnSubmit_Click: aborting - no operation selected or no values to write (ValueArray length={ValueArray?.Length ?? -1})");
                    return;
                }

                WriteValuesToExcel();
                LogUtility.LogDebug("GLSegmentDiscovery.BtnSubmit_Click: WriteValuesToExcel completed successfully");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                CommonMethods.EnableExcelSettings();
            }
        }
        private bool ValidatePrerequisites()
        {
            if (AppState.Instance.SelectedCube == null || AppState.Instance.SelectedLedger == null)
            {
                LogUtility.LogDebug("GLSegmentDiscovery.ValidatePrerequisites: validation failed - no Cube/Ledger selected");
                AppOverlayControl.ShowWarning("Please select a valid Cube and Ledger before proceeding.");
                return false;
            }

            if (AppState.Instance.DefaultSegment == null || AppState.Instance.SegmentPickedIndex < 0)
            {
                LogUtility.LogDebug("GLSegmentDiscovery.ValidatePrerequisites: validation failed - no Default Segment set");
                AppOverlayControl.ShowWarning("Please set a Default Segment in dropdown before proceeding.");
                return false;
            }

            var repository = new DataRepository();
            AllSegments = repository.GetSegments(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);

            if (AllSegments == null || AllSegments.Count == 0)
            {
                LogUtility.LogDebug($"GLSegmentDiscovery.ValidatePrerequisites: validation failed - no segments found for CubeId={AppState.Instance.SelectedCube.CubeId}, LedgerId={AppState.Instance.SelectedLedger.LedgerId}");
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
                LogUtility.LogDebug($"GLSegmentDiscovery.ValidatePrerequisites: validation failed - no Segment Values found for segment={AppState.Instance.DefaultSegment}");
                AppOverlayControl.ShowWarning("No Segment Values found for the selected Segment.");
                return false;
            }

            LogUtility.LogDebug($"GLSegmentDiscovery.ValidatePrerequisites: validation passed - segmentCount={AllSegments.Count}, segmentValueCount={AllSegmentValues.Count}");
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
                LogUtility.LogDebug("GLSegmentDiscovery.ValidateOperationSelected: validation failed - no operation/direction selected");
                AppOverlayControl.ShowWarning("Choose any operation for the segment discoverer and try again.");
                return false;
            }
            return true;
        }

        private void LoadSegmentData()
        {
            FillDetailstoArray();
        }

        private void WriteValuesToExcel()
        {
            var direction = GetSelectedDirection();
            var config = new WriteConfig
            {
                IsInsert = rbInsert.IsChecked == true,
                IsFormula = AddinModule.CurrentInstance.RibAsFormula.Pressed,
                SheetName = OrbActCell.Worksheet.Name,
                SegmentName = AppState.Instance.DefaultSegment,
                ParentChecked = chkParent.IsChecked == true,
                ChildChecked = chkChild.IsChecked == true
            };

            LogUtility.LogDebug($"GLSegmentDiscovery.WriteValuesToExcel invoked - direction={direction}, isInsert={config.IsInsert}, isFormula={config.IsFormula}, sheet={config.SheetName}, segment={config.SegmentName}, valueCount={ValueArray?.Length ?? 0}");

            PerformInsertIfNeeded(direction, config);
            WriteValuesByDirection(direction, config);

            LogUtility.LogDebug("GLSegmentDiscovery.WriteValuesToExcel: finished writing values to Excel");
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
                var selectedRange = AppState.Instance.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            }
            else if (direction == DirectionType.Up)
            {
                OrbActCell.Offset[-ValueArray.Length, 0].Resize[ValueArray.Length, 1].Select();
                var selectedRange = AppState.Instance.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            }
            else if (direction == DirectionType.Right || direction == DirectionType.RightAll)
            {
                OrbActCell.Offset[0, 1].Resize[1, ValueArray.Length].Select();
                var selectedRange = AppState.Instance.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            else if (direction == DirectionType.Left)
            {
                OrbActCell.Offset[0, -ValueArray.Length].Resize[1, ValueArray.Length].Select();
                var selectedRange = AppState.Instance.ExcelApp.Selection as Excel.Range;
                selectedRange.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }
        }
        private void WriteValuesByDirection(DirectionType direction, WriteConfig config)
        {
            switch (direction)
            {
                case DirectionType.Down or DirectionType.DownAll:
                    WriteVerticalForward(config);
                    break;
                case DirectionType.Right or DirectionType.RightAll:
                    WriteHorizontalForward(config);
                    break;
                case DirectionType.Up:
                    WriteVerticalBackward(config);
                    break;
                case DirectionType.Left:
                    WriteHorizontalBackward(config);
                    break;
            }
        }
        private void WriteVerticalForward(WriteConfig config)
        {
            for (int i = 0; i < ValueArray.Length; i++)
            {
                var targetCell = GetVerticalCell(i);
                WriteCellValue(targetCell.Offset[1, 0], config, ValueArray[i], targetCell, DirectionType.Down);
            }
        }

        private void WriteHorizontalForward(WriteConfig config)
        {
            for (int i = 0; i < ValueArray.Length; i++)
            {
                var targetCell = GetHorizontalCell(i);
                WriteCellValue(targetCell.Offset[0, 1], config, ValueArray[i], targetCell, DirectionType.Right);
            }
        }

        private void WriteVerticalBackward(WriteConfig config)
        {
            for (int i = ValueArray.Length - 1; i >= 0; i--)
            {
                var targetCell = GetVerticalCell(-(i + 1));
                var formulaRef = targetCell.Offset[1, 0];
                WriteCellValue(targetCell, config, ValueArray[i], formulaRef, DirectionType.Up);
            }
        }

        private void WriteHorizontalBackward(WriteConfig config)
        {
            for (int i = ValueArray.Length - 1; i >= 0; i--)
            {
                var targetCell = GetHorizontalCell(-(i + 1));
                var formulaRef = targetCell.Offset[0, 1];
                WriteCellValue(targetCell, config, ValueArray[i], formulaRef, DirectionType.Left);
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
            LogUtility.LogDebug("GLSegmentDiscovery.Window_Loaded invoked");

            OrbSelection = AppState.Instance.ExcelApp.Selection as Excel.Range;
            OrbFirstCell = OrbSelection.Cells[1, 1] as Excel.Range;
            OrbActCell = AppState.Instance.ExcelApp.ActiveCell;


            txtRangeSelected.Text = "'" + (OrbSelection.Parent is Excel.Worksheet worksheet ? worksheet.Name : string.Empty) + "'!" + OrbSelection.Address[RowAbsolute: true, ColumnAbsolute: true, External: false];
            txtSegmentValue.Text = OrbActCell.Value2 != null ? OrbActCell.Value2.ToString() : string.Empty;

            LogUtility.LogDebug($"GLSegmentDiscovery.Window_Loaded: rangeSelected={txtRangeSelected.Text}, segmentValue={txtSegmentValue.Text}");

            ResetControls();

            if (txtSegmentValue.Text == string.Empty)
            {
                LogUtility.LogDebug("GLSegmentDiscovery.Window_Loaded: segment value is empty - disabling Submit button");
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
                LogUtility.LogException(ex);
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
                LogUtility.LogException(ex);
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
            {
                LogUtility.LogDebug($"GLSegmentDiscovery.GetFilteredSegmentValues: searchValue '{searchValue}' not found in segment values - returning empty array");
                return Array.Empty<string>();
            }

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
