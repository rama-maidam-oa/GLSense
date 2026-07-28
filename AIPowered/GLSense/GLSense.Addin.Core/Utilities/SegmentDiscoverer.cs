// SegmentDiscoverer.cs in GLSense.Addin.Core
// Port of GLSense\Utilities\SegmentDiscoverer.cs (FinalWorkingCode) for Group D
// (Segment/Period discoverers). Preserves the original's Excel-range-manipulation logic
// (hierarchy expansion, sheet-cloning safety threshold, offset-based fill) verbatim -
// only the plumbing below was re-pointed:
//   - GLSense.Helpers/.Models/.Repositories/.Views -> GLSense.Addin.Core.* equivalents.
//   - GLSense.Service.ServiceLocator.SegmentDataService -> Services.DataServiceLocator.
//     SegmentDataService (Group C's data-service layer; see DataServiceLocator.cs header
//     for why it was renamed from the old project's own "ServiceLocator").
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp (this project's AppState has
//     no ExcelApp field - Excel access always comes from the host via ServiceLocator).
//   - LogUtility.* -> ServiceLocator.Logger.*.
//   - WinForms MessageBoxIcon/MessageBoxButtons -> WPF MessageBoxImage/MessageBoxButton
//     (CommonFunctions.GLSenseMessage's actual signature in this project).
//   - AddinModule.CurrentInstance.RibAsFormula.Pressed -> ServiceLocator.RibbonController.
//     GetControlPressed("RibAsFormula") (new read-back method added to IRibbonController
//     for this group - see IRibbonController.cs; no existing Group A/B/C mechanism read a
//     ribbon toggle's *current* state back across the AppDomain boundary, only set it).
//   - CreateAndShowProgressWindow: WpfAppManager.InvokeOnWpfThread only takes an Action
//     (no Func<T> overload) in this project, so the GLWaitWindow is captured via a local
//     variable assigned from inside the delegate - the same pattern AddinEntry.
//     LedgerChanged already uses for the same reason. win.ShowWithOwner(hwnd) (old
//     GLWaitWindow) -> win.Show() (this project's BaseWindow-derived GLWaitWindow sets
//     the Excel owner automatically via ServiceLocator.ExcelHandle).
//   - ExcelWindowHelper.ActivateExcelMainWindow ported alongside this file (Helpers\
//     ExcelWindowHelper.cs) since nothing else needed it yet.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Services;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Utilities
{
#nullable enable
    public static class SegmentDiscoverer
    {
        private const int MaxSheetsThreshold = 200; // configurable safety threshold; Excel does not have a strict hard limit
        private static GLWaitWindow? Win { get; set; }
        private static Excel.Application? ExcelApp { get; set; }
        private static Excel.Workbook? HrWorbook { get; set; }
        private static Excel.Worksheet? HrWorksheet { get; set; }

        private static Excel.Range? CellActive { get; set; }
        private static string? Action { get; set; }

        private const string Hierarchy1Level = "Hierarchy1Level";
        private const string Explode1Level = "Explode1Level";

        private static ObservableCollection<SegmentValueModel>? SegmentValues;
        private static CancellationHelper? _ctsHelper;
        private static CancellationToken Token => _ctsHelper?.GetToken() ?? default;
        private static string? Title { get; set; }
        private static string? Msg { get; set; }

        private static readonly IReadOnlyDictionary<string, string> TitleLookup
                = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Property"] = "Segment Properties",
                    ["HierarchyAll"] = "Expanding Segment Values",
                    [Hierarchy1Level] = "Expanding Segment Values",
                    ["ExplodeAll"] = "Exploding Segment Values",
                    [Explode1Level] = "Exploding Segment Values"
                };
        private static readonly IReadOnlyDictionary<string, string> MsgLookup
                = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Property"] = "Getting Segment Value Properties",
                    ["HierarchyAll"] = "Expanding Segment Values By All Levels",
                    [Hierarchy1Level] = "Expanding Segment Value By 1 level",
                    ["ExplodeAll"] = "Exploding Segment Values By All Level",
                    [Explode1Level] = "Exploding Segment By Values 1 Level"
                };
        // Rest remains the same...
        private static string GetTitle(string action) =>
            TitleLookup.TryGetValue(action, out var title) ? title : "Segment Properties...";

        private static string GetMsg(string action) =>
            MsgLookup.TryGetValue(action, out var title) ? title : "Processing...";
        // byColumns: new option surfaced by GLExpandOptions.xaml (replaces the old
        // RibExpandAll/RibbonExpand1Level ribbon-menu pair). Only meaningful for
        // "HierarchyAll"/Hierarchy1Level (routed to FillSegmentHierarchies below) - every
        // other ActionType ignores it and keeps its existing behavior, so all other
        // callers can omit it.
        public static async Task SegmentAction(string ActionType, bool byColumns = false)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentDiscoverer.SegmentAction: started. ActionType='{ActionType}', byColumns={byColumns}");
            Action = ActionType;
            _ctsHelper = new CancellationHelper();
            try
            {
                CommonMethods.DisableExcelSettings();

                Title = GetTitle(ActionType);
                Msg = GetMsg(ActionType);

                if (ServiceLocator.ExcelApp == null)
                {
                    await ShowErrorMessage("Unable to get excel instance.");
                    return;
                }

                if (!AppState.Instance.IsLoginCompleted)
                {
                    await ShowErrorMessage("Please login to the instance.");
                    return;
                }

                if (AppState.Instance.DefaultSegment == null || AppState.Instance.DefaultSegment.Length == 0 || AppState.Instance.SegmentPickedIndex < 0)
                {
                    await ShowWarnMessage("Choose a segment from the segment dropdown of the ribbon tab, respective to the selected segment value.");
                    return;
                }

                Win = CreateAndShowProgressWindow(_ctsHelper);
                await InitializeProgressWindowAsync();

                SegmentValues = new ObservableCollection<SegmentValueModel>();
                SegmentValues = LoadSegmentValues(AppState.Instance.SelectedLedger.LedgerName);

                if (SegmentValues == null || !SegmentValues.Any())
                {
                    await ShowWarnMessage("Failed in fetching the segment values.");
                    return;
                }

                Token.ThrowIfCancellationRequested();

                ExcelApp = ServiceLocator.ExcelApp;
                CellActive = ExcelApp.ActiveCell;
                HrWorbook = ExcelApp.ActiveWorkbook;
                HrWorksheet = CellActive.Worksheet;

                if (CellActive.Value2 == null)
                {
                    await ShowWarnMessage("The first cell of the selection cannot be empty.");
                    return;
                }

                Token.ThrowIfCancellationRequested();
                string rngValue = CellActive.Value2.ToString();

                var value1 = SegValueModel(rngValue);
                Token.ThrowIfCancellationRequested();
                if (value1 == null)
                {
                    await ShowWarnMessage($"The selected item \"{rngValue}\" does not exists in the segment \"{AppState.Instance.DefaultSegment}\".");
                    return;
                }

                switch (ActionType)
                {
                    case "Property":
                        await FillSegmentProperties();
                        break;
                    case "HierarchyAll":
                    case Hierarchy1Level:
                        await FillSegmentHierarchies(byColumns);
                        break;
                    case "ExplodeAll":
                    case Explode1Level:
                        await ExplodeSegment();
                        break;
                }

                await MessageProgressWindowAsync("Excel refreshing the formulas.");

                HrWorksheet?.Select();
            }
            catch (OperationCanceledException)
            {
                await ShowCancelledAsync();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.SegmentAction(ActionType='{ActionType}'): failed");
            }
            finally
            {
                try
                {
                    if (_ctsHelper != null && !_ctsHelper.IsCancellationRequested)
                        _ctsHelper.Cancel();

                    _ctsHelper?.Dispose();  // always safe - handles all cases
                }
                catch (Exception ex)
                {
                    // Swallow dispose exceptions (Excel COM weirdness) but still log for
                    // diagnostics - same convention as BalanceRefresh.CleanupAsync.
                    ServiceLocator.Logger?.LogWarn($"SegmentDiscoverer.SegmentAction: exception disposing CancellationHelper (ignored): {ex.Message}");
                }

                await SafelyCloseWindowAsync();
                CommonMethods.EnableExcelSettings();
            }
        }
        //Helper to fill segment properties
        #region DTO used for snapshot (no COM references)
        private sealed class RowInput
        {
            public string? RawValue { get; set; }
            public string? SheetName { get; set; }
            public string? AddressA1 { get; set; }
        }
        #endregion
        private static async Task FillSegmentProperties()
        {
            try
            {
                var selection = await ValidateSelectionAsync();
                if (selection == null) return;
                Token.ThrowIfCancellationRequested();
                await MessageProgressWindowAsync("Filling segment property details.");

                bool writeAsFormula = ServiceLocator.RibbonController?.GetControlPressed("RibAsFormula") ?? false;
                string defaultSegment = AppState.Instance.DefaultSegment;

                foreach (Excel.Range area in selection.Areas)
                {
                    await ProcessAreaAsync(area, writeAsFormula, defaultSegment);
                }
            }
            catch (Exception ex)
            {
                await HandleUnexpectedErrorAsync(ex);
            }
        }
        private static async Task<Excel.Range?> ValidateSelectionAsync()
        {
            if (ExcelApp?.Selection is not Excel.Range selection || selection.Columns.Count > 1)
            {
                await ShowErrorMessage("Invalid selection to print segment properties." + Environment.NewLine +
                                       "Selection can be multi rows but with single column!");
                return null;
            }

            var rng = selection.Cells[1, 1] as Excel.Range;
            if (rng?.Row != CellActive?.Row)
            {
                await ShowErrorMessage("The selection should be from top To bottom.");
                return null;
            }

            return selection;
        }
        private static async Task ProcessAreaAsync(Excel.Range area, bool writeAsFormula, string defaultSegment)
        {
            var inputs = await ReadAreaInputsAsync(area);
            if (inputs.Length == 0) return;

            object[,] data = await ComputeSegmentPropertiesAsync(inputs, writeAsFormula, defaultSegment);
            Token.ThrowIfCancellationRequested();

            await WriteResultsAsync(area, data, writeAsFormula);
        }
        private static async Task<RowInput[]> ReadAreaInputsAsync(Excel.Range area)
        {
            int rows = area.Rows.Count;
            var inputs = new RowInput[rows];
            Token.ThrowIfCancellationRequested();
            await MessageProgressWindowAsync("Reading input data…");

            for (int r = 1; r <= rows; r++)
            {
                Token.ThrowIfCancellationRequested();
                var cell = area.Cells[r, 1] as Excel.Range;
                // Value2 is NOT mandatorily a string - a cell containing e.g. 1000 comes
                // back as a double, a checkbox/boolean cell as a bool, etc. The old
                // "is string" check silently skipped every non-text cell (numeric segment
                // values included), so nothing was discovered for rows like "1000".
                // NormalizeCellValue converts whatever Value2 actually is into the same
                // plain-string form SegValueModel expects to match against.
                string? rawValue = NormalizeCellValue(cell?.Value2);
                if (!string.IsNullOrEmpty(rawValue))
                {
                    inputs[r - 1] = new RowInput
                    {
                        RawValue = rawValue,
                        SheetName = cell!.Worksheet.Name,
                        AddressA1 = cell.get_Address(RowAbsolute: false, ColumnAbsolute: true,
                                                     ReferenceStyle: Excel.XlReferenceStyle.xlA1)
                    };
                }

                if (r % 200 == 0)
                {
                    await MessageProgressWindowAsync($"Reading input… {r}/{rows}");
                    await Task.Yield();
                }
            }

            return inputs;
        }

        // Excel.Range.Value2 returns whatever native type the cell actually holds -
        // string, double (numbers AND dates), bool, or an error code (int/other) - never
        // guaranteed to be a string. This normalizes any of those into the plain string
        // form SegValueModel looks values up by (e.g. a numeric segment value of 1000
        // must become "1000", not be discarded because it wasn't already a string).
        private static string? NormalizeCellValue(object? value)
        {
            switch (value)
            {
                case null:
                    return null;
                case string s:
                    return s;
                case double d:
                    // Fixed-point, no thousands separators, no trailing zeros, and no
                    // scientific notation even for large segment/account codes.
                    return d.ToString("0.################", System.Globalization.CultureInfo.InvariantCulture);
                case bool b:
                    return b ? "TRUE" : "FALSE";
                default:
                    return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        private static async Task<object[,]> ComputeSegmentPropertiesAsync(
            RowInput[] inputs, bool writeAsFormula, string defaultSegment)
        {
            await MessageProgressWindowAsync("Calculating values…");
            await Task.Yield();

            return await Task.Run(() =>
            {
                var data = new object[inputs.Length, 3];

                for (int i = 0; i < inputs.Length; i++)
                {
                    Token.ThrowIfCancellationRequested();
                    if (inputs[i] == null || string.IsNullOrEmpty(inputs[i].RawValue))
                        continue;

                    var inp = inputs[i];

                    var sModel = SegValueModel(inp.RawValue ?? string.Empty);
                    if (sModel == null) continue;

                    if (writeAsFormula)
                    {
                        string itm = $"'{inp.SheetName}'!{inp.AddressA1},\"{defaultSegment}\"";
                        string defaulLedger = $"\"{AppState.Instance.SelectedLedger.LedgerName}\"";
                        data[i, 0] = $"=GLSense_GetSegmentDesc({itm},FALSE, {defaulLedger})";
                        data[i, 1] = $"=GLSense_GetSegmentEnabledFlag({itm}, {defaulLedger})";
                        data[i, 2] = $"=GLSense_GetSegmentSummaryFlag({itm}, {defaulLedger})";
                    }
                    else
                    {
                        data[i, 0] = sModel.Description;
                        data[i, 1] = sModel.EnabledFlag;
                        data[i, 2] = sModel.SummaryFlag;
                    }
                }

                return data;
            });
        }
        private static async Task WriteResultsAsync(Excel.Range area, object[,] data, bool writeAsFormula)
        {
            await MessageProgressWindowAsync("Writing results to Excel…");

            var target = area.Offset[0, 1].Resize[area.Rows.Count, 3];
            target.NumberFormat = AppConstants.General;

            Token.ThrowIfCancellationRequested();

            if (writeAsFormula)
                target.Formula = data;
            else
                target.Value2 = data;

            await MessageProgressWindowAsync("Filling segment property details completed.");
            await Task.Yield();
        }

        private static async Task FillSegmentHierarchies(bool byColumns = false)
        {
            bool oneLevel = Action == Hierarchy1Level;
            const string DummyToken = "GLSDummy";

            try
            {
                await MessageProgressWindowAsync("Filling segment hierarchy details.");

                if (ExcelApp?.Selection is not Excel.Range selection) return;

                foreach (Excel.Range area in selection.Areas)
                {
                    Token.ThrowIfCancellationRequested();
                    var validatedValues = byColumns
                        ? ValidateAreaValuesByColumnAsync(area, DummyToken)
                        : ValidateAreaValuesAsync(area, DummyToken);
                    if (validatedValues.Count == 0) continue;

                    await ExpandSummaryAccountsAsync(validatedValues, area, oneLevel, byColumns);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.FillSegmentHierarchies: failed");
            }
        }
        private static List<string> ValidateAreaValuesAsync(
            Excel.Range area, string dummyToken)
        {
            var listRangeValue = new List<string>(area.Rows.Count);

            try
            {
                for (int i = 1; i <= area.Rows.Count; i++)
                {
                    Token.ThrowIfCancellationRequested();
                    var cell = area.Cells[i, 1] as Excel.Range;
                    var v = cell?.Value2;

                    if (v == null)
                    {
                        listRangeValue.Add(dummyToken);
                        continue;
                    }

                    string value = v.ToString().Trim();
                    listRangeValue.Add(SegmentValueExists(value) ? value : dummyToken);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.ValidateAreaValuesAsync: failed");
            }

            return listRangeValue;
        }

        // Column-wise counterpart of ValidateAreaValuesAsync: reads across row 1 of the
        // area (one value per column) instead of down column 1 (one value per row). Used
        // when GLExpandOptions' "By Columns" RadioButton is selected.
        private static List<string> ValidateAreaValuesByColumnAsync(
            Excel.Range area, string dummyToken)
        {
            var listRangeValue = new List<string>(area.Columns.Count);

            try
            {
                for (int j = 1; j <= area.Columns.Count; j++)
                {
                    Token.ThrowIfCancellationRequested();
                    var cell = area.Cells[1, j] as Excel.Range;
                    var v = cell?.Value2;

                    if (v == null)
                    {
                        listRangeValue.Add(dummyToken);
                        continue;
                    }

                    string value = v.ToString().Trim();
                    listRangeValue.Add(SegmentValueExists(value) ? value : dummyToken);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.ValidateAreaValuesByColumnAsync: failed");
            }

            return listRangeValue;
        }
        private static async Task ExpandSummaryAccountsAsync(
            List<string> validatedValues,
            Excel.Range area,
            bool oneLevel,
            bool byColumns = false)
        {
            if (area.Cells[1, 1] is not Excel.Range firstCell) return;

            if (byColumns)
            {
                int startCol = firstCell.Column;
                int rowNum = firstCell.Row;
                int rowCount = area.Rows.Count;
                bool multiRow = rowCount > 1;

                for (int i = 0; i < validatedValues.Count; i++)
                {
                    Token.ThrowIfCancellationRequested();
                    string value = validatedValues[i];

                    if (value == "GLSDummy")
                    {
                        startCol++;
                        continue;
                    }

                    if (!IsSummaryAccount(value))
                    {
                        startCol++;
                        continue;
                    }

                    await InsertHierarchyExpansionByColumn(value, rowNum, startCol, rowCount, multiRow, oneLevel);
                    int insertedCount = await GetInsertedChildCountAsync(value, oneLevel);
                    startCol += insertedCount + 1;
                }
                return;
            }

            int startRow = firstCell.Row;
            int columnNum = firstCell.Column;
            int columnCount = area.Columns.Count;
            bool multiColumn = columnCount > 1;

            for (int i = 0; i < validatedValues.Count; i++)
            {
                Token.ThrowIfCancellationRequested();
                string value = validatedValues[i];

                if (value == "GLSDummy")
                {
                    startRow++;
                    continue;
                }

                if (!IsSummaryAccount(value))
                {
                    startRow++;
                    continue;
                }

                await InsertHierarchyExpansion(value, startRow, columnNum, columnCount, multiColumn, oneLevel);
                int insertedCount = await GetInsertedChildCountAsync(value, oneLevel);
                startRow += insertedCount + 1;
            }
        }
        private static async Task InsertHierarchyExpansion(
        string value, int startRow, int columnNum, int columnCount,
        bool multiColumn, bool oneLevel)
        {
            Token.ThrowIfCancellationRequested();
            var children = await GetHierarchyChildrenAsync(value, oneLevel);
            if (children?.Count == 0) return;

            string title = AppState.Instance.DefaultSegment + " (" + value + ")";
            await MessageProgressWindowAsync($"Filling segment hierarchy values for segment {title}.");
            await Task.Yield();

            InsertRowsAndFillData(children, startRow, columnNum, columnCount, multiColumn);
        }

        // Column-wise counterpart of InsertHierarchyExpansion: same child-fetch/progress
        // steps, but hands off to InsertColumnsAndFillData instead of
        // InsertRowsAndFillData.
        private static async Task InsertHierarchyExpansionByColumn(
        string value, int rowNum, int startCol, int rowCount,
        bool multiRow, bool oneLevel)
        {
            Token.ThrowIfCancellationRequested();
            var children = await GetHierarchyChildrenAsync(value, oneLevel);
            if (children?.Count == 0) return;

            string title = AppState.Instance.DefaultSegment + " (" + value + ")";
            await MessageProgressWindowAsync($"Filling segment hierarchy values for segment {title}.");
            await Task.Yield();

            InsertColumnsAndFillData(children, rowNum, startCol, rowCount, multiRow);
        }

        private static async Task<List<string>> GetHierarchyChildrenAsync(
            string value, bool oneLevel)
        {
            try
            {
                var match = SegmentValues.FirstOrDefault(sv =>
                    sv.SegmentValue.Equals(value, StringComparison.OrdinalIgnoreCase));
                return await LoadHierarchySegmentValuesAsync(match, oneLevel);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.GetHierarchyChildrenAsync(value='{value}', oneLevel={oneLevel}): failed");
                return new List<string>();
            }
        }

        private static void InsertRowsAndFillData(
            List<string>? children, int startRow, int columnNum,
            int columnCount, bool multiColumn)
        {
            if (children == null || children.Count == 0) return;
            Token.ThrowIfCancellationRequested();
            var actRowFirstCol = HrWorksheet?.Cells[startRow, columnNum] as Excel.Range;
            var insertBlock = actRowFirstCol?.Offset[1, 0].Resize[children.Count, columnCount];

            insertBlock?.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Bulk fill first column
            var targetColRange = actRowFirstCol?.Offset[1, 0].Resize[children.Count, 1];
            if (targetColRange == null) return;
            targetColRange.NumberFormat = "@";
            object[,] data = new object[children.Count, 1];
            for (int j = 0; j < children.Count; j++)
                data[j, 0] = children[j];
            targetColRange.Value2 = data;
            Token.ThrowIfCancellationRequested();
            // Copy additional columns if multi-column
            if (multiColumn)
            {
                var copyRange = actRowFirstCol?.Offset[0, 1].Resize[1, columnCount - 1];
                var pasteRange = actRowFirstCol?.Offset[1, 1].Resize[children.Count, columnCount - 1];
                copyRange?.Copy(pasteRange);
                if (ExcelApp != null)
                    ExcelApp.CutCopyMode = (Excel.XlCutCopyMode)0;
            }
        }

        private static void InsertColumnsAndFillData(
            List<string>? children, int rowNum, int startCol,
            int rowCount, bool multiRow)
        {
            if (children == null || children.Count == 0) return;
            Token.ThrowIfCancellationRequested();
            var actColFirstRow = HrWorksheet?.Cells[rowNum, startCol] as Excel.Range;
            var insertBlock = actColFirstRow?.Offset[0, 1].Resize[rowCount, children.Count];

            insertBlock?.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

            // Bulk fill first row
            var targetRowRange = actColFirstRow?.Offset[0, 1].Resize[1, children.Count];
            if (targetRowRange == null) return;
            targetRowRange.NumberFormat = "@";
            object[,] data = new object[1, children.Count];
            for (int j = 0; j < children.Count; j++)
                data[0, j] = children[j];
            targetRowRange.Value2 = data;
            Token.ThrowIfCancellationRequested();
            // Copy additional rows if multi-row
            if (multiRow)
            {
                var copyRange = actColFirstRow?.Offset[1, 0].Resize[rowCount - 1, 1];
                var pasteRange = actColFirstRow?.Offset[1, 1].Resize[rowCount - 1, children.Count];
                copyRange?.Copy(pasteRange);
                if (ExcelApp != null)
                    ExcelApp.CutCopyMode = (Excel.XlCutCopyMode)0;
            }
        }

        // Renamed from GetInsertedRowCountAsync: orientation-agnostic (just returns how
        // many children were inserted, whether as rows or as columns), reused by both the
        // row-wise and column-wise ExpandSummaryAccountsAsync branches above.
        private static async Task<int> GetInsertedChildCountAsync(string value, bool oneLevel)
        {
            try
            {
                var match = SegmentValues.FirstOrDefault(sv =>
                    sv.SegmentValue.Equals(value, StringComparison.OrdinalIgnoreCase));

                var children = await LoadHierarchySegmentValuesAsync(match, oneLevel);
                return children?.Count ?? 0;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.GetInsertedChildCountAsync(value='{value}', oneLevel={oneLevel}): failed");
                return 0;
            }
        }

        private static async Task<List<string>> LoadHierarchySegmentValuesAsync(
                SegmentValueModel selectedHierarchy,
                bool oneLevel)
        {
            if (selectedHierarchy == null)
                return new List<string>();

            await MessageProgressWindowAsync("Fetching hierarchy data... ");

            if (DataRepository.SegmentValuesHierarchyExists(selectedHierarchy))
                return new List<string>();

            var hierarchyData = await HierarhyApiAsync(selectedHierarchy, Token);
            if (string.IsNullOrWhiteSpace(hierarchyData))
                return new List<string>();

            DataRepository.SaveHierarchyToCache(selectedHierarchy, hierarchyData);
            Token.ThrowIfCancellationRequested();
            var records = DeserializeRecords(hierarchyData);
            if (records == null || records.Length == 0)
                return new List<string>();

            var listValues = BuildListValues(records, oneLevel);

            return listValues;
        }
        private static object[]? DeserializeRecords(string hierarchyData)
        {
            try
            {
                using JsonDocument document = JsonDocument.Parse(hierarchyData);
                JsonElement root = document.RootElement;

                if (root.ValueKind != JsonValueKind.Object)
                    return null;

                if (!root.TryGetProperty("status", out JsonElement statusElem) ||
                    !string.Equals(statusElem.GetString(), AppConstants.Success, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                if (!root.TryGetProperty("records", out JsonElement recordsElem) ||
                    recordsElem.ValueKind != JsonValueKind.Array)
                {
                    return null;
                }

                // Convert JsonElement array to object[]
                var array = recordsElem.EnumerateArray();
                var list = new List<object>();
                foreach (JsonElement element in array)
                {
                    Token.ThrowIfCancellationRequested();

                    var obj = element.CloneToObject();
                    if (obj != null)
                    {
                        list.Add(obj); // Only add non-null objects to avoid CS8604
                    }
                }

                return list.ToArray();
            }
            catch (JsonException ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to parse hierarchy data");
                ServiceLocator.Logger?.LogRawJson("SegmentDiscoverer.DeserializeRecords", hierarchyData);
                return null;
            }
        }

        // Extension helper
        private static object? CloneToObject(this JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number => element.TryGetInt64(out long l) ? l : element.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null => null,
                JsonValueKind.Object => element.EnumerateObject()
                    .ToDictionary(p => p.Name, p => p.Value.CloneToObject()),
                JsonValueKind.Array => element.EnumerateArray()
                    .Select(e => e.CloneToObject()).ToArray(),
                _ => null
            };
        }
        private static List<string> BuildListValues(object[] records, bool oneLevel)
        {
            var listValues = new List<string>();

            foreach (var record in records)
            {
                if (record is not Dictionary<string, object> rec)
                    continue;

                int level = GetLevel(rec);
                string? segmentValue = GetSegmentValue(rec);

                if (string.IsNullOrEmpty(segmentValue))
                    continue;

                if (oneLevel)
                {
                    if (level == 1)
                        listValues.Add(" " + segmentValue);
                }
                else
                {
                    if (level >= 1)
                    {
                        string sp = new(' ', level);
                        listValues.Add(sp + segmentValue);
                    }
                }
            }

            return listValues;
        }

        private static int GetLevel(Dictionary<string, object> rec)
        {
            if (!rec.TryGetValue("lvl", out var lvlObj))
                return 0;

            try
            {
                return Convert.ToInt32(lvlObj);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.GetLevel: failed converting 'lvl' value '{lvlObj}' to int - defaulting to 0");
                return 0;
            }
        }

        private static string? GetSegmentValue(Dictionary<string, object> rec)
        {
            return rec.TryGetValue("segmentValue", out var segValObj)
                ? Convert.ToString(segValObj)
                : null;
        }
        private static async Task<string> HierarhyApiAsync(
    SegmentValueModel selectedHierarchy,
    CancellationToken token)
        {
            try
            {
                string apiUrl =
                    $"{AppState.Instance.LoginUrl}/rest/secure/finance/segment-hierarchy" +
                    $"?segmentValueSetId={selectedHierarchy.SegmentValueSetId}" +
                    $"&parentSegmentValue={WebUtility.UrlEncode(selectedHierarchy.SegmentValue.Trim())}" +
                    $"&cubeId={selectedHierarchy.CubeId}";

                token.ThrowIfCancellationRequested();

                string response =
                    await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", token);

                token.ThrowIfCancellationRequested();

                ValidateTransportResponse(response);

                var result =
                    ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn($"Hierarchy API failed: {apiUrl}");
                    ServiceLocator.Logger?.LogWarn($"Response: {response}");

                    await ShowWarnMessage(result.ErrorMessage ??
                                          "Hierarchy API returned failure status.");

                    return string.Empty;
                }

                return response;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn(
                    "User cancelled operation. Fetching hierarchy data interrupted.");
                return string.Empty;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.HierarhyApiAsync(segmentValue='{selectedHierarchy?.SegmentValue}'): failed");
                return string.Empty;
            }
        }
        private static void ValidateTransportResponse(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
                throw new InvalidOperationException("Empty API response.");

            if (response.IndexOf("(401)", StringComparison.OrdinalIgnoreCase) >=0)
                throw new UnauthorizedAccessException("Session expired.");

            if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >=0)
                throw new InvalidOperationException(response);
        }

        private static bool SegmentValueExists(string sValue)
        {
            try
            {
                if (sValue == null || string.IsNullOrWhiteSpace(sValue.Trim()))
                    return false;

                sValue = sValue.Replace("--", "").Replace("~", "");

                var match = SegmentValues.FirstOrDefault(sv => sv.SegmentValue.Equals(sValue, StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.SegmentValueExists(sValue='{sValue}'): failed");
            }
            return false;
        }

        private static bool IsSummaryAccount(string sValue)
        {
            try
            {
                if (sValue == null || string.IsNullOrWhiteSpace(sValue.Trim()))
                    return false;

                sValue = sValue.Replace("--", "").Replace("~", "");

                var match = SegmentValues.FirstOrDefault(sv =>
                            sv.SegmentName.Equals(AppState.Instance.DefaultSegment, StringComparison.OrdinalIgnoreCase) &&
                            sv.SegmentValue.Equals(sValue, StringComparison.OrdinalIgnoreCase) &&
                            !sv.SummaryFlag.Equals("RG", StringComparison.OrdinalIgnoreCase));

                if (match != null && match.SummaryFlag == "Y")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.IsSummaryAccount(sValue='{sValue}'): failed");
            }
            return false;
        }
        private static async Task ExplodeSegment()
        {
            bool oneLevel = Action == Explode1Level;
            Token.ThrowIfCancellationRequested();
            try
            {
                await MessageProgressWindowAsync("Exploding segment value.");

                var activeRange = await GetAndValidateSelectionAsync();
                if (activeRange == null) return;

                string segmentValue = GetAndNormalizeSegmentValue(activeRange);
                if (!await ValidateSummaryAccountAsync(segmentValue)) return;

                var segList = await LoadSegmentsAsync(segmentValue, oneLevel);
                if (segList == null || segList.Count == 0)
                {
                    await ShowErrorMessage($"No child items exists for the selected segment value \"{segmentValue}\"");
                    return;
                }

                if (!EnsureSheetLimit(segList.Count)) return;

                await CreateSheetsForSegmentsAsync(segmentValue, segList, activeRange);
            }
            catch (Exception ex)
            {
                await LogAndShowException(ex);
            }
        }
        private static async Task<Excel.Range?> GetAndValidateSelectionAsync()
        {
            if (ExcelApp?.Selection is not Excel.Range selection || selection.Columns.Count > 1 || selection.Rows.Count > 1)
            {
                await ShowErrorMessage("Invalid selection for explode option.\r\n" +
                     "Selection should be a single cell with the parent/summary account value in it.");
                return null;
            }


            if (HrWorksheet == null || selection.Cells[1, 1] is not Excel.Range activeRange)
            {
                await ShowErrorMessage("Either worksheet or current range is null.");
                return null;
            }

            return activeRange;
        }

        private static string GetAndNormalizeSegmentValue(Excel.Range activeRange)
        {
            string segmentValue = activeRange.Value2?.ToString() ?? string.Empty;
            return segmentValue.Replace("--", "").Replace("~", "").Trim();
        }

        private static async Task<bool> ValidateSummaryAccountAsync(string segmentValue)
        {
            if (IsSummaryAccount(segmentValue))
                return true;

            await ShowErrorMessage("Selected value should be a parent/summary account.");
            return false;
        }

        private static async Task<List<string>> LoadSegmentsAsync(
            string segmentValue,
            bool oneLevel)
        {
            Token.ThrowIfCancellationRequested();
            try
            {
                var match = SegmentValues
                    .FirstOrDefault(sv => sv.SegmentValue.Equals(segmentValue, StringComparison.OrdinalIgnoreCase));

                return await LoadHierarchySegmentValuesAsync(match, oneLevel);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.LoadSegmentsAsync(segmentValue='{segmentValue}', oneLevel={oneLevel}): failed");
                await ShowErrorMessage("An exception encountered!" + Environment.NewLine + ex.Message);
                return new List<string>();
            }
        }

        private static bool EnsureSheetLimit(int toAdd)
        {
            int existingCount = (HrWorbook?.Worksheets.Count) ?? 0;
            if (existingCount + toAdd <= MaxSheetsThreshold)
                return true;

            _ = ShowErrorMessage($"Cannot create {toAdd} sheets: workbook would exceed the safety threshold of {MaxSheetsThreshold} worksheets.");
            return false;
        }

        private static async Task CreateSheetsForSegmentsAsync(
            string segmentValue,
            List<string> segList,
            Excel.Range activeRange)
        {
            await MessageProgressWindowAsync("Creating worksheets.");

            for (int i = 0; i < segList.Count; i++)
            {
                Token.ThrowIfCancellationRequested();
                string rawName = $"{segmentValue}({segList[i].Trim()})";
                string sanitizedName = CommonFunctions.SanitizeSheetName(rawName);

                await MessageProgressWindowAsync(
                $"Creating worksheets {i + 1} of {segList.Count}, worksheet name = \"({sanitizedName})\".");
                await Task.Yield();
                await CreateSingleSheetAsync(i, segList[i], sanitizedName, activeRange);
            }

            await MessageProgressWindowAsync("Creating worksheets task completed.");
        }

        private static async Task CreateSingleSheetAsync(
            int index,
            string segmentValue,
            string sanitizedName,
            Excel.Range activeRange)
        {

            Token.ThrowIfCancellationRequested();

            var dispatcher = System.Windows.Application.Current?.Dispatcher
                     ?? System.Windows.Threading.Dispatcher.CurrentDispatcher;

            await dispatcher.InvokeAsync(() =>
            {
                if (SheetExists(HrWorbook, sanitizedName))
                {
                    Excel.Worksheet? existing = GetWorksheetByName(HrWorbook, sanitizedName);
                    existing?.Delete();
                }

                int afterIndex = (HrWorksheet?.Index ?? 1) + index;
                int worksheetCount = HrWorbook?.Worksheets.Count ?? 0;
                afterIndex = Math.Min(afterIndex, worksheetCount);
                Excel.Worksheet? afterSheet = HrWorbook?.Worksheets[afterIndex] as Excel.Worksheet;

                HrWorksheet?.Copy(Type.Missing, afterSheet);

                if (ExcelApp?.ActiveSheet is not Excel.Worksheet newSheet) return;

                try
                {
                    newSheet.Name = sanitizedName;
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.CreateSingleSheetAsync: failed setting sheet name to '{sanitizedName}' - retrying with a re-sanitized name");
                    string uniqueName = CommonFunctions.SanitizeSheetName(sanitizedName);
                    newSheet.Name = uniqueName;
                }

                Excel.Range? targetCell = newSheet.Range[activeRange.Address];
                targetCell.NumberFormat = "@";
                targetCell.Value2 = segmentValue.Trim();
            });
        }

        private static async Task LogAndShowException(Exception ex)
        {
            ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.ExplodeSegment: failed");
            await ShowErrorMessage("An exception encountered!" + Environment.NewLine + ex.Message);
        }


        //Standard helpers

        private static Excel.Worksheet? GetWorksheetByName(Excel.Workbook? wb, string sheetName)
        {
            if (wb == null || string.IsNullOrEmpty(sheetName)) return null;
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return ws;

                ExcelComHelper.SafeRelease(ws, "Worksheet");
            }
            return null;
        }

        private static bool SheetExists(Excel.Workbook? wb, string sheetName)
        {
            if (wb == null || string.IsNullOrEmpty(sheetName)) return false;
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return true;

                ExcelComHelper.SafeRelease(ws, "Worksheet");
            }
            return false;
        }
        private static SegmentValueModel SegValueModel(string sValue)
        {
            if (sValue == null || string.IsNullOrWhiteSpace(sValue.Trim()))
                return new SegmentValueModel();

            sValue = sValue.Replace("--", "").Replace("~", "");

            return SegmentValues.FirstOrDefault(sv =>
                sv.SegmentName.Equals(AppState.Instance.DefaultSegment, StringComparison.OrdinalIgnoreCase) &&
                sv.SegmentValue.Equals(sValue, StringComparison.OrdinalIgnoreCase) &&
                !sv.SummaryFlag.Equals("RG", StringComparison.OrdinalIgnoreCase));
        }
        private static async Task ShowErrorMessage(string message)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }
        private static async Task ShowWarnMessage(string message)
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                message,
                MessageBoxImage.Warning,
                MessageBoxButton.OK);
        }
        private static async Task ShowCancelledAsync()
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                "Operation cancelled!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private static async Task HandleUnexpectedErrorAsync(Exception ex)
        {
            ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.FillSegmentProperties: unexpected error");
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                $"An unexpected error occurred.{Environment.NewLine}{ex.Message}",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }
        private static GLWaitWindow? CreateAndShowProgressWindow(CancellationHelper cts)
        {
            GLWaitWindow? win = null;
            try
            {
                // WpfAppManager.InvokeOnWpfThread only takes an Action (no return value),
                // so capture the created window from inside the delegate - same pattern
                // AddinEntry.LedgerChanged uses for GLWaitWindow.
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        // Use the passed-in cts, don't create a new one
                        win = new GLWaitWindow(cts);
                        win.Show();
                        win.StartMonitoring();
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.CreateAndShowProgressWindow: failed on WPF thread");
                        win = null;
                    }
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.CreateAndShowProgressWindow: InvokeOnWpfThread failed");
            }
            return win;
        }

        private static Task InitializeProgressWindowAsync()
        {
            // Basic guards
            if (Win == null || Win.Dispatcher == null)
                return Task.CompletedTask;

            try
            {
                // Fire-and-forget: this is invoked from contexts that may run on a
                // thread with no captured SynchronizationContext, so awaiting the
                // dispatch would risk resuming subsequent Excel COM calls on an
                // arbitrary ThreadPool thread instead of the calling thread.
                _ = Win.Dispatcher.InvokeAsync(
                    () =>
                    {
                        Win.SetProcessTitle(Title);
                        Win.SetProcessMessage(Msg);
                    },
                    DispatcherPriority.Normal);

                return Task.CompletedTask;
            }
            catch (TaskCanceledException)
            {
                // Dispatcher shutting down; nothing to do
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                // Last-resort logging if something unexpected happens
                ServiceLocator.Logger?.LogException(ex, "SegmentDiscoverer.InitializeProgressWindowAsync: failed");
                return Task.CompletedTask;
            }
        }
        private static Task MessageProgressWindowAsync(string message)
        {
            // Basic guards
            if (Win == null || Win.Dispatcher == null)
                return Task.CompletedTask;

            try
            {
                // Fire-and-forget: do not await the dispatcher operation itself.
                // Awaiting here would introduce a suspend point that can let the
                // caller resume on a different thread (e.g. a background worker
                // with no captured SynchronizationContext), which is unsafe when
                // the code right after the await touches Excel COM objects.
                _ = Win.Dispatcher.InvokeAsync(
                    () =>
                    {
                        Win.SetProcessMessage(message);
                    },
                    DispatcherPriority.Normal);

                return Task.CompletedTask;
            }
            catch (TaskCanceledException)
            {
                // Dispatcher shutting down; nothing to do
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                // Last-resort logging if something unexpected happens
                ServiceLocator.Logger?.LogException(ex, $"SegmentDiscoverer.MessageProgressWindowAsync(message='{message}'): failed");
                return Task.CompletedTask;
            }
        }
        private static async Task SafelyCloseWindowAsync()
        {
            if (Win == null)
                return;

            try
            {
                if (Win.Dispatcher.CheckAccess())  // Already on UI thread
                {
                    Win.RequestClose();
                }
                else
                {
                    await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error closing wait window: {ex.Message}");
            }
            finally
            {
                ExcelWindowHelper.ActivateExcelMainWindow(ServiceLocator.ExcelApp);
                Win = null;
            }
        }
        private static ObservableCollection<SegmentValueModel> LoadSegmentValues(string ledgerName)
        {
            try
            {
                var task = Task.Run(() =>
                {
                    var dataService = DataServiceLocator.SegmentDataService;
                    return dataService.GetSegmentValues(ledgerName);
                });

                if (task.Wait(TimeSpan.FromSeconds(180)))
                {
                    return task.Result;
                }
                else
                {
                    throw new TimeoutException("Timeout loading segment values from service");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to load segment values");
                return new ObservableCollection<SegmentValueModel>();
            }
        }
    }
#nullable disable
}
