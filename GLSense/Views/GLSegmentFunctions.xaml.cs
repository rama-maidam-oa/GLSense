using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using GLSense.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Views
{
    /// <summary>
    /// Interaction logic for GLSegmentFunctions.xaml
    /// </summary>
    public partial class GLSegmentFunctions : DpiAwareWindow, IWarningHost
    {
        private string _formulaName;
        private sealed class LedgerInfo
        {
            public List<string> FuncArgs { get; set; }
            public List<string> FuncValues { get; set; }
            public string LedgerName { get; set; }
            public long LedgerId { get; set; }
            public long CoaId { get; set; }
            public dynamic LedgerRecord { get; set; }
        }

        private readonly GLSegmentFuncsViewModel vm;
        public GLSegmentFunctions(string funcName)
        {
            LogUtility.LogDebug($"GLSegmentFunctions.ctor invoked - funcName={funcName}");
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);
            _formulaName= funcName;
            switch (funcName)
            {
                case "ENABLEDFLAG":
                    this.TxtFuncName.Text = "Get Segment Enabled Flag";
                    ParentChildSection.Visibility = Visibility.Collapsed;
                    AttributesSection.Visibility = Visibility.Collapsed;
                    IncludeValuesSection.Visibility = Visibility.Collapsed;
                    break;
                case "SUMMARYFLAG":
                    this.TxtFuncName.Text = "Get Segment Summary Flag";
                    ParentChildSection.Visibility = Visibility.Collapsed;
                    AttributesSection.Visibility = Visibility.Collapsed;
                    IncludeValuesSection.Visibility = Visibility.Collapsed;
                    break;
                case "DESCRIPTION":
                    this.TxtFuncName.Text = "Get Segment Description";
                    ParentChildSection.Visibility = Visibility.Collapsed;
                    AttributesSection.Visibility = Visibility.Collapsed;
                    break;
                case "NEXTSEGMENT":
                    this.TxtFuncName.Text = "Get Next Segment";
                    AttributesSection.Visibility = Visibility.Collapsed;
                    IncludeValuesSection.Visibility = Visibility.Collapsed;
                    break;
                case "PREVIOUSSEGMENT":
                    this.TxtFuncName.Text = "Get Previous Segment";
                    AttributesSection.Visibility = Visibility.Collapsed;
                    IncludeValuesSection.Visibility = Visibility.Collapsed;
                    ChkIncludeParent.Content = "Previous Parent";
                    ChkIncludeChild.Content = "Previous Child";
                    break;
                case "DFF":
                    this.TxtFuncName.Text = "Get Segment Descriptive Flex Field";
                    ParentChildSection.Visibility = Visibility.Collapsed;
                    IncludeValuesSection.Visibility = Visibility.Collapsed;
                    break;
                case "ACCOUNTTYPE":
                    this.TxtFuncName.Text = "Get Segment Account Type";
                    ParentChildSection.Visibility = Visibility.Collapsed;
                    AttributesSection.Visibility = Visibility.Collapsed;
                    IncludeValuesSection.Visibility = Visibility.Collapsed;
                    break;
            }

            vm = new GLSegmentFuncsViewModel(Dispatcher, funcName)
            {
                ExcelApp = AppState.Instance.ExcelApp.Application, // Pass the Excel application instance to the ViewModel
                ShowWarningAction = (msg) => Dispatcher.Invoke(() => AppOverlayControl.ShowWarning(msg)),
                ShowWarningAsyncAction = async (msg) => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.ShowWarningAsync(msg)),
                ShowBusyAction = async (txt, cancel) =>
                        await Dispatcher.InvokeAsync(async () =>
                            await AppOverlayControl.ShowBusyasynTask(txt, cancel)),
                HideBusyAsyncAction = async () => await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync())
            };
            DataContext = vm;

            Loaded += Window_Loaded;
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentFunctions.Window_Loaded invoked");
            var (Address, rng) = GetActiveCellInfo();
            GlobalStateViewModel.Instance.ReferenceText = Address;
            LogUtility.LogDebug($"GLSegmentFunctions.Window_Loaded: active cell address={Address}");

            if (!HasValidCubeAndLedger())
            {
                LogUtility.LogDebug("GLSegmentFunctions.Window_Loaded: validation failed - no valid cube/ledger selected, aborting load");
                return;
            }

            var ledgerInfo = await ProcessFormulaAsync(rng);

            string ledgerName = string.IsNullOrWhiteSpace(ledgerInfo.LedgerName)
                    ? AppState.Instance.SelectedLedger.LedgerName
                    : ledgerInfo.LedgerName;

            LogUtility.LogDebug($"GLSegmentFunctions.Window_Loaded: loading data - cubeId={AppState.Instance.SelectedCube.CubeId}, ledgerId={ledgerInfo.LedgerId}, coaId={ledgerInfo.CoaId}, ledgerName={ledgerName}");
            await vm.LoadDataAsync(AppState.Instance.SelectedCube.CubeId, ledgerInfo.LedgerId,ledgerInfo.CoaId, ledgerInfo.FuncArgs, ledgerInfo.FuncValues);
            await Dispatcher.InvokeAsync(() => cmbLedgers.Text = ledgerName);
            if (_formulaName == "ACCOUNTTYPE" || _formulaName == "DFF")
            {
                if (rng != null && rng.Value2 != null)
                {
                    await Dispatcher.InvokeAsync(() => txtResult.Text = rng.Value2.ToString() ?? string.Empty);
                }
            }
            LogUtility.LogDebug("GLSegmentFunctions.Window_Loaded completed");
        }
        private async Task<LedgerInfo> ProcessFormulaAsync(Excel.Range rng)
        {
            LogUtility.LogDebug("GLSegmentFunctions.ProcessFormulaAsync invoked");
            if (!IsSegmentFormula(rng))
            {
                LogUtility.LogDebug("GLSegmentFunctions.ProcessFormulaAsync: cell does not contain a recognized segment formula - using selected ledger defaults");
                return new LedgerInfo
                {
                    FuncArgs = new List<string>(),
                    FuncValues = new List<string>(),
                    LedgerName = AppState.Instance.SelectedLedger.LedgerName,
                    LedgerId = AppState.Instance.SelectedLedger.LedgerId,
                    CoaId = AppState.Instance.SelectedLedger.CoaId,
                    LedgerRecord = null
                };
            }

            var ledgerInfo = ValidateLedgerFromFormula(rng);
            if (ledgerInfo.LedgerRecord == null)
            {
                LogUtility.LogDebug($"GLSegmentFunctions.ProcessFormulaAsync: ledger from formula not found - ledgerName={ledgerInfo.LedgerName}");
                await ShowLedgerNotFoundWarningAsync(ledgerInfo.LedgerName);
                return new LedgerInfo
                {
                    FuncArgs = new List<string>(),
                    FuncValues = new List<string>(),
                    LedgerName = AppState.Instance.SelectedLedger.LedgerName,
                    LedgerId = AppState.Instance.SelectedLedger.LedgerId,
                    CoaId = AppState.Instance.SelectedLedger.CoaId,
                    LedgerRecord = null
                };
            }

            await LoadLedgerDataIfNeededAsync(ledgerInfo.LedgerRecord);
            return ledgerInfo;
        }
        private async Task ShowLedgerNotFoundWarningAsync(string ledgerName)
        {
            string message = $"Ledger \"{ledgerName}\" in the formula does not exist in the selected cube!" +
                            Environment.NewLine + "Setting default values.";
            await AppOverlayControl.ShowWarningAsync(message);
        }
        private async Task LoadLedgerDataIfNeededAsync(dynamic ledgerRecord)
        {
            LogUtility.LogDebug("GLSegmentFunctions.LoadLedgerDataIfNeededAsync invoked");
            var segmentCount = DataRepository.GetTableItemsCount(
                AppState.Instance.SelectedCube.CubeId,
                ledgerRecord.LedgerId,
                "SEGMENTS");
            LogUtility.LogDebug($"GLSegmentFunctions.LoadLedgerDataIfNeededAsync: segmentCount={segmentCount}");

            if (segmentCount == 0)
            {
                LogUtility.LogDebug("GLSegmentFunctions.LoadLedgerDataIfNeededAsync: no segments cached, loading with progress");
                await LoadLedgerSegmentsWithProgressAsync(ledgerRecord);
            }
        }

        private async Task LoadLedgerSegmentsWithProgressAsync(dynamic ledgerRecord)
        {
            LogUtility.LogDebug($"GLSegmentFunctions.LoadLedgerSegmentsWithProgressAsync invoked - ledgerId={ledgerRecord.LedgerId}, cubeId={AppState.Instance.SelectedCube.CubeId}");
            var ctsHelper = new CancellationHelper();
            var token = ctsHelper.GetToken();

            try
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    AppOverlayControl.ShowBusyasyn(
                        "Please wait while we fetch the ledger data...  (click Cancel to stop)",
                        async () => {
                            if (!ctsHelper.IsCancellationRequested)
                            {
                                ctsHelper.Cancel();
                                LogUtility.LogWarn("Loading cancelled by user");
                            }
                            await Task.CompletedTask;
                        });
                    return Task.CompletedTask;
                });

                LogUtility.LogDebug("GLSegmentFunctions.LoadLedgerSegmentsWithProgressAsync: calling CommonFunctions.FillResponsibilitiesAsync");
                await CommonFunctions.FillResponsibilitiesAsync(
                    ledgerRecord.LedgerId,
                    AppState.Instance.SelectedCube.CubeId,
                    token);
                LogUtility.LogDebug("GLSegmentFunctions.LoadLedgerSegmentsWithProgressAsync: CommonFunctions.FillResponsibilitiesAsync completed successfully");
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Loading ledger segments operation was cancelled.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentFunctions.LoadLedgerSegmentsWithProgressAsync");
            }
            finally
            {
                try
                {
                    if (ctsHelper != null && !ctsHelper.IsCancellationRequested)
                        ctsHelper.Cancel();

                    ctsHelper?.Dispose();
                }
                catch (Exception ex)
                {
                    // Swallow dispose exceptions (Excel COM weirdness) but still log for diagnostics.
                    LogUtility.LogWarn($"GLSegmentFunctions.LoadLedgerSegmentsWithProgressAsync: exception disposing CancellationHelper (ignored): {ex.Message}");
                }
                await Dispatcher.InvokeAsync(async () => await AppOverlayControl.HideBusyAsync());
                ctsHelper.Dispose();
            }
        }
        private static bool IsSegmentFormula(Excel.Range rng)
        {
            if (!(bool)rng.HasFormula) return false;
            return rng.Formula.ToString().IndexOf("GLSense_GetSegmentEnabledFlag", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rng.Formula.ToString().IndexOf("GLSense_GetSegmentSummaryFlag", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rng.Formula.ToString().IndexOf("GLSense_GetSegmentDesc", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rng.Formula.ToString().IndexOf("GLSense_GetNextSegment", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rng.Formula.ToString().IndexOf("GLSense_GetPreviousSegment", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rng.Formula.ToString().IndexOf("GLSense_GetAccountType", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   rng.Formula.ToString().IndexOf("GLSense_GetSegmentDFF", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        private LedgerInfo ValidateLedgerFromFormula(Excel.Range rng)
        {
            LogUtility.LogDebug("GLSegmentFunctions.ValidateLedgerFromFormula invoked");
            string formula = rng.Formula.ToString();
            var funcArgs = CommonFunctions.FormulaParameters(formula);
            var funcValues = CommonFunctions.FormulaValues(formula);

            // Defensive defaults
            if (funcArgs == null) funcArgs = new List<string>();
            if (funcValues == null) funcValues = new List<string>();

            // Determine whether the last parameter is actually a ledger name.
            // New formulas include ledger as the last argument; older formulas do not.
            string ledgerName = string.Empty;
            dynamic ledgerRecord = null;

            if (funcValues.Count > 0)
            {
                string candidate = GetLedgerName(funcValues);
                if (!string.IsNullOrEmpty(candidate))
                {
                    ledgerRecord = AppState.Instance.SelectedCube.Ledgers
                        .FirstOrDefault(x => x.LedgerName == candidate);

                    if (ledgerRecord != null)
                    {
                        // Formula contains ledger explicitly
                        ledgerName = candidate;
                        LogUtility.LogDebug($"GLSegmentFunctions.ValidateLedgerFromFormula: ledger matched from formula - ledgerName={ledgerName}");
                    }
                }
            }

            if (ledgerRecord == null)
            {
                // Ledger not present in formula -> use default ledger for processing.
                var def = DefaultLedgerRecord();
                // Ensure lists are non-null
                // Append synthetic ledger argument so downstream code indexes remain valid
                string syntheticLedgerArg = def != null ? $"\"{def.LedgerName}\"" : "\"\"";
                funcArgs.Add(syntheticLedgerArg);
                funcValues.Add(syntheticLedgerArg);
                LogUtility.LogDebug($"GLSegmentFunctions.ValidateLedgerFromFormula: no explicit ledger matched in formula - falling back to default ledger (found={def != null})");

                return new LedgerInfo
                {
                    FuncArgs = funcArgs,
                    FuncValues = funcValues,
                    LedgerName = string.Empty, // signal to caller to use selected ledger name
                    LedgerId = def != null ? GetId(def) : AppState.Instance.SelectedLedger.LedgerId,
                    CoaId = def != null ? GetCoa(def) : AppState.Instance.SelectedLedger.CoaId,
                    LedgerRecord = def
                };
            }

            // Ledger was present and matched
            return new LedgerInfo
            {
                FuncArgs = funcArgs,
                FuncValues = funcValues,
                LedgerName = ledgerName,
                LedgerId = GetId(ledgerRecord),
                CoaId = GetCoa(ledgerRecord),
                LedgerRecord = ledgerRecord
            };
        }
        private string GetLedgerName(List<string> funcValues)
        {
            string ledgerName = AppState.Instance.SelectedLedger.LedgerName;

            if (string.IsNullOrWhiteSpace(_formulaName) || funcValues == null)
                return ledgerName;

            int expectedArgCount = _formulaName switch
            {
                "ENABLEDFLAG" => 3,
                "SUMMARYFLAG" => 3,
                "ACCOUNTTYPE" => 3,
                "DESCRIPTION" => 4,
                "NEXTSEGMENT" => 5,
                "PREVIOUSSEGMENT" => 5,
                "DFF" => 4,
                _ => 0
            };

            // If formula contains the new ledger parameter
            if (expectedArgCount > 0 &&
                funcValues.Count == expectedArgCount)
            {
                string lastArg = funcValues.LastOrDefault();

                if (!string.IsNullOrWhiteSpace(lastArg))
                {
                    ledgerName = lastArg.Replace("\"", "");
                }
            }

            return ledgerName.Trim();
        }
        // Helper to read LedgerId property (handles different model types/casing)
        private static long GetId(dynamic rec)
        {
            if (rec == null) return 0;
            try
            {
                // Try common property names
                var type = rec.GetType();
                var p = type.GetProperty("LedgerId") ?? type.GetProperty("ledgerId");
                if (p != null) return Convert.ToInt64(p.GetValue(rec));
            }
            catch { /* Ignore exceptions */ }
            try { return Convert.ToInt64(rec.LedgerId); } catch { /* Ignore exceptions */ }
            try { return Convert.ToInt64(rec.ledgerId); } catch { /* Ignore exceptions */ }
            return 0;
        }

        // Helper to read CoaId/Coaid property
        private static long GetCoa(dynamic rec)
        {
            if (rec == null) return 0;
            try
            {
                var type = rec.GetType();
                var p = type.GetProperty("Coaid") ?? type.GetProperty("coaId");
                if (p != null) return Convert.ToInt64(p.GetValue(rec));
            }
            catch { /* Ignore exceptions */ }
            try { return Convert.ToInt64(rec.Coaid); } catch { /* Ignore exceptions */ }
            try { return Convert.ToInt64(rec.coaId); } catch { /* Ignore exceptions */ }
            return 0;
        }
        private static LedgerRecord DefaultLedgerRecord()
        {
            return AppState.Instance.SelectedCube.Ledgers
                .FirstOrDefault(x => x.LedgerId == AppState.Instance.SelectedLedger.LedgerId);
        }
        private static (string Address, Excel.Range Range) GetActiveCellInfo()
        {
            var rng = AppState.Instance.ExcelApp.ActiveCell;
            string sheetName = ((Excel.Worksheet)rng.Parent).Name;
            string cellAddress = rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            return ($"'{sheetName}'!{cellAddress}", rng);
        }
        private static bool HasValidCubeAndLedger()
        {
            return AppState.Instance.SelectedCube != null && AppState.Instance.SelectedLedger != null;
        }
        public void CellSelectionWarning(string message)
        {
            LogUtility.LogDebug($"GLSegmentFunctions.CellSelectionWarning invoked - message={message}");
            try
            {
                AppOverlayControl.ShowWarning(message);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentFunctions.CellSelectionWarning");
            }
        }
        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentFunctions.BtnSubmit_Click invoked");
            if (!ValidateCellReference())
            {
                LogUtility.LogDebug("GLSegmentFunctions.BtnSubmit_Click: validation failed - no cell reference selected");
                return;
            }

            try
            {
                var rng = ProcessCellReference();
                if (rng == null)
                {
                    LogUtility.LogDebug("GLSegmentFunctions.BtnSubmit_Click: cell reference did not resolve to a valid range, aborting");
                    return;
                }

                FormatAndWriteCell(rng);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentFunctions.BtnSubmit_Click");
            }
        }

        private bool ValidateCellReference()
        {
            if (string.IsNullOrWhiteSpace(CellReference.Text))
            {
                LogUtility.LogDebug("GLSegmentFunctions.ValidateCellReference: validation failed - CellReference.Text is empty");
                AppOverlayControl.ShowWarning("Please select a cell reference for entering formula.");
                return false;
            }
            return true;
        }

        private Excel.Range ProcessCellReference()
        {
            LogUtility.LogDebug($"GLSegmentFunctions.ProcessCellReference invoked - cellReference={CellReference.Text}");
            var cleanedRef = CommonFunctions.RemoveInDirect(CellReference.Text);
            if (cleanedRef == null)
            {
                LogUtility.LogDebug("GLSegmentFunctions.ProcessCellReference: validation failed - cell reference does not resolve to a valid cell");
                AppOverlayControl.ShowWarning("The specified cell reference does not refer to a valid cell in the current workbook.");
            }
            return cleanedRef;
        }

        private void FormatAndWriteCell(Excel.Range rng)
        {
            LogUtility.LogDebug("GLSegmentFunctions.FormatAndWriteCell invoked");
            rng.NumberFormat = AppConstants.General;

            if (vm.WriteFormulaToCell(rng))
            {
                LogUtility.LogDebug("GLSegmentFunctions.FormatAndWriteCell: formula written successfully, closing window");
                Close();
            }
            else
            {
                // Regression fix: WriteFormulaToCell already raises a specific warning via
                // ShowWarningAction for every failure path it has - missing-mandatory-field
                // messages from ValidateMandatoryFields (e.g. "Segment name is mandatory.")
                // as well as the actual-exception message from its catch block. This
                // generic "Failed to write formula to cell." call used to run
                // unconditionally right after, immediately overwriting/replacing that
                // specific message in the shared AppOverlayControl toast with a useless
                // generic one - so the user only ever saw the generic message no matter
                // what actually went wrong. Do not show a second, generic message here;
                // the ViewModel's own message is the one that should reach the user.
                LogUtility.LogWarn("GLSegmentFunctions.FormatAndWriteCell: failed to write formula to cell");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            LogUtility.LogDebug("GLSegmentFunctions.BtnClose_Click invoked");
            Close();
        }
    }
}

