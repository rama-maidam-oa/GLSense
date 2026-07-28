// DD_EP.cs in GLSense.Addin.Core
// Port of GLSense\Drilldowns\DD_ExcelPrecedents.cs (FinalWorkingCode), class
// DrilldownXlPrecedents. File named DD_EP.cs to match this project's sibling
// Drilldowns\DD_BL.cs/DD_JL.cs/DD_SL.cs naming convention ("DD_" + a two-letter
// drilldown-type code; "EP" = Excel Precedents).
//
// This is the CONTENT side of the double-click ("precedent drilldown") feature - the
// host-side SheetBeforeDoubleClick classification/dispatch is a separate, later pass (see
// AddinEntry.cs OnExcelEvent's "SheetBeforeDoubleClick" TODO comment, which references
// this exact class by name). That later pass is expected to call this class exactly like
// AddinEntry.cs already calls DrilldownBl/DrilldownJl/DrilldownSl for the ribbon-driven
// drilldowns:
//     var runProcess = new DrilldownXlPrecedents(external);
//     await runProcess.ProcessEPDrilldown();
// where `external` is a fully-qualified external address string (e.g.
// "[Book1.xlsx]Sheet1!$A$1", via Excel.Range.Address[External:=true] /
// ExcelExternalRef.BuildExternalAddress) built by the host from the live Range the
// double-click event handed it.
//
// Re-pointed vs. the original (business logic/precedent-walking algorithm unchanged):
//   - namespace GLSense.Drilldowns -> GLSense.Addin.Core.Drilldowns; GLSense.Helpers/
//     .Models/.Utilities/.Views -> GLSense.Addin.Core.* equivalents.
//   - Constructor: DrilldownXlPrecedents(Excel.Application xlapp, string rngAddress) ->
//     DrilldownXlPrecedents(string rngAddress) - the Excel.Application parameter is
//     dropped entirely; ExcelApp is now a property that always reads
//     ServiceLocator.ExcelApp (this project's AppState has no ExcelApp field).
//   - LogUtility.* (static) -> ServiceLocator.Logger?.* (instance via context).
//   - System.Windows.Forms MessageBoxIcon/MessageBoxButtons -> System.Windows
//     MessageBoxImage/MessageBoxButton (this project has no WinForms reference; the enum
//     member names used here - Exclamation/Error/OK - exist under both).
//   - GLWaitWindow now derives from BaseWindow: no more Win.SetExcelOwner((IntPtr)
//     ExcelApp.Hwnd) call - BaseWindow sets the Excel owner automatically via
//     ServiceLocator.ExcelHandle/ModalToExcel.
//   - AddinModule.RibbonHelper.ApplyState("ApplySheetActiveState") ->
//     ServiceLocator.RibbonController?.SetState("ApplySheetActiveState") (see
//     DDDatatoWorksheet.cs/GLCubeDetails.xaml.cs/GLLogin.xaml.cs for the same mapping
//     already established elsewhere in this project).
//   - ClsFormulaParser (Helpers), AppConstants.glBal, GLWaitWindow, CancellationHelper,
//     ExcelExternalRef.ResolveRangeWithContext, CommonFunctions.GetBalancesCountInCells/
//     MultiFormulaValues are all already ported - used as-is, no changes needed.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Drilldowns
{
    public class DrilldownXlPrecedents
    {
        private static Excel.Application ExcelApp => ServiceLocator.ExcelApp;
        private Excel.Workbook EpWorbook { get; set; }
        private Excel.Worksheet EpWorksheet { get; set; }
        private Excel.Range EpRange { get; set; }
        private CancellationHelper _ctsHelper;
        private CancellationToken Token => _ctsHelper?.GetToken() ?? default;
        private ExternalResolveResult ExternalResolveResult { get; set; } = new ExternalResolveResult();
        private readonly string _epAddress;
        private static GLWaitWindow Win { get; set; }

        public DrilldownXlPrecedents(string rngAddress)
        {
            _epAddress = rngAddress;
        }

        public async Task ProcessEPDrilldown()
        {
            _ctsHelper = new CancellationHelper();

            ServiceLocator.Logger?.LogDebug($"DrilldownXlPrecedents.ProcessEPDrilldown started. address={_epAddress}");

            try
            {
                await ExecuteDrilldownProcess();
            }
            catch (OperationCanceledException)
            {
                await HandleCancellation();
            }
            catch (Exception ex)
            {
                await HandleException(ex);
            }
            finally
            {
                await CleanupResources();
            }
        }

        private async Task ExecuteDrilldownProcess()
        {
            CommonMethods.DisableExcelSettings();

            Token.ThrowIfCancellationRequested();

            ExternalResolveResult = ExcelExternalRef.ResolveRangeWithContext(_epAddress);
            EpRange = ExternalResolveResult.Range;
            EpWorksheet = ExternalResolveResult.Worksheet;
            EpWorbook = ExternalResolveResult.Workbook;

            Token.ThrowIfCancellationRequested();

            if (EpRange == null) return;

            var precedents = GetAllPrecedents(EpRange);
            ServiceLocator.Logger?.LogDebug($"DrilldownXlPrecedents.ExecuteDrilldownProcess: found {precedents?.Count ?? 0} precedent cell(s).");
            if (!HasValidPrecedents(precedents))
            {
                ServiceLocator.Logger?.LogDebug("DrilldownXlPrecedents.ExecuteDrilldownProcess: no valid precedents found, aborting.");
                ShowNoPrecedentsMessage();
                return;
            }

            await InitializeProgressWindow(_ctsHelper);
            await ProcessPrecedentsAsync(precedents);
            ServiceLocator.RibbonController?.SetState("ApplySheetActiveState");
        }
        private static bool HasValidPrecedents(List<string> precedents)
        {
            return precedents != null && precedents.Count > 0;
        }

        private static void ShowNoPrecedentsMessage()
        {
            CommonFunctions.GLSenseMessage(
                "No balance formula cells found for the selected formula.",
                MessageBoxImage.Exclamation);
        }

        private async Task InitializeProgressWindow(CancellationHelper cts)
        {
            try
            {
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    Win = new GLWaitWindow(cts);
                    Win.Show();
                    Win.StartMonitoring();
                });

                await UpdateProgressWindow("Getting dependent formulas", "Processing request...");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                await SafelyCloseWindowAsync();
                throw;
            }
        }

        private static Task UpdateProgressWindow(string title, string message)
        {
            // Fire-and-forget: progress UI update only. Do not introduce a
            // suspend point here — callers proceed to touch Excel COM objects
            // immediately after awaiting this method.
            _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessTitle(title));
            _ = Win.Dispatcher.InvokeAsync(() => Win.SetProcessMessage(message));
            return Task.CompletedTask;
        }

        private async Task ProcessPrecedentsAsync(List<string> precedents)
        {
            await CreateDrillReferencesAsync(EpRange, precedents);
        }

        private static async Task HandleCancellation()
        {
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                "Operation Cancelled by the user!",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private static async Task HandleException(Exception ex)
        {
            ServiceLocator.Logger?.LogException(ex);
            await SafelyCloseWindowAsync();
            CommonFunctions.GLSenseMessage(
                $"An unexpected error occurred.{Environment.NewLine}{ex.Message}",
                MessageBoxImage.Error,
                MessageBoxButton.OK);
        }

        private async Task CleanupResources()
        {
            await SafelyCloseWindowAsync();
            CancelAndDisposeTokenSource();
            CommonMethods.EnableExcelSettings();
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
                await Win.Dispatcher.InvokeAsync(() => Win.RequestClose());
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error closing wait window: {ex.Message}");
            }
        }
        private void CancelAndDisposeTokenSource()
        {
            try
            {
                if (_ctsHelper != null && !_ctsHelper.IsCancellationRequested)
                    _ctsHelper.Cancel();

                _ctsHelper?.Dispose();  // ALWAYS SAFE - handles ALL cases
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogWarn($"CancelAndDisposeTokenSource: failed disposing CancellationHelper (non-fatal): {ex.Message}");
            }
        }
        private async Task CreateDrillReferencesAsync(Excel.Range cellrng, List<string> listRng)
        {
            // Runs synchronously on the calling (STA) thread — Excel COM objects are
            // apartment-affinitized and must not be touched from a ThreadPool thread.
            CreateDrillReferencesCore(cellrng, listRng);
            await Task.CompletedTask;
        }

        private void CreateDrillReferencesCore(Excel.Range cellrng, List<string> listRng)
        {
            Excel.Worksheet ws = null;
            try
            {
                ws = (Excel.Worksheet)ExcelApp.Worksheets.Add();
                ws.Tab.Color = ColorTranslator.ToOle(Color.FromArgb(131, 204, 235));

                SetupDrillHeader(ws, cellrng);
                int rowIndex = ProcessPrecedentCells(ws, listRng);
                AutoFitColumns(ws, rowIndex);
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownXlPrecedents.CreateDrillReferencesCore");
            }
            finally
            {
                ws?.Activate();
            }
        }

        private static void SetupDrillHeader(Excel.Worksheet ws, Excel.Range cellrng)
        {
            var drillCell = ws.Range["A1"];
            var drillCellRng = $"'{cellrng.Worksheet.Name}'!{cellrng.Address}";

            drillCell.Hyperlinks.Add(
                 Anchor: drillCell,
                    Address: string.Empty,          // No external URL
                    SubAddress: drillCellRng,       // Target cell reference
                    ScreenTip: Type.Missing,        // Optional
                    TextToDisplay: "Goto Definition"
                );
            drillCell.Font.Size = 12;

            SetupHeaderCell(ws.Range["A3"], "Reference Cell");
            SetupHeaderCell(ws.Range["B3"], "Balance Value");
        }

        private static void SetupHeaderCell(Excel.Range cell, string text)
        {
            cell.Value = text;
            cell.Font.Bold = true;
            cell.Font.Italic = true;
            cell.Font.Size = 11;
            cell.Font.ColorIndex = 2;
            cell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(21, 96, 130));
        }

        private int ProcessPrecedentCells(Excel.Worksheet ws, List<string> listRng)
        {
            int rowIndex = 3;
            foreach (string str in listRng)
            {
                try
                {
                    var formulaCell = ExcelApp.Range[str];
                    if (formulaCell == null) continue;

                    _ = ProcessFormulaCell(str, formulaCell, ref rowIndex, ws);
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, $"Error processing cell {str}");
                }
            }
            return rowIndex;
        }

        private static string ProcessFormulaCell(string cellAddress, Excel.Range formulaCell, ref int rowIndex, Excel.Worksheet ws)
        {
            string updatedFormula = string.Empty;

            if (CommonFunctions.GetBalancesCountInCells(formulaCell.Formula.ToString()) >= 2)
            {
                var formulaStr = formulaCell.Formula.ToString();
                var fncFormulas = CommonFunctions.MultiFormulaValues(formulaStr, "Functions");
                if (fncFormulas != null && fncFormulas.Count > 0)
                {
                    foreach (string fncargs in fncFormulas)
                    {
                        rowIndex++;
                        updatedFormula = ProcessSingleFormula(cellAddress, fncargs, rowIndex, ws);
                    }
                }
            }
            else
            {
                rowIndex++;
                updatedFormula = ProcessSingleFormula(cellAddress, formulaCell.Formula.ToString(), rowIndex, ws);
            }

            return updatedFormula;
        }

        private static string ProcessSingleFormula(string cellAddress, string formula, int rowIndex, Excel.Worksheet ws)
        {

            var fncparser = new ClsFormulaParser(formula);

            var updatedFormula = fncparser.Formula_Values();
            ws.Range[$"A{rowIndex}"].Value = cellAddress;
            ws.Range[$"B{rowIndex}"].Value = updatedFormula;
            ws.Range[$"B{rowIndex}"].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            return updatedFormula;

        }

        private static void AutoFitColumns(Excel.Worksheet ws, int rowIndex)
        {
            if (rowIndex > 3)
            {
                var formatRange = ws.Range[$"A3:B{rowIndex}"];
                formatRange.EntireColumn.AutoFit();
            }
        }

        public List<string> GetAllPrecedents(Excel.Range rng)
        {
            ServiceLocator.Logger?.LogDebug($"DrilldownXlPrecedents.GetAllPrecedents started for range {rng?.Address}.");
            var allPrecedents = new List<string>();
            try
            {
                GetPrecedents(rng, allPrecedents);
                return allPrecedents;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error in GetAllPrecedent");
                return new List<string>();
            }
        }

        private void GetPrecedents(Excel.Range rngToCheck, List<string> allPrecedents)
        {
            try
            {
                if (rngToCheck.Worksheet.ProtectContents) return;

                var formulaCells = GetFormulaCells(rngToCheck);
                if (formulaCells != null)
                {
                    foreach (Excel.Range cell in formulaCells.Cells)
                    {
                        Token.ThrowIfCancellationRequested();
                        GetCellPrecedents(cell, allPrecedents);
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DrilldownXlPrecedents.GetPrecedents");
            }
        }

        private static Excel.Range GetFormulaCells(Excel.Range rngToCheck)
        {
            try
            {
                object countLargeObj = rngToCheck.Cells.CountLarge;
                double countLarge = Convert.ToDouble(countLargeObj);
                bool isMultiCell = countLarge > 1;

                if (isMultiCell)
                {
                    return rngToCheck.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                }
                else if ((bool)rngToCheck.HasFormula)
                {
                    return rngToCheck;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Error getting formula cells");
            }
            return null;
        }

        private void GetCellPrecedents(Excel.Range rngCell, List<string> allPrecedents)
        {
            const int MAX_ARROWS = 50;  // Prevent infinite loops
            const int MAX_LINKS = 20;   // Prevent infinite loops

            int arrow = 0;
            bool hasMoreArrows = true;

            while (arrow < MAX_ARROWS && hasMoreArrows)
            {
                Token.ThrowIfCancellationRequested();

                arrow++;
                bool newArrow = true;
                int link = 0;

                while (link < MAX_LINKS)
                {
                    Token.ThrowIfCancellationRequested();

                    link++;
                    var precedentRange = NavigateToPrecedent(rngCell, arrow, link);

                    if (precedentRange == null) break;

                    string precedentAddress = GetCellAddress(precedentRange);
                    if (precedentAddress == GetCellAddress(rngCell))
                    {
                        break;
                    }

                    newArrow = false;
                    ProcessPrecedentRange(allPrecedents, precedentAddress);

                    // Recurse with depth protection
                    if (allPrecedents.Count < 1000)  // Prevent stack overflow
                    {
                        GetPrecedents(precedentRange, allPrecedents);
                    }
                }
                hasMoreArrows = !newArrow;
            }
        }

        private static Excel.Range NavigateToPrecedent(Excel.Range rngCell, int arrow, int link)
        {
            try
            {
                rngCell.ShowPrecedents();
                return (Excel.Range)rngCell.NavigateArrow(true, arrow, link);
            }
            catch
            {
                // Expected/normal: NavigateArrow throws once there are no more arrows/links to
                // walk for the given (arrow, link) pair - this is how GetCellPrecedents detects
                // the end of the precedent chain, so it fires routinely and isn't worth logging
                // per-iteration (would spam the log heavily under Debug mode on large sheets).
                return null;
            }
        }

        private void ProcessPrecedentRange(List<string> allPrecedents, string precedentAddress)
        {
            if (allPrecedents.Contains(precedentAddress)) return;

            var xlRange = ExcelApp.Range[precedentAddress];
            if (xlRange.Cells.Count >= 2)
            {
                foreach (Excel.Range lpRange in xlRange)
                {
                    Token.ThrowIfCancellationRequested();
                    ProcessBalanceFormulaCell(lpRange, allPrecedents);
                }
            }
            else
            {
                ProcessBalanceFormulaCell(xlRange, allPrecedents);
            }
        }

        private static void ProcessBalanceFormulaCell(Excel.Range range, List<string> allPrecedents)
        {
            if (range == null || !(bool)range.HasFormula || !range.Formula.ToString().Contains(AppConstants.glBal))
                return;

            string address = GetCellAddress(range);
            if (!allPrecedents.Contains(address))
            {
                allPrecedents.Add(address);
            }
        }

        public static string GetCellAddress(Excel.Range rng)
        {
            if (rng == null) return string.Empty;

            string sheetName = rng.Worksheet.Name;
            if (sheetName.Contains(" ") || sheetName.Contains("'"))
            {
                sheetName = $"'{sheetName.Replace("'", "''")}'";
            }

            return $"{sheetName}!{rng.Address[true, true, Excel.XlReferenceStyle.xlA1, false]}";
        }
    }
}
