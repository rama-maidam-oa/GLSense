using GLSense.Helpers;
using GLSense.Repositories;
using GLSense.Service;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Models
{
    public class BalanceDto
    {
        public Balances[] balanceDtos { get; set; }
        static string NormalizeStrings(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;

            // Convert \" to ", \\ to \ (handles common escaped inputs)
            // Regex.Unescape handles most escaped sequences
            var unescaped = Regex.Unescape(s);

            // Trim whitespace and any leading/trailing quotes
            var trimmed = unescaped.Trim().Trim('"', '“', '”', '\'');

            return trimmed;
        }

        public static object CreateBalanceDto(string cellAddress)
        {
            LogUtility.LogDebug($"BalanceDto.CreateBalanceDto: cellAddress={cellAddress}");
            try
            {
                var allBalances = ExtractBalancesFromExcelRange(cellAddress);
                LogUtility.LogDebug($"BalanceDto.CreateBalanceDto: extracted {allBalances?.Count ?? 0} balance(s) from '{cellAddress}'");
                return CreateBalanceDtoFromBalances(allBalances);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error creating BalanceDto from Excel range '{cellAddress}'");
                return null;
            }
        }
        private static List<Balances> ExtractBalancesFromExcelRange(string cellAddress)
        {
            var formulaCells = GetExcelRangeFromAddress(cellAddress);
            if (formulaCells == null || formulaCells.Count == 0)
                return new List<Balances>();

            return ProcessAllAreasInRange(formulaCells);
        }

        private static Excel.Range GetExcelRangeFromAddress(string cellAddress)
        {
            var result = ExcelExternalRef.ResolveRangeWithContext(cellAddress);
            return result?.Range;
        }

        private static List<Balances> ProcessAllAreasInRange(Excel.Range range)
        {
            var allBalances = new List<Balances>();
            var areas = range.Areas;

            for (int areaIndex = 1; areaIndex <= areas.Count; areaIndex++)
            {
                Excel.Range area = areas[areaIndex];
                ProcessSingleArea(area, allBalances);
            }

            return allBalances;
        }

        private static void ProcessSingleArea(Excel.Range area, List<Balances> allBalances)
        {
            object rawFormula = area.Formula;

            if (rawFormula is object[,] formulasArray)
            {
                ProcessMultiCellArea(area, formulasArray, allBalances);
            }
            else
            {
                ProcessSingleCellArea(area, rawFormula?.ToString(), allBalances);
            }
        }

        private static void ProcessMultiCellArea(Excel.Range area, object[,] formulasArray, List<Balances> allBalances)
        {
            int rows = formulasArray.GetLength(0);
            int cols = formulasArray.GetLength(1);

            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= cols; col++)
                {
                    string formula = formulasArray[row, col]?.ToString();
                    if (ShouldProcessFormula(formula))
                    {
                        ProcessFormulaCell(area, row, col, formula, allBalances);
                    }
                }
            }
        }

        private static void ProcessSingleCellArea(Excel.Range area, string formula, List<Balances> allBalances)
        {
            if (ShouldProcessFormula(formula))
            {
                ProcessCellFormula(area, formula, allBalances);
            }
        }

        private static void ProcessFormulaCell(Excel.Range area, int row, int col, string formula, List<Balances> allBalances)
        {
            Excel.Range cell = (Excel.Range)area.Cells[row, col];
            ProcessCellFormula(cell, formula, allBalances);
        }

        private static bool ShouldProcessFormula(string formula)
        {
            const string targetFunction = AppConstants.glBal;

            return !string.IsNullOrWhiteSpace(formula) &&
                   formula.IndexOf(targetFunction, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static object CreateBalanceDtoFromBalances(List<Balances> balances)
        {
            if (balances.Count == 0)
                return null;

            return new BalanceDto { balanceDtos = balances.ToArray() };
        }
        private static void ProcessCellFormula(Excel.Range cell, string formula, List<Balances> allBalances)
        {
            try
            {
                ProcessFormulaInternal(cell, formula, allBalances);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Failed processing cell {cell?.Address[true, true, Excel.XlReferenceStyle.xlA1, false]}");
            }
        }
        private static void ProcessFormulaInternal(Excel.Range cell, string formula, List<Balances> allBalances)
        {
            int functionCount = CommonFunctions.GetBalancesCountInCells(formula);

            switch (functionCount)
            {
                case 0:
                    // Shouldn't happen because we already filtered on IndexOf, do nothing.
                    break;

                case 1:
                    ProcessSingleBalanceFunction(cell, formula, allBalances);
                    break;

                default: // > 1
                    ProcessMultipleBalanceFunctions(cell, formula, allBalances);
                    break;
            }
        }
        private static void ProcessSingleBalanceFunction(Excel.Range cell, string formula, List<Balances> allBalances)
        {
            var funcValues = CommonFunctions.FormulaValues(formula);
            if (!HasSufficientParameters(funcValues, cell))
                return;

            AddBalanceFromParameters(cell, funcValues, allBalances);
        }
        private static void ProcessMultipleBalanceFunctions(Excel.Range cell, string formula, List<Balances> allBalances)
        {
            var functionFormulas = CommonFunctions.MultiFormulaValues(formula, "Functions");
            if (functionFormulas == null || functionFormulas.Count == 0)
                return;

            foreach (var singleFunction in functionFormulas)
            {
                ProcessSingleFunctionInMultiFormula(cell, singleFunction, allBalances);
            }
        }
        private static void ProcessSingleFunctionInMultiFormula(Excel.Range cell, string functionFormula, List<Balances> allBalances)
        {
            var funcValues = CommonFunctions.MultiFormulaValues(functionFormula, "Arguments_WithValues");
            if (funcValues == null)
                return;

            AddBalanceFromParameters(cell, funcValues, allBalances);
        }
        private static bool HasSufficientParameters(List<string> funcValues, Excel.Range cell)
        {
            const int minimumRequiredParameters = 11;

            if (funcValues != null && funcValues.Count >= minimumRequiredParameters)
                return true;

            LogInsufficientParametersError(cell);
            return false;
        }
        private static void AddBalanceFromParameters(Excel.Range cell, List<string> parameters, List<Balances> allBalances)
        {
            var cellAddress = GetR1C1Address(cell);
            var result = CreateFromXllParameters(cellAddress, parameters);

            if (result is Balances balance)
            {
                allBalances.Add(balance);
            }
        }
        private static string GetR1C1Address(Excel.Range cell)
        {
            return cell.Address[true, true, Excel.XlReferenceStyle.xlR1C1, false];
        }

        private static string GetA1Address(Excel.Range cell)
        {
            return cell.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
        }

        private static void LogInsufficientParametersError(Excel.Range cell)
        {
            LogUtility.LogError(
                $"Insufficient parameters in GLSense_GetBalance formula at {GetA1Address(cell)}");
        }
        public static object CreateFromXllParameters(string cellRef, List<string> args)
        {
            LogUtility.LogDebug($"BalanceDto.CreateFromXllParameters: cellRef={cellRef}, argCount={args?.Count ?? 0}");
            string periodName = NormalizeStrings(args[3]);

            var normalizedBT = NormalizeStrings(args[4]);

            if (new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "CTD", "JED", "JEDP", "JEDU" }.Contains(normalizedBT))
            {
                periodName = ResolveCtdPeriodName(args[3]);
            }

            // First standard parameters
            var balance = new Balances
            {
                excelCell = NormalizeStrings(cellRef),
                cellSign = NormalizeStrings(args[0]),
                periodName = periodName,
                balanceType = NormalizeStrings(args[4]),
                currencyCode = NormalizeStrings(args[5]),
                translatedFlag = NormalizeStrings(args[6]),
                actualFlag = NormalizeStrings(args[7]),
                jeSourceName = NormalizeStrings(args[9]),
                jeCategoryName = NormalizeStrings(args[10])
            };

            //Ledger ID List - handle multiple ledger IDs separated by commas
            string ledgerName = args[1];

            var ledgerRecord = AppState.Instance.SelectedCube.Ledgers;

            long ledgerId = 0;
            List<LedgerRecord> matchedLedgersForCurrency = null;

            if (ledgerRecord == null || ledgerRecord.Count == 0)
            {
                balance.ledgerIdList = null;
            }
            else
            {
                string ldgerNameNormalized = NormalizeStrings(ledgerName);
                var ledgerNames = ldgerNameNormalized.ToString().Split(';').Select(name => name.Trim());
                var matchingLedgers = ledgerRecord.Where(l => ledgerNames.Contains(l.LedgerName)).ToList();
                balance.ledgerIdList = matchingLedgers.Any()
                                        ? matchingLedgers.Select(l => (object)l.LedgerId).ToArray()
                                        : null;

                balance.coaid = matchingLedgers.FirstOrDefault()?.Coaid.ToString(); // Safe null check with ?.
                ledgerId = matchingLedgers.FirstOrDefault()?.LedgerId ?? 0; // Default to 0 if no match
                matchedLedgersForCurrency = matchingLedgers;
            }

            // isFunctionalCurrency must be derived from the ledger(s) named in this formula's
            // own parameters (matchedLedgersForCurrency, resolved above from args[1]) - NOT
            // from whatever ledger happens to be selected in the ribbon dropdown
            // (AppState.Instance.SelectedLedger). The ribbon selection has no relationship to
            // which ledger(s) this particular formula call is actually evaluating. When the
            // formula names multiple ledgers, each can have a different functional currency,
            // so every matched ledger must be checked individually - true if ANY of them has
            // a functional currency equal to this balance's currency code, false only if none do.
            if (matchedLedgersForCurrency != null && matchedLedgersForCurrency.Count > 0)
            {
                balance.isFunctionalCurrency = matchedLedgersForCurrency.Any(l => l.CurrencyCode == balance.currencyCode);
            }
            else
            {
                balance.isFunctionalCurrency = true; // Default to true if no matching ledger found
            }

            EnsureLedgerDataLoaded(AppState.Instance.SelectedCube.CubeId, ledgerId);

            //Safer converter for Activity type if short form used manually
            string Activity = NormalizeStrings(args[2]);

            if (!string.IsNullOrEmpty(Activity))
            {
                var repo = new DataRepository();
                var activities = repo.GetActivities(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);

                var matchedActivity = activities.FirstOrDefault(a =>
                    string.Equals(a.DisplayName, Activity, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(a.ShortName, Activity, StringComparison.OrdinalIgnoreCase));

                if (matchedActivity != null)
                {
                    balance.activity = matchedActivity.ShortName; // Always return ShortName
                }
                else
                {
                    balance.activity = NormalizeStrings(args[2]);
                }
            }
            else
            {
                balance.activity = NormalizeStrings(args[2]);
            }

            string ActualFlag = NormalizeStrings(args[7]);

            string budEncum = string.Empty;

            if (ActualFlag == "BUDGET" || ActualFlag == "B")
            {
                budEncum = NormalizeStrings(args[8]); // Budget Name
                balance.budgetName = budEncum;
                balance.encumbranceName = string.Empty;
            }
            else if (ActualFlag == "ENCUMBRANCE" || ActualFlag == "E" || ActualFlag == "ACTUAL+ENCUMBRANCE" || ActualFlag == "A+E")
            {
                budEncum = NormalizeStrings(args[8]); // Encumbrance Name
                balance.encumbranceName = budEncum;

                if (ledgerId != 0 && !string.IsNullOrEmpty(budEncum))
                {
                    var repo = new DataRepository();
                    var encumbrances = repo.GetEncumbrances(AppState.Instance.SelectedCube.CubeId, ledgerId);

                    var encumbranceNames = budEncum
                        .Split(';')
                        .Select(e => NormalizeStrings(e))
                        .Where(e => !string.IsNullOrWhiteSpace(e))
                        .ToList();

                    var encumbranceIds = new List<object>();

                    foreach (var encName in encumbranceNames)
                    {
                        var matchedEncumbrance = encumbrances.FirstOrDefault(e =>
                            string.Equals(e.EncumbranceType, encName, StringComparison.OrdinalIgnoreCase));
                        if (matchedEncumbrance != null)
                        {
                            encumbranceIds.Add(matchedEncumbrance.EncumbranceTypeId);
                        }
                    }

                    balance.encumbranceTypeIdList = encumbranceIds.Count > 0 ? encumbranceIds.ToArray() : null;
                }

                balance.budgetName = string.Empty;
            }

            // Process segments
            var processedSegments = ProcessSegments(args.Skip(11).ToArray(), ledgerName);
            balance.segments = processedSegments.Count > 0 ? processedSegments.ToArray() : null;

            return balance;
        }
        private static void EnsureLedgerDataLoaded(long cubeId, long ledgerId)
        {
            LogUtility.LogDebug($"BalanceDto.EnsureLedgerDataLoaded: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
                {
                    LogUtility.LogDebug("BalanceDto.EnsureLedgerDataLoaded: login not completed or no cube selected, skipping.");
                    return;
                }

                if (DataRepository.GetTableItemsCount(cubeId, ledgerId, "SEGMENTS") > 0)
                {
                    LogUtility.LogDebug("BalanceDto.EnsureLedgerDataLoaded: SEGMENTS already populated, skipping remote fetch.");
                    return;
                }
                using var ctsHelper = new CancellationHelper();
                CommonFunctions.FillResponsibilitiesAsync(ledgerId, cubeId, ctsHelper.GetToken())
                    .GetAwaiter()
                    .GetResult();
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogError("Ledger data loading was cancelled by the user.");
            }
            catch (Exception)
            {
                // Already logged inside CommonFunctions.FillResponsibilitiesAsync before rethrow.
            }
        }
        private static string ResolveCtdPeriodName(string rawPeriodName)
        {
            if (string.IsNullOrWhiteSpace(rawPeriodName))
                return string.Empty;

            var parts = rawPeriodName.Split(new[] { "~" }, StringSplitOptions.None)
                .Select(part => ResolvePeriodToken(part))
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .ToArray();

            return parts.Length > 0 ? string.Join("~", parts) : ResolvePeriodToken(rawPeriodName);
        }

        private static string ResolvePeriodToken(string token)
        {
            var cleanToken = NormalizeStrings(token);
            if (string.IsNullOrWhiteSpace(cleanToken))
                return string.Empty;

            return cleanToken.Replace("\"", string.Empty)
                             .Replace("“", string.Empty)
                             .Replace("”", string.Empty)
                             .Trim();
        }

        private static List<Segment> ProcessSegments(string[] segVals, string ledgerName)
        {
            var segments = new List<Segment>();
            var segmentValues = LoadSegmentValues(ledgerName);
            var segmentModels = LoadSegments(ledgerName);

            if (segmentModels == null || segmentModels.Count == 0 || segmentValues == null || segmentValues.Count == 0)
                return segments;

            var distinctAppCols = segmentModels
                .Select(s => s.ApplicationColumnName)
                .Distinct()
                .ToList();

            if (distinctAppCols.Count == 0)
                return segments;

            var segmentsToProcess = GetSegmentsToProcess(segVals, distinctAppCols.Count);

            ProcessSegmentsBasedOnFirstValue(segVals, distinctAppCols, segmentValues, segments, segmentsToProcess);

            return segments;
        }
        private static int GetSegmentsToProcess(string[] segVals, int distinctAppColCount)
        {
            return Math.Min(distinctAppColCount, segVals.Length);
        }
        private static void ProcessSegmentsBasedOnFirstValue(string[] segVals, List<string> distinctAppCols,
                ObservableCollection<SegmentValueModel> segmentValues, List<Segment> segments, int segmentsToProcess)
        {
            var firstSegValue = segVals.FirstOrDefault();

            if (ShouldProcessAsCombinedSegments(firstSegValue))
            {
                ProcessCombinedSegments(firstSegValue, distinctAppCols, segmentValues, segments, segmentsToProcess);
            }
            else
            {
                ProcessIndividualSegments(segVals, distinctAppCols, segmentValues, segments, segmentsToProcess);
            }
        }
        private static bool ShouldProcessAsCombinedSegments(string segValue)
        {
            return !string.IsNullOrWhiteSpace(segValue) && segValue.Contains(";");
        }
        private static void ProcessCombinedSegments(string combinedValue, List<string> distinctAppCols,
            ObservableCollection<SegmentValueModel> segmentValues, List<Segment> segments, int segmentsToProcess)
        {
            var segmentParts = combinedValue.Split(';');
            var partsToProcess = Math.Min(segmentParts.Length, segmentsToProcess);

            for (int i = 0; i < partsToProcess; i++)
            {
                ProcessSegmentPart(segmentParts[i], i, distinctAppCols, segmentValues, segments);
            }
        }
        private static void ProcessIndividualSegments(string[] segVals, List<string> distinctAppCols,
    ObservableCollection<SegmentValueModel> segmentValues, List<Segment> segments, int segmentsToProcess)
        {
            for (int i = 0; i < segmentsToProcess; i++)
            {
                ProcessSegmentPart(segVals[i], i, distinctAppCols, segmentValues, segments);
            }
        }
        private static void ProcessSegmentPart(string segValue, int index, List<string> distinctAppCols,
    ObservableCollection<SegmentValueModel> segmentValues, List<Segment> segments)
        {
            if (string.IsNullOrEmpty(segValue))
                return;

            var segName = GetSegmentName(index, distinctAppCols);
            var segSetId = GetSegmentValueSetId(segName, segmentValues);

            var segment = CreateSegment(segValue, index, segmentValues, segSetId, segName);
            if (segment != null)
            {
                segments.Add(segment);
            }
        }
        private static string GetSegmentName(int index, List<string> distinctAppCols)
        {
            if (index < 0 || index >= distinctAppCols.Count)
                return string.Empty;

            return distinctAppCols[index];
        }

        private static long GetSegmentValueSetId(string segName, ObservableCollection<SegmentValueModel> segmentValues)
        {
            var segmentValue = segmentValues.FirstOrDefault(s => s.ApplicationColumnName == segName);
            return segmentValue?.SegmentValueSetId ?? 0;
        }
        private static Segment CreateSegment(string segmentValue, int segmentIndex, ObservableCollection<SegmentValueModel> segValues, long segSetId, string segName)
        {
            if (string.IsNullOrEmpty(NormalizeStrings(segmentValue))) return null;

            return new Segment
            {
                segmentValueSetId = segSetId,
                segmentNumber = segmentIndex + 1,
                segmentValues = CreateSegmentValues(segmentValue, segValues, segName).ToArray()
            };
        }

        private static List<SegmentValue> CreateSegmentValues(string value, ObservableCollection<SegmentValueModel> segValues, string segName)
        {
            var segmentValues = new List<SegmentValue>();

            if (string.IsNullOrEmpty(value))
                return segmentValues;

            foreach (var item in value.Split(','))
            {
                var segmentValue = CreateSingleSegmentValue(item, segValues, segName);
                if (segmentValue != null)
                {
                    segmentValues.Add(segmentValue);
                }
            }

            return segmentValues;
        }
        private static SegmentValue CreateSingleSegmentValue(string item, ObservableCollection<SegmentValueModel> segValues, string segName)
        {
            var cleanValue = NormalizeStrings(item.Replace("~", "").Trim());
            if (string.IsNullOrEmpty(cleanValue))
                return null;

            if (cleanValue.Contains("%"))
            {
                return CreateLikeSegmentValue(cleanValue);
            }

            if (cleanValue.Contains("|"))
            {
                return CreateRangeSegmentValue(cleanValue);
            }

            if (cleanValue.Contains("--"))
            {
                return CreateNotInSegmentValue(cleanValue, segValues, segName);
            }

            return CreateInSegmentValue(cleanValue, segValues, segName);
        }

        private static SegmentValue CreateLikeSegmentValue(string cleanValue)
        {
            return new SegmentValue
            {
                @operator = "LIKE",
                values = new[] { cleanValue },
                summaryEnabled = false
            };
        }
        private static SegmentValue CreateRangeSegmentValue(string originalItem)
        {
            var rangeValues = originalItem.Replace("--", "").Split('|');
            if (rangeValues.Length != 2)
                return null;

            return new SegmentValue
            {
                @operator = originalItem.Contains("--") ? "NOTBETWEEN" : "BETWEEN",
                values = rangeValues,
                summaryEnabled = false
            };
        }
        private static SegmentValue CreateNotInSegmentValue(string originalItem, ObservableCollection<SegmentValueModel> segValues, string segName)
        {
            originalItem = originalItem.Replace("--", "");
            return CreateComparisonSegmentValue("NOTIN", originalItem, segValues, segName);
        }

        private static SegmentValue CreateInSegmentValue(string cleanValue, ObservableCollection<SegmentValueModel> segValues, string segName)
        {
            return CreateComparisonSegmentValue("IN", cleanValue, segValues, segName);
        }

        private static SegmentValue CreateComparisonSegmentValue(string operatorName, string cleanValue, ObservableCollection<SegmentValueModel> segValues, string segName)
        {
            var segmentValue = new SegmentValue
            {
                @operator = operatorName,
                values = new[] { cleanValue },
                summaryEnabled = GetSummaryEnabledStatus(cleanValue, segValues, segName)
            };
            return segmentValue;
        }

        private static bool GetSummaryEnabledStatus(string segmentValueText, ObservableCollection<SegmentValueModel> segValues, string segName)
        {
            var firstMatch = segValues.FirstOrDefault(s =>
                s.ApplicationColumnName == segName &&
                s.SegmentValue == segmentValueText
            );

            if (firstMatch == null)
                return false;

            return !string.Equals(firstMatch.SummaryFlag, "N", StringComparison.OrdinalIgnoreCase);
        }
        private static ObservableCollection<SegmentValueModel> LoadSegmentValues(string ledgerName)
        {
            LogUtility.LogDebug($"BalanceDto.LoadSegmentValues: ledgerName={ledgerName}");
            try
            {
                var task = Task.Run(() =>
                {
                    var dataService = ServiceLocator.SegmentDataService;
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
                LogUtility.LogException(ex, "Failed to load segment values");
                return new ObservableCollection<SegmentValueModel>();
            }
        }
        private static ObservableCollection<SegmentModel> LoadSegments(string ledgerName)
        {
            LogUtility.LogDebug($"BalanceDto.LoadSegments: ledgerName={ledgerName}");
            try
            {
                var task = Task.Run(() =>
                {
                    var dataService = ServiceLocator.SegmentDataService;
                    return dataService.GetSegments(ledgerName);
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
                LogUtility.LogException(ex, "Failed to load segment values");
                return new ObservableCollection<SegmentModel>();
            }
        }
    }
    public class Balances
    {
        public string excelCell { get; set; }
        public string cellSign { get; set; }
        public object[] ledgerIdList { get; set; }
        public string activity { get; set; }
        public string periodName { get; set; }
        public string balanceType { get; set; }
        public string currencyCode { get; set; }
        public bool isFunctionalCurrency { get; set; }
        public string translatedFlag { get; set; }
        public string actualFlag { get; set; }
        public string budgetName { get; set; }
        public string encumbranceName { get; set; }
        public object[] encumbranceTypeIdList { get; set; }
        public string jeSourceName { get; set; }
        public string jeCategoryName { get; set; }
        public string coaid { get; set; }
        public Segment[] segments { get; set; }
    }

    public class Segment
    {
        public long segmentValueSetId { get; set; }
        public int segmentNumber { get; set; }
        public SegmentValue[] segmentValues { get; set; }
    }

    public class SegmentValue
    {
        public string @operator { get; set; }
        public bool summaryEnabled { get; set; }
        public string[] values { get; set; }
    }
}
