using GLSense.Bindings;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Service;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace GLSense.ViewModels
{
    public class GLPeriodDetails : INotifyPropertyChanged, IFieldDependencyProvider
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        public ObservableCollection<GenericLedgerModel> Ledgers { get; set; } = new ObservableCollection<GenericLedgerModel>();
        public ObservableCollection<PeriodModel> Periods { get; set; } = new ObservableCollection<PeriodModel>();

        public ComboFieldBindings LedgerField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings PeriodField { get; set; } = new ComboFieldBindings();

        private readonly Dispatcher _dispatcher;

        private readonly string _formulaName;

        private const string Start = "START";

        public Action<string> ShowWarningAction { get; set; }
        public Func<string, Task> ShowWarningAsyncAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

        // ===== Check Sign, Zeroes and Text Factor =====
        private bool _isAdjacentChecked;
        public bool IsAdjacentChecked
        {
            get => _isAdjacentChecked;
            set
            {
                _isAdjacentChecked = value;
                OnPropertyChanged(nameof(IsAdjacentChecked));
                ResultOutPut();
            }
        }

        private string _resultText;
        public string ResultText
        {
            get => _resultText;
            set
            {
                _resultText = value;
                OnPropertyChanged(nameof(ResultText));
            }
        }
        private PeriodModel _selectedPeriod;
        public PeriodModel SelectedPeriod
        {
            get => _selectedPeriod;
            set
            {
                _selectedPeriod = value;
                OnPropertyChanged(nameof(SelectedPeriod));
            }
        }
        public GLPeriodDetails(Dispatcher dispatcher, string formulaName)
        {
            _dispatcher = dispatcher;
            _formulaName = formulaName;

            LedgerField.DependencyProvider = this;
            LedgerField.Type = ComboFieldBindings.FieldType.Ledger;

            PeriodField.DependencyProvider = this;
            PeriodField.Type = ComboFieldBindings.FieldType.Period;
        }
        public async Task LoadDataAsync(List<string> FuncArgs = null, List<string> FuncValues = null)
        {
            LogUtility.LogDebug($"GLPeriodDetails.LoadDataAsync: FormulaName={_formulaName}, FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            try
            {
                var ledgers = await Task.Run(() =>
                {
                    var repository = new DataRepository();
                    return repository.GetConfiguratorLedgers(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.CoaId, true);
                });

                await _dispatcher.InvokeAsync(() =>
                {
                    if (Ledgers == null)
                        Ledgers = new ObservableCollection<GenericLedgerModel>();
                    else
                        Ledgers.Clear();

                    foreach (var ledger in ledgers)
                    {
                        Ledgers.Add(ledger);
                    }
                });

                if (FuncArgs != null && FuncArgs.Count > 0)
                    await ApplyFormulaParams(FuncArgs, FuncValues);
                else
                    await ApplyDefaultLedgerSelection();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLPeriodDetails.LoadDataAsync");
                ShowWarningAction?.Invoke($"Error loading data: {ex.Message}");
            }
            finally
            {
                if (HideBusyAsyncAction != null)
                {
                    await HideBusyAsyncAction();
                }
            }
        }
        private async Task ApplyDefaultLedgerSelection()
        {
            LogUtility.LogDebug("GLPeriodDetails.ApplyDefaultLedgerSelection: entry");
            if (AppState.Instance.SelectedLedger == null)
            {
                LogUtility.LogWarn("GLPeriodDetails.ApplyDefaultLedgerSelection: AppState.Instance.SelectedLedger is null, aborting default ledger selection.");
                return;
            }
            var match = Ledgers.FirstOrDefault(x => x.LedgerId == AppState.Instance.SelectedLedger.LedgerId);
            if (match != null)
            {
                LedgerField.ComboValue = match;
                LedgerField.RefValue = null;
                LedgerField.RefreshEnableState();

                await LoadPeriodsForLedger(match);
            }
            else
            {
                LogUtility.LogWarn($"GLPeriodDetails.ApplyDefaultLedgerSelection: no ledger match found for LedgerId={AppState.Instance.SelectedLedger.LedgerId}.");
            }

            if (_formulaName is Start or "END")
            {
                IsAdjacentChecked = true;
                OnPropertyChanged(nameof(IsAdjacentChecked));
            }
        }
        private void SetPeriodFromMatch(string periodName)
        {
            var periodMatch = Periods.FirstOrDefault(x => x.PeriodName == periodName);
            if (periodMatch != null)
            {
                PeriodField.ComboValue = periodMatch;
                PeriodField.RefValue = null;
            }
        }
        private GenericLedgerModel FindLedgerMatch(string ledgerName)
        {
            return Ledgers.FirstOrDefault(x => x.LedgerName == ledgerName);
        }
        private void ProcessLedgerField(string funcArg, GenericLedgerModel match)
        {
            if (ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", "")))
            {
                LedgerField.ComboValue = null;
                LedgerField.IsValueFromRefEdit = true;
                LedgerField.RefValue = funcArg.Replace("\"", "");
            }
            else
            {
                LedgerField.ComboValue = match;
                LedgerField.RefValue = null;
            }
            LedgerField.RefreshEnableState();
        }
        private async Task LoadPeriodsForLedger(GenericLedgerModel ledger)
        {
            if (ledger == null)
                return;

            LogUtility.LogDebug($"GLPeriodDetails.LoadPeriodsForLedger: LedgerName={ledger.LedgerName}");
            try
            {
                if (ShowBusyAction != null)
                    await ShowBusyAction("Loading periods...", null);

                var periods = await Task.Run(() => ServiceLocator.PeriodDataService.GetPeriodsForLedger(ledger.LedgerName));
                LogUtility.LogDebug($"GLPeriodDetails.LoadPeriodsForLedger: loaded {periods?.Count ?? 0} period(s) for LedgerName={ledger.LedgerName}");

                await _dispatcher.InvokeAsync(() =>
                {
                    Periods.Clear();
                    if (periods?.Count > 0)
                    {
                        foreach (var p in periods)
                        {
                            Periods.Add(p);
                        }
                        OnPropertyChanged(nameof(Periods));
                    }
                });

                await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
            }
            finally
            {
                if (HideBusyAsyncAction != null)
                    await HideBusyAsyncAction();
            }
        }
        private void ProcessPeriodField(string funcArg, string funcValue)
        {
            string periodName = funcValue.Replace("\"", "");

            if (ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", "")))
            {
                PeriodField.ComboValue = null;
                PeriodField.IsValueFromRefEdit = true;
                PeriodField.RefValue = funcArg.Replace("\"", "");
            }
            else
                SetPeriodFromMatch(periodName);

            PeriodField.RefreshEnableState();
        }
        private bool IsStartOrEndFormula()
        {
            return _formulaName == Start || _formulaName == "END";
        }
        private void ProcessAdjacentCheck(string funcArg, string funcValue)
        {
            bool boolValue = ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", ""))
                ? ConvertToBool(GetCellValueIfRange(funcArg.Replace("\"", "")))
                : ConvertToBool(funcValue.Replace("\"", ""));

            IsAdjacentChecked = boolValue;
            OnPropertyChanged(nameof(IsAdjacentChecked));
        }
        private async Task ApplyFormulaParams(List<string> FuncArgs, List<string> FuncValues)
        {
            LogUtility.LogDebug($"GLPeriodDetails.ApplyFormulaParams: FormulaName={_formulaName}, FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            string ledgerName = string.Empty;

            if (FuncValues.Count >= 2)
            {
                ledgerName = FuncValues[1].Replace("\"", "");
            }
            else
            {
                ledgerName = AppState.Instance.SelectedLedger.LedgerName;
            }

            if (string.IsNullOrEmpty(ledgerName))
            {
                LogUtility.LogWarn("GLPeriodDetails.ApplyFormulaParams: resolved ledgerName is empty, aborting.");
                return;
            }

            var ledgerMatch = FindLedgerMatch(ledgerName);
            if (ledgerMatch == null)
            {
                LogUtility.LogWarn($"GLPeriodDetails.ApplyFormulaParams: no ledger match found for '{ledgerName}'.");
                return;
            }

            ProcessLedgerField(ledgerName, ledgerMatch);
            await LoadPeriodsForLedger(ledgerMatch);
            await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
            ProcessPeriodField(FuncArgs[0], FuncValues[0]);

            if (IsStartOrEndFormula())
            {
                await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                ProcessAdjacentCheck(FuncArgs[2], FuncValues[2]);
            }

            ResultOutPut();
        }
        private static bool ConvertToBool(string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;

            // Remove surrounding quotes if any
            value = value.Trim().Trim('"');

            // Parse ignoring case
            if (bool.TryParse(value, out bool result))
            {
                return result;
            }

            // Fallback: custom checks just in case
            var lowerValue = value.ToLowerInvariant();
            if (lowerValue == "true")
                return true;
            if (lowerValue == "false")
                return false;

            // Default fallback if neither true nor false
            return false;
        }
        public bool IsRefEnabled(ComboFieldBindings field)
        {
            if (field.IsValueFromRefEdit)
                return true;       // RefEdit always enabled when value comes from ref

            return field.ComboValue == null;
        }

        public bool IsComboEnabled(ComboFieldBindings field)
        {
            if (field.IsValueFromRefEdit)
                return false;      // Combo never enabled when RefEdit drives it

            return string.IsNullOrEmpty(field.RefValue);
        }
        public async Task OnFieldDependencyChanged(ComboFieldBindings field)
        {
            LogUtility.LogDebug($"GLPeriodDetails.OnFieldDependencyChanged: FieldType={field?.Type}, ComboValue={field?.ComboValue}, RefValue={field?.RefValue}");
            // Only one control enabled at a time
            LedgerField.RefreshEnableState();
            PeriodField.RefreshEnableState();

            if (field.Type == ComboFieldBindings.FieldType.Ledger)
            {
                ResetLedgerDependentFields();

                if (field.ComboValue is GenericLedgerModel Ledger)
                {
                    await LoadPeriodsForLedger(Ledger);
                }
            }
            // Offset population based on Date selection
            if (field.Type == ComboFieldBindings.FieldType.Period && field.ComboValue != null)
            {
                SelectedPeriod = (PeriodModel)field.ComboValue;
                OnPropertyChanged(nameof(SelectedPeriod));
            }
            ResultOutPut();
        }
        private void ResetField(ComboFieldBindings field)
        {
            ResetFieldToEmpty(field);
            if (field.Type == ComboFieldBindings.FieldType.Ledger)
            {
                ResetLedgerDependentFields();
            }
            SelectedPeriod = null;
            OnPropertyChanged(nameof(SelectedPeriod));
        }
        private static void ResetFieldToEmpty(ComboFieldBindings field)
        {
            field.IsValueFromRefEdit = false;  // ⭐ RESET FLAG
            field.ComboText = string.Empty;
            field.ComboValue = null;
            field.RefValue = null;
            field.RefreshEnableState();
        }
        private void ResetLedgerDependentFields()
        {
            Periods = new ObservableCollection<PeriodModel>();
            PeriodField.IsValueFromRefEdit = false;
            PeriodField.ComboText = string.Empty;
            PeriodField.ComboValue = null;
            PeriodField.RefValue = null;
            PeriodField.RefreshEnableState();
            SelectedPeriod = null;
            ResultText = string.Empty;
        }
        private void ProcessRefEditValue(ComboFieldBindings field, string newText)
        {
            string cellValue = GetCellValueIfRange(newText);

            if (string.IsNullOrEmpty(cellValue))
            {
                HandleEmptyCellReference(field, newText);
                return;
            }

            if (TryMatchFieldValue(field, cellValue))
                field.IsValueFromRefEdit = true;
            else
                HandleMatchFailure(field, cellValue);
        }
        private bool TryMatchFieldValue(ComboFieldBindings field, string cellValue)
        {
            return field.Type switch
            {
                ComboFieldBindings.FieldType.Ledger => MatchLedger(field, cellValue),
                ComboFieldBindings.FieldType.Period => MatchPeriod(field, cellValue),
                _ => false
            };
        }

        private bool MatchLedger(ComboFieldBindings field, string cellValue)
        {
            var match = Ledgers.FirstOrDefault(x => x.LedgerName.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                field.ComboValue = match;
                field.RefreshEnableState();
                return true;
            }
            HandleMatchFailure(field, cellValue);
            return false;
        }

        private bool MatchPeriod(ComboFieldBindings field, string cellValue)
        {
            var match = Periods.FirstOrDefault(x => x.PeriodName.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                field.ComboValue = match;
                field.RefreshEnableState();
                return true;
            }
            HandleMatchFailure(field, cellValue);
            return false;
        }
        private void HandleEmptyCellReference(ComboFieldBindings field, string newText)
        {
            ResetField(field);
            ShowWarningAction?.Invoke($"The referenced cell \"{newText}\" is empty");
        }

        private void HandleMatchFailure(ComboFieldBindings field, string cellValue)
        {
            ResetField(field);
            string message = field.Type switch
            {
                ComboFieldBindings.FieldType.Ledger => $"Ledger \"{cellValue}\" not found in available ledgers.",
                ComboFieldBindings.FieldType.Period => $"Period \"{cellValue}\" not found in available periods.",
                _ => "Item not found."
            };
            ShowWarningAction?.Invoke(message);
        }

        public void OnRefEditTextChanged(ComboFieldBindings field, string newText)
        {
            LogUtility.LogDebug($"GLPeriodDetails.OnRefEditTextChanged: FieldType={field?.Type}, newText={newText}");
            if (string.IsNullOrWhiteSpace(newText))
            {
                ResetField(field);
                return;
            }

            if (string.IsNullOrWhiteSpace(field.RefValue))
            {
                ResultOutPut();
                return;
            }


            if (!string.IsNullOrWhiteSpace(field.RefValue) && field.IsRefEnabled)
            {
                field.IsValueFromRefEdit = true;
            }

            ProcessRefEditValue(field, newText);
            ResultOutPut();
        }
        private string GetCellValueIfRange(string refText)
        {
            if (string.IsNullOrWhiteSpace(refText)) return null;
            try
            {
                return ExcelApp?.Range[refText]?.Value2?.ToString();
            }
            catch
            {
                return null;
            }
        }
        private string GetResultText()
        {
            if (!HasValidInputs())
                return string.Empty;

            var selPeriod = GetSelectedPeriod();
            if (selPeriod == null)
                return string.Empty;

            var (startPeriod, endPeriod) = GetStartEndPeriods(selPeriod);

            return _formulaName switch
            {
                "NUM" => selPeriod.PeriodNum.ToString(),
                "YEAR" => selPeriod.PeriodYear.ToString(),
                "QTR" => selPeriod.QuarterNum.ToString(),
                Start => startPeriod ?? string.Empty,
                "END" => endPeriod ?? string.Empty,
                _ => string.Empty
            };
        }
        private bool HasValidInputs()
        {
            return HasValidField(LedgerField) &&
                   HasValidField(PeriodField) &&
                   Periods?.Count > 0;
        }
        private static bool HasValidField(ComboFieldBindings field)
        {
            return field.ComboValue != null ||
                   !(field.RefValue == null || string.IsNullOrWhiteSpace(field.RefValue.ToString()));
        }
        private PeriodModel GetSelectedPeriod()
        {
            var selPeriod = (PeriodModel)PeriodField.ComboValue;
            return selPeriod;
        }
        private void ResultOutPut()
        {
            try
            {
                ResultText = GetResultText();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLPeriodDetails.ResultOutPut");
            }
            finally
            {
                OnPropertyChanged(nameof(ResultText));
            }
        }
        private (string StartPeriod, string EndPeriod) GetStartEndPeriods(PeriodModel SelectedPeriod)
        {

            if (Periods == null || SelectedPeriod == null)
                return (null, null);

            var filteredPeriods = Periods
                .Where(p => p.PeriodYear == SelectedPeriod.PeriodYear)
                .ToList();

            if (!filteredPeriods.Any())
                return (null, null);

            IEnumerable<PeriodModel> candidatePeriods;

            if (IsAdjacentChecked)
            {
                // Include all periods regardless of AdjustmentPeriodFlag
                candidatePeriods = filteredPeriods;
            }
            else
            {
                // Exclude periods with AdjustmentPeriodFlag == "Y"
                candidatePeriods = filteredPeriods
                     .Where(p => p.AdjustmentPeriodFlag != "Y");

            }

            var startPeriod = candidatePeriods.OrderBy(p => p.PeriodNum).First();
            var endPeriod = candidatePeriods.OrderByDescending(p => p.PeriodNum).First();

            return (startPeriod.PeriodName, endPeriod.PeriodName);
        }
        public bool WriteFormulaToCell(Microsoft.Office.Interop.Excel.Range rng)
        {
            LogUtility.LogDebug("GLPeriodDetails.WriteFormulaToCell: entry");
            try
            {
                //--- Step 1: Mandatory field validations

                var ledgerVal = GetFieldValue(LedgerField, "LedgerName");
                if (string.IsNullOrWhiteSpace(ledgerVal))
                {
                    ShowWarningAction?.Invoke("Ledger is a mandatory field.");
                    return false;
                }

                var periodVal = GetFieldValue(PeriodField, "PeriodName");
                if (string.IsNullOrWhiteSpace(periodVal))
                {
                    ShowWarningAction?.Invoke("Period is a mandatory field.");
                    return false;
                }

                //--- Step 8: Build formula parts
                var formulaParts = new List<string>
                    {
                        FormatFormulaArg(periodVal),
                        FormatFormulaArg(ledgerVal)

                    };

                //--- Step 9: Construct and write formula
                string strFormula = string.Empty;
                if (!string.IsNullOrWhiteSpace(_formulaName))
                {

                    switch (_formulaName)
                    {
                        case "NUM":
                            strFormula = "GLSense_GetPeriodNum";
                            break;
                        case "YEAR":
                            strFormula = "GLSense_GetPeriodYear";
                            break;
                        case "QTR":
                            strFormula = "GLSense_GetPeriodQuarter";
                            break;
                        case Start:
                            formulaParts.Add(IsAdjacentChecked ? "TRUE" : "FALSE");
                            strFormula = "GLSense_GetPeriodStart";
                            break;
                        case "END":
                            formulaParts.Add(IsAdjacentChecked ? "TRUE" : "FALSE");
                            strFormula = "GLSense_GetPeriodEnd";
                            break;
                    }
                }
                rng.NumberFormat = AppConstants.General;
                var finalFormula = "=@" + $"{strFormula}(" + string.Join(",", formulaParts) + ")";
                rng.Formula = finalFormula;
                LogUtility.LogDebug($"GLPeriodDetails.WriteFormulaToCell: wrote formula '{finalFormula}'");
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLPeriodDetails.WriteFormulaToCell");
                ShowWarningAction?.Invoke("Exception encountered while writing formula to excel cell." + Environment.NewLine + ex.Message);
                return false;
            }

        }
        private static string FormatFormulaArg(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                // Truly empty argument → ""
                return "\"\"";
            }

            // Cell reference (contains ! or $) → do not quote
            if (value.Contains("!") || value.Contains("$"))
            {
                return value;
            }

            // Already a concatenation (e.g., period & "~" & endPeriod)
            if (value.Contains("&"))
            {
                return value;
            }

            // Check if value is numeric (int, float, double, etc.)
            if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                // Numeric values - return as is (no quotes)
                return value;
            }

            // Otherwise, quote everything (removing embedded quotes if any)
            return $"\"{value.Replace("\"", "")}\"";
        }
        private static string GetFieldValue(ComboFieldBindings field, string propertyName)
        {
            if (field == null) return string.Empty;

            // Step 1: Highest priority — RefValue
            var refVal = field.RefValue;
            if (!string.IsNullOrWhiteSpace(refVal))
            {
                return refVal.Trim();
            }

            // Step 2: Handle ComboValue (can be String or model)
            var comboVal = field.ComboValue;

            if (comboVal == null) return string.Empty;

            // If it's a simple string
            if (comboVal is string str)
            {
                return str.Trim();
            }

            // Step 3: If it's a model — find a suitable display name

            var val = comboVal.GetType()
                              .GetProperties()
                              .FirstOrDefault(p => string.Equals(p.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                              ?.GetValue(comboVal)?.ToString().Trim();

            return !string.IsNullOrEmpty(val) ? val : comboVal.ToString();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // ---------------- Excel interop for refedit controls ----------------
        private Microsoft.Office.Interop.Excel.Application _excelApp;
        public Microsoft.Office.Interop.Excel.Application ExcelApp
        {
            get => _excelApp;
            set
            {
                _excelApp = value;
                OnPropertyChanged(nameof(ExcelApp));
            }
        }
    }
}
