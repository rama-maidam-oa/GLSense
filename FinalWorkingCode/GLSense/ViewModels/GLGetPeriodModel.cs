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
    public class GLGetPeriodModel : INotifyPropertyChanged, IFieldDependencyProvider
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        public ObservableCollection<GenericLedgerModel> Ledgers { get; set; } = new ObservableCollection<GenericLedgerModel>();
        public ObservableCollection<PeriodModel> Periods { get; set; } = new ObservableCollection<PeriodModel>();
        public ObservableCollection<int> OffsetValues { get; set; } = new ObservableCollection<int>();

        public ComboFieldBindings LedgerField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings PeriodField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings OffsetField { get; set; } = new ComboFieldBindings();

        private readonly Dispatcher _dispatcher;

        public Action<string> ShowWarningAction { get; set; }
        public Func<string, Task> ShowWarningAsyncAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

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
                UpdateOffsetValues();
            }
        }
        public GLGetPeriodModel(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;

            LedgerField.DependencyProvider = this;
            LedgerField.Type = ComboFieldBindings.FieldType.Ledger;

            PeriodField.DependencyProvider = this;
            PeriodField.Type = ComboFieldBindings.FieldType.Period;

            OffsetField.DependencyProvider = this;
            OffsetField.Type = ComboFieldBindings.FieldType.Offset;
        }
        public async Task LoadDataAsync(List<string> FuncArgs = null, List<string> FuncValues = null)
        {
            LogUtility.LogDebug($"GLGetPeriodModel.LoadDataAsync: FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
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
                {
                    await ApplyFormulaParams(FuncArgs, FuncValues);
                }
                else
                {
                    await ApplyDefaultLedgerSelection();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodModel.LoadDataAsync");
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
            if (AppState.Instance.SelectedLedger == null) return;
            var match = Ledgers.FirstOrDefault(x => x.LedgerId == AppState.Instance.SelectedLedger.LedgerId);
            if (match != null)
            {
                LedgerField.ComboValue = match;
                LedgerField.RefValue = null;
                LedgerField.RefreshEnableState();

                await LoadPeriodsForLedger(match);
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
        private PeriodModel FindPeriodMatch(string periodName)
        {
            return Periods.FirstOrDefault(x => x.PeriodName == periodName);
        }
        private void ProcessPeriodField(string funcArg, PeriodModel match)
        {
            if (ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", "")))
            {
                PeriodField.ComboValue = null;
                PeriodField.IsValueFromRefEdit = true;
                PeriodField.RefValue = funcArg.Replace("\"", "");
            }
            else
            {
                PeriodField.ComboValue = match;
                PeriodField.RefValue = null;
            }
            PeriodField.RefreshEnableState();
        }
        private Task LoadAndProcessPeriods(GenericLedgerModel ledger)
        {
            return LoadPeriodsForLedger(ledger);
        }
        private async Task LoadPeriodsForLedger(GenericLedgerModel ledger)
        {
            if (ledger == null)
                return;

            LogUtility.LogDebug($"GLGetPeriodModel.LoadPeriodsForLedger: LedgerName={ledger.LedgerName}");
            try
            {
                if (ShowBusyAction != null)
                    await ShowBusyAction("Loading periods...", null);

                var periods = await Task.Run(() => ServiceLocator.PeriodDataService.GetPeriodsForLedger(ledger.LedgerName));
                LogUtility.LogDebug($"GLGetPeriodModel.LoadPeriodsForLedger: loaded {periods?.Count ?? 0} period(s) for LedgerName={ledger.LedgerName}");

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
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"GLGetPeriodModel.LoadPeriodsForLedger: LedgerName={ledger.LedgerName}");
                throw;
            }
            finally
            {
                if (HideBusyAsyncAction != null)
                    await HideBusyAsyncAction();
            }
        }
        private void ProcessOffsetField(string funcArg, string funcValue, string periodName)
        {
            string offset = funcValue.Replace("\"", "");
            if (ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", "")))
            {
                OffsetField.ComboValue = null;
                OffsetField.IsValueFromRefEdit = true;
                OffsetField.RefValue = funcArg.Replace("\"", "");
            }
            else
            {
                var selectedPeriodIndex = Periods
                    .Select((p, i) => new { Period = p, Index = i })
                    .FirstOrDefault(x => x.Period.PeriodName == periodName);
                if (selectedPeriodIndex != null)
                {
                    OffsetField.ComboValue = Convert.ToInt32(offset);
                    OffsetField.RefValue = null;
                }
            }
            OffsetField.RefreshEnableState();
        }
        private async Task ApplyFormulaParams(List<string> FuncArgs, List<string> FuncValues)
        {
            LogUtility.LogDebug($"GLGetPeriodModel.ApplyFormulaParams: FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            try
            {
                string ledgerName = AppState.Instance.SelectedLedger.LedgerName;

                if (FuncValues.Count >= 3)
                {
                    ledgerName = FuncValues[2].Replace("\"", "");
                }

                GenericLedgerModel ledgerMatch = FindLedgerMatch(ledgerName);
                if (ledgerMatch == null)
                {
                    LogUtility.LogWarn($"GLGetPeriodModel.ApplyFormulaParams: no ledger match found for '{ledgerName}'.");
                    return;
                }

                ProcessLedgerField(ledgerName, ledgerMatch);
                await LoadAndProcessPeriods(ledgerMatch);
                await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                PeriodModel periodModel = FindPeriodMatch(FuncValues[0].Replace("\"", ""));
                if (periodModel == null)
                {
                    LogUtility.LogWarn($"GLGetPeriodModel.ApplyFormulaParams: no period match found for '{FuncValues[0]}'.");
                    return;
                }

                ProcessPeriodField(FuncArgs[0], periodModel);
                await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                ProcessOffsetField(FuncArgs[1], FuncValues[1], periodModel.PeriodName);

                ResultOutPut();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodModel.ApplyFormulaParams");
            }
        }
        // IFieldDependencyProvider implementations
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
            LogUtility.LogDebug($"GLGetPeriodModel.OnFieldDependencyChanged: FieldType={field?.Type}, ComboValue={field?.ComboValue}, RefValue={field?.RefValue}");
            if (field.ComboValue == null && field.IsValueFromRefEdit)
            {
                ResetField(field);  // This will re-enable both controls
                return;
            }

            // Only one control enabled at a time
            LedgerField.RefreshEnableState();
            PeriodField.RefreshEnableState();
            OffsetField.RefreshEnableState();

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
                PopulateOffsets();
            }
            ResultOutPut();
        }
        private void ResetField(ComboFieldBindings field)
        {
            field.IsValueFromRefEdit = false;  // ⭐ RESET FLAG
            field.ComboText = string.Empty;
            field.ComboValue = null;
            field.RefValue = null;
            field.RefreshEnableState();
            SelectedPeriod = null;
            OnPropertyChanged(nameof(SelectedPeriod));
        }
        private void ResetLedgerDependentFields()
        {
            ResetField(PeriodField);
            ResetField(OffsetField);
            Periods = new ObservableCollection<PeriodModel>();
            OffsetValues.Clear();
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
                ComboFieldBindings.FieldType.Offset => MatchOffset(field, cellValue),
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
        private bool MatchOffset(ComboFieldBindings field, string cellValue)
        {
            if (int.TryParse(cellValue, out int number))
            {
                if (OffsetValues.Contains(number))
                {
                    field.ComboValue = number;
                    field.RefreshEnableState();
                    return true;
                }
                HandleMatchFailure(field, cellValue);
                return false;
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
                ComboFieldBindings.FieldType.Offset => $"Offset \"{cellValue}\" not found in available offsets.",
                ComboFieldBindings.FieldType.Period => $"Period \"{cellValue}\" not found in available periods.",
                _ => "Item not found."
            };
            ShowWarningAction?.Invoke(message);
        }
        public void OnRefEditTextChanged(ComboFieldBindings field, string newText)
        {
            if (string.IsNullOrWhiteSpace(newText))
            {
                ResetField(field);
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

        private void PopulateOffsets()
        {
            OffsetValues.Clear();

            if (!Periods.Any())
            {
                ShowWarningAction?.Invoke("No periods are available for the selected ledger.");
                OffsetField.ComboValue = null;
                return;
            }

            if (SelectedPeriod == null)
            {
                ShowWarningAction?.Invoke("No period is selected.");
                OffsetField.ComboValue = null;
                return;
            }

            var selectedPeriodIndex = Periods
                .Select((p, i) => new { Period = p, Index = i })
                .FirstOrDefault(x => x.Period.PeriodName == SelectedPeriod.PeriodName);

            if (selectedPeriodIndex == null)
            {

                ShowWarningAction?.Invoke($"No periods found for the selected date.");
                OffsetField.ComboValue = null;
                return;
            }

            int index = selectedPeriodIndex.Index;

            for (int i = -index; i < Periods.Count - index; i++)
                OffsetValues.Add(i);

            OffsetField.ComboValue = null;
            OffsetField.RefValue = null; // Ensure RefEdit cleared
            OffsetField.RefreshEnableState();
        }

        private void UpdateOffsetValues()
        {
            OffsetValues.Clear();

            if (SelectedPeriod == null)
            {
                return;
            }

            if (Periods == null || Periods.Count == 0)
            {
                ShowWarningAction?.Invoke("No periods available.");
                return;
            }

            // Find selected period containing the date
            var selPeriod = Periods.FirstOrDefault(p =>
                SelectedPeriod.PeriodName == p.PeriodName);

            if (selPeriod == null)
            {
                PeriodField.ComboValue = null;
                PeriodField.RefValue = null;
                SelectedPeriod = null;
                return;
            }

            int index = Periods.IndexOf(selPeriod);

            for (int i = -index; i < Periods.Count - index; i++)
            {
                OffsetValues.Add(i);
            }

            // Auto-select OffsetField if RefValue exists
            if (!string.IsNullOrWhiteSpace(OffsetField.RefValue))
            {
                string rngValue = GetCellValueIfRange(OffsetField.RefValue);
                if (int.TryParse(rngValue, out int offset) && OffsetValues.Contains(offset))
                {
                    OffsetField.ComboValue = offset;
                    OffsetField.RefreshEnableState();
                    return;
                }
            }
            OffsetField.RefreshEnableState();
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
        private bool HasValidInputs()
        {
            return HasValidField(LedgerField) &&
                   HasValidField(PeriodField) &&
                   HasValidField(OffsetField) &&
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
        private string GetResultText()
        {
            if (!HasValidInputs())
                return string.Empty;

            var selPeriod = GetSelectedPeriod();
            if (selPeriod == null)
                return string.Empty;

            // Convert periods to list of period names (ordered by date)
            var periodNames = Periods
                .OrderBy(p => p.StartDate)
                .Select(p => p.PeriodName)
                .ToList();

            // Find index of current period
            int currentIndex = periodNames.IndexOf(selPeriod.PeriodName);

            if (currentIndex == -1)
            {
                return string.Empty;
            }

            int? SelectedOffset = OffsetField.ComboValue as int?;

            int targetIndex = currentIndex + SelectedOffset.Value;

            if (targetIndex >= 0 && targetIndex < periodNames.Count)
            {
                string periodName = periodNames[targetIndex];

                return periodName;
            }

            return string.Empty;
        }
        private void ResultOutPut()
        {
            try
            {
                ResultText = GetResultText();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodModel.ResultOutPut");
            }
            finally
            {
                OnPropertyChanged(nameof(ResultText));
            }
        }

        public bool WriteFormulaToCell(Microsoft.Office.Interop.Excel.Range rng)
        {
            LogUtility.LogDebug("GLGetPeriodModel.WriteFormulaToCell: entry");
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

                var offsetVal = GetFieldValue(OffsetField, string.Empty);
                if (string.IsNullOrWhiteSpace(offsetVal))
                {
                    ShowWarningAction?.Invoke("Offset is a mandatory field.");
                    return false;
                }

                //--- Step 8: Build formula parts
                var formulaParts = new List<string>
                    {
                        FormatFormulaArg(periodVal),
                        FormatFormulaArg(offsetVal),
                        FormatFormulaArg(ledgerVal)

                    };

                //--- Step 9: Construct and write formula
                var finalFormula = "=@" + "GLSense_GetPeriod(" + string.Join(",", formulaParts) + ")";
                rng.Formula = finalFormula;
                LogUtility.LogDebug($"GLGetPeriodModel.WriteFormulaToCell: wrote formula '{finalFormula}'");
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLGetPeriodModel.WriteFormulaToCell");
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
