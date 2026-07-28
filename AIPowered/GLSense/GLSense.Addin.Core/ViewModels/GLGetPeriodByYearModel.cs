// GLGetPeriodByYearModel.cs in GLSense.Addin.Core
// Port of GLSense\ViewModels\GLGetPeriodByYearModel.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers) - backs the GLGetPeriodByYear view (GLSense_GetPeriodByYear
// formula: period year + period num).
// Re-pointed the same way as GLGetPeriodModel.cs (see that file's header for the full
// mapping). Also dropped an unused "using ControlzEx.Standard;" from the original -
// nothing in this class references ControlzEx, and GLSense.Addin.Core does not
// reference it.
using GLSense.Addin.Core.Bindings;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Services;
using GLSense.Addin.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace GLSense.Addin.Core.ViewModels
{
    public class GLGetPeriodByYearModel : INotifyPropertyChanged, IFieldDependencyProvider
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;
        public ObservableCollection<GenericLedgerModel> Ledgers { get; set; } = new ObservableCollection<GenericLedgerModel>();
        public ObservableCollection<PeriodModel> Periods { get; set; } = new ObservableCollection<PeriodModel>();
        public ObservableCollection<int> PeriodYears { get; set; } = new ObservableCollection<int>();
        public ObservableCollection<int> PeriodNums { get; set; } = new ObservableCollection<int>();

        public ComboFieldBindings LedgerField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings PeriodYearField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings PeriodNumField { get; set; } = new ComboFieldBindings();

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

        public GLGetPeriodByYearModel(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;

            LedgerField.DependencyProvider = this;
            LedgerField.Type = ComboFieldBindings.FieldType.Ledger;

            PeriodYearField.DependencyProvider = this;
            PeriodYearField.Type = ComboFieldBindings.FieldType.PeriodYear;

            PeriodNumField.DependencyProvider = this;
            PeriodNumField.Type = ComboFieldBindings.FieldType.PeriodNum;
        }
        public async Task LoadDataAsync(List<string> FuncArgs = null, List<string> FuncValues = null)
        {
            ServiceLocator.Logger?.LogDebug($"GLGetPeriodByYearModel.LoadDataAsync started. FuncArgs count={FuncArgs?.Count ?? 0}");
            try
            {
                ServiceLocator.Logger?.LogDebug($"GLGetPeriodByYearModel.LoadDataAsync: calling DataRepository.GetConfiguratorLedgers for CubeId={AppState.Instance.SelectedCube?.CubeId}, CoaId={AppState.Instance.SelectedLedger?.CoaId}");
                var ledgers = await Task.Run(() =>
                {
                    var repository = new DataRepository();
                    return repository.GetConfiguratorLedgers(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.CoaId, true);
                });
                ServiceLocator.Logger?.LogDebug($"GLGetPeriodByYearModel.LoadDataAsync: GetConfiguratorLedgers returned {ledgers?.Count ?? 0} ledger(s).");

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
                ServiceLocator.Logger?.LogException(ex, "GLGetPeriodByYearModel.LoadDataAsync");
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

                ResetFieldBinding(PeriodYearField, PeriodYears);
                PopulatePeriodYears();
            }
        }
        private GenericLedgerModel FindLedgerMatch(string ledgerName)
        {
            return Ledgers.FirstOrDefault(x => x.LedgerName == ledgerName);
        }
        private void ProcessLedgerField(string funcArg, GenericLedgerModel match)
        {
            if (ExcelRangeHelper.IsRealRange(funcArg))
            {
                LedgerField.ComboValue = null;
                LedgerField.IsValueFromRefEdit = true;
                LedgerField.RefValue = funcArg;
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

            try
            {
                if (ShowBusyAction != null)
                    await ShowBusyAction("Loading periods...", null);

                ServiceLocator.Logger?.LogDebug($"GLGetPeriodByYearModel.LoadPeriodsForLedger: loading periods for ledger \"{ledger.LedgerName}\"");
                var periods = await Task.Run(() => DataServiceLocator.PeriodDataService.GetPeriodsForLedger(ledger.LedgerName));
                ServiceLocator.Logger?.LogDebug($"GLGetPeriodByYearModel.LoadPeriodsForLedger: GetPeriodsForLedger(\"{ledger.LedgerName}\") returned {periods?.Count ?? 0} period(s).");

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

                await _dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            }
            finally
            {
                if (HideBusyAsyncAction != null)
                    await HideBusyAsyncAction();
            }
        }
        private void ProcessPeriodYearField(string funcArg)
        {
            try
            {
                if (ExcelRangeHelper.IsRealRange(funcArg))
                {
                    PeriodYearField.ComboValue = null;
                    PeriodYearField.IsValueFromRefEdit = true;
                    PeriodYearField.RefValue = funcArg;
                }
                else
                {
                    if (PeriodYears != null && PeriodYears.Count > 0 && int.TryParse(funcArg, out int periodYearValue) && PeriodYears.Contains(periodYearValue))
                    {
                        PeriodYearField.ComboValue = periodYearValue;
                        PeriodYearField.RefValue = null;
                    }
                }
                PeriodYearField.RefreshEnableState();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error processing period year field with argument \"{funcArg}\": {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }

        }
        private void ProcessPeriodNumField(string funcArg)
        {
            try
            {
                if (ExcelRangeHelper.IsRealRange(funcArg))
                {
                    PeriodNumField.ComboValue = null;
                    PeriodNumField.IsValueFromRefEdit = true;
                    PeriodNumField.RefValue = funcArg;
                }
                else
                {
                    if (PeriodNums != null && PeriodNums.Count > 0 && int.TryParse(funcArg, out int periodNumValue) && PeriodNums.Contains(periodNumValue))
                    {
                        PeriodNumField.ComboValue = periodNumValue;
                        PeriodNumField.RefValue = null;
                    }
                }
                PeriodNumField.RefreshEnableState();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error processing period num field with argument \"{funcArg}\": {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
        }
        private async Task ApplyFormulaParams(List<string> FuncArgs, List<string> FuncValues)
        {
            ServiceLocator.Logger?.LogDebug($"GLGetPeriodByYearModel.ApplyFormulaParams: FuncArgs=[{string.Join(",", FuncArgs ?? new List<string>())}], FuncValues=[{string.Join(",", FuncValues ?? new List<string>())}]");
            try
            {
                string ledgerName;

                if (FuncValues.Count >= 3)
                {
                    ledgerName = FuncValues[2].Replace("\"", "");
                }
                else
                {
                    ledgerName = AppState.Instance.SelectedLedger.LedgerName;
                }

                GenericLedgerModel ledgerMatch = FindLedgerMatch(ledgerName);
                if (ledgerMatch == null) return;

                ProcessLedgerField(ledgerName.Replace("\"", ""), ledgerMatch);
                await LoadPeriodsForLedger(ledgerMatch);
                await _dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
                ProcessPeriodYearField(FuncArgs[0].Replace("\"", ""));
                await _dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
                ProcessPeriodNumField(FuncArgs[1].Replace("\"", ""));
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error applying formula parameters in GLGetPeriodByYearModel: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
            finally
            {
                ResultOutPut();
            }
        }
        private static void ResetFieldBinding(ComboFieldBindings field, ObservableCollection<int> Model)
        {
            Model.Clear();
            field.ComboText = string.Empty;
            field.ComboValue = null;
            field.RefValue = null;
            field.IsValueFromRefEdit = false;
            field.RefreshEnableState();
        }
        private void ResetLedgerDependentFields()
        {
            ResetFieldBinding(PeriodYearField, PeriodYears);
            ResetFieldBinding(PeriodNumField, PeriodNums);
            Periods = new ObservableCollection<PeriodModel>();
            ResultText = string.Empty;
        }
        // IFieldDependencyProvider implementations
        public bool IsRefEnabled(ComboFieldBindings field)
        {
            if (field.IsValueFromRefEdit)
                return true;

            return field.ComboValue == null;
        }

        public bool IsComboEnabled(ComboFieldBindings field)
        {
            if (field.IsValueFromRefEdit)
                return false;

            return string.IsNullOrEmpty(field.RefValue);
        }
        public async Task OnFieldDependencyChanged(ComboFieldBindings field)
        {
            ServiceLocator.Logger?.LogDebug($"GLGetPeriodByYearModel.OnFieldDependencyChanged: field type={field?.Type}");
            // Only one control enabled at a time
            LedgerField.RefreshEnableState();
            PeriodYearField.RefreshEnableState();
            PeriodNumField.RefreshEnableState();

            if (field.Type == ComboFieldBindings.FieldType.Ledger)
            {
                ResetLedgerDependentFields();

                if (field.ComboValue is GenericLedgerModel Ledger)
                {
                    await LoadPeriodsForLedger(Ledger);

                    ResetFieldBinding(PeriodYearField, PeriodYears);
                    PopulatePeriodYears();
                }
            }
            if (field.Type == ComboFieldBindings.FieldType.PeriodYear && field.ComboValue != null)
            {
                ResetFieldBinding(PeriodNumField, PeriodNums);
                PopulatePeriodNums();
            }
            ResultOutPut();
        }
        private static void ResetField(ComboFieldBindings field)
        {
            field.IsValueFromRefEdit = false;
            field.ComboValue = null;
            field.RefValue = null;
            field.RefreshEnableState();
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
            switch (field.Type)
            {
                case ComboFieldBindings.FieldType.Ledger: return MatchLedger(field, cellValue);
                case ComboFieldBindings.FieldType.PeriodYear: return MatchPeriodYear(field, cellValue);
                case ComboFieldBindings.FieldType.PeriodNum: return MatchPeriodNum(field, cellValue);
                default: return false;
            }
        }
        private bool MatchLedger(ComboFieldBindings field, string cellValue)
        {
            var match = Ledgers.FirstOrDefault(x => x.LedgerName.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                field.ComboValue = match;
                field.RefreshEnableState();
                ResetFieldBinding(PeriodYearField, PeriodYears);
                PopulatePeriodYears();
                return true;
            }
            HandleMatchFailure(field, cellValue);
            return false;
        }
        private bool MatchPeriodYear(ComboFieldBindings field, string cellValue)
        {
            if (int.TryParse(cellValue, out int number))
            {
                if (PeriodYears.Contains(number))
                {
                    field.ComboValue = number;
                    field.RefreshEnableState();
                    ResetFieldBinding(PeriodNumField, PeriodNums);
                    PopulatePeriodNums();
                    return true;
                }
                HandleMatchFailure(field, cellValue);
                return false;
            }
            HandleMatchFailure(field, cellValue);
            return false;
        }
        private bool MatchPeriodNum(ComboFieldBindings field, string cellValue)
        {
            if (int.TryParse(cellValue, out int number))
            {
                if (PeriodNums.Contains(number))
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
            ServiceLocator.Logger?.LogWarn($"GLGetPeriodByYearModel.HandleEmptyCellReference: referenced cell \"{newText}\" is empty for field type={field?.Type}");
            ShowWarningAction?.Invoke($"The referenced cell \"{newText}\" is empty");
        }
        private void HandleMatchFailure(ComboFieldBindings field, string cellValue)
        {
            ResetField(field);
            string message;
            switch (field.Type)
            {
                case ComboFieldBindings.FieldType.Ledger: message = $"Ledger \"{cellValue}\" not found in available ledgers."; break;
                case ComboFieldBindings.FieldType.PeriodYear: message = $"Year \"{cellValue}\" not found in available period years."; break;
                case ComboFieldBindings.FieldType.PeriodNum: message = $"Num \"{cellValue}\" not found in available period nums."; break;
                default: message = "Item not found."; break;
            }
            ServiceLocator.Logger?.LogWarn($"GLGetPeriodByYearModel.HandleMatchFailure: {message}");
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

        private void PopulatePeriodYears()
        {
            PeriodYears.Clear();
            var distinctYears = Periods.Select(p => p.PeriodYear).Distinct().OrderBy(y => y);
            foreach (var year in distinctYears)
            {
                PeriodYears.Add(year);
            }
        }
        private void PopulatePeriodNums()
        {

            if (PeriodYearField.ComboValue == null) return;

            int periodYear = (int)PeriodYearField.ComboValue;

            PeriodNums.Clear();

            var nums = Periods
                       .Where(p => p.PeriodYear == periodYear)
                       .Select(p => p.PeriodNum)
                       .Distinct()
                       .OrderBy(n => n);

            foreach (var num in nums)
            {
                PeriodNums.Add(num);
            }
        }

        private string GetCellValueIfRange(string refText)
        {
            if (string.IsNullOrWhiteSpace(refText)) return null;
            try
            {
                return ExcelApp?.Range[refText]?.Value2?.ToString();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLGetPeriodByYearModel.GetCellValueIfRange");
                return null;
            }
        }

        private void ResultOutPut()
        {
            try
            {
                ResultText = string.Empty;

                if (LedgerField.ComboValue == null && (LedgerField.RefValue == null || string.IsNullOrWhiteSpace(LedgerField.RefValue.ToString())))
                {
                    return;
                }
                if (PeriodYearField.ComboValue == null && (PeriodYearField.RefValue == null || string.IsNullOrWhiteSpace(PeriodYearField.RefValue.ToString())))
                {
                    return;
                }
                if ((PeriodNumField.ComboValue == null || string.IsNullOrWhiteSpace(PeriodNumField.ComboValue.ToString())) &&
                    (PeriodNumField.RefValue == null || string.IsNullOrWhiteSpace(PeriodNumField.RefValue.ToString())))
                {
                    return;
                }
                if (PeriodYears == null || PeriodYears.Count == 0 || PeriodNums == null || PeriodNums.Count == 0)
                {
                    return;
                }

                var period = Periods
                    .FirstOrDefault(p => p.PeriodYear == (int)PeriodYearField.ComboValue && p.PeriodNum == (int)PeriodNumField.ComboValue);

                ResultText = period?.PeriodName ?? string.Empty;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
            }
            finally
            {
                OnPropertyChanged(nameof(ResultText));
            }
        }

        public bool WriteFormulaToCell(Microsoft.Office.Interop.Excel.Range rng)
        {
            ServiceLocator.Logger?.LogDebug("GLGetPeriodByYearModel.WriteFormulaToCell started.");
            try
            {
                //--- Step 1: Mandatory field validations

                var ledgerVal = GetFieldValue(LedgerField, "LedgerName");
                if (string.IsNullOrWhiteSpace(ledgerVal))
                {
                    ShowWarningAction?.Invoke("Ledger is a mandatory field.");
                    return false;
                }

                var periodYearVal = GetFieldValue(PeriodYearField, string.Empty);
                if (string.IsNullOrWhiteSpace(periodYearVal))
                {
                    ShowWarningAction?.Invoke("Period year is a mandatory field.");
                    return false;
                }

                var periodNumVal = GetFieldValue(PeriodNumField, string.Empty);
                if (string.IsNullOrWhiteSpace(periodNumVal))
                {
                    ShowWarningAction?.Invoke("Period num is a mandatory field.");
                    return false;
                }

                //--- Step 8: Build formula parts
                var formulaParts = new List<string>
                    {
                        FormatFormulaArg(periodYearVal),
                        FormatFormulaArg(periodNumVal),
                        FormatFormulaArg(ledgerVal)

                    };

                //--- Step 9: Construct and write formula
                var finalFormula = "=@" + "GLSense_GetPeriodByYear(" + string.Join(",", formulaParts) + ")";
                rng.Formula = finalFormula;

                return true;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex);
                ShowWarningAction?.Invoke("Exception encountered while writing formula to excel cell." + Environment.NewLine + ex.Message);
                return false;
            }

        }
        private static string FormatFormulaArg(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "\"\"";
            }

            if (value.Contains("!") || value.Contains("$"))
            {
                return value;
            }

            if (value.Contains("&"))
            {
                return value;
            }

            if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                return value;
            }

            return $"\"{value.Replace("\"", "")}\"";
        }
        private static string GetFieldValue(ComboFieldBindings field, string propertyName)
        {
            if (field == null) return string.Empty;

            var refVal = field.RefValue;
            if (!string.IsNullOrWhiteSpace(refVal))
            {
                return refVal.Trim();
            }

            var comboVal = field.ComboValue;

            if (comboVal == null) return string.Empty;

            if (comboVal is string str)
            {
                return str.Trim();
            }

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
