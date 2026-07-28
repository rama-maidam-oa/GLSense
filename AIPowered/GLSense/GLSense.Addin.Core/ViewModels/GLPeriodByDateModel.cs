// GLPeriodByDateModel.cs in GLSense.Addin.Core
// Port of GLSense\ViewModels\GLPeriodByDateModel.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers) - backs the GLGetPeriodByDate view (GLSense_GetPeriodByDate
// formula: date + ledger + numeric offset).
// Re-pointed the same way as GLGetPeriodModel.cs (see that file's header for the full
// mapping). Also dropped an unused "using static System.Windows.Forms.AxHost;" from the
// original (dead leftover import, not referenced anywhere in the class body) - this
// project's Addin.Core does not reference System.Windows.Forms from ViewModels.
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
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace GLSense.Addin.Core.ViewModels
{
    public class GLPeriodByDateModel : INotifyPropertyChanged, IFieldDependencyProvider
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        public ObservableCollection<GenericLedgerModel> Ledgers { get; set; } = new ObservableCollection<GenericLedgerModel>();
        public ObservableCollection<PeriodModel> Periods { get; set; } = new ObservableCollection<PeriodModel>();

        public ObservableCollection<int> OffsetValues { get; set; } = new ObservableCollection<int>();

        public ComboFieldBindings LedgerField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings DateField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings OffsetField { get; set; } = new ComboFieldBindings();

        private readonly Dispatcher _dispatcher;

        public Action<string> ShowWarningAction { get; set; }
        public Func<string, Task> ShowWarningAsyncAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

        private DateTime? _selectedDate;
        public DateTime? SelectedDate
        {
            get => _selectedDate;
            set
            {
                if (_selectedDate != value)
                {
                    _selectedDate = value;
                    OnPropertyChanged(nameof(SelectedDate));
                    DateField.ComboValue = value;

                    if (value.HasValue)
                    {
                        UpdateOffsetValues();
                    }
                    else
                    {
                        ClearOffsetSelection();
                    }
                }
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
        public GLPeriodByDateModel(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;

            LedgerField.DependencyProvider = this;
            LedgerField.Type = ComboFieldBindings.FieldType.Ledger;

            DateField.DependencyProvider = this;
            DateField.Type = ComboFieldBindings.FieldType.Date;

            OffsetField.DependencyProvider = this;
            OffsetField.Type = ComboFieldBindings.FieldType.Offset;

        }
        public async Task LoadDataAsync(List<string> FuncArgs = null, List<string> FuncValues = null)
        {
            ServiceLocator.Logger?.LogDebug($"GLPeriodByDateModel.LoadDataAsync started. FuncArgs count={FuncArgs?.Count ?? 0}");
            try
            {
                ServiceLocator.Logger?.LogDebug($"GLPeriodByDateModel.LoadDataAsync: calling DataRepository.GetConfiguratorLedgers for CubeId={AppState.Instance.SelectedCube?.CubeId}, CoaId={AppState.Instance.SelectedLedger?.CoaId}");
                var ledgers = await Task.Run(() =>
                {
                    var repository = new DataRepository();
                    return repository.GetConfiguratorLedgers(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.CoaId, true);
                });
                ServiceLocator.Logger?.LogDebug($"GLPeriodByDateModel.LoadDataAsync: GetConfiguratorLedgers returned {ledgers?.Count ?? 0} ledger(s).");

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
                ServiceLocator.Logger?.LogException(ex, "GLPeriodByDateModel.LoadDataAsync");
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
        private Task LoadAndProcessPeriods(GenericLedgerModel ledger)
        {
            return LoadPeriodsForLedger(ledger);
        }
        private async Task LoadPeriodsForLedger(GenericLedgerModel ledger)
        {
            if (ledger == null)
                return;

            try
            {
                if (ShowBusyAction != null)
                    await ShowBusyAction("Loading periods...", null);

                ServiceLocator.Logger?.LogDebug($"GLPeriodByDateModel.LoadPeriodsForLedger: loading periods for ledger \"{ledger.LedgerName}\"");
                var periods = await Task.Run(() => DataServiceLocator.PeriodDataService.GetPeriodsForLedger(ledger.LedgerName));
                ServiceLocator.Logger?.LogDebug($"GLPeriodByDateModel.LoadPeriodsForLedger: GetPeriodsForLedger(\"{ledger.LedgerName}\") returned {periods?.Count ?? 0} period(s).");

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
        private void SetDateFromMatch(string date)
        {
            if (date.IndexOf("DATE(", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                DateTime? result = TryParseDateFunction(date);
                if (result.HasValue)
                {
                    SelectedDate = result.Value;
                    OnPropertyChanged(nameof(SelectedDate));
                    DateField.ComboValue = result.Value;
                    DateField.IsValueFromRefEdit = false;
                }
                else
                {
                    DateField.ComboValue = null;
                }
            }
            else
            {
                if (TryParseExcelDate(date, out DateTime dt))
                {
                    SelectedDate = dt;
                    OnPropertyChanged(nameof(SelectedDate));
                    DateField.ComboValue = dt;
                    DateField.IsValueFromRefEdit = false;
                }
                else
                {
                    DateField.ComboValue = null;
                }
            }
        }
        private void ProcessDateField(string funcArg, string funcValue)
        {
            string date = funcValue.Replace("\"", "");

            if (ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", "")))
            {
                DateField.ComboValue = null;
                DateField.IsValueFromRefEdit = true;
                DateField.RefValue = funcArg.Replace("\"", "");
            }
            else
                SetDateFromMatch(date);

            DateField.RefreshEnableState();
        }
        private void ProcessOffsetField(string funcArg, string funcValue)
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
                    .FirstOrDefault(x => SelectedDate >= x.Period.StartDate.Date && SelectedDate <= x.Period.EndDate.Date);
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
            ServiceLocator.Logger?.LogDebug($"GLPeriodByDateModel.ApplyFormulaParams: FuncArgs=[{string.Join(",", FuncArgs ?? new List<string>())}], FuncValues=[{string.Join(",", FuncValues ?? new List<string>())}]");
            string ledgerName = FuncValues[1].Replace("\"", "");
            if (string.IsNullOrEmpty(ledgerName)) return;

            var ledgerMatch = FindLedgerMatch(ledgerName);
            if (ledgerMatch == null) return;

            ProcessLedgerField(FuncArgs[1], ledgerMatch);
            await LoadAndProcessPeriods(ledgerMatch);
            await _dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            ProcessDateField(FuncArgs[0], FuncValues[0]);
            await _dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
            ProcessOffsetField(FuncArgs[2], FuncValues[2]);

            ResultOutPut();
        }
        private static DateTime? TryParseDateFunction(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return null;

            try
            {
                var match = Regex.Match(dateString, @"DATE\((\d{4}),(\d{1,2}),(\d{1,2})\)");

                if (match.Success)
                {
                    int year = int.Parse(match.Groups[1].Value);
                    int month = int.Parse(match.Groups[2].Value);
                    int day = int.Parse(match.Groups[3].Value);

                    return new DateTime(year, month, day, 0, 0, 0, DateTimeKind.Utc);
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogError($"Error parsing date function: {ex.Message}");
            }

            return null;
        }
        private void ResetField(ComboFieldBindings field)
        {
            ResetFieldToEmpty(field);
            if (field.Type == ComboFieldBindings.FieldType.Date)
            {
                SelectedDate = null;
                OnPropertyChanged(nameof(SelectedDate));
            }
            field.RefreshEnableState();
        }
        private static void ResetFieldToEmpty(ComboFieldBindings field)
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
                case ComboFieldBindings.FieldType.Offset: return MatchOffset(field, cellValue);
                case ComboFieldBindings.FieldType.Date: return MatchDate(field, cellValue);
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
        private bool MatchDate(ComboFieldBindings field, string cellValue)
        {
            if (TryParseExcelDate(cellValue, out DateTime dt))
            {
                field.ComboValue = dt;
                field.RefreshEnableState();
                return true;
            }
            HandleMatchFailure(field, cellValue);
            return false;
        }
        private void HandleEmptyCellReference(ComboFieldBindings field, string newText)
        {
            ResetField(field);
            ServiceLocator.Logger?.LogWarn($"GLPeriodByDateModel.HandleEmptyCellReference: referenced cell \"{newText}\" is empty for field type={field?.Type}");
            ShowWarningAction?.Invoke($"The referenced cell \"{newText}\" is empty");
        }
        private void HandleMatchFailure(ComboFieldBindings field, string cellValue)
        {
            ResetField(field);
            string message;
            switch (field.Type)
            {
                case ComboFieldBindings.FieldType.Ledger: message = $"Ledger \"{cellValue}\" not found in available ledgers."; break;
                case ComboFieldBindings.FieldType.Offset: message = $"Offset \"{cellValue}\" not found in available offsets."; break;
                case ComboFieldBindings.FieldType.Date: message = $"Date \"{cellValue}\" is not a valid date."; break;
                default: message = "Item not found."; break;
            }
            ServiceLocator.Logger?.LogWarn($"GLPeriodByDateModel.HandleMatchFailure: {message}");
            ShowWarningAction?.Invoke(message);
        }
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
            ServiceLocator.Logger?.LogDebug($"GLPeriodByDateModel.OnFieldDependencyChanged: field type={field?.Type}");
            LedgerField.RefreshEnableState();
            DateField.RefreshEnableState();
            OffsetField.RefreshEnableState();

            if (field.Type == ComboFieldBindings.FieldType.Ledger)
            {
                ClearOffsetSelection();
                DateField.IsValueFromRefEdit = false;
                DateField.ComboText = string.Empty;
                DateField.ComboValue = null;
                DateField.RefValue = null;
                DateField.RefreshEnableState();
                SelectedDate = null;
                ResultText = string.Empty;

                if (field.ComboValue is GenericLedgerModel Ledger)
                {
                    await LoadPeriodsForLedger(Ledger);
                }
            }

            if (field.Type == ComboFieldBindings.FieldType.Date && field.ComboValue is DateTime dt)
            {
                SelectedDate = dt;
                OnPropertyChanged(nameof(SelectedDate));
                PopulateOffsets(dt);
            }

            ResultOutPut();
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

        private static bool TryParseExcelDate(string input, out DateTime result)
        {
            result = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(input))
                return false;

            // If numeric -> treat as OADate
            if (double.TryParse(input, out double oa))
            {
                try
                {
                    result = DateTime.FromOADate(oa);
                    return true;
                }
                catch
                {
                    // ignore and continue to other parse attempts
                }
            }

            // Try normal DateTime.Parse
            if (DateTime.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                return true;

            // Try strict formats
            string[] formats = new[]
            {
                "dd-MMM-yyyy", "d-MMM-yyyy", "dd/MM/yyyy", "d/M/yyyy",
                "MM/dd/yyyy", "M/d/yyyy", "dd-MM-yyyy", "d-M-yyyy",
                "yyyy-MM-dd", "yyyy/MM/dd", "dd MMM yyyy", "d MMM yyyy",
                "dd.MM.yyyy", "d.M.yyyy"
            };


            if (DateTime.TryParseExact(input,
                                       formats,
                                       CultureInfo.InvariantCulture,
                                       DateTimeStyles.AllowWhiteSpaces,
                                       out result))
                return true;

            return false;
        }

        private void PopulateOffsets(DateTime selectedDate)
        {
            OffsetValues.Clear();

            if (!Periods.Any())
            {
                ShowWarningAction?.Invoke("No periods are available for the selected ledger.");
                OffsetField.ComboValue = null;
                return;
            }

            var selectedPeriodIndex = Periods
                .Select((p, i) => new { Period = p, Index = i })
                .FirstOrDefault(x => selectedDate.Date >= x.Period.StartDate.Date && selectedDate <= x.Period.EndDate.Date);

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

        private void ClearOffsetSelection()
        {
            OffsetValues.Clear();
            OffsetField.IsValueFromRefEdit = false;
            OffsetField.ComboValue = null;
            OffsetField.RefValue = null;
            OffsetField.RefreshEnableState();
        }

        private void UpdateOffsetValues()
        {
            OffsetValues.Clear();

            if (!SelectedDate.HasValue)
            {
                ClearOffsetSelection();
                return;
            }

            if (Periods == null || Periods.Count == 0)
            {
                ShowWarningAction?.Invoke("No periods available.");
                return;
            }

            // Find selected period containing the date
            var selPeriod = Periods.FirstOrDefault(p =>
                SelectedDate.Value.Date >= p.StartDate.Date &&
                SelectedDate.Value.Date <= p.EndDate.Date);

            if (selPeriod == null)
            {
                DateTime? firstStartDate = null;
                DateTime? lastEndDate = null;
                string msgStr = string.Empty;
                if (Periods != null && Periods.Count > 0)
                {
                    firstStartDate = Periods.Min(p => p.StartDate);
                    lastEndDate = Periods.Max(p => p.EndDate);

                    msgStr = "No period found for selected date." + Environment.NewLine + $"Calendar is between \"{firstStartDate:dd-MMM-yyyy}\" and \"{lastEndDate:dd-MMM-yyyy}\".";
                }
                ShowWarningAction?.Invoke(msgStr);
                DateField.ComboValue = null;
                DateField.RefValue = null;
                SelectedDate = null;
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
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLPeriodByDateModel.GetCellValueIfRange");
                return null;
            }
        }
        private bool HasValidInputs()
        {
            return HasValidField(LedgerField) &&
                   HasValidField(DateField) &&
                   HasValidField(OffsetField) &&
                   SelectedDate.HasValue &&
                   Periods?.Count > 0;
        }
        private static bool HasValidField(ComboFieldBindings field)
        {
            return field.ComboValue != null ||
                   !(field.RefValue == null || string.IsNullOrWhiteSpace(field.RefValue.ToString()));
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
        private PeriodModel GetSelectedPeriod()
        {
            return Periods.FirstOrDefault(p =>
                    SelectedDate.Value.Date >= p.StartDate.Date &&
                    SelectedDate.Value.Date <= p.EndDate.Date);
        }
        private void ResultOutPut()
        {
            try
            {
                ResultText = GetResultText();
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
            ServiceLocator.Logger?.LogDebug("GLPeriodByDateModel.WriteFormulaToCell started.");
            try
            {
                //--- Step 1: Mandatory field validations

                var ledgerVal = GetFieldValue(LedgerField, "LedgerName");
                if (string.IsNullOrWhiteSpace(ledgerVal))
                {
                    ShowWarningAction?.Invoke("Ledger is a mandatory field.");
                    return false;
                }

                var dateVal = GetFieldValue(DateField, string.Empty);
                if (string.IsNullOrWhiteSpace(dateVal))
                {
                    ShowWarningAction?.Invoke("Date is a mandatory field.");
                    return false;
                }
                else
                {
                    if (!dateVal.Contains("$") && TryParseExcelDate(dateVal, out DateTime dt))
                    {
                        dateVal = $"DATE({dt.Year},{dt.Month},{dt.Day})";
                    }
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
                        FormatFormulaArg(dateVal).Replace("\"", ""),
                        FormatFormulaArg(ledgerVal),
                        FormatFormulaArg(offsetVal)
                    };

                //--- Step 9: Construct and write formula
                var finalFormula = "=@" + "GLSense_GetPeriodByDate(" + string.Join(",", formulaParts) + ")";
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
