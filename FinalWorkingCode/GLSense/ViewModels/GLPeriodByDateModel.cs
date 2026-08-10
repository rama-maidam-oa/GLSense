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
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using static System.Windows.Forms.AxHost;

namespace GLSense.ViewModels
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
            LogUtility.LogDebug($"GLPeriodByDateModel.LoadDataAsync: FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
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
                LogUtility.LogException(ex, "GLPeriodByDateModel.LoadDataAsync");
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
            LogUtility.LogDebug("GLPeriodByDateModel.ApplyDefaultLedgerSelection: entry");
            if (AppState.Instance.SelectedLedger == null)
            {
                LogUtility.LogWarn("GLPeriodByDateModel.ApplyDefaultLedgerSelection: AppState.Instance.SelectedLedger is null, aborting default ledger selection.");
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
                LogUtility.LogWarn($"GLPeriodByDateModel.ApplyDefaultLedgerSelection: no ledger match found for LedgerId={AppState.Instance.SelectedLedger.LedgerId}.");
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
                LedgerField.ComboValue=null;
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

            LogUtility.LogDebug($"GLPeriodByDateModel.LoadPeriodsForLedger: LedgerName={ledger.LedgerName}");
            try
            {
                if (ShowBusyAction != null)
                    await ShowBusyAction("Loading periods...", null);

                var periods = await Task.Run(() => ServiceLocator.PeriodDataService.GetPeriodsForLedger(ledger.LedgerName));
                LogUtility.LogDebug($"GLPeriodByDateModel.LoadPeriodsForLedger: loaded {periods?.Count ?? 0} period(s) for LedgerName={ledger.LedgerName}");

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
            LogUtility.LogDebug($"GLPeriodByDateModel.ApplyFormulaParams: FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            string ledgerName = FuncValues[1].Replace("\"", "");
            if (string.IsNullOrEmpty(ledgerName))
            {
                LogUtility.LogWarn("GLPeriodByDateModel.ApplyFormulaParams: resolved ledgerName is empty, aborting.");
                return;
            }

            var ledgerMatch = FindLedgerMatch(ledgerName);
            if (ledgerMatch == null)
            {
                LogUtility.LogWarn($"GLPeriodByDateModel.ApplyFormulaParams: no ledger match found for '{ledgerName}'.");
                return;
            }

            ProcessLedgerField(FuncArgs[1], ledgerMatch);
            await LoadAndProcessPeriods(ledgerMatch);
            await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
            ProcessDateField(FuncArgs[0], FuncValues[0]);
            await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
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
                LogUtility.LogException(ex, "GLPeriodByDateModel.TryParseDateFunction");
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
            field.IsValueFromRefEdit = false;  // ⭐ RESET FLAG
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
            return field.Type switch
            {
                ComboFieldBindings.FieldType.Ledger => MatchLedger(field, cellValue),
                ComboFieldBindings.FieldType.Offset => MatchOffset(field, cellValue),
                ComboFieldBindings.FieldType.Date => MatchDate(field, cellValue),
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
            ShowWarningAction?.Invoke($"The referenced cell \"{newText}\" is empty");
        }
        private void HandleMatchFailure(ComboFieldBindings field, string cellValue)
        {
            ResetField(field);
            string message = field.Type switch
            {
                ComboFieldBindings.FieldType.Ledger => $"Ledger \"{cellValue}\" not found in available ledgers.",
                ComboFieldBindings.FieldType.Offset => $"Offset \"{cellValue}\" not found in available offsets.",
                ComboFieldBindings.FieldType.Date => $"Date \"{cellValue}\" is not a valid date.",
                _ => "Item not found."
            };
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
            LogUtility.LogDebug($"GLPeriodByDateModel.OnFieldDependencyChanged: FieldType={field?.Type}, ComboValue={field?.ComboValue}, RefValue={field?.RefValue}");
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
            LogUtility.LogDebug($"GLPeriodByDateModel.OnRefEditTextChanged: FieldType={field?.Type}, newText={newText}");
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

            // 1?? If numeric ? treat as OADate
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

            // 2?? Try normal DateTime.Parse
            if (DateTime.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                return true;

            // 3?? Try your strict formats

            string[] formats = [
                            "dd-MMM-yyyy", "d-MMM-yyyy", "dd/MM/yyyy", "d/M/yyyy",
                            "MM/dd/yyyy", "M/d/yyyy", "dd-MM-yyyy", "d-M-yyyy",
                            "yyyy-MM-dd", "yyyy/MM/dd", "dd MMM yyyy", "d MMM yyyy",
                            "dd.MM.yyyy", "d.M.yyyy"
                        ];


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
            catch
            {
                // refText isn't a resolvable range (e.g. still mid-typing) - can be ignored as expected.
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
                LogUtility.LogException(ex, "GLPeriodByDateModel.ResultOutPut");
            }
            finally
            {
                OnPropertyChanged(nameof(ResultText));
            }
        }

        public bool WriteFormulaToCell(Microsoft.Office.Interop.Excel.Range rng)
        {
            LogUtility.LogDebug("GLPeriodByDateModel.WriteFormulaToCell: entry");
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
                LogUtility.LogDebug($"GLPeriodByDateModel.WriteFormulaToCell: wrote formula '{finalFormula}'");
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLPeriodByDateModel.WriteFormulaToCell");
                ShowWarningAction?.Invoke("Exception encountered while writing formula to excel cell." + Environment.NewLine + ex.Message);
                return false;
            }

        }
        private static string FormatFormulaArg(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                // Truly empty argument ? ""
                return "\"\"";
            }

            // Cell reference (contains ! or $) ? do not quote
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
