using GLSense.Bindings;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace GLSense.ViewModels
{
    public class GLDailyRatesViewModel : INotifyPropertyChanged, IFieldDependencyProviderNonAsync
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        public ObservableCollection<string> Currencies { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<CurrencyModel> CurrenciesModel { get; set; }

        public ComboFieldBindings FromCurrencyField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings ToCurrencyField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings DateField { get; set; } = new ComboFieldBindings();

        private readonly Dispatcher _dispatcher;

        public Action<string> ShowWarningAction { get; set; }

        private string _conversionType;
        public string ConversionType
        {
            get => _conversionType;
            set
            {
                _conversionType = value;
                OnPropertyChanged(nameof(ConversionType));
            }
        }

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

        public GLDailyRatesViewModel(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;

            FromCurrencyField.DependencyProviderNonAsync = this;
            FromCurrencyField.Type = ComboFieldBindings.FieldType.Currency;

            ToCurrencyField.DependencyProviderNonAsync = this;
            ToCurrencyField.Type = ComboFieldBindings.FieldType.Currency;

            DateField.DependencyProviderNonAsync = this;
            DateField.Type = ComboFieldBindings.FieldType.Date;

        }
        public async Task LoadDataAsync( List<string> FuncArgs = null)
        {
            LogUtility.LogDebug($"GLDailyRatesViewModel.LoadDataAsync: FuncArgs.Count={FuncArgs?.Count ?? 0}");
            await Task.Run(async () =>
            {
                var repository = new DataRepository();
                var currenciesData = repository.GetCurrencies(AppState.Instance.SelectedCube.CubeId, AppState.Instance.SelectedLedger.LedgerId);
                LogUtility.LogDebug($"GLDailyRatesViewModel.LoadDataAsync: loaded {currenciesData?.Count ?? 0} currency(ies) for CubeId={AppState.Instance.SelectedCube.CubeId}, LedgerId={AppState.Instance.SelectedLedger.LedgerId}");

                await _dispatcher.InvokeAsync(() =>
                {
                    if (CurrenciesModel == null)
                        CurrenciesModel = new ObservableCollection<CurrencyModel>();
                    else
                    {
                        CurrenciesModel.Clear();
                    }

                    CurrenciesModel = currenciesData;

                    foreach (var currency in CurrenciesModel)
                    {
                        Currencies.Add(currency.CurrencyCode);
                    }

                    if (FuncArgs != null && FuncArgs.Count > 0)
                    {
                        ApplyFormulaParams(FuncArgs);
                    }
                    else
                    {
                        FromCurrencyField.ComboValue = AppState.Instance.SelectedLedger.CurrencyCode;
                    }
                });

            });
        }
        private void ProcessCurrencyField(ComboFieldBindings field, string funcArg)
        {
            if (string.IsNullOrWhiteSpace(funcArg))
            {
                ResetField(field);
                return;
            }

            if (ExcelRangeHelper.IsRealRange(funcArg))
            {
                field.ComboValue = null;
                field.IsValueFromRefEdit = true;
                field.RefValue = funcArg;
            }
            else
            {
                bool exists = Currencies.Any(c => string.Equals(c, funcArg, StringComparison.OrdinalIgnoreCase));
                if (exists)
                {
                    field.ComboValue = funcArg;
                    field.RefValue = null;
                }

            }
            field.RefreshEnableState();
        }
        private void ProcessDateField(ComboFieldBindings field, string funcArg)
        {
            if (string.IsNullOrWhiteSpace(funcArg))
            {
                SelectedDate = null;
                ResetField(field);
                return;
            }

            if (ExcelRangeHelper.IsRealRange(funcArg))
            {
                field.ComboValue = null;
                field.IsValueFromRefEdit = true;
                field.RefValue = funcArg;
            }
            else
            {
                if (TryParseExcelDate(funcArg, out DateTime dt))
                {
                    SelectedDate = dt;
                    field.ComboValue = dt;
                    field.RefValue = null;
                }

            }
            field.RefreshEnableState();
        }
        private void ProcessConversionTypeField(string funcArg)
        {
            if (string.IsNullOrWhiteSpace(funcArg))
            {
                ConversionType = string.Empty;
                return;
            }

            if (ExcelRangeHelper.IsRealRange(funcArg))
            {
                string rngValue = GetCellValueIfRange(funcArg);
                ConversionType = rngValue;
            }
            else
            {
                ConversionType = funcArg;
            }
        }
        private void ApplyFormulaParams(List<string> FuncArgs)
        {
            LogUtility.LogDebug($"GLDailyRatesViewModel.ApplyFormulaParams: FuncArgs.Count={FuncArgs?.Count ?? 0}");

            if (Currencies == null || Currencies.Count == 0)
            {
                LogUtility.LogWarn($"GLDailyRatesViewModel.ApplyFormulaParams: no currencies loaded for ledger \"{AppState.Instance.SelectedLedger.LedgerName}\", aborting formula param processing.");
                ShowWarningAction?.Invoke($"Unable to fetch the segments list for the ledger \"{AppState.Instance.SelectedLedger.LedgerName}\"");
                return;
            }

            ProcessCurrencyField(FromCurrencyField, FuncArgs.Count > 0 ? FuncArgs[0].Replace("\"", "") : string.Empty);
            ProcessCurrencyField(ToCurrencyField, FuncArgs.Count > 1 ? FuncArgs[1].Replace("\"", "") : string.Empty);
            ProcessConversionTypeField(FuncArgs.Count > 2 ? FuncArgs[2].Replace("\"", "") : string.Empty);
            ProcessDateField(DateField, FuncArgs.Count > 3 ? FuncArgs[3].Replace("\"", "") : string.Empty);
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

        public void OnFieldDependencyChanged(ComboFieldBindings field)
        {
            //No implementation needed for now
        }
        private static void ResetField(ComboFieldBindings field)
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
        private void HandleEmptyCellReference(ComboFieldBindings field, string newText)
        {
            ResetField(field);
            ShowWarningAction?.Invoke($"The referenced cell \"{newText}\" is empty");
        }
        private bool TryMatchFieldValue(ComboFieldBindings field, string cellValue)
        {
            return field.Type switch
            {
                ComboFieldBindings.FieldType.Currency => MatchCurrency(field, cellValue),
                ComboFieldBindings.FieldType.Date => MatchDate(field, cellValue),
                _ => false
            };
        }
        private bool MatchCurrency(ComboFieldBindings field, string cellValue)
        {
            bool exists = Currencies.Any(c => string.Equals(c, cellValue, StringComparison.OrdinalIgnoreCase));

            if (exists)
            {
                field.IsValueFromRefEdit = true;     // ⭐ IMPORTANT
                field.ComboValue = cellValue;
                field.RefreshEnableState();
                return true;
            }

            HandleMatchFailure(field, cellValue);
            return false;
        }
        private bool MatchDate(ComboFieldBindings field, string cellValue)
        {
            if (TryParseExcelDate(cellValue, out DateTime dt))
            {

                field.IsValueFromRefEdit = true;     // ⭐ IMPORTANT
                field.ComboValue = dt;   // valid → update DatePicker
                SelectedDate = dt;
                OnPropertyChanged(nameof(SelectedDate));
                field.RefreshEnableState();
                return true;
            }

            HandleMatchFailure(field, cellValue);
            return false;
        }
        private void HandleMatchFailure(ComboFieldBindings field, string cellValue)
        {
            ResetField(field);
            string message = field.Type switch
            {
                ComboFieldBindings.FieldType.Currency => $"Currency \"{cellValue}\" not found in available currencies list.",
                ComboFieldBindings.FieldType.Date => $"Date \"{cellValue}\" is inavlid date.",
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
                return;
            }


            if (!string.IsNullOrWhiteSpace(field.RefValue) && field.IsRefEnabled)
            {
                field.IsValueFromRefEdit = true;
            }

            ProcessRefEditValue(field, newText);
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
        public bool WriteFormulaToCell(Microsoft.Office.Interop.Excel.Range rng)
        {
            LogUtility.LogDebug("GLDailyRatesViewModel.WriteFormulaToCell: entry");
            try
            {
                //--- Step 1: Mandatory field validations

                var fromCurrency = GetFieldValue(FromCurrencyField, string.Empty);
                if (string.IsNullOrWhiteSpace(fromCurrency))
                {
                    ShowWarningAction?.Invoke("Currency from is a mandatory field.");
                    return false;
                }

                var toCurrency = GetFieldValue(ToCurrencyField, string.Empty);
                if (string.IsNullOrWhiteSpace(toCurrency))
                {
                    ShowWarningAction?.Invoke("Currency to is a mandatory field.");
                    return false;
                }

                var conversionDate = GetFieldValue(DateField, string.Empty);
                if (string.IsNullOrWhiteSpace(conversionDate))
                {
                    ShowWarningAction?.Invoke("Date is a mandatory field.");
                    return false;
                }
                else
                {
                    if (!conversionDate.Contains("$") && TryParseExcelDate(conversionDate, out DateTime dt))
                    {
                        conversionDate = dt.ToString("yyyy-MM-dd"); // or any desired date format
                    }
                }

                var conversionType = string.IsNullOrWhiteSpace(ConversionType) ? string.Empty : ConversionType;
                if (string.IsNullOrWhiteSpace(conversionType))
                {
                    ShowWarningAction?.Invoke("Conversion type is a mandatory field.");
                    return false;
                }

                //--- Step 8: Build formula parts
                var formulaParts = new List<string>
                    {
                        FormatFormulaArg(fromCurrency),
                        FormatFormulaArg(toCurrency),
                        FormatFormulaArg(conversionType),
                        FormatFormulaArg(conversionDate)
                    };


                var finalFormula = "=@" + $"GLSense_GetDailyRate(" + string.Join(",", formulaParts) + ")";
                rng.Formula = finalFormula;
                LogUtility.LogDebug($"GLDailyRatesViewModel.WriteFormulaToCell: wrote formula '{finalFormula}'");
                return true;
            }
            catch (Exception ex)
            {
                ShowWarningAction?.Invoke("Exception encountered while writing formula to excel cell." + Environment.NewLine + ex.Message);
                LogUtility.LogException(ex, "GLDailyRatesViewModel.WriteFormulaToCell");
                return false;
            }

        }
        private static bool TryParseExcelDate(string input, out DateTime result)
        {
            result = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(input))
                return false;

            // 1️⃣ If numeric → treat as OADate
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

            // 2️⃣ Try normal DateTime.Parse
            if (DateTime.TryParse(input, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                return true;

            // 3️⃣ Try your strict formats
            string[] formats = { "dd-MMM-yyyy", "d-MMM-yyyy", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd" };

            if (DateTime.TryParseExact(input,
                                       formats,
                                       CultureInfo.InvariantCulture,
                                       DateTimeStyles.AllowWhiteSpaces,
                                       out result))
                return true;

            return false;
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

            if (field.Type == ComboFieldBindings.FieldType.Segment)
            {
                return comboVal is SegmentModel seg ? seg.SegmentName.Trim() : comboVal.ToString();
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
