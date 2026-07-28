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
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace GLSense.ViewModels
{
    public class GLSegmentFuncsViewModel : INotifyPropertyChanged, IFieldDependencyProviderNonAsync
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        private readonly Dispatcher _dispatcher;
        private readonly string _formulaName;
        private bool _isApplyingFormulaParams;

        private const string NextSegment = "NEXTSEGMENT";
        private const string PreviousSegment = "PREVIOUSSEGMENT";
        private const string StringTrue = "TRUE";
        private const string StringFalse = "FALSE";
        private const string DescriptionFormula = "DESCRIPTION";

        public ObservableCollection<GenericLedgerModel> Ledgers { get; set; } = new ObservableCollection<GenericLedgerModel>();

        public ComboFieldBindings LedgerField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings SegmentField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings SegmentValueField { get; set; } = new ComboFieldBindings();
        public ComboFieldBindings AttributeField { get; set; } = new ComboFieldBindings();


        public Action<string> ShowWarningAction { get; set; }
        public Func<string, Task> ShowWarningAsyncAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

        private ObservableCollection<SegmentModel> _segments = new ObservableCollection<SegmentModel>();
        public ObservableCollection<SegmentModel> Segments
        {
            get => _segments;
            set
            {
                _segments = value;
                OnPropertyChanged();
            }
        }

        private SegmentModel _selectedSegment;
        public SegmentModel SelectedSegment
        {
            get => _selectedSegment;
            set
            {
                _selectedSegment = value;
                OnPropertyChanged();
                LoadSegmentValues();
            }
        }

        private ObservableCollection<SegmentValueModel> _segmentValues = new ObservableCollection<SegmentValueModel>();
        public ObservableCollection<SegmentValueModel> SegmentValues
        {
            get => _segmentValues;
            set
            {
                _segmentValues = value;
                OnPropertyChanged();
            }
        }

        private SegmentValueModel _selectedSegmentValue;
        public SegmentValueModel SelectedSegmentValue
        {
            get => _selectedSegmentValue;
            set
            {
                _selectedSegmentValue = value;
                OnPropertyChanged();
            }
        }

        public static ObservableCollection<AttributeTypeModel> AttributeTypes => AttributeTypeService.GetAttributesType();

        private AttributeTypeModel _selectedAttributeType;
        public AttributeTypeModel SelectedAttributeType
        {
            get => _selectedAttributeType;
            set => SetProperty(ref _selectedAttributeType, value);
        }

        // Search/filter
        private string _attributeText;
        public string AttributeText
        {
            get => _attributeText;
            set
            {
                _attributeText = value;
                OnPropertyChanged();
            }
        }

        // ===== Check parent checked =====
        private bool _isParentChecked;
        public bool IsParentChecked
        {
            get => _isParentChecked;
            set
            {
                _isParentChecked = value;
                OnPropertyChanged(nameof(IsParentChecked));
                ResultOutPut();
            }
        }
        // ===== Check child checked =====
        private bool _isChildChecked;
        public bool IsChildChecked
        {
            get => _isChildChecked;
            set
            {
                _isChildChecked = value;
                OnPropertyChanged(nameof(IsChildChecked));
                ResultOutPut();
            }
        }
        // ===== Check include parent value=====
        private bool _isParentValueChecked;
        public bool IsParentValueChecked
        {
            get => _isParentValueChecked;
            set
            {
                _isParentValueChecked = value;
                OnPropertyChanged(nameof(IsParentValueChecked));
                ResultOutPut();
            }
        }

        //Result Output
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

        public GLSegmentFuncsViewModel(Dispatcher dispatcher, string formulaName)
        {
            _dispatcher = dispatcher;
            _formulaName = formulaName;

            LedgerField.DependencyProviderNonAsync = this;
            LedgerField.Type = ComboFieldBindings.FieldType.Ledger;

            SegmentField.DependencyProviderNonAsync = this;
            SegmentField.Type = ComboFieldBindings.FieldType.Segment;

            SegmentValueField.DependencyProviderNonAsync = this;
            SegmentValueField.Type = ComboFieldBindings.FieldType.SegmentValue;

            AttributeField.DependencyProviderNonAsync = this;
            AttributeField.Type = ComboFieldBindings.FieldType.Attribute;

            if (_formulaName == NextSegment || _formulaName == PreviousSegment)
            {
                IsParentChecked = true;
                IsChildChecked = true;
            }
        }

        public async Task LoadDataAsync(long cubeId, long ledgerId, long CoaId, List<string> FuncArgs = null, List<string> FuncValues = null)
        {
            LogUtility.LogDebug($"GLSegmentFuncsViewModel.LoadDataAsync: cubeId={cubeId}, ledgerId={ledgerId}, CoaId={CoaId}, FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            try
            {
                if (ShowBusyAction != null)
                {
                    await ShowBusyAction("Loading ledgers...", null);
                }

                var ledgers = await Task.Run(() =>
                {
                    var repository = new DataRepository();
                    return repository.GetConfiguratorLedgers(cubeId, CoaId, true);
                });
                LogUtility.LogDebug($"GLSegmentFuncsViewModel.LoadDataAsync: loaded {ledgers?.Count ?? 0} ledger(s) for cubeId={cubeId}, CoaId={CoaId}");

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

                if (FuncArgs == null || FuncArgs.Count == 0)
                {
                    await ApplyDefaultLedgerSelectionAsync();
                }
                else
                {
                    string ledgerName = GetLedgerName(FuncValues);

                    await EnsureFormulaLedgerDataAsync(ledgerName);
                    await ApplyFormulaParamsAsync(ledgerName, FuncArgs, FuncValues);
                }

                if (HideBusyAsyncAction != null)
                {
                    await HideBusyAsyncAction();
                }

            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentFuncsViewModel.LoadDataAsync");
                if (HideBusyAsyncAction != null)
                {
                    await HideBusyAsyncAction();
                }
                throw;
            }
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
        private async Task ApplyDefaultLedgerSelectionAsync()
        {
            if (AppState.Instance.SelectedLedger == null) return;
            var match = Ledgers.FirstOrDefault(x => x.LedgerId == AppState.Instance.SelectedLedger.LedgerId);
            if (match != null)
            {
                LedgerField.ComboValue = match;
                LedgerField.RefValue = null;
                LedgerField.RefreshEnableState();
                await RefreshSegmentsForLedgerAsync(match);
            }
        }

        private async Task EnsureFormulaLedgerDataAsync(string ledgerName)
        {
            LogUtility.LogDebug($"GLSegmentFuncsViewModel.EnsureFormulaLedgerDataAsync: ledgerName={ledgerName}");

            if (string.IsNullOrWhiteSpace(ledgerName))
                return;

            GenericLedgerModel ledgerMatch = FindLedgerMatch(ledgerName);
            if (ledgerMatch == null)
            {
                LogUtility.LogWarn($"GLSegmentFuncsViewModel.EnsureFormulaLedgerDataAsync: no ledger match found for '{ledgerName}'.");
                return;
            }

            bool needsLoad = DataRepository.GetTableItemsCount(
                AppState.Instance.SelectedCube.CubeId,
                ledgerMatch.LedgerId,
                "SEGMENTS") == 0;

            if (needsLoad)
            {
                LogUtility.LogDebug($"GLSegmentFuncsViewModel.EnsureFormulaLedgerDataAsync: SEGMENTS not loaded for LedgerId={ledgerMatch.LedgerId}, triggering ReLoadDataAsync.");
                await ReLoadDataAsync(ledgerMatch);
            }
        }
        private void LoadSegmentValues()
        {
            LogUtility.LogDebug($"GLSegmentFuncsViewModel.LoadSegmentValues: SelectedSegment={SelectedSegment?.SegmentName}");
            SegmentValueField.ComboValue = null;
            SegmentValueField.RefValue = null;
            SelectedSegmentValue = null;

            if (SelectedSegment != null)
            {
                var segmentValues = DataRepository.GetSegmentValues(SelectedSegment);
                LogUtility.LogDebug($"GLSegmentFuncsViewModel.LoadSegmentValues: loaded {segmentValues?.Count ?? 0} segment value(s) for SegmentName={SelectedSegment.SegmentName}");
                if (SegmentValues == null)
                    SegmentValues = new ObservableCollection<SegmentValueModel>();
                else
                {
                    SegmentValues.Clear();
                }

                SegmentValues = segmentValues;
            }
            else
            {
                SegmentValues?.Clear();
            }
        }
        private bool EarlyExit(string ledgerName, List<string> FuncValues)
        {
            if (FuncValues == null || FuncValues.Count == 0)
            {
                LogUtility.LogWarn("GLSegmentFuncsViewModel.EarlyExit: FuncValues is null or empty, returning early.");
                return true;
            }

            if (ledgerName == null || string.IsNullOrEmpty(ledgerName) || ledgerName.Length == 0)
            {
                LogUtility.LogWarn("GLSegmentFuncsViewModel.EarlyExit: Ledger is null in the formula.");
                ShowWarningAction?.Invoke("Ledger is null in the formula");
                return true;
            }

            string sName = FuncValues[1].Replace("\"", "");


            if (sName == null || sName.Length == 0)
            {
                LogUtility.LogWarn("GLSegmentFuncsViewModel.EarlyExit: Segment is null in the formula.");
                ShowWarningAction?.Invoke("Segment is null in the formula");
                return true;
            }
            string sVal = FuncValues[0].Replace("\"", "");
            if (sVal == null || sVal.Length == 0)
            {
                LogUtility.LogWarn("GLSegmentFuncsViewModel.EarlyExit: Segment value is null in the formula.");
                ShowWarningAction?.Invoke("Segment value is null in the formula");
                return true;
            }
            return false;
        }
        private void ProcessSegmentField(string funcArg, string funcValue)
        {
            string cleanValue = funcValue.Replace("\"", "");
            if (string.IsNullOrEmpty(cleanValue))
            {
                ShowWarningAction?.Invoke("Segment is null in the formula");
                return;
            }

            if (ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", "")))
                SetSegmentAsReference(funcArg);
            else if (_formulaName == "ACCOUNTTYPE")
                SetSegmentFromIndex(cleanValue);
            else
                SetSegmentFromList(cleanValue);
        }
        private void SetSegmentAsReference(string funcArg)
        {
            SegmentField.ComboValue = null;
            SegmentField.IsValueFromRefEdit = true;
            SegmentField.RefValue = funcArg.Replace("\"", "");
        }
        private void SetSegmentFromList(string cleanValue)
        {
            var match = Segments.FirstOrDefault(x => x.SegmentName == cleanValue);
            if (match != null)
            {
                SegmentField.ComboValue = match;
                SegmentField.RefValue = null;
                return;
            }

            if (Segments == null || Segments.Count == 0)
                ShowWarningAction?.Invoke($"Unable to fetch the segments list for the ledger \"{AppState.Instance.SelectedLedger.LedgerName}\"");
            else
                ShowWarningAction?.Invoke($"Segment \"{cleanValue}\" not found in available segments.");
        }
        // ACCOUNTTYPE's formula carries the segment's 1-based dropdown position instead of
        // its name (see FormulaParameters/GetSelectedSegmentIndex) - reopening the picker
        // window from an existing cell has to reverse that: parse the number and select the
        // segment sitting at that position in Segments, instead of matching by name.
        private void SetSegmentFromIndex(string cleanValue)
        {
            if (!int.TryParse(cleanValue.Trim(), out int oneBasedIndex) ||
                oneBasedIndex < 1 || Segments == null || oneBasedIndex > Segments.Count)
            {
                if (Segments == null || Segments.Count == 0)
                    ShowWarningAction?.Invoke($"Unable to fetch the segments list for the ledger \"{AppState.Instance.SelectedLedger.LedgerName}\"");
                else
                    ShowWarningAction?.Invoke($"Segment index \"{cleanValue}\" is not valid for the currently loaded segments.");
                return;
            }

            SegmentField.ComboValue = Segments[oneBasedIndex - 1];
            SegmentField.RefValue = null;
        }
        private void ProcessSegmentValueField(string funcArg, string funcValue)
        {
            string cleanValue = funcValue.Replace("\"", "");
            if (string.IsNullOrEmpty(cleanValue))
            {
                ShowWarningAction?.Invoke("Segment value is null in the formula");
                return;
            }

            if (ExcelRangeHelper.IsRealRange(funcArg.Replace("\"", "")))
                SetSegmentValueAsReference(funcArg);
            else
                SetSegmentValuesFromList(cleanValue);
        }
        private void SetSegmentValueAsReference(string funcArg)
        {
            SegmentValueField.ComboValue = null;
            SegmentValueField.IsValueFromRefEdit = true;
            SegmentValueField.RefValue = funcArg.Replace("\"", "");
        }
        private void SetSegmentValuesFromList(string cleanValue)
        {
            var match = SegmentValues.FirstOrDefault(x => x.SegmentValue == cleanValue);
            if (match != null)
            {
                SegmentValueField.ComboValue = match;
                SegmentValueField.RefValue = null;
                return;
            }

            if (SegmentValues == null || SegmentValues.Count == 0)
                ShowWarningAction?.Invoke($"Unable to fetch segment values for ledger \"{AppState.Instance.SelectedLedger.LedgerName}\"");
            else
                ShowWarningAction?.Invoke($"Segment value \"{cleanValue}\" not found in available segment values.");
        }
        private async Task ApplyFormulaParamsAsync(string ledgerName, List<string> FuncArgs, List<string> FuncValues)
        {
            LogUtility.LogDebug($"GLSegmentFuncsViewModel.ApplyFormulaParamsAsync: ledgerName={ledgerName}, FuncArgs.Count={FuncArgs?.Count ?? 0}, FuncValues.Count={FuncValues?.Count ?? 0}");
            try
            {
                _isApplyingFormulaParams = true;

                if (EarlyExit(ledgerName,FuncValues)) return;

                GenericLedgerModel ledgerMatch = FindLedgerMatch(ledgerName);
                if (ledgerMatch == null)
                {
                    LogUtility.LogWarn($"GLSegmentFuncsViewModel.ApplyFormulaParamsAsync: no ledger match found for '{ledgerName}'.");
                    return;
                }

                ProcessLedgerField(FuncArgs[FuncArgs.Count - 1], ledgerMatch);
                await RefreshSegmentsForLedgerAsync(ledgerMatch);
                await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                ProcessSegmentField(FuncArgs[1], FuncValues[1]);
                await _dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);
                ProcessSegmentValueField(FuncArgs[0], FuncValues[0]);

                if (_formulaName == "DFF")
                {
                    if (ExcelRangeHelper.IsRealRange(FuncArgs[2].Replace("\"", "")))
                    {
                        AttributeField.ComboValue = null;
                        AttributeField.IsValueFromRefEdit = true;
                        AttributeField.RefValue = FuncArgs[2].Replace("\"", "");
                    }
                    else
                    {
                        string aName = FuncValues[2].Replace("\"", "");
                        var match = AttributeTypes.FirstOrDefault(x => x.Value == aName);
                        if (match != null)
                        {
                            AttributeField.ComboValue = match;
                            AttributeField.RefValue = null;
                        }
                    }
                }

                if (_formulaName == DescriptionFormula)
                {
                    string rawValue = FuncValues[2].Replace("\"", "").Trim();
                    bool.TryParse(rawValue, out bool pInclude);
                    IsParentValueChecked = pInclude;
                }
                if (_formulaName == NextSegment || _formulaName == PreviousSegment)
                {
                    string rawValue = FuncValues[2].Replace("\"", "").Trim();
                    bool.TryParse(rawValue, out bool pInclude);
                    IsParentChecked = pInclude;

                    string rawValue1 = FuncValues[3].Replace("\"", "").Trim();
                    bool.TryParse(rawValue1, out bool pInclude1);
                    IsChildChecked = pInclude1;
                }

                ResultOutPut();
            }
            finally
            {
                _isApplyingFormulaParams = false;
            }
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
        private GenericLedgerModel FindLedgerMatch(string ledgerName)
        {
            return Ledgers.FirstOrDefault(x => x.LedgerName == ledgerName);
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
            LogUtility.LogDebug($"GLSegmentFuncsViewModel.OnFieldDependencyChanged: FieldType={field?.Type}, ComboValue={field?.ComboValue}, RefValue={field?.RefValue}");
            if (field.Type == ComboFieldBindings.FieldType.Ledger && field.ComboValue != null && field.ComboValue is GenericLedgerModel ledger)
            {
                ResetLedgerDependentFields();
                if (!_isApplyingFormulaParams)
                {
                    _ = RefreshSegmentsForLedgerAsync(ledger);
                }
            }
            if (field.Type == ComboFieldBindings.FieldType.Segment && field.ComboValue != null && field.ComboValue is SegmentModel segment)
            {
                SelectedSegment = segment;
                OnPropertyChanged(nameof(SelectedSegment));
            }
            if (field.Type == ComboFieldBindings.FieldType.SegmentValue && field.ComboValue != null && field.ComboValue is SegmentValueModel segmentValue)
            {
                SelectedSegmentValue = segmentValue;
                OnPropertyChanged(nameof(SelectedSegmentValue));
            }
            if (field.Type == ComboFieldBindings.FieldType.Attribute && field.ComboValue != null && field.ComboValue is AttributeTypeModel attributeType)
            {
                SelectedAttributeType = attributeType;
                OnPropertyChanged(nameof(SelectedAttributeType));
                AttributeText = attributeType.Value;
                OnPropertyChanged(nameof(AttributeText));
            }
            ResultOutPut();
        }
        private async Task RefreshSegmentsForLedgerAsync(GenericLedgerModel ledger)
        {
            LogUtility.LogDebug($"GLSegmentFuncsViewModel.RefreshSegmentsForLedgerAsync: LedgerId={ledger?.LedgerId}, LedgerName={ledger?.LedgerName}");
            if (ledger == null || AppState.Instance.SelectedCube == null)
            {
                LogUtility.LogWarn("GLSegmentFuncsViewModel.RefreshSegmentsForLedgerAsync: ledger or SelectedCube is null, returning early.");
                return;
            }

            bool showBusy = false;
            try
            {
                Segments = new ObservableCollection<SegmentModel>();

                bool needsLoad = DataRepository.GetTableItemsCount(AppState.Instance.SelectedCube.CubeId, ledger.LedgerId, "SEGMENTS") == 0;
                showBusy = !_isApplyingFormulaParams && !needsLoad && ShowBusyAction != null;

                if (showBusy)
                {
                    await ShowBusyAction("Loading segments...", null);
                }

                if (needsLoad)
                {
                    LogUtility.LogDebug($"GLSegmentFuncsViewModel.RefreshSegmentsForLedgerAsync: SEGMENTS not loaded for LedgerId={ledger.LedgerId}, triggering ReLoadDataAsync.");
                    await ReLoadDataAsync(ledger);
                }

                var repository = new DataRepository();
                var segments = repository.GetSegments(AppState.Instance.SelectedCube.CubeId, ledger.LedgerId);
                LogUtility.LogDebug($"GLSegmentFuncsViewModel.RefreshSegmentsForLedgerAsync: loaded {segments?.Count ?? 0} segment(s) for LedgerId={ledger.LedgerId}");
                Segments = segments ?? new ObservableCollection<SegmentModel>();

                if (showBusy && HideBusyAsyncAction != null)
                {
                    await HideBusyAsyncAction();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentFuncsViewModel.RefreshSegmentsForLedgerAsync");
                ShowWarningAction?.Invoke($"Error refreshing segments for ledger \"{ledger.LedgerName}\": {ex.Message}");
                if (showBusy && HideBusyAsyncAction != null)
                {
                    await HideBusyAsyncAction();
                }
            }
        }
        private async Task ReLoadDataAsync(GenericLedgerModel ledger)
        {
            LogUtility.LogDebug($"GLSegmentFuncsViewModel.ReLoadDataAsync: LedgerId={ledger?.LedgerId}, LedgerName={ledger?.LedgerName}");
            if (ledger == null)
                return;

            CancellationHelper ctsHelper = new();

            try
            {
                if (ShowBusyAction != null)
                {
                    await ShowBusyAction("Fetching and transferring ledger data...",
                          async () => { ctsHelper.Cancel(); await Task.CompletedTask; });
                }

                LogUtility.LogDebug($"GLSegmentFuncsViewModel.ReLoadDataAsync: calling CommonFunctions.FillResponsibilitiesAsync for LedgerId={ledger.LedgerId}");
                await CommonFunctions.FillResponsibilitiesAsync(ledger.LedgerId, AppState.Instance.SelectedCube.CubeId, ctsHelper.GetToken());
                LogUtility.LogDebug($"GLSegmentFuncsViewModel.ReLoadDataAsync: FillResponsibilitiesAsync completed for LedgerId={ledger.LedgerId}");
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn($"GLSegmentFuncsViewModel.ReLoadDataAsync: operation cancelled by user for LedgerId={ledger.LedgerId}.");
                if (ShowWarningAsyncAction != null)
                    await ShowWarningAsyncAction("Operation cancelled by user.");
                return;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentFuncsViewModel.ReLoadDataAsync");
                if (ShowWarningAsyncAction != null)
                    await ShowWarningAsyncAction($"Error fetching ledger data: {ex.Message}");
            }
            finally
            {
                if (HideBusyAsyncAction != null)
                {
                    await HideBusyAsyncAction();
                }
            }
        }
        private void ClearField(ComboFieldBindings field)
        {
            ResetField(field);
            if (field.Type == ComboFieldBindings.FieldType.Ledger)
            {
                ResetLedgerDependentFields();
            }
            UpdateSelectedProperty(field.Type);
            field.RefreshEnableState();
        }
        private void ResetField(ComboFieldBindings field)
        {
            field.ComboValue = null;
            field.RefValue = null;
            field.ComboText = string.Empty;
            field.IsValueFromRefEdit = false;
        }
        private void ResetLedgerDependentFields()
        {
            ClearField(SegmentField);
            ClearField(SegmentValueField);
            ClearField(AttributeField);
            Segments = new ObservableCollection<SegmentModel>();
            SegmentValues = new ObservableCollection<SegmentValueModel>();
            ResultText = string.Empty;
        }
        private void UpdateSelectedProperty(ComboFieldBindings.FieldType fieldType)
        {
            switch (fieldType)
            {
                case ComboFieldBindings.FieldType.Segment:
                    SelectedSegment = null;
                    OnPropertyChanged(nameof(SelectedSegment));
                    SegmentValues?.Clear();
                    break;
                case ComboFieldBindings.FieldType.SegmentValue:
                    SelectedSegmentValue = null;
                    OnPropertyChanged(nameof(SelectedSegmentValue));
                    break;
                case ComboFieldBindings.FieldType.Attribute:
                    SelectedAttributeType = null;
                    OnPropertyChanged(nameof(SelectedAttributeType));
                    AttributeText = string.Empty;
                    OnPropertyChanged(nameof(AttributeText));
                    break;
            }
        }

        private void ProcessRefEditValue(ComboFieldBindings field, string newText)
        {
            string cellValue = GetCellValueIfRange(newText);
            if (string.IsNullOrEmpty(cellValue))
            {
                HandleEmptyCell(field, newText);
                return;
            }

            if (TrySetFieldFromLookup(field, cellValue))
                field.IsValueFromRefEdit = true;
            else
                HandleLookupFailure(field, cellValue);
        }
        private void HandleEmptyCell(ComboFieldBindings field, string newText)
        {
            field.IsValueFromRefEdit = false;  // ⭐ RESET FLAG
            field.ComboValue = null;
            field.RefValue = null;
            field.RefreshEnableState();
            ShowWarningAction?.Invoke($"The referenced cell \"{newText}\" is empty");
        }
        private void HandleLookupFailure(ComboFieldBindings field, string cellValue)
        {
            field.IsValueFromRefEdit = false;  // ⭐ RESET FLAG
            field.ComboValue = null;
            field.RefValue = null;
            field.RefreshEnableState();
            ShowWarningAction?.Invoke($"Value \"{cellValue}\" not found in available options.");
        }
        private bool TrySetFieldFromLookup(ComboFieldBindings field, string cellValue)
        {
            return field.Type switch
            {
                ComboFieldBindings.FieldType.Ledger => SetLedgerFromLookup(cellValue),
                ComboFieldBindings.FieldType.Segment => SetSegmentFromLookup(cellValue),
                ComboFieldBindings.FieldType.SegmentValue => SetSegmentValueFromLookup(cellValue),
                ComboFieldBindings.FieldType.Attribute => SetAttributeFromLookup(cellValue),
                _ => false
            };
        }
        private bool SetLedgerFromLookup(string cellValue)
        {
            var match = Ledgers.FirstOrDefault(x => x.LedgerName.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                LedgerField.IsValueFromRefEdit = true;     // ⭐ IMPORTANT
                LedgerField.ComboValue = match;
                LedgerField.RefreshEnableState();
                return true;
            }
            HandleMatchFailure(LedgerField, cellValue);
            return false;
        }
        private bool SetSegmentFromLookup(string cellValue)
        {
            // ACCOUNTTYPE's Segment field is picked by 1-based dropdown position, not name
            // (see FormulaParameters/GetSelectedSegmentIndex) - a live cell reference (typed
            // or selected while the window is open, not just on formula reopen) needs the
            // same index-based resolution instead of name matching, or a referenced cell
            // containing e.g. "1" would fail with "not found in available options" even
            // though 1 is a perfectly valid segment position.
            SegmentModel match = _formulaName == "ACCOUNTTYPE"
                ? (int.TryParse(cellValue?.Trim(), out int oneBasedIndex) &&
                   oneBasedIndex >= 1 && Segments != null && oneBasedIndex <= Segments.Count
                    ? Segments[oneBasedIndex - 1]
                    : null)
                : Segments.FirstOrDefault(x => x.SegmentName.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                SegmentField.IsValueFromRefEdit = true;     // ⭐ IMPORTANT
                SegmentField.ComboValue = match;
                SelectedSegment = match;
                OnPropertyChanged(nameof(SelectedSegment));
                SegmentField.RefreshEnableState();
                return true;
            }
            HandleMatchFailure(SegmentField, cellValue);
            return false;
        }
        private bool SetSegmentValueFromLookup(string cellValue)
        {
            var match = SegmentValues.FirstOrDefault(x => x.SegmentValue.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                SegmentValueField.IsValueFromRefEdit = true;     // ⭐ IMPORTANT
                SegmentValueField.ComboValue = match;
                SelectedSegmentValue = match;
                OnPropertyChanged(nameof(SelectedSegmentValue));
                SegmentValueField.RefreshEnableState();
                return true;
            }
            HandleMatchFailure(SegmentValueField, cellValue);
            return false;
        }
        private bool SetAttributeFromLookup(string cellValue)
        {
            var match = AttributeTypes.FirstOrDefault(x => x.Value.Equals(cellValue, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                AttributeField.IsValueFromRefEdit = true;     // ⭐ IMPORTANT
                AttributeField.ComboValue = match;
                SelectedAttributeType = match;
                AttributeText = match.Value;
                OnPropertyChanged(nameof(SelectedAttributeType));
                AttributeField.RefreshEnableState();
                return true;
            }
            HandleMatchFailure(AttributeField, cellValue);
            return false;
        }
        private void HandleMatchFailure(ComboFieldBindings field, string cellValue)
        {
            ResetField(field);
            if (field.Type == ComboFieldBindings.FieldType.Ledger)
            {
                ResetLedgerDependentFields();
            }
            string message = field.Type switch
            {
                ComboFieldBindings.FieldType.Ledger => $"Ledger \"{cellValue}\" not found in available ledgers.",
                ComboFieldBindings.FieldType.Segment => $"Segment \"{cellValue}\" not found in available segments.",
                ComboFieldBindings.FieldType.SegmentValue => $"Segment Value \"{cellValue}\" not found in available segment values.",
                ComboFieldBindings.FieldType.Attribute => $"Attribute \"{cellValue}\" not found in available attributes.",
                _ => "Item not found."
            };
            ShowWarningAction?.Invoke(message);
        }
        public void OnRefEditTextChanged(ComboFieldBindings field, string newText)
        {
            if (string.IsNullOrWhiteSpace(newText))
            {
                ClearField(field);
                return;
            }

            if (!string.IsNullOrWhiteSpace(field.RefValue))
                ProcessRefEditValue(field, newText);


            if (!string.IsNullOrWhiteSpace(field.RefValue) && field.IsRefEnabled)
            {
                field.IsValueFromRefEdit = true;
            }

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
        private bool HasValidSegmentInputs()
        {
            return HasValidField(SegmentField) && HasValidField(SegmentValueField);
        }
        private static bool HasValidField(ComboFieldBindings field)
        {
            if (field.ComboValue != null) return true;
            if (!string.IsNullOrWhiteSpace(field.RefValue)) return true;
            return false;
        }
        private string GetFormulaResult(SegmentValueModel selSegmentValue)
        {
            return _formulaName switch
            {
                "ENABLEDFLAG" => selSegmentValue.EnabledFlag,
                "SUMMARYFLAG" => selSegmentValue.SummaryFlag,
                DescriptionFormula => FormatDescription(selSegmentValue),
                NextSegment => GetAdjacentSegmentValue(true),
                PreviousSegment => GetAdjacentSegmentValue(false),
                _ => string.Empty
            };
        }
        private string FormatDescription(SegmentValueModel selSegmentValue)
        {
            string desc = selSegmentValue.Description;
            if (IsParentValueChecked)
            {
                desc = selSegmentValue.SegmentValue + " - " + desc;
            }
            return desc;
        }
        private string GetAdjacentSegmentValue(bool isNext)
        {
            var sel = GetAdjacentSegment(SegmentValues, SelectedSegmentValue, IsParentChecked, IsChildChecked, isNext);
            return sel != null ? sel.SegmentValue : string.Empty;
        }
        private void ResultOutPut()
        {
            try
            {
                ResultText = string.Empty;

                if (!HasValidSegmentInputs()) return;

                SegmentModel selSegment = (SegmentModel)SegmentField.ComboValue;
                SegmentValueModel selSegmentValue = (SegmentValueModel)SegmentValueField.ComboValue;

                if (selSegment == null || selSegmentValue == null) return;

                ResultText = GetFormulaResult(selSegmentValue);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLSegmentFuncsViewModel.ResultOutPut");
            }
            finally
            {
                OnPropertyChanged(nameof(ResultText));
            }
        }
        private static SegmentValueModel GetAdjacentSegment(
            ObservableCollection<SegmentValueModel> allSegments,
            SegmentValueModel selected,
            bool includeParent,
            bool includeChild,
            bool isNext)
        {
            if (allSegments == null || selected == null || allSegments.Count == 0)
                return null;

            int index = allSegments.IndexOf(selected);
            if (index < 0)
                return null;

            int step = isNext ? 1 : -1;
            int i = index + step;

            while (i >= 0 && i < allSegments.Count)
            {
                var candidate = allSegments[i];
                bool isSummary = string.Equals(candidate.SummaryFlag, "Y", StringComparison.OrdinalIgnoreCase);

                if (includeParent && includeChild)
                    return candidate;

                if (includeParent && !includeChild && isSummary)
                    return candidate;

                if (!includeParent && includeChild && !isSummary)
                    return candidate;

                i += step;
            }

            return null;
        }
        private string BuildFormulaString()
        {
            string strFormula = BuildFormulaName();
            List<string> formulaParts = FormulaParameters();

            string finalFormula = "=@" + $"{strFormula}(" + string.Join(",", formulaParts) + ")";

            return finalFormula;
        }
        private string BuildFormulaName()
        {
            if (!string.IsNullOrWhiteSpace(_formulaName))
            {
                switch (_formulaName)
                {
                    case "ENABLEDFLAG":
                        return "GLSense_GetSegmentEnabledFlag";
                    case "SUMMARYFLAG":
                        return "GLSense_GetSegmentSummaryFlag";
                    case DescriptionFormula:
                        return "GLSense_GetSegmentDesc";
                    case "NEXTSEGMENT":
                        return "GLSense_GetNextSegment";
                    case "PREVIOUSSEGMENT":
                        return "GLSense_GetPreviousSegment";
                    case "ACCOUNTTYPE":
                        return "GLSense_GetAccountType";
                    case "DFF":
                        return "GLSense_GetSegmentDFF";
                }
            }

            return string.Empty;
        }
        private List<string> FormulaParameters()
        {
            string segmentVal = GetFieldValue(SegmentField, "Segment");
            string segmentValue = GetFieldValue(SegmentValueField, "SegmentValue");
            string attributeVal = string.Empty;

            if (_formulaName == "DFF")
            {
                attributeVal = GetFieldValue(AttributeField, "Attribute");
            }

            // ACCOUNTTYPE's second argument is now the segment's 1-based position within
            // the Segment dropdown (Segments), not its name. Same reference-vs-value
            // priority as every other field (GetFieldValue's Step 1 checks RefValue before
            // ComboValue): if the user is currently referencing a cell, embed that LIVE
            // reference so the formula keeps recalculating from it (Excel resolves it at
            // calc time, and the UDF parses whatever numeric value the cell currently
            // holds); only when there's no active reference do we resolve and bake in a
            // static index from the combo selection (see GetAccountTypeSegmentArg). The
            // index is emitted as a bare number (not FormatFormulaArg-quoted) so it isn't
            // treated as text, matching how NextParent/NextChild's TRUE/FALSE are emitted
            // unquoted below.
            List<string> formulaParts;
            if (_formulaName == "ACCOUNTTYPE")
            {
                formulaParts =
                [
                    FormatFormulaArg(segmentValue),
                    GetAccountTypeSegmentArg()
                ];
            }
            else
            {
                formulaParts =
                [
                    FormatFormulaArg(segmentValue),
                                FormatFormulaArg(segmentVal)
                ];
            }

            if (!string.IsNullOrWhiteSpace(_formulaName))
            {
                switch (_formulaName)
                {
                    case DescriptionFormula:
                        formulaParts.Add(IsParentValueChecked ? StringTrue : StringFalse);
                        break;
                    case "NEXTSEGMENT":
                    case "PREVIOUSSEGMENT":
                        formulaParts.Add(IsParentChecked ? StringTrue : StringFalse);
                        formulaParts.Add(IsChildChecked ? StringTrue : StringFalse);
                        break;
                    case "DFF":
                        formulaParts.Add(FormatFormulaArg(attributeVal));
                        break;
                }
            }

            string ledgerVal = GetFieldValue(LedgerField, "Ledger");
            formulaParts.Add(FormatFormulaArg(ledgerVal));

            return formulaParts;
        }
        // Returns the 1-based position of the currently selected/resolved segment within
        // Segments (the Segment dropdown's ItemsSource), or -1 if none is selected. Works
        // whether the segment was chosen directly from the combo (WPF binds SelectedItem to
        // ComboValue, so it's the exact same SegmentModel instance from Segments) or resolved
        // from a cell reference (SetSegmentFromLookup sets ComboValue = the matching
        // SegmentModel from Segments too) - either way ComboValue is one of the same object
        // instances Segments.IndexOf can find.
        private int GetSelectedSegmentIndex()
        {
            if (SegmentField.ComboValue is not SegmentModel selected)
                return -1;

            int idx = Segments.IndexOf(selected);
            return idx >= 0 ? idx + 1 : -1;
        }
        // Reference wins over a direct/combo selection, same priority GetFieldValue applies
        // to every other field: if SegmentField currently has an active cell reference,
        // embed it as-is (unquoted, like any other reference) so the formula keeps
        // recalculating live from that cell - reopening the picker later re-derives the
        // right combo selection from whatever numeric value the cell holds via
        // SetSegmentAsReference -> SetSegmentFromLookup. Only when there's no reference do
        // we resolve the combo's current selection into a static 1-based index.
        private string GetAccountTypeSegmentArg()
        {
            if (!string.IsNullOrWhiteSpace(SegmentField.RefValue))
                return SegmentField.RefValue.Trim();

            int segmentIndex = GetSelectedSegmentIndex();
            return segmentIndex > 0 ? segmentIndex.ToString() : FormatFormulaArg(string.Empty);
        }
        public bool WriteFormulaToCell(Microsoft.Office.Interop.Excel.Range rng)
        {
            LogUtility.LogDebug("GLSegmentFuncsViewModel.WriteFormulaToCell: entry");
            try
            {
                if (!ValidateMandatoryFields()) return false;

                string finalFormula = BuildFormulaString();

                rng.Formula = finalFormula;
                LogUtility.LogDebug($"GLSegmentFuncsViewModel.WriteFormulaToCell: wrote formula '{finalFormula}'");
                return true;
            }
            catch (Exception ex)
            {
                ShowWarningAction?.Invoke("Exception encountered while writing formula to excel cell." + Environment.NewLine + ex.Message);
                LogUtility.LogException(ex, "GLSegmentFuncsViewModel.WriteFormulaToCell");
                return false;
            }

        }
        private bool ValidateMandatoryFields()
        {
            if (string.IsNullOrWhiteSpace(_formulaName))
            {
                ShowWarningAction?.Invoke("Formula name is not specified.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(GetFieldValue(LedgerField, "Ledger")))
            {
                ShowWarningAction?.Invoke("Ledger is mandatory.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(GetFieldValue(SegmentField, "Segment")))
            {
                ShowWarningAction?.Invoke("Segment name is mandatory.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(GetFieldValue(SegmentValueField, "SegmentValue")))
            {
                ShowWarningAction?.Invoke("Segment value is mandatory.");
                return false;
            }

            if (_formulaName == "DFF" && string.IsNullOrWhiteSpace(GetFieldValue(AttributeField, "Attribute")))
            {
                ShowWarningAction?.Invoke("Attribute is mandatory for DFF formula.");
                return false;
            }

            return true;
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
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        protected bool SetProperty<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
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
