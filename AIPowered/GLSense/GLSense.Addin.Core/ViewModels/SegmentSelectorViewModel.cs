// SegmentSelectorViewModel.cs in GLSense.Addin.Core
// Port of GLSense\ViewModels\SegmentSelectorViewModel.cs (FinalWorkingCode) - shared
// ViewModel used by both GLSegmentValues ("val" mode) and GLSegmentRef ("Ref" mode) in the
// old project (Group H - Balance Configurator pane + LOVs/Roller/Account dialogs). Both
// windows are now ported and share this class unchanged - GLSegmentValues always
// constructs it with windowName="val" and GLSegmentRef with windowName="Ref". This class
// was originally ported in full (verbatim logic, including the "Ref"-only branches) ahead
// of GLSegmentRef's own port specifically so that later addition wouldn't require touching
// this file again - see Views\GLSegmentRef.xaml.cs for that follow-up and
// Views\GLAccountsRef.xaml.cs for how its EditRequested hook is now wired up.
// Re-pointed the same way as every other already-ported ViewModel in this project (see
// GLDailyRatesViewModel.cs header for the general mapping): GLSense.Helpers ->
// GLSense.Addin.Core.Helpers (ExcelRangeHelper/ApiHelper/ApiResponseHelper/JsonGlobals/
// CancellationHelper); GLSense.Models -> GLSense.Addin.Core.Models; GLSense.Repositories.
// DataRepository -> GLSense.Addin.Core.Repositories.DataRepository; GLSense.Service.
// SearchTypeService -> GLSense.Addin.Core.Models.SearchTypeService (already ported
// alongside SearchTypeModel); GLSense.Utilities.AppState -> GLSense.Addin.Core.AppState;
// GLSense.Utilities.UserConfig -> GLSense.Addin.Core.Utilities.UserConfig; LogUtility.* ->
// ServiceLocator.Logger?.*. Does NOT derive from GLSense.Base.NotifyBase (never ported
// into this project) - implements INotifyPropertyChanged directly instead, exactly like
// the old class already did. No logic changes vs. the original.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.ViewModels
{
    public class SegmentSelectorViewModel(Dispatcher dispatcher, string iWindow, string svals) : INotifyPropertyChanged
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        public event Action<ScrollToTopMessage> ScrollToTopRequested;

        private readonly Dispatcher _dispatcher = dispatcher;
        private readonly string _segValues = svals;
        private readonly string _windowName = iWindow;

        // Actions for window overlay controls
        public Action<string> ShowWarningAction { get; set; }
        public Action<string> ShowInfoAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

        // Set by the View so it can resettle its SizeToContent window once
        // PagedSegmentValues actually has real data - the initial load runs via
        // SelectedSegment's setter firing "_ = LoadSegmentValuesAsync();" fire-and-
        // forget, detached from Window_Loaded's own await chain, so BaseWindow.
        // OnLoaded's SizeToContent resettle always ran against an empty dgLeft/dgRight.
        // Invoked from UpdatePagingAndGrid() below, which is the single choke point
        // every paging/filter/search update already funnels through - harmless to also
        // fire on later interactions (SizeToContent resettle is a cheap no-op once the
        // window is already the right size). See CLAUDE.md section 1.4b.
        public System.Action DataLoadedAction { get; set; }

        private void ScrollDataGridsToTop()
        {
            ScrollToTopRequested?.Invoke(new ScrollToTopMessage
            {
                ScrollLeft = true,
                ScrollRight = true,
                Trigger = "PagingUpdated"
            });
        }

        public static ObservableCollection<SearchTypeModel> SearchTypes => SearchTypeService.GetSearchTypes();

        private SearchTypeModel _selectedSearchType = SearchTypeService.GetDefaultSearchType();
        public SearchTypeModel SelectedSearchType
        {
            get => _selectedSearchType;
            set
            {
                if (SetProperty(ref _selectedSearchType, value))
                {
                    _currentPage = 1;
                    UpdatePagingAndGrid();
                }
            }
        }

        // Data
        private ObservableCollection<SegmentModel> _segments = new ObservableCollection<SegmentModel>();
        public ObservableCollection<SegmentModel> Segments
        {
            get => _segments;
            set => SetProperty(ref _segments, value);
        }

        private SegmentModel _selectedSegment;
        public SegmentModel SelectedSegment
        {
            get => _selectedSegment;
            set
            {
                if (_windowName == "Ref" && _selectedSegment != null)
                {
                    // Save old selections
                    _selectedSegment.SelectedValues =
                        new ObservableCollection<SegmentSelectionModel>(_selectedRight);

                    OnPropertyChanged(nameof(SelectedItemsRight));
                }

                if (SetProperty(ref _selectedSegment, value))
                {
                    // Load async (safe only if method does not touch UI)
                    _ = LoadSegmentValuesAsync();

                    if (_windowName == "Ref")
                    {
                        if (_selectedSegment != null)
                        {
                            _selectedRight =
                                new ObservableCollection<SegmentSelectionModel>(_selectedSegment.SelectedValues);
                        }
                        else
                        {
                            _selectedRight = new ObservableCollection<SegmentSelectionModel>();
                        }

                        OnPropertyChanged(nameof(SelectedItemsRight));
                    }
                }
            }
        }


        // Paging state
        private int _pageSize = UserConfig.RecordsPerPage;
        public int PageSize
        {
            get => _pageSize;
            set
            {
                if (value > 0 && _pageSize != value)
                {
                    _pageSize = value;
                    _currentPage = 1;
                    UpdatePagingAndGrid();
                    OnPropertyChanged(nameof(PageSize));
                }
            }
        }

        private int _currentPage = 1;
        public int CurrentPage
        {
            get => _currentPage;
            set
            {
                if (value >= 1 && value <= TotalPages)
                {
                    _currentPage = value;
                    UpdatePagingAndGrid();
                    OnPropertyChanged(nameof(CurrentPage));
                }
            }
        }

        private int _totalRecords;
        public int TotalRecords
        {
            get => _totalRecords;
            set => SetProperty(ref _totalRecords, value);
        }

        private int _totalPages;
        public int TotalPages
        {
            get => _totalPages;
            set => SetProperty(ref _totalPages, value);
        }

        public string PageRangeText
        {
            get
            {
                if (_totalRecords == 0) return "0 - 0";
                int startIndex = ((_currentPage - 1) * _pageSize) + 1;
                int endIndex = Math.Min(_currentPage * _pageSize, _totalRecords);
                return $"{startIndex} - {endIndex}";
            }
        }

        // Segment Values
        private List<SegmentValueModel> _allSegmentValues = new List<SegmentValueModel>();
        private List<SegmentValueModel> _pagedSegmentValues = new List<SegmentValueModel>();
        public ReadOnlyCollection<SegmentValueModel> PagedSegmentValues => _pagedSegmentValues.AsReadOnly();

        // Selection models
        private ObservableCollection<SegmentSelectionModel> _selectedRight = new ObservableCollection<SegmentSelectionModel>();
        public ObservableCollection<SegmentSelectionModel> SelectedItemsRight
        {
            get => _selectedRight;
            set => SetProperty(ref _selectedRight, value);
        }

        // Search/filter
        private string _searchText;
        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                {
                    _currentPage = 1;
                    UpdatePagingAndGrid();
                }
            }
        }

        // Hierarchy
        private ObservableCollection<SegmentValueModel> _hierarchyItems = new ObservableCollection<SegmentValueModel>();
        public ObservableCollection<SegmentValueModel> HierarchyItems
        {
            get => _hierarchyItems;
            set => SetProperty(ref _hierarchyItems, value);
        }

        private SegmentValueModel _selectedHierarchy;
        public SegmentValueModel SelectedHierarchy
        {
            get => _selectedHierarchy;
            set
            {
                if (SetProperty(ref _selectedHierarchy, value))
                {
                    Task.Run(async () => await LoadHierarchySegmentValuesAsync());
                }
            }
        }

        // Multi-row checkbox state
        private bool _isMultipleRowsEnabled;
        public bool IsMultipleRowsEnabled
        {
            get => _isMultipleRowsEnabled;
            set => SetProperty(ref _isMultipleRowsEnabled, value);
        }

        private bool _isMultipleRowsChecked;
        public bool IsMultipleRowsChecked
        {
            get => _isMultipleRowsChecked;
            set => SetProperty(ref _isMultipleRowsChecked, value);
        }

        // Left grid enabled state (for refedit window)
        private bool _isLeftGridEnabled = true;
        public bool IsLeftGridEnabled
        {
            get => _isLeftGridEnabled;
            set => SetProperty(ref _isLeftGridEnabled, value);
        }

        // Excel interop for refedit controls
        private Excel.Application _excelApp;
        public Excel.Application ExcelApp
        {
            get => _excelApp;
            set => SetProperty(ref _excelApp, value);
        }
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
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

        // ----------- Segment, hierarchy loaders ----------------
        public async Task LoadSegmentsAsync(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadSegmentsAsync: started. windowName={_windowName}, cubeId={cubeId}, ledgerId={ledgerId}");
            await Task.Run(() =>
            {
                var repository = new DataRepository();
                var segs = repository.GetSegments(cubeId, ledgerId);
                ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadSegmentsAsync: DataRepository.GetSegments returned {segs?.Count ?? 0} segment(s).");
                _dispatcher.Invoke(() => ProcessSegments(segs));
            });
        }

        private void ProcessSegments(IEnumerable<SegmentModel> segs)
        {
            foreach (var s in Segments)
                s.PropertyChanged -= OnSegmentValueChanged;

            Segments.Clear();

            int index = 0;
            foreach (var s in segs)
            {
                s.SegmentName = s.SegmentName.Trim();
                InitializeSegment(s, index++);
            }

            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.ProcessSegments: initialized {Segments.Count} segment(s).");
            SelectInitialSegment();
        }

        private void InitializeSegment(SegmentModel s, int index)
        {
            if (_windowName == "Ref" && !string.IsNullOrWhiteSpace(_segValues))
            {
                var resolvedValues = ResolveSegmentValueText(_segValues);
                var parts = resolvedValues
                    .Split(new[] { ';' }, StringSplitOptions.None)
                    .Select(x => x.Trim().Trim('"'))
                    .ToList();

                ProcessRefSegment(s, parts, index);
            }
            else
            {
                ApplyDefaultSegment(s);
            }

            ApplyEnableState(s);
            s.PropertyChanged += OnSegmentValueChanged;
            Segments.Add(s);
        }

        private void ProcessRefSegment(SegmentModel s, List<string> parts, int index)
        {
            if (parts.Count <= index)
            {
                ClearSegmentValues(s);
                return;
            }

            var part = parts[index];
            s.IsVisible = true;
            s.SelectedValues = new ObservableCollection<SegmentSelectionModel>();

            if (string.IsNullOrWhiteSpace(part))
            {
                ClearSegmentValues(s);
                return;
            }

            ParseAndSetSegmentValues(s, part);
        }

        private void ParseAndSetSegmentValues(SegmentModel s, string part)
        {
            var isReference = IsExcelReference(part);
            s.Reference = isReference ? part : string.Empty;
            s.Value = isReference ? string.Empty : part;

            var entries = part.Split(',')
                             .Select(x => x.Trim())
                             .Where(x => !string.IsNullOrEmpty(x));

            foreach (var entry in entries)
                AddParsedSelection(s, entry);
        }

        private static bool IsExcelReference(string part) =>
            part.Contains("$") || part.Contains("!") || part.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase);

        private string ResolveSegmentValueText(string segmentValues)
        {
            if (string.IsNullOrWhiteSpace(segmentValues))
                return string.Empty;

            var cleanedValue = segmentValues.Trim();
            if (!ExcelRangeHelper.IsRealRange(cleanedValue) || ExcelApp == null)
                return cleanedValue;

            try
            {
                var resolved = ExcelApp.Range[cleanedValue]?.Value2;
                return resolved?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentSelectorViewModel.ResolveSegmentValueText");
                return cleanedValue;
            }
        }

        private void AddParsedSelection(SegmentModel s, string entry)
        {
            if (entry.Contains('|'))
            {
                var pair = entry.Split('|');
                s.SelectedValues.Add(new SegmentSelectionModel
                {
                    Value1 = pair[0].Trim(),
                    Value2 = pair.Length > 1 ? pair[1].Trim() : "",
                    Segment = s.SegmentName
                });
            }
            else
            {
                s.SelectedValues.Add(new SegmentSelectionModel
                {
                    Value1 = entry,
                    Value2 = null,
                    Segment = s.SegmentName
                });
            }
        }
        private void SelectInitialSegment()
        {
            if (Segments.Count > 0)
            {
                if (AppState.Instance.SegmentPickedIndex >= 0 &&
                    AppState.Instance.SegmentPickedIndex < Segments.Count)
                {
                    SelectedSegment = Segments[AppState.Instance.SegmentPickedIndex];
                }
                else
                {
                    SelectedSegment = Segments[0];
                }
            }
        }

        private void ApplyDefaultSegment(SegmentModel s)
        {
            var defaultValue = s.DefaultValue?.Trim();
            s.Value = defaultValue;
            s.Reference = string.Empty;
            s.IsVisible = true;
            s.SelectedValues = new ObservableCollection<SegmentSelectionModel>();

            if (!string.IsNullOrWhiteSpace(defaultValue))
            {
                s.SelectedValues.Add(new SegmentSelectionModel
                {
                    Value1 = s.Value,
                    Value2 = null,
                    Segment = s.SegmentName
                });
            }
        }

        private void ClearSegmentValues(SegmentModel s)
        {
            s.Value = string.Empty;
            s.Reference = string.Empty;
            ApplyEnableState(s);
        }


        public async Task LoadSegmentValuesAsync(SegmentModel segModel = null, SegmentValueModel segValModel = null, bool fromHierarchy = false)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadSegmentValuesAsync: started. segment={(segModel ?? SelectedSegment)?.SegmentName}, fromHierarchy={fromHierarchy}, hierarchyValue={segValModel?.SegmentValue}");
            await Task.Run(() =>
            {
                _allSegmentValues.Clear();
                ObservableCollection<SegmentValueModel> vals = null;
                var repository = new DataRepository();
                if (fromHierarchy)
                {
                    vals = repository.GetSegmentValuesHierarchy(segValModel);
                }
                else
                {
                    vals = DataRepository.GetSegmentValues(segModel ?? SelectedSegment);
                }

                // Convert ObservableCollection to List explicitly (remove shorthand)
                _allSegmentValues = vals?.ToList() ?? new List<SegmentValueModel>();
                ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadSegmentValuesAsync: loaded {_allSegmentValues.Count} segment value(s) (fromHierarchy={fromHierarchy}).");
            });

            await _dispatcher.InvokeAsync(() =>
            {
                UpdateHierarchyCombo(fromHierarchy);
                _currentPage = 1;
                UpdatePagingAndGrid();
            });
        }

        private void UpdateHierarchyCombo(bool fromHierarchy)
        {
            if (!fromHierarchy)
            {
                var summaryAccounts = _allSegmentValues
                    .Where(x => string.Equals(x.SummaryFlag, "Y", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(x => x.SegmentValue)
                    .ToList();
                HierarchyItems = new ObservableCollection<SegmentValueModel>(summaryAccounts);
            }
        }
        public async Task LoadHierarchySegmentValuesAsync()
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: started. SelectedHierarchy={SelectedHierarchy?.SegmentValue}");
            if (SelectedHierarchy == null)
            {
                ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: SelectedHierarchy is null, resetting to page 1 and returning.");
                _currentPage = 1;
                await _dispatcher.InvokeAsync(() => UpdatePagingAndGrid());
                return;
            }

            CancellationHelper ctsHelper = new();
            CancellationToken token = ctsHelper.GetToken();

            try
            {
                if (ShowBusyAction != null)
                {
                    await ShowBusyAction.Invoke("Fetching hierarchy data... (click Cancel to stop)",
                        async () => {
                            if (!ctsHelper.IsCancellationRequested)
                            {
                                ctsHelper.Cancel();
                            }
                            await Task.CompletedTask;
                        });
                }

                bool hierarchyBool = DataRepository.SegmentValuesHierarchyExists(SelectedHierarchy);
                ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: DataRepository.SegmentValuesHierarchyExists({SelectedHierarchy.SegmentValue})={hierarchyBool}");

                if (!hierarchyBool)
                {
                    ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: hierarchy cache miss, calling HierarhyApiAsync for {SelectedHierarchy.SegmentValue}.");
                    var hierarchyData = await HierarhyApiAsync(SelectedHierarchy, token);
                    if (!string.IsNullOrWhiteSpace(hierarchyData))
                    {
                        DataRepository.SaveHierarchyToCache(SelectedHierarchy, hierarchyData);
                        ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: cached hierarchy data for {SelectedHierarchy.SegmentValue}.");
                    }
                    else
                    {
                        ServiceLocator.Logger?.LogWarn($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: HierarhyApiAsync returned empty data for {SelectedHierarchy.SegmentValue}.");
                    }
                }
                await LoadSegmentValuesAsync(null, SelectedHierarchy, true);
                ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: completed.");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: user cancelled operation. Loading hierarchy segment values interrupted.");
            }
            catch (Exception ex)
            {
                LogError(ex, "SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync");
            }
            finally
            {
                try
                {
                    if (!ctsHelper.IsCancellationRequested)
                        ctsHelper.Cancel();

                    ctsHelper.Dispose();
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogWarn($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: failed disposing CancellationHelper (non-fatal): {ex.Message}");
                }
                if (HideBusyAsyncAction != null)
                {
                    await HideBusyAsyncAction.Invoke();
                }
            }
        }

        private void ApplyEnableState(SegmentModel s)
        {
            if (_windowName.IndexOf("Ref") >= 0)
            {
                if (!string.IsNullOrWhiteSpace(s.Reference))
                {
                    // Has reference -> disable textbox, enable refedit
                    s.IsTextEnabled = false;
                    s.IsRefEditEnabled = true;
                }
                else if (s.SelectedValues != null && s.SelectedValues.Any())
                {
                    // Has values -> enable textbox, disable refedit
                    s.IsTextEnabled = true;
                    s.IsRefEditEnabled = false;
                }
                else
                {
                    // Both empty -> enable both
                    s.IsTextEnabled = true;
                    s.IsRefEditEnabled = true;
                }
            }
        }

        // ------- Paging and Search Logic -----------
        private void UpdatePagingAndGrid()
        {
            var filtered = ApplySearchFilter(_allSegmentValues);
            _totalRecords = filtered.Count;
            _totalPages = _pageSize <= 0 ? 1 : (int)Math.Ceiling(_totalRecords / (double)_pageSize);
            if (_totalPages == 0) _totalPages = 1;
            if (_currentPage < 1) _currentPage = 1;
            if (_currentPage > _totalPages) _currentPage = _totalPages;
            int skipCount = (_currentPage - 1) * _pageSize;

            // Remove shorthand; take explicit ToList()
            _pagedSegmentValues = filtered.Skip(skipCount).Take(_pageSize).ToList();

            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.UpdatePagingAndGrid: totalRecords={_totalRecords}, totalPages={_totalPages}, currentPage={_currentPage}, pagedCount={_pagedSegmentValues.Count}");

            OnPropertyChanged(nameof(PagedSegmentValues));
            OnPropertyChanged(nameof(CurrentPage));
            OnPropertyChanged(nameof(TotalPages));
            OnPropertyChanged(nameof(TotalRecords));
            OnPropertyChanged(nameof(PageRangeText));

            // Scroll to top after data is loaded
            ScrollDataGridsToTop();

            DataLoadedAction?.Invoke();
        }

        private List<SegmentValueModel> ApplySearchFilter(List<SegmentValueModel> source)
        {
            var searchText = (_searchText ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(searchText))
                return source;

            var searchType = _selectedSearchType;
            return searchType.Value switch
            {
                "StartsWith" => FilterMatches(source, searchText, StringMatch.StartsWith),
                "DoesNotStartWith" => FilterMatches(source, searchText, StringMatch.DoesNotStartWith),
                "EndsWith" => FilterMatches(source, searchText, StringMatch.EndsWith),
                "DoesNotEndWith" => FilterMatches(source, searchText, StringMatch.DoesNotEndWith),
                "Contains" => FilterMatches(source, searchText, StringMatch.Contains),
                "NotContains" => FilterMatches(source, searchText, StringMatch.NotContains),
                "Equals" => FilterMatches(source, searchText, StringMatch.Equals),
                "NotEquals" => FilterMatches(source, searchText, StringMatch.NotEquals),
                _ => source
            };
        }

        private static List<SegmentValueModel> FilterMatches(List<SegmentValueModel> source, string searchText, StringMatch matchType)
        {
            // Return explicit List instead of shorthand
            return source.Where(x => MatchesCriteria(x, searchText, matchType)).ToList();
        }

        private static bool MatchesCriteria(SegmentValueModel item, string searchText, StringMatch matchType)
        {
            var valueMatch = HasMatch(item.SegmentValue, searchText, matchType);
            var descMatch = HasMatch(item.Description, searchText, matchType);

            return IsPositiveMatch(matchType)
                ? valueMatch || descMatch
                : valueMatch && descMatch;
        }

        private static bool HasMatch(string text, string searchText, StringMatch matchType)
        {
            return text switch
            {
                null => IsNullMatch(matchType),
                _ => matchType switch
                {
                    StringMatch.StartsWith => text.StartsWith(searchText, StringComparison.CurrentCultureIgnoreCase),
                    StringMatch.DoesNotStartWith => !text.StartsWith(searchText, StringComparison.CurrentCultureIgnoreCase),
                    StringMatch.EndsWith => text.EndsWith(searchText, StringComparison.CurrentCultureIgnoreCase),
                    StringMatch.DoesNotEndWith => !text.EndsWith(searchText, StringComparison.CurrentCultureIgnoreCase),
                    StringMatch.Contains => text.IndexOf(searchText, StringComparison.CurrentCultureIgnoreCase) >= 0,
                    StringMatch.NotContains => text.IndexOf(searchText, StringComparison.CurrentCultureIgnoreCase) == -1,
                    StringMatch.Equals => string.Equals(text, searchText, StringComparison.CurrentCultureIgnoreCase),
                    StringMatch.NotEquals => !string.Equals(text, searchText, StringComparison.CurrentCultureIgnoreCase),
                    _ => false
                }
            };
        }

        // Regular static methods (no 'this' parameter)
        private static bool IsPositiveMatch(StringMatch matchType) =>
            matchType is StringMatch.StartsWith or StringMatch.EndsWith or
            StringMatch.Contains or StringMatch.Equals;

        private static bool IsNullMatch(StringMatch matchType) =>
            matchType is StringMatch.DoesNotStartWith or StringMatch.DoesNotEndWith or
            StringMatch.NotContains or StringMatch.NotEquals;

        private enum StringMatch
        {
            StartsWith, DoesNotStartWith, EndsWith, DoesNotEndWith,
            Contains, NotContains, Equals, NotEquals
        }



        // Paging commands (to be wired from UI)
        public void GoFirstPage()
        {
            ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.GoFirstPage invoked.");
            _currentPage = 1;
            UpdatePagingAndGrid();
        }

        public void GoPreviousPage()
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.GoPreviousPage invoked. currentPage={_currentPage}");
            if (_currentPage > 1) _currentPage--;
            UpdatePagingAndGrid();
        }

        public void GoNextPage()
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.GoNextPage invoked. currentPage={_currentPage}, totalPages={_totalPages}");
            if (_currentPage < _totalPages) _currentPage++;
            UpdatePagingAndGrid();
        }

        public void GoLastPage()
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.GoLastPage invoked. totalPages={_totalPages}");
            _currentPage = _totalPages;
            UpdatePagingAndGrid();
        }

        // Called after changing page size
        public void ApplyPageSize(int size)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.ApplyPageSize invoked. size={size}");
            if (size > 0)
            {
                _pageSize = size;
                _currentPage = 1;
                UpdatePagingAndGrid();
            }
        }

        // --------- Selection Logic ------------
        private static string GetEffectiveValue(SegmentValueModel item)
        {
            var val = item.SegmentValue;
            if (item.IsModified)
            {
                val += "~";
            }
            return val;
        }

        public void AddSelection(IList<SegmentValueModel> selectedItems)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddSelection: started. selectedItems count={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count == 0)
            {
                ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.AddSelection: no items selected to add.");
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select at least one item to add."));
                }
                return;
            }

            int addedCount = 0;
            foreach (var seg in selectedItems)
            {
                var val = GetEffectiveValue(seg);
                var exists = _selectedRight.Any(r => r.Value1 == val && r.Segment == seg.SegmentName);
                if (exists)
                {
                    if (selectedItems.Count == 1) ShowWarningAction?.Invoke($"'{val}' already exists in the right grid.");
                    continue;
                }
                _selectedRight.Add(new SegmentSelectionModel { Value1 = val, Value2 = null, Segment = seg.SegmentName });
                OnPropertyChanged(nameof(SelectedItemsRight));
                addedCount++;
            }

            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddSelection: added {addedCount} of {selectedItems.Count} item(s).");
            UpdateMultiRowState();
        }

        public void RemoveSelection(IList<SegmentSelectionModel> selectedItems)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.RemoveSelection: started. selectedItems count={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count == 0)
            {
                ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.RemoveSelection: no items selected to remove.");
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select one or more items to remove."));
                }
                return;
            }

            foreach (var sel in selectedItems.ToList())
            {
                _selectedRight.Remove(sel);
                OnPropertyChanged(nameof(SelectedItemsRight));
            }

            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.RemoveSelection: removed {selectedItems.Count} item(s).");
            UpdateMultiRowState();
        }

        public void AddBetweenSelection(IList<SegmentValueModel> selectedItems, bool isExclude)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddBetweenSelection: started. selectedItems count={selectedItems?.Count ?? 0}, isExclude={isExclude}");
            if (selectedItems == null || selectedItems.Count < 2)
            {
                ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.AddBetweenSelection: fewer than 2 items selected, aborting.");
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select two or more items (the first -> Value1, the last -> Value2)."));
                }
                return;
            }

            var seg1 = selectedItems[0];
            var seg2 = selectedItems[selectedItems.Count - 1];
            var val1 = GetEffectiveValue(seg1);
            var val2 = GetEffectiveValue(seg2);
            if (isExclude)
            {
                val1 = "--" + val1;
                val2 = "--" + val2;
            }
            var exists = _selectedRight.Any(r => r.Value1 == val1 && r.Value2 == val2 && r.Segment == seg1.SegmentName);
            if (exists)
            {
                ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddBetweenSelection: range '{val1} - {val2}' already exists, skipping.");
                ShowWarningAction?.Invoke($"Range '{val1} - {val2}' already exists.");
                return;
            }
            _selectedRight.Add(new SegmentSelectionModel { Value1 = val1, Value2 = val2, Segment = seg1.SegmentName });
            OnPropertyChanged(nameof(SelectedItemsRight));
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddBetweenSelection: added range '{val1} - {val2}'.");

            UpdateMultiRowState();
        }

        public void AddNotBetweenSelection(IList<SegmentValueModel> selectedItems)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddNotBetweenSelection: started. selectedItems count={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count < 2)
            {
                ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.AddNotBetweenSelection: fewer than 2 items selected, aborting.");
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select two or more items (the first -> Value1, the last -> Value2)."));
                }
                return;
            }

            var seg1 = selectedItems[0];
            var seg2 = selectedItems[selectedItems.Count - 1];
            var val1 = "--" + GetEffectiveValue(seg1);
            var val2 = "--" + GetEffectiveValue(seg2);
            var exists = _selectedRight.Any(r => r.Value1 == val1 && r.Value2 == val2 && r.Segment == seg1.SegmentName);
            if (exists)
            {
                ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddNotBetweenSelection: range '{val1} - {val2}' already exists, skipping.");
                ShowWarningAction?.Invoke($"Range '{val1} - {val2}' already exists.");
                return;
            }
            _selectedRight.Add(new SegmentSelectionModel { Value1 = val1, Value2 = val2, Segment = seg1.SegmentName });
            OnPropertyChanged(nameof(SelectedItemsRight));
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddNotBetweenSelection: added range '{val1} - {val2}'.");

            UpdateMultiRowState();
        }

        public void AddExcludeSelection(IList<SegmentValueModel> selectedItems)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddExcludeSelection: started. selectedItems count={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count == 0)
            {
                ServiceLocator.Logger?.LogDebug("SegmentSelectorViewModel.AddExcludeSelection: no items selected to exclude.");
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select one or more items to exclude."));
                }
                return;
            }

            int addedCount = 0;
            foreach (var seg in selectedItems)
            {
                var baseVal = GetEffectiveValue(seg);
                var excludeVal = "--" + baseVal;
                var exists = _selectedRight.Any(r => r.Value1 == excludeVal && r.Segment == seg.SegmentName);
                if (exists)
                {
                    if (selectedItems.Count == 1) ShowWarningAction?.Invoke($"Excluded '{excludeVal}' already exists.");
                    continue;
                }
                _selectedRight.Add(new SegmentSelectionModel { Value1 = excludeVal, Value2 = null, Segment = seg.SegmentName });
                OnPropertyChanged(nameof(SelectedItemsRight));
                addedCount++;
            }

            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.AddExcludeSelection: added {addedCount} of {selectedItems.Count} item(s).");
            UpdateMultiRowState();
        }

        public void ClearDefaults()
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.ClearDefaults: started. Segments count={Segments?.Count ?? 0}");
            if (Segments == null) return;

            foreach (var s in Segments)
                ClearSegmentDefault(s);

            UpdateMultiRowState();
        }

        private void ClearSegmentDefault(SegmentModel s)
        {
            var defaultValue = s.DefaultValue?.Trim();
            if (string.IsNullOrWhiteSpace(defaultValue)) return;

            RemoveDefaultSelections(s, defaultValue);
            RemoveDefaultFromValue(s, defaultValue);
            ApplyEnableState(s);
        }

        private static void RemoveDefaultSelections(SegmentModel s, string defaultValue)
        {
            var toRemove = s.SelectedValues
                .Where(sl => sl.Value1 != null &&
                           sl.Value1.Equals(defaultValue, StringComparison.OrdinalIgnoreCase) &&
                           string.IsNullOrEmpty(sl.Value2))
                .ToList();

            foreach (var item in toRemove)
                s.SelectedValues.Remove(item);
        }

        private static void RemoveDefaultFromValue(SegmentModel s, string defaultValue)
        {
            if (string.IsNullOrEmpty(s.Value)) return;

            var parts = s.Value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => !p.Equals(defaultValue, StringComparison.OrdinalIgnoreCase))
                .ToList();

            s.Value = parts.Count > 0 ? string.Join(",", parts) : string.Empty;
        }


        public string GetAllSegmentValues()
        {
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.GetAllSegmentValues: started. Segments count={Segments?.Count ?? 0}");
            if (Segments == null || Segments.Count == 0) return string.Empty;

            var result = new List<string>();

            foreach (var s in Segments)
            {
                string segVal = string.Empty;

                // Priority 1: reference (if available)
                if (!string.IsNullOrWhiteSpace(s.Reference))
                {
                    segVal = s.Reference.Trim();
                }
                // Priority 2: value (if reference is empty)
                else if (!string.IsNullOrWhiteSpace(s.Value))
                {
                    segVal = $"\"{s.Value.Trim()}\"";
                }

                result.Add(segVal);
            }

            var joined = string.Join(";", result);
            ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.GetAllSegmentValues: built value string for {result.Count} segment(s), length={joined.Length}.");
            return joined;
        }

        // ---------- API (hierarchy) async call helper ----------
        public async Task<string> HierarhyApiAsync(SegmentValueModel selectedHierarchy, CancellationToken token)
        {
            string apiUrl = string.Empty;
            try
            {

                apiUrl =
                    $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}segment-hierarchy" +
                    $"?segmentValueSetId={selectedHierarchy.SegmentValueSetId}" +
                    $"&parentSegmentValue={WebUtility.UrlEncode(selectedHierarchy.SegmentValue.Trim())}" +
                    $"&cubeId={selectedHierarchy.CubeId}";

                token.ThrowIfCancellationRequested();

                ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.HierarhyApiAsync: calling ApiHelper.ServerAPI (POST) {apiUrl}");
                string response =
                    await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", token);

                token.ThrowIfCancellationRequested();

                ValidateTransportResponse(response);

                var result = ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    ServiceLocator.Logger?.LogWarn($"SegmentSelectorViewModel.HierarhyApiAsync: hierarchy API failed: {apiUrl}");
                    ServiceLocator.Logger?.LogRawJson("SegmentSelectorViewModel.HierarhyApiAsync", response);

                    if (HideBusyAsyncAction != null)
                    {
                        await HideBusyAsyncAction.Invoke();
                    }
                    if (ShowWarningAction != null)
                    {
                        await _dispatcher.InvokeAsync(() => ShowWarningAction.Invoke(result.ErrorMessage ??
                                          "Hierarchy API returned failure status."));
                    }

                    return string.Empty;
                }

                ServiceLocator.Logger?.LogDebug($"SegmentSelectorViewModel.HierarhyApiAsync: succeeded for {apiUrl}");
                return response;
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn(
                    "SegmentSelectorViewModel.HierarhyApiAsync: user cancelled operation. Fetching hierarchy data interrupted.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"SegmentSelectorViewModel.HierarhyApiAsync (apiUrl={apiUrl})");
            }
            return string.Empty;
        }
        private static void ValidateTransportResponse(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
            {
                ServiceLocator.Logger?.LogWarn("SegmentSelectorViewModel.ValidateTransportResponse: empty API response.");
                throw new InvalidOperationException("Empty API response.");
            }

            if (response.IndexOf("(401)", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ServiceLocator.Logger?.LogWarn("SegmentSelectorViewModel.ValidateTransportResponse: session expired (401) response.");
                throw new UnauthorizedAccessException("Session expired.");
            }

            if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ServiceLocator.Logger?.LogWarn("SegmentSelectorViewModel.ValidateTransportResponse: server returned error/HTML response.");
                ServiceLocator.Logger?.LogRawJson("SegmentSelectorViewModel.ValidateTransportResponse", response);
                throw new InvalidOperationException(response);
            }
        }

        // ---------------- Multi-row and Segment Ref logic ----------------
        private void UpdateMultiRowState()
        {
            if (_windowName.Contains("Ref"))
                UpdateRefWindowState();
            else
                UpdateNonRefWindowState();
        }

        private void UpdateRefWindowState()
        {
            if (SelectedSegment == null) return;

            // Reference takes priority over any right-grid selection. Without this check
            // first, ValidateAndApplyReferenceValue's reference-driven display (Value set
            // to the resolved cell value, then still reference-owned per its own
            // ApplyEnableState call) got clobbered right back to "manual value" mode the
            // instant this method next ran, because setting seg.Value there also
            // repopulates _selectedRight (via HandleValueChange -> SyncActiveSegmentGrid),
            // which used to make the "_selectedRight.Any()" branch below win and flip
            // IsTextEnabled back to true/IsRefEditEnabled back to false - re-enabling the
            // Value box even though a Reference is still active. See CLAUDE.md's
            // GLSegmentManager reference-mode section for the full symptom/fix writeup.
            if (!string.IsNullOrWhiteSpace(SelectedSegment.Reference))
            {
                SelectedSegment.IsTextEnabled = false;
                SelectedSegment.IsRefEditEnabled = true;
            }
            else if (_selectedRight.Any())
            {
                SelectedSegment.Value = BuildValueFromSelections();
                SelectedSegment.IsTextEnabled = true;
                SelectedSegment.IsRefEditEnabled = false;
                // Testing feedback: this branch is only ever reached from an actual
                // runtime mutation of the right-hand grid (Add/Between/NotBetween/
                // Exclude/Remove button handlers, or a direct Value textbox edit) - never
                // from the initial default-value parse or a plain segment switch (see
                // SegmentModel.IsUserSelected's own comment). So once we're here, this
                // segment's Value no longer represents its untouched factory default;
                // flip the flag so the Segments list subtitle says "Selected: X" instead
                // of "Default: X" from now on.
                SelectedSegment.IsUserSelected = true;
            }
            else
            {
                UpdateRefControlsFromReference();
            }

            OnPropertyChanged(nameof(SelectedSegment));
        }

        private void UpdateNonRefWindowState()
        {
            if (SelectedItemsRight == null || !SelectedItemsRight.Any())
            {
                IsMultipleRowsEnabled = false;
                IsMultipleRowsChecked = false;
                return;
            }

            var distinctSegments = SelectedItemsRight.Select(r => r.Segment).Distinct().Count();
            IsMultipleRowsEnabled = distinctSegments == 1;
            IsMultipleRowsChecked = IsMultipleRowsEnabled && SelectedItemsRight.Count > 0;
        }

        private string BuildValueFromSelections() =>
            string.Join(",", _selectedRight.Select(r =>
                !string.IsNullOrEmpty(r.Value2) ? $"{r.Value1}|{r.Value2}" : r.Value1));

        private void UpdateRefControlsFromReference()
        {
            if (!string.IsNullOrWhiteSpace(SelectedSegment.Reference))
            {
                SelectedSegment.IsTextEnabled = false;
                SelectedSegment.IsRefEditEnabled = true;
            }
            else
            {
                SelectedSegment.IsTextEnabled = true;
                SelectedSegment.IsRefEditEnabled = true;
                SelectedSegment.Value = string.Empty;
            }
        }


        private void OnSegmentValueChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is not SegmentModel seg) return;

            switch (e.PropertyName)
            {
                case nameof(SegmentModel.Value):
                    HandleValueChange(seg);
                    break;
                case nameof(SegmentModel.Reference):
                    HandleReferenceChange(seg);
                    break;
            }
        }

        private void HandleValueChange(SegmentModel seg)
        {
            if (string.IsNullOrWhiteSpace(seg.Value))
            {
                ClearActiveSegmentGrid(seg);
            }
            else
            {
                var selections = ParseValueToSelections(seg.Value, seg.SegmentName);
                seg.SelectedValues = selections;
                SyncActiveSegmentGrid(seg, selections);
            }
        }
        private void HandleReferenceChange(SegmentModel seg)
        {
            // When reference changes, clear value and update selections based on reference
            if (!string.IsNullOrWhiteSpace(seg.Reference))
            {
                seg.Value = string.Empty;
                seg.SelectedValues.Clear();
                ApplyEnableState(seg);

                // Issue 5 (Segment Manager polish): resolve the referenced cell and, if it
                // holds a valid value (or list of values) for this segment, mirror it into
                // the (disabled) Value box; otherwise warn the user instead of silently
                // leaving the Value box blank.
                ValidateAndApplyReferenceValue(seg);
            }
            else
            {
                // Reference cleared - always clear Value/SelectedValues too, rather than
                // "restoring" SelectedValues from whatever seg.Value currently holds. That
                // restore made sense before ValidateAndApplyReferenceValue existed, back
                // when Value could only get populated by something the user typed
                // directly. Now Value may just be a mirror of whatever the (just-cleared)
                // Reference resolved to (see ValidateAndApplyReferenceValue below), so
                // keeping it around would silently repopulate the Value box and right
                // grid with data tied to a reference that no longer exists. Testing
                // feedback: clearing the RefEdit box should clear the Value box/grid too.
                seg.Value = string.Empty;
                seg.SelectedValues.Clear();
                ApplyEnableState(seg);
            }

            // Sync right grid if this is the active segment
            if (_selectedSegment == seg)
                UpdateMultiRowState();
        }

        // Issue 5 (Segment Manager polish): reads the cell(s) behind seg.Reference and
        // checks it against the currently loaded segment values. Three outcomes:
        //   - cell is empty -> warn "no data in cell"
        //   - cell has a value/list that doesn't match a valid segment value -> warn which
        //     token(s) weren't found
        //   - cell has a valid value/list -> mirror it into seg.Value (shown, disabled, in
        //     the Value box - see GLSegmentManager.xaml's Value TextBox IsEnabled binding)
        // Only meaningful for the currently selected segment, since _allSegmentValues below
        // only ever holds the currently-loaded segment's values (see LoadSegmentValuesAsync)
        // - matches the same "only act on the active segment" guard already used at the
        // bottom of HandleReferenceChange/HandleValueChange.
        private void ValidateAndApplyReferenceValue(SegmentModel seg)
        {
            if (seg != _selectedSegment) return;
            if (ExcelApp == null || string.IsNullOrWhiteSpace(seg.Reference)) return;

            string cellValue;
            try
            {
                if (!ExcelRangeHelper.IsRealRange(seg.Reference)) return;
                var resolved = ExcelApp.Range[seg.Reference]?.Value2;
                cellValue = resolved?.ToString();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentSelectorViewModel.ValidateAndApplyReferenceValue: failed reading referenced cell");
                return;
            }

            if (string.IsNullOrWhiteSpace(cellValue))
            {
                ShowWarningAction?.Invoke($"No data in the referenced cell for '{seg.SegmentName}'.");
                return;
            }

            var candidates = ParseReferenceCellValues(cellValue);
            if (candidates.Count == 0)
            {
                ShowWarningAction?.Invoke($"No data in the referenced cell for '{seg.SegmentName}'.");
                return;
            }

            var validValues = new HashSet<string>(
                _allSegmentValues.Select(v => v.SegmentValue),
                StringComparer.OrdinalIgnoreCase);

            var invalid = candidates.Where(c => !validValues.Contains(c)).ToList();
            if (invalid.Count > 0)
            {
                ShowWarningAction?.Invoke($"'{string.Join(", ", invalid)}' not found in segment '{seg.SegmentName}'.");
                return;
            }

            // All candidates are valid segment values - mirror them into the (disabled)
            // Value box. This flows through HandleValueChange -> SyncActiveSegmentGrid like
            // a manual selection would, populating the right-hand grid too; UpdateRefWindowState's
            // Reference-first check (above) keeps the Value box itself disabled afterward.
            seg.Value = string.Join(",", candidates);
        }

        // Splits a resolved cell value into individual segment-value tokens. Handles a
        // plain single value, comma-separated multiple values, and a leading/trailing
        // double-hyphen or single-hyphen wrapper (e.g. "--01,02--" or "-01-"), which is how
        // some upstream sheets prefix/suffix a segment value list.
        private static List<string> ParseReferenceCellValues(string cellValue)
        {
            var trimmed = cellValue.Trim();

            foreach (var marker in new[] { "--", "-" })
            {
                if (trimmed.Length > marker.Length * 2 &&
                    trimmed.StartsWith(marker, StringComparison.Ordinal) &&
                    trimmed.EndsWith(marker, StringComparison.Ordinal))
                {
                    trimmed = trimmed.Substring(marker.Length, trimmed.Length - marker.Length * 2);
                    break;
                }
            }

            return trimmed
                .Split(',')
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrEmpty(x))
                .ToList();
        }

        private void ClearActiveSegmentGrid(SegmentModel seg)
        {
            if (_selectedSegment != null && _selectedSegment == seg)
            {
                _selectedRight.Clear();
                UpdateMultiRowState();
            }
        }

        private void SyncActiveSegmentGrid(SegmentModel seg, ObservableCollection<SegmentSelectionModel> selections)
        {
            if (_selectedSegment != null && _selectedSegment == seg)
            {
                _selectedRight = new ObservableCollection<SegmentSelectionModel>(selections);
                OnPropertyChanged(nameof(SelectedItemsRight));
                UpdateMultiRowState();
            }
        }

        private ObservableCollection<SegmentSelectionModel> ParseValueToSelections(string value, string segmentName)
        {
            var parts = value.Split(new[] { ',' }, StringSplitOptions.None)
                            .Select(x => x.Trim()).ToList();

            var selections = new ObservableCollection<SegmentSelectionModel>();
            foreach (var part in parts)
            {
                if (part.Contains('|'))
                {
                    var rangeParts = part.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    selections.Add(new SegmentSelectionModel
                    {
                        Value1 = rangeParts[0].Trim(),
                        Value2 = rangeParts.Length > 1 ? rangeParts[1].Trim() : "",
                        Segment = segmentName
                    });
                }
                else
                {
                    selections.Add(new SegmentSelectionModel
                    {
                        Value1 = part,
                        Value2 = "",
                        Segment = segmentName
                    });
                }
            }
            return selections;
        }


        private static void LogError(Exception ex, string context = "SegmentSelectorViewModel")
        {
            ServiceLocator.Logger?.LogException(ex, context);
        }
    }
}
