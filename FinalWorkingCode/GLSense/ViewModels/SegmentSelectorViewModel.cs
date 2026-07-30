using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Service;
using GLSense.Utilities;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.ViewModels
{
    public class SegmentSelectorViewModel(Dispatcher dispatcher, string iWindow, string svals) : INotifyPropertyChanged
    {
        public event Action<ScrollToTopMessage> ScrollToTopRequested;

        private readonly Dispatcher _dispatcher = dispatcher;
        private readonly string _segValues = svals;
        private readonly string _windowName = iWindow;

        // Actions for window overlay controls
        public Action<string> ShowWarningAction { get; set; }
        public Action<string> ShowInfoAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

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
                LogUtility.LogDebug($"SegmentSelectorViewModel.SelectedSearchType (set): {_selectedSearchType?.Value} -> {value?.Value}");
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
                LogUtility.LogDebug($"SegmentSelectorViewModel.SelectedSegment (set): SegmentName={value?.SegmentName}");
                if (_windowName == "Ref" && _selectedSegment != null)
                {
                    // Save old selections
                    _selectedSegment.SelectedValues =
                        new ObservableCollection<SegmentSelectionModel>(_selectedRight);

                    OnPropertyChanged(nameof(SelectedItemsRight));
                }

                if (SetProperty(ref _selectedSegment, value))
                {
                    // A hierarchy selected against the PREVIOUS segment no longer means
                    // anything once the segment changes - HierarchyItems below gets
                    // repopulated for the new segment, but nothing was clearing the
                    // still-selected value itself, so the Hierarchy combo (bound to
                    // SelectedHierarchy) kept showing the old segment's stale selection.
                    // Clear the backing field directly (rather than going through the
                    // SelectedHierarchy property setter) so this doesn't also kick off
                    // LoadHierarchySegmentValuesAsync's hierarchy-data fetch, which would
                    // otherwise run concurrently with - and race against - the
                    // LoadSegmentValuesAsync() call below for the newly selected segment.
                    if (_selectedHierarchy != null)
                    {
                        _selectedHierarchy = null;
                        OnPropertyChanged(nameof(SelectedHierarchy));
                    }

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
                LogUtility.LogDebug($"SegmentSelectorViewModel.PageSize (set): {_pageSize} -> {value}");
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
                LogUtility.LogDebug($"SegmentSelectorViewModel.CurrentPage (set): {_currentPage} -> {value} (TotalPages={TotalPages})");
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
                LogUtility.LogDebug($"SegmentSelectorViewModel.SearchText (set): '{_searchText}' -> '{value}'");
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
                    LogUtility.LogDebug($"SegmentSelectorViewModel.SelectedHierarchy (set): SegmentValue={value?.SegmentValue}");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.LoadSegmentsAsync: CubeId={cubeId}, LedgerId={ledgerId}");
            await Task.Run(() =>
            {
                var repository = new DataRepository();
                var segs = repository.GetSegments(cubeId, ledgerId);
                LogUtility.LogDebug($"SegmentSelectorViewModel.LoadSegmentsAsync: loaded {segs?.Count ?? 0} segment(s).");
                _dispatcher.Invoke(() => ProcessSegments(segs));
            });
        }

        private void ProcessSegments(IEnumerable<SegmentModel> segs)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.ProcessSegments: segs.Count={segs?.Count() ?? 0}");
            foreach (var s in Segments)
                s.PropertyChanged -= OnSegmentValueChanged;

            Segments.Clear();

            int index = 0;
            foreach (var s in segs)
            {
                s.SegmentName = s.SegmentName.Trim();
                InitializeSegment(s, index++);
            }

            SelectInitialSegment();
        }

        private void InitializeSegment(SegmentModel s, int index)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.InitializeSegment: SegmentName={s?.SegmentName}, index={index}, WindowName={_windowName}");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.ProcessRefSegment: SegmentName={s?.SegmentName}, index={index}, parts.Count={parts?.Count ?? 0}");
            if (parts.Count <= index)
            {
                LogUtility.LogWarn($"SegmentSelectorViewModel.ProcessRefSegment: no part found at index={index} for SegmentName={s.SegmentName} (parts.Count={parts.Count}). Clearing segment values.");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.ResolveSegmentValueText: segmentValues='{segmentValues}'");
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
                LogUtility.LogException(ex, "SegmentSelectorViewModel.ResolveSegmentValueText");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.SelectInitialSegment: Segments.Count={Segments?.Count ?? 0}, AppState.Instance.SegmentPickedIndex={AppState.Instance.SegmentPickedIndex}");
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
            else
            {
                LogUtility.LogWarn("SegmentSelectorViewModel.SelectInitialSegment: no segments available to select.");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.LoadSegmentValuesAsync: SegmentName={(segModel ?? SelectedSegment)?.SegmentName}, fromHierarchy={fromHierarchy}, HierarchySegmentValue={segValModel?.SegmentValue}");
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
                LogUtility.LogDebug($"SegmentSelectorViewModel.LoadSegmentValuesAsync: loaded {_allSegmentValues.Count} value(s).");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: entry. SelectedHierarchy={SelectedHierarchy?.SegmentValue}");
            if (SelectedHierarchy == null)
            {
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
                LogUtility.LogDebug($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: hierarchy cached={hierarchyBool} for SegmentValue={SelectedHierarchy?.SegmentValue}");

                if (!hierarchyBool)
                {
                    var hierarchyData = await HierarhyApiAsync(SelectedHierarchy, token);
                    if (!string.IsNullOrWhiteSpace(hierarchyData))
                    {
                        DataRepository.SaveHierarchyToCache(SelectedHierarchy, hierarchyData);
                        LogUtility.LogDebug($"SegmentSelectorViewModel.LoadHierarchySegmentValuesAsync: saved hierarchy data to cache for SegmentValue={SelectedHierarchy?.SegmentValue}.");
                    }
                }
                await LoadSegmentValuesAsync(null, SelectedHierarchy, true);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("User cancelled operation. Loading hierarchy segment values interrupted.");
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
            finally
            {
                try
                {
                    if (!ctsHelper.IsCancellationRequested)
                        ctsHelper.Cancel();

                    ctsHelper.Dispose();  // ✅ ALWAYS SAFE - handles ALL cases
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"LoadHierarchySegmentValuesAsync: failed disposing CancellationHelper (non-fatal): {ex.Message}");
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
                    // Has reference → disable textbox, enable refedit
                    s.IsTextEnabled = false;
                    s.IsRefEditEnabled = true;
                }
                else if (s.SelectedValues != null && s.SelectedValues.Any())
                {
                    // Has values → enable textbox, disable refedit
                    s.IsTextEnabled = true;
                    s.IsRefEditEnabled = false;
                }
                else
                {
                    // Both empty → enable both
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.UpdatePagingAndGrid: TotalRecords={_totalRecords}, TotalPages={_totalPages}, CurrentPage={_currentPage}, PagedCount={_pagedSegmentValues.Count}");

            OnPropertyChanged(nameof(PagedSegmentValues));
            OnPropertyChanged(nameof(CurrentPage));
            OnPropertyChanged(nameof(TotalPages));
            OnPropertyChanged(nameof(TotalRecords));
            OnPropertyChanged(nameof(PageRangeText));

            // Scroll to top after data is loaded
            ScrollDataGridsToTop();
        }

        private List<SegmentValueModel> ApplySearchFilter(List<SegmentValueModel> source)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.ApplySearchFilter: source.Count={source?.Count ?? 0}, SearchText='{_searchText}', SearchType={_selectedSearchType?.Value}");
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
            LogUtility.LogDebug("SegmentSelectorViewModel.GoFirstPage: entry");
            _currentPage = 1;
            UpdatePagingAndGrid();
        }

        public void GoPreviousPage()
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.GoPreviousPage: CurrentPage={_currentPage}");
            if (_currentPage > 1) _currentPage--;
            UpdatePagingAndGrid();
        }

        public void GoNextPage()
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.GoNextPage: CurrentPage={_currentPage}, TotalPages={_totalPages}");
            if (_currentPage < _totalPages) _currentPage++;
            UpdatePagingAndGrid();
        }

        public void GoLastPage()
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.GoLastPage: TotalPages={_totalPages}");
            _currentPage = _totalPages;
            UpdatePagingAndGrid();
        }

        // Called after changing page size
        public void ApplyPageSize(int size)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.ApplyPageSize: {_pageSize} -> {size}");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.AddSelection: selectedItemCount={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count == 0)
            {
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select at least one item to add."));
                }
                return;
            }

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
            }

            UpdateMultiRowState();
        }

        public void RemoveSelection(IList<SegmentSelectionModel> selectedItems)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.RemoveSelection: selectedItemCount={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count == 0)
            {
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

            UpdateMultiRowState();
        }

        public void AddBetweenSelection(IList<SegmentValueModel> selectedItems, bool isExclude)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.AddBetweenSelection: selectedItemCount={selectedItems?.Count ?? 0}, isExclude={isExclude}");
            if (selectedItems == null || selectedItems.Count < 2)
            {
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select two or more items (the first → Value1, the last → Value2)."));
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
                ShowWarningAction?.Invoke($"Range '{val1} - {val2}' already exists.");
                return;
            }
            _selectedRight.Add(new SegmentSelectionModel { Value1 = val1, Value2 = val2, Segment = seg1.SegmentName });
            OnPropertyChanged(nameof(SelectedItemsRight));

            UpdateMultiRowState();
        }

        public void AddNotBetweenSelection(IList<SegmentValueModel> selectedItems)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.AddNotBetweenSelection: selectedItemCount={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count < 2)
            {
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select two or more items (the first → Value1, the last → Value2)."));
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
                ShowWarningAction?.Invoke($"Range '{val1} - {val2}' already exists.");
                return;
            }
            _selectedRight.Add(new SegmentSelectionModel { Value1 = val1, Value2 = val2, Segment = seg1.SegmentName });
            OnPropertyChanged(nameof(SelectedItemsRight));

            UpdateMultiRowState();
        }

        public void AddExcludeSelection(IList<SegmentValueModel> selectedItems)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.AddExcludeSelection: selectedItemCount={selectedItems?.Count ?? 0}");
            if (selectedItems == null || selectedItems.Count == 0)
            {
                if (ShowWarningAction != null)
                {
                    _dispatcher.Invoke(() => ShowWarningAction.Invoke("Please select one or more items to exclude."));
                }
                return;
            }

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
            }

            UpdateMultiRowState();
        }

        public void ClearDefaults()
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.ClearDefaults: Segments.Count={Segments?.Count ?? 0}");
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
            LogUtility.LogDebug($"SegmentSelectorViewModel.GetAllSegmentValues: entry. Segments.Count={Segments?.Count ?? 0}");
            if (Segments == null || Segments.Count == 0) return string.Empty;

            var result = new List<string>();

            foreach (var s in Segments)
            {
                string segVal = string.Empty;

                // 🥇 Priority 1: reference (if available)
                if (!string.IsNullOrWhiteSpace(s.Reference))
                {
                    segVal = s.Reference.Trim();
                }
                // 🥈 Priority 2: value (if reference is empty)
                else if (!string.IsNullOrWhiteSpace(s.Value))
                {
                    segVal = $"\"{s.Value.Trim()}\"";
                }

                result.Add(segVal);
            }

            var joined = string.Join(";", result);
            LogUtility.LogDebug($"SegmentSelectorViewModel.GetAllSegmentValues: result='{joined}'");
            return joined;
        }

        // ---------- API (hierarchy) async call helper ----------
        public async Task<string> HierarhyApiAsync(SegmentValueModel selectedHierarchy, CancellationToken token)
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.HierarhyApiAsync: entry. SegmentValue={selectedHierarchy?.SegmentValue}, SegmentValueSetId={selectedHierarchy?.SegmentValueSetId}, CubeId={selectedHierarchy?.CubeId}");
            try
            {

                string apiUrl =
                    $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}segment-hierarchy" +
                    $"?segmentValueSetId={selectedHierarchy.SegmentValueSetId}" +
                    $"&parentSegmentValue={WebUtility.UrlEncode(selectedHierarchy.SegmentValue.Trim())}" +
                    $"&cubeId={selectedHierarchy.CubeId}";

                token.ThrowIfCancellationRequested();

                LogUtility.LogDebug($"SegmentSelectorViewModel.HierarhyApiAsync: calling ServerAPI. apiUrl={apiUrl}");
                string response =
                    await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", token);
                LogUtility.LogDebug($"SegmentSelectorViewModel.HierarhyApiAsync: ServerAPI responded. response.Length={response?.Length ?? 0}");

                token.ThrowIfCancellationRequested();

                ValidateTransportResponse(response);

                var result = ApiResponseHelper.Parse<JsonElement>(response, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    LogUtility.LogWarn($"Hierarchy API failed: {apiUrl}");
                    LogUtility.LogWarn($"Response: {response}");

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

                return response;
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn(
                    "User cancelled operation. Fetching hierarchy data interrupted.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SegmentSelectorViewModel.HierarhyApiAsync");
            }
            return string.Empty;
        }
        private static void ValidateTransportResponse(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
                throw new InvalidOperationException("Empty API response.");

            if (response.IndexOf("(401)", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new UnauthorizedAccessException("Session expired.");

            if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase) ||
                response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException(response);
        }

        // ---------------- Multi-row and Segment Ref logic ----------------
        private void UpdateMultiRowState()
        {
            LogUtility.LogDebug($"SegmentSelectorViewModel.UpdateMultiRowState: WindowName={_windowName}, SelectedItemsRight.Count={_selectedRight?.Count ?? 0}");
            if (_windowName.Contains("Ref"))
                UpdateRefWindowState();
            else
                UpdateNonRefWindowState();
        }

        private void UpdateRefWindowState()
        {
            if (SelectedSegment == null) return;

            if (_selectedRight.Any())
            {
                SelectedSegment.Value = BuildValueFromSelections();
                SelectedSegment.IsTextEnabled = true;
                SelectedSegment.IsRefEditEnabled = false;
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

            LogUtility.LogDebug($"SegmentSelectorViewModel.OnSegmentValueChanged: SegmentName={seg.SegmentName}, PropertyName={e.PropertyName}");
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
            }
            else
            {
                // Reference cleared - restore from value or clear
                if (!string.IsNullOrWhiteSpace(seg.Value))
                {
                    seg.SelectedValues = ParseValueToSelections(seg.Value, seg.SegmentName);
                }
                else
                {
                    seg.SelectedValues.Clear();
                }
                ApplyEnableState(seg);
            }

            // Sync right grid if this is the active segment
            if (_selectedSegment == seg)
                UpdateMultiRowState();
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


        private static void LogError(Exception ex, [System.Runtime.CompilerServices.CallerMemberName] string context = "")
        {
            LogUtility.LogException(ex, $"SegmentSelectorViewModel.{context}");
        }
    }
}
