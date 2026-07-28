// SimpleSegmentViewModel.cs in GLSense.Addin.Core
// Port of GLSense\ViewModels\SimpleSegmentViewModel.cs (FinalWorkingCode) -
// GLRollerGroups's ViewModel (Group H - Balance Configurator pane + LOVs/Roller/Account
// dialogs, last remaining piece). Re-pointed the same way as every other already-ported
// ViewModel in this project (see GLDailyRatesViewModel.cs header for the general mapping):
// GLSense.Models -> GLSense.Addin.Core.Models (SegmentModel/ScrollToTopMessage/
// ISegmentRow/TitleRow/SegmentDataRow/SegmentSelectionModel); GLSense.Repositories.
// DataRepository -> GLSense.Addin.Core.Repositories.DataRepository; GLSense.Service.
// SearchTypeService -> GLSense.Addin.Core.Models.SearchTypeService (already ported
// alongside SearchTypeModel, per Models\PeriodModels.cs header); LogUtility.* ->
// ServiceLocator.Logger?.*. Does NOT derive from GLSense.Base.NotifyBase (never ported
// into this project) - implements INotifyPropertyChanged directly instead, exactly like
// the old class already did. No logic changes vs. the original.
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
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
#nullable enable
    public class SimpleSegmentViewModel : INotifyPropertyChanged
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        public event Action<ScrollToTopMessage>? ScrollToTopRequested;

        private void ScrollDataGridsToTop()
        {
            ScrollToTopRequested?.Invoke(new ScrollToTopMessage
            {
                ScrollLeft = true,
                ScrollRight = true,
                Trigger = "PagingUpdated"
            });
        }
        public Action<string>? ShowWarningAction { get; set; }
        private readonly Dispatcher? _dispatcher;

        private ObservableCollection<SegmentModel>? _segments;
        public ObservableCollection<SegmentModel>? Segments
        {
            get => _segments;
            set
            {
                _segments = value;
                OnPropertyChanged();
            }
        }

        private SegmentModel? _selectedSegment;
        public SegmentModel? SelectedSegment
        {
            get => _selectedSegment;
            set
            {
                _selectedSegment = value;
                OnPropertyChanged();
                LoadSegmentValues();
            }
        }

        private List<ISegmentRow>? _allRows;
        private ObservableCollection<ISegmentRow>? _rows;
        public ObservableCollection<ISegmentRow>? Rows
        {
            get => _rows;
            set
            {
                _rows = value;
                OnPropertyChanged();
            }
        }

        // Selection models
        private ObservableCollection<SegmentSelectionModel>? _selectedRight;
        public ObservableCollection<SegmentSelectionModel>? SelectedItemsRight
        {
            get => _selectedRight;
            set
            {
                _selectedRight = value;
                OnPropertyChanged();
            }
        }

        // Multi-row checkbox state
        private bool _isMultipleRowsEnabled;
        public bool IsMultipleRowsEnabled
        {
            get => _isMultipleRowsEnabled;
            set
            {
                _isMultipleRowsEnabled = value;
                OnPropertyChanged();
            }
        }

        private bool _isMultipleRowsChecked;
        public bool IsMultipleRowsChecked
        {
            get => _isMultipleRowsChecked;
            set
            {
                _isMultipleRowsChecked = value;
                OnPropertyChanged();
            }
        }

        public static ObservableCollection<SearchTypeModel> SearchTypes => SearchTypeService.GetSearchTypes();

        private SearchTypeModel? _selectedSearchType;
        public SearchTypeModel? SelectedSearchType
        {
            get => _selectedSearchType;
            set => SetProperty(ref _selectedSearchType, value);
        }

        // Search/filter
        private string? _searchText;
        public string? SearchText
        {
            get => _searchText;
            set
            {
                _searchText = value;
                OnPropertyChanged();
                ApplySearchFilter();
            }
        }

        public int SegmentPickedIndex { get; set; } = -1;

        public SimpleSegmentViewModel(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
            _selectedRight = new ObservableCollection<SegmentSelectionModel>();
            Segments = new ObservableCollection<SegmentModel>();
            Rows = new ObservableCollection<ISegmentRow>();
            _selectedSearchType = SearchTypeService.GetDefaultSearchType();
            ServiceLocator.Logger?.LogDebug("SimpleSegmentViewModel.ctor: instance constructed.");
        }

        private void ApplySearchFilter()
        {
            try
            {
                var searchText = (_searchText ?? string.Empty).Trim();
                ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.ApplySearchFilter: searchText=\"{searchText}\", searchType={_selectedSearchType?.Value}");
                if (string.IsNullOrEmpty(searchText))
                {
                    Rows = new ObservableCollection<ISegmentRow>(_allRows);
                    ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.ApplySearchFilter: empty search text, restored {Rows.Count} unfiltered row(s).");
                    return;
                }

                if (_allRows == null)
                {
                    ServiceLocator.Logger?.LogDebug("SimpleSegmentViewModel.ApplySearchFilter: _allRows is null, nothing to filter.");
                    return;
                }

                FilterState filterState = new(searchText);
                ProcessRows(_allRows, filterState);
                Rows = new ObservableCollection<ISegmentRow>(filterState.FilteredRows);
                ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.ApplySearchFilter: filtered to {Rows.Count} row(s) out of {_allRows.Count}.");
            }
            catch (Exception ex)
            {
                LogError(ex, "SimpleSegmentViewModel.ApplySearchFilter");
            }
        }

        private void ProcessRows(IEnumerable<ISegmentRow> rows, FilterState state)
        {
            foreach (var row in rows)
            {
                if (row is TitleRow titleRow)
                    HandleTitleRow(titleRow, state);
                else if (row is SegmentDataRow dataRow)
                    HandleDataRow(dataRow, state);
            }
        }

        private static void HandleTitleRow(TitleRow titleRow, FilterState state)
        {
            state.ResetForNewTitle(titleRow);
        }

        private void HandleDataRow(SegmentDataRow dataRow, FilterState state)
        {
            if (!ShouldIncludeDataRow(dataRow, state))
                return;

            EnsureTitleIncluded(state);
            state.FilteredRows.Add(dataRow);
        }

        private bool ShouldIncludeDataRow(SegmentDataRow dataRow, FilterState state)
        {
            var searchType = _selectedSearchType?.Value;
            if (searchType == null)
                return false;

            return MatchesSearchFilter(dataRow, state.SearchText, searchType);
        }

        private static void EnsureTitleIncluded(FilterState state)
        {
            if (!state.IncludeCurrentTitle && state.CurrentTitle != null)
            {
                state.FilteredRows.Add(state.CurrentTitle);
                state.IncludeCurrentTitle = true;
            }
        }

        private sealed class FilterState(string searchText)
        {
            public string SearchText { get; } = searchText;
            public TitleRow? CurrentTitle { get; set; }
            public bool IncludeCurrentTitle { get; set; }
            public List<ISegmentRow> FilteredRows { get; } = new List<ISegmentRow>();

            public void ResetForNewTitle(TitleRow titleRow)
            {
                CurrentTitle = titleRow;
                IncludeCurrentTitle = false;
            }
        }


        private static bool MatchesSearchFilter(SegmentDataRow dataRow, string searchText, string searchType)
        {
            try
            {
                return searchType switch
                {
                    "StartsWith" => MatchesAnyField(dataRow, searchText, (s, t) => s?.StartsWith(t, StringComparison.CurrentCultureIgnoreCase) == true),
                    "DoesNotStartWith" => MatchesNone(dataRow, searchText, (s, t) => s?.StartsWith(t, StringComparison.CurrentCultureIgnoreCase) == true),
                    "EndsWith" => MatchesAnyField(dataRow, searchText, (s, t) => s?.EndsWith(t, StringComparison.CurrentCultureIgnoreCase) == true),
                    "DoesNotEndWith" => MatchesNone(dataRow, searchText, (s, t) => s?.EndsWith(t, StringComparison.CurrentCultureIgnoreCase) == true),
                    "Contains" => MatchesAnyField(dataRow, searchText, (s, t) => s?.IndexOf(t, StringComparison.CurrentCultureIgnoreCase) >= 0),
                    "NotContains" => MatchesNone(dataRow, searchText, (s, t) => s?.IndexOf(t, StringComparison.CurrentCultureIgnoreCase) >= 0),
                    "Equals" => MatchesAnyField(dataRow, searchText, (s, t) => string.Equals(s, t, StringComparison.CurrentCultureIgnoreCase)),
                    "NotEquals" => !MatchesAnyField(dataRow, searchText, (s, t) => string.Equals(s, t, StringComparison.CurrentCultureIgnoreCase)),
                    _ => true
                };
            }
            catch (Exception ex)
            {
                LogError(ex, "SimpleSegmentViewModel.MatchesSearchFilter");
                return false;
            }
        }

        private static bool MatchesAnyField(SegmentDataRow dataRow, string searchText, Func<string?, string?, bool> matchPredicate)
        {
            return matchPredicate(dataRow.SegmentValue, searchText) || matchPredicate(dataRow.Description, searchText);
        }

        private static bool MatchesNone(SegmentDataRow dataRow, string searchText, Func<string?, string?, bool> matchPredicate)
        {
            return !matchPredicate(dataRow.SegmentValue, searchText) && !matchPredicate(dataRow.Description, searchText);
        }


        // ----------- Segment, hierarchy loaders ----------------
        public async Task LoadSegmentsAsync(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.LoadSegmentsAsync: started. cubeId={cubeId}, ledgerId={ledgerId}");
            try
            {
                await Task.Run(() =>
                {
                    var repository = new DataRepository();
                    var segs = repository.GetSegments(cubeId, ledgerId);
                    ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.LoadSegmentsAsync: DataRepository.GetSegments returned {segs?.Count ?? 0} segment(s).");
                    _dispatcher?.Invoke(() =>
                    {
                        Segments?.Clear();
                        if (segs != null)
                        {
                            foreach (var s in segs)
                            {
                                Segments?.Add(s);
                            }
                        }
                        if (Segments?.Count > 0)
                        {
                            // Upper-bound check added alongside CLAUDE.md section 28's fix
                            // (Window_Loaded now actually populates SegmentPickedIndex from
                            // AppState.Instance.SegmentPickedIndex, so this branch is no
                            // longer dead code). Matches SegmentSelectorViewModel.
                            // SelectInitialSegment's own "&& SegmentPickedIndex <
                            // Segments.Count" guard - without it, a RibSegS selection made
                            // while a DIFFERENT cube/ledger (with a longer segment list) was
                            // active could leave SegmentPickedIndex pointing past the end of
                            // THIS cube/ledger's (shorter) segment list and throw an
                            // IndexOutOfRangeException here instead of just falling back to
                            // the first segment.
                            if (SegmentPickedIndex >= 0 && SegmentPickedIndex < Segments.Count)
                            {
                                ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.LoadSegmentsAsync: selecting segment at SegmentPickedIndex={SegmentPickedIndex}.");
                                SelectedSegment = Segments[SegmentPickedIndex];
                            }
                            else
                            {
                                ServiceLocator.Logger?.LogDebug("SimpleSegmentViewModel.LoadSegmentsAsync: no valid SegmentPickedIndex, defaulting to first segment.");
                                SelectedSegment = Segments[0];
                            }
                        }
                        else
                        {
                            ServiceLocator.Logger?.LogWarn($"SimpleSegmentViewModel.LoadSegmentsAsync: no segments returned for cubeId={cubeId}, ledgerId={ledgerId}.");
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                LogError(ex, "SimpleSegmentViewModel.LoadSegmentsAsync");
            }
        }

        private void LoadSegmentValues()
        {
            try
            {
                if (SelectedSegment == null)
                {
                    ServiceLocator.Logger?.LogDebug("SimpleSegmentViewModel.LoadSegmentValues: SelectedSegment is null, clearing Rows.");
                    Rows?.Clear();
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.LoadSegmentValues: started for segment \"{SelectedSegment.SegmentName}\".");
                var items = DataRepository.GetSegmentValues_RG(SelectedSegment);

                if (items == null)
                {
                    ServiceLocator.Logger?.LogWarn($"SimpleSegmentViewModel.LoadSegmentValues: DataRepository.GetSegmentValues_RG returned null for segment \"{SelectedSegment.SegmentName}\".");
                    Rows?.Clear();
                    return;
                }

                var tempRows = new List<ISegmentRow>();
                TitleRow? currentTitle = null;
                foreach (var item in items)
                {
                    if (string.Equals(item.SummaryFlag, "RG", StringComparison.OrdinalIgnoreCase))
                    {
                        // This is a title row
                        currentTitle = new TitleRow { Title = item.Description };
                        tempRows.Add(currentTitle);
                    }
                    else
                    {
                        // Normal data row under current title
                        if (currentTitle != null)
                        {
                            tempRows.Add(new SegmentDataRow
                            {
                                SegmentValue = item.SegmentValue,
                                Description = item.Description,
                                SegmentName = item.SegmentName,
                                SegmentValueSetId = item.SegmentValueSetId,
                                SummaryFlag = item.SummaryFlag
                            });
                        }
                        else
                        {
                            // No title yet - add as normal row or handle as needed
                            tempRows.Add(new SegmentDataRow
                            {
                                SegmentValue = item.SegmentValue,
                                Description = item.Description,
                                SegmentName = item.SegmentName,
                                SegmentValueSetId = item.SegmentValueSetId,
                                SummaryFlag = item.SummaryFlag
                            });
                        }
                    }
                }

                _allRows = tempRows;
                Rows = new ObservableCollection<ISegmentRow>(_allRows);
                ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.LoadSegmentValues: built {_allRows.Count} row(s) for segment \"{SelectedSegment.SegmentName}\".");
                ApplySearchFilter();
                ScrollDataGridsToTop();
            }
            catch (Exception ex)
            {
                LogError(ex, "SimpleSegmentViewModel.LoadSegmentValues");
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        protected bool SetProperty<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
        // ---------------- Multi-row logic ----------------
        private void UpdateMultiRowState()
        {
            try
            {
                if (SelectedItemsRight == null || SelectedItemsRight.Count == 0)
                {
                    IsMultipleRowsEnabled = false;
                    IsMultipleRowsChecked = false;
                    return;
                }

                var distinctSegments = SelectedItemsRight.Select(r => r.Segment).Distinct().Count();
                if (distinctSegments == 1)
                {
                    IsMultipleRowsEnabled = true;
                    if (SelectedItemsRight.Count > 0) IsMultipleRowsChecked = true;
                }
                else
                {
                    IsMultipleRowsEnabled = false;
                    IsMultipleRowsChecked = false;
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "SimpleSegmentViewModel.UpdateMultiRowState");
            }
        }

        public void AddSelection(IList<SegmentDataRow> selectedItems)
        {
            ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.AddSelection: started. selectedItems count={selectedItems?.Count ?? 0}");
            try
            {
                if (selectedItems == null)
                {
                    ServiceLocator.Logger?.LogWarn("SimpleSegmentViewModel.AddSelection: selectedItems was null.");
                    ShowWarning("Please select at least one item to add.");
                    return;
                }

                if (!IsValidSelection(selectedItems))
                    return;

                var newItems = selectedItems
                    .Where(IsValidToAdd)
                    .Where(NotAlreadyExists)
                    .Select(seg => new SegmentSelectionModel
                    {
                        Value1 = seg.SegmentValue,
                        Value2 = null,
                        Segment = seg.SegmentName
                    })
                    .ToList();

                if (newItems.Count > 0)
                {
                    foreach (var item in newItems)
                        _selectedRight?.Add(item);
                    UpdateMultiRowState();
                    ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.AddSelection: added {newItems.Count} item(s) to right grid.");
                }
                else
                {
                    ServiceLocator.Logger?.LogDebug("SimpleSegmentViewModel.AddSelection: no new items to add (all invalid or duplicates).");
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "SimpleSegmentViewModel.AddSelection");
            }
        }

        private bool IsValidSelection(IList<SegmentDataRow> selectedItems)
        {
            if (selectedItems == null || selectedItems.Count == 0)
            {
                ShowWarning("Please select at least one item to add.");
                return false;
            }
            return true;
        }

        private bool IsValidToAdd(SegmentDataRow seg)
        {
            if (seg.SummaryFlag == "RG")
            {
                ShowWarning("Title cannot be added as a selected segment value.");
                return false;
            }
            return true;
        }

        private bool NotAlreadyExists(SegmentDataRow seg)
        {
            var exists = _selectedRight!.Any(r => r.Value1 == seg.SegmentValue && r.Segment == seg.SegmentName);
            if (exists)
            {
                ShowSingleItemWarning(seg.SegmentValue);
            }
            return !exists;
        }

        private void ShowWarning(string message)
        {
            _dispatcher?.Invoke(() => ShowWarningAction?.Invoke(message));
        }

        private void ShowSingleItemWarning(string value)
        {
            if (ShowWarningAction != null)
            {
                _dispatcher?.Invoke(() => ShowWarningAction.Invoke($"'{value}' already exists in the right grid."));
            }
        }


        public void RemoveSelection(IList<SegmentSelectionModel> selectedItems)
        {
            ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.RemoveSelection: started. selectedItems count={selectedItems?.Count ?? 0}");
            try
            {
                if (selectedItems == null || selectedItems.Count == 0)
                {
                    ServiceLocator.Logger?.LogDebug("SimpleSegmentViewModel.RemoveSelection: no items selected to remove.");
                    if (ShowWarningAction != null)
                    {
                        _dispatcher?.Invoke(() =>
                        {
                            ShowWarningAction.Invoke("Please select one or more items to remove.");
                        });
                    }
                    return;
                }

                foreach (var sel in selectedItems.ToList())
                {
                    _selectedRight?.Remove(sel);
                }

                UpdateMultiRowState();
                ServiceLocator.Logger?.LogDebug($"SimpleSegmentViewModel.RemoveSelection: removed {selectedItems.Count} item(s) from right grid.");
            }
            catch (Exception ex)
            {
                LogError(ex, "SimpleSegmentViewModel.RemoveSelection");
            }
        }

        // ---------------- Excel interop for refedit controls ----------------
        private Microsoft.Office.Interop.Excel.Application? _excelApp;
        public Microsoft.Office.Interop.Excel.Application? ExcelApp
        {
            get => _excelApp;
            set
            {
                _excelApp = value;
                OnPropertyChanged(nameof(ExcelApp));
            }
        }

        // Error logging method
        private static void LogError(Exception ex, string context = "SimpleSegmentViewModel")
        {
            ServiceLocator.Logger?.LogException(ex, context);
        }
    }
#nullable disable
}
