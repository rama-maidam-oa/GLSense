using ControlzEx.Standard;
using GLSense.Helpers;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.ViewModels
{
    public class GLLovViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<LedgerModel> Ledgers { get; set; }
        public ObservableCollection<LovRow> LOVRows { get; set; }

        private readonly Dispatcher _dispatcher;
        public Action<string> ShowWarningAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

        // -----------Segments collections-----------------------
        private ObservableCollection<SegmentModel> _segments;
        public ObservableCollection<SegmentModel> Segments
        {
            get => _segments;
            set
            {
                _segments = value;
                OnPropertyChanged();
            }
        }
        // ------------------------------------------------------

        // -------------------Selected Ledger--------------------
        private LedgerModel _lOV_SelectedLedger;
        public LedgerModel LOV_SelectedLedger
        {
            get => _lOV_SelectedLedger;
            set
            {
                LogUtility.LogDebug($"GLLovViewModel.LOV_SelectedLedger (set): LedgerId={value?.LedgerId}, LedgerName={value?.LedgerName}");
                _lOV_SelectedLedger = value;
                OnPropertyChanged();
                LoadLovRows();
            }
        }
        // ------------------------------------------------------
        // -------------------Selected Lov--------------------
        private LovRow _selectedLov;
        public LovRow SelectedLov
        {
            get => _selectedLov;
            set
            {
                LogUtility.LogDebug($"GLLovViewModel.SelectedLov (set): Name={value?.Name}, Category={value?.Category}");
                _selectedLov = value;
                OnPropertyChanged(nameof(SelectedLov)); // if using INotifyPropertyChanged
            }
        }
        // ------------------------------------------------------
        private Excel.Application _excelApp;
        public Excel.Application ExcelApp
        {
            get => _excelApp;
            set
            {
                _excelApp = value;
                OnPropertyChanged(nameof(ExcelApp));
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        public GLLovViewModel(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
            Ledgers = new ObservableCollection<LedgerModel>();
            LOVRows = new ObservableCollection<LovRow>();
            Segments = new ObservableCollection<SegmentModel>();
        }

        // DATA LOADING:
        public async Task LoadDataAsync(long selectedCubeId, long? defaultLedgerId)
        {
            LogUtility.LogDebug($"GLLovViewModel.LoadDataAsync: selectedCubeId={selectedCubeId}, defaultLedgerId={defaultLedgerId}");
            // 1. Load Ledgers for cube, set SelectedLedger to the default
            var repository = new DataRepository();
            var allLedgers = await Task.Run(() => repository.GetLedgers(selectedCubeId));
            LogUtility.LogDebug($"GLLovViewModel.LoadDataAsync: loaded {allLedgers?.Count ?? 0} ledger(s) for CubeId={selectedCubeId}");
            await _dispatcher.InvokeAsync(() =>
            {
                Ledgers.Clear();
                foreach (var l in allLedgers)
                {
                    Ledgers.Add(l);
                }
                if (AppState.Instance.SelectedLedger != null)
                {
                    LOV_SelectedLedger = AppState.Instance.SelectedLedger;
                }
                else
                {
                    LOV_SelectedLedger = Ledgers.FirstOrDefault(x => x.LedgerId == defaultLedgerId);
                }

            });
        }

        private void LoadLovRows()
        {
            LogUtility.LogDebug($"GLLovViewModel.LoadLovRows: LedgerId={LOV_SelectedLedger?.LedgerId}");
            Task.Run(async () =>
            {
                try
                {
                    await LoadLovRowsAsync();
                }
                catch (Exception ex)
                {
                    LogError(ex);
                }
            });
        }
        private bool IsLedgerDataExist()
        {
            var intCount = DataRepository.GetTableItemsCount(AppState.Instance.SelectedCube.CubeId, LOV_SelectedLedger.LedgerId, "SEGMENTS");
            LogUtility.LogDebug($"GLLovViewModel.IsLedgerDataExist: SEGMENTS count={intCount} for LedgerId={LOV_SelectedLedger.LedgerId}");
            return intCount != 0;
        }
        public async Task LoadLovRowsAsync()
        {
            LogUtility.LogDebug($"GLLovViewModel.LoadLovRowsAsync: entry. LedgerId={LOV_SelectedLedger?.LedgerId}");
            var rows = new List<LovRow>();
            bool ledgerDataExist = IsLedgerDataExist();

            CancellationHelper ctsHelper = new();
            CancellationToken token = ctsHelper.GetToken();

            try
            {
                if (!ledgerDataExist && ShowBusyAction != null)
                {
                    // Use InvokeAsync so the dispatcher call is awaitable (Invoke returns void)
                    // Dispatched to the UI thread since ShowBusyAction mutates bound busy-overlay state.
                    await _dispatcher.InvokeAsync(async () =>
                    {
                        await ShowBusyAction.Invoke("Fetching ledger data... (click Cancel to stop)",
                            async () =>
                            {
                                if (!ctsHelper.IsCancellationRequested)
                                {
                                    ctsHelper.Cancel();
                                }
                                await Task.CompletedTask;
                            });
                    });
                    await CommonFunctions.FillResponsibilitiesAsync(LOV_SelectedLedger.LedgerId, AppState.Instance.SelectedCube.CubeId, token);
                }

                var repository = new DataRepository();

                // a. All SEGMENTs (and their counts from SEGMENT_VALUES)
                // Dispatched to the UI thread since the Segments setter raises PropertyChanged.
                await _dispatcher.InvokeAsync(() =>
                {
                    Segments = repository.GetSegments(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId);
                });

                foreach (var seg in Segments)
                {
                    var count = DataRepository.GetSegmentItemsCount(seg);
                    rows.Add(new LovRow { Name = seg.SegmentName, ItemsCount = count, Category = "Segment" });
                }

                // b. Activity (DB Table)
                var activityCount = DataRepository.GetTableItemsCount(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId, "ACTIVITY");
                if (activityCount > 0)
                {
                    rows.Add(new LovRow { Name = "Activity", ItemsCount = activityCount, Category = "Activity" });
                }

                // c. Actual Flag (Hardcoded)
                var actualFlags = new List<string> { "Actual", "Budget", "Encumbrance", "Actual+Encumbrance" };
                rows.Add(new LovRow { Name = "Actual Flag", ItemsCount = actualFlags.Count, Category = "Actual Flag" });

                // d. Balance Type (Hardcoded)
                var balTypes = new List<string> { "PTD", "YTD", "QTD", "PJTD", "CTD", "JED", "JEDP", "JEDU" };
                rows.Add(new LovRow { Name = "Balance Type", ItemsCount = balTypes.Count, Category = "Balance Type" });

                // e. Budgets (DB Table)
                var budgetCount = DataRepository.GetTableItemsCount(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId, "BUDGETS");
                if (budgetCount > 0)
                {
                    rows.Add(new LovRow { Name = "Budgets", ItemsCount = budgetCount, Category = "Budgets" });
                }

                // f. Currencies (DB Table)
                var currenciesCount = DataRepository.GetTableItemsCount(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId, "CURRENCIES");
                if (currenciesCount > 0)
                {
                    rows.Add(new LovRow { Name = "Currencies", ItemsCount = currenciesCount, Category = "Currencies" });
                }

                // g. Currency Type (Hardcoded)
                var currencyTypes = new List<string> { "Total", "Entered", "Translated", "Converted" };
                rows.Add(new LovRow { Name = "Currency Type", ItemsCount = currencyTypes.Count, Category = "Currency Type" });

                // h. Encumbrances (DB Table)
                var encumCount = DataRepository.GetTableItemsCount(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId, "ENCUMBRANCES");
                if (encumCount > 0)
                {
                    rows.Add(new LovRow { Name = "Encumbrances", ItemsCount = encumCount, Category = "Encumbrances" });
                }

                // i. Journal Categories (DB Table)
                var jcCount = DataRepository.GetTableItemsCount(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId, "JOURNALCATEGORIES");
                if (jcCount > 0)
                {
                    rows.Add(new LovRow { Name = "Journal Categories", ItemsCount = jcCount, Category = "Journal Categories" });
                }

                // j. Journal Sources (DB Table)
                var jsCount = DataRepository.GetTableItemsCount(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId, "JOURNALSOURCES");
                if (jsCount > 0)
                {
                    rows.Add(new LovRow { Name = "Journal Sources", ItemsCount = jsCount, Category = "Journal Sources" });
                }

                // k. Periods (DB Table)
                var periodsCount = DataRepository.GetTableItemsCount(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId, "PERIODS");
                if (periodsCount > 0)
                {
                    rows.Add(new LovRow { Name = "Periods", ItemsCount = periodsCount, Category = "Periods" });
                }
                // Fill ViewModel collection on UI thread
                await _dispatcher.InvokeAsync(() =>
                {
                    LOVRows.Clear();
                    foreach (var r in rows)
                    {
                        LOVRows.Add(r);
                    }
                });

                if (!ledgerDataExist && HideBusyAsyncAction != null)
                {
                    var task = HideBusyAsyncAction.Invoke();
                    await task;
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Ledger data fetch operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GLLovViewModel.LoadLovRowsAsync");
                if (ShowWarningAction != null)
                {
                    await _dispatcher.InvokeAsync(() =>
                    {
                        ShowWarningAction.Invoke("An error occurred while fetching ledger data: " + ex.Message);
                    });
                }
            }
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static void LogError(Exception ex)
        {
            LogUtility.LogException(ex, "GLLovViewModel.LoadLovRows");
        }
    }
}
