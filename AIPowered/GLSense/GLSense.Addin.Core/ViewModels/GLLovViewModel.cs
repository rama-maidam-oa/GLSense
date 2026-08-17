// GLLovViewModel.cs in GLSense.Addin.Core
// Port of GLSense\ViewModels\GLLovViewModel.cs (FinalWorkingCode) - GLLOVs's ViewModel
// (Group H - Balance Configurator pane + LOVs/Roller/Account dialogs, last remaining
// piece). Re-pointed the same way as every other already-ported ViewModel in this project
// (see GLDailyRatesViewModel.cs header for the general mapping): GLSense.Models ->
// GLSense.Addin.Core.Models; GLSense.Repositories.DataRepository ->
// GLSense.Addin.Core.Repositories.DataRepository; GLSense.Utilities.AppState ->
// GLSense.Addin.Core.AppState; GLSense.Helpers.CancellationHelper ->
// GLSense.Addin.Core.Helpers.CancellationHelper; GLSense.Utilities.CommonFunctions ->
// GLSense.Addin.Core.Utilities.CommonFunctions; LogUtility.* ->
// ServiceLocator.Logger?.*. Does NOT derive from GLSense.Base.NotifyBase (never ported
// into this project - see Models\PeriodModels.cs header for the established rationale) -
// implements INotifyPropertyChanged directly instead, exactly like the old class already
// did. Adds a GlobalState passthrough (GlobalStateViewModel.Instance) so GLLOVs.xaml can
// bind its "Reference:" ExcelRefEditControl the same way every other already-ported
// dialog does (see GLDailyRates.xaml's CellReference binding) - the old code set
// excelRefEdit.Text directly in code-behind instead, since GlobalStateViewModel didn't
// exist yet at that point in the old project.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.ViewModels
{
    public class GLLovViewModel : INotifyPropertyChanged
    {
        [ComVisible(false)]
        public GlobalStateViewModel GlobalState => GlobalStateViewModel.Instance;

        public ObservableCollection<LedgerModel> Ledgers { get; set; }
        public ObservableCollection<LovRow> LOVRows { get; set; }

        private readonly Dispatcher _dispatcher;
        public Action<string> ShowWarningAction { get; set; }
        public Func<string, Func<Task>, Task> ShowBusyAction { get; set; }
        public Func<Task> HideBusyAsyncAction { get; set; }

        // Set by the View so it can resettle its SizeToContent window once LOVRows
        // actually has real data - LoadLovRows() is fired fire-and-forget from
        // LOV_SelectedLedger's setter (detached from Window_Loaded's own await chain),
        // so BaseWindow.OnLoaded's resettle always runs against an empty grid. See
        // CLAUDE.md section 1.4b (GLCubeDetails) for the full history of this pattern.
        public Action DataLoadedAction { get; set; }

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
                _selectedLov = value;
                OnPropertyChanged(nameof(SelectedLov));
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
            ServiceLocator.Logger?.LogDebug("GLLovViewModel.ctor: instance constructed.");
        }

        // DATA LOADING:
        public async Task LoadDataAsync(long selectedCubeId, long? defaultLedgerId)
        {
            ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadDataAsync: started. selectedCubeId={selectedCubeId}, defaultLedgerId={(defaultLedgerId.HasValue ? defaultLedgerId.Value.ToString() : "null")}");
            try
            {
                // 1. Load Ledgers for cube, set SelectedLedger to the default
                var repository = new DataRepository();
                ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadDataAsync: calling DataRepository.GetLedgers for cubeId={selectedCubeId}");
                var allLedgers = await Task.Run(() => repository.GetLedgers(selectedCubeId));
                ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadDataAsync: DataRepository.GetLedgers returned {allLedgers?.Count ?? 0} ledger(s).");
                await _dispatcher.InvokeAsync(() =>
                {
                    Ledgers.Clear();
                    foreach (var l in allLedgers)
                    {
                        Ledgers.Add(l);
                    }

                    // Sets the backing field directly (not the LOV_SelectedLedger property)
                    // specifically to skip its setter's LoadLovRows() call - that call is
                    // fire-and-forget (Task.Run, never awaited by its caller), which meant
                    // this method previously returned as soon as the ledger dropdown was
                    // populated while the actual grid content kept loading in the background
                    // *after* PrepareAsync had already gone on to call ShowDialog. Explicitly
                    // awaiting LoadLovRowsAsync() below instead makes this method genuinely
                    // block until the grid has real data. A later user-driven ledger change
                    // (window already open) still goes through the normal property setter
                    // further down in this class, so that fire-and-forget-with-busy-overlay
                    // UX is unaffected. Ported from FinalWorkingCode's identical fix.
                    if (AppState.Instance.SelectedLedger != null)
                    {
                        _lOV_SelectedLedger = AppState.Instance.SelectedLedger;
                        ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadDataAsync: using AppState.SelectedLedger \"{LOV_SelectedLedger?.LedgerName}\" (LedgerId={LOV_SelectedLedger?.LedgerId}).");
                    }
                    else
                    {
                        _lOV_SelectedLedger = Ledgers.FirstOrDefault(x => x.LedgerId == defaultLedgerId);
                        if (LOV_SelectedLedger == null)
                            ServiceLocator.Logger?.LogWarn($"GLLovViewModel.LoadDataAsync: no ledger found matching defaultLedgerId={defaultLedgerId} among {Ledgers.Count} loaded ledger(s).");
                        else
                            ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadDataAsync: defaulted to ledger \"{LOV_SelectedLedger.LedgerName}\" (LedgerId={LOV_SelectedLedger.LedgerId}).");
                    }
                    OnPropertyChanged(nameof(LOV_SelectedLedger));
                });

                await LoadLovRowsAsync();

                ServiceLocator.Logger?.LogDebug("GLLovViewModel.LoadDataAsync: completed.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLLovViewModel.LoadDataAsync");
                throw;
            }
        }

        private void LoadLovRows()
        {
            ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadLovRows: triggered for ledger \"{LOV_SelectedLedger?.LedgerName}\" (LedgerId={LOV_SelectedLedger?.LedgerId}).");
            Task.Run(async () =>
            {
                try
                {
                    await LoadLovRowsAsync();
                }
                catch (Exception ex)
                {
                    LogError(ex, "GLLovViewModel.LoadLovRows (background task)");
                }
            });
        }
        private bool IsLedgerDataExist()
        {
            var intCount = DataRepository.GetTableItemsCount(AppState.Instance.SelectedCube.CubeId, LOV_SelectedLedger.LedgerId, "SEGMENTS");
            ServiceLocator.Logger?.LogDebug($"GLLovViewModel.IsLedgerDataExist: SEGMENTS count for LedgerId={LOV_SelectedLedger.LedgerId} = {intCount}.");
            return intCount != 0;
        }
        public async Task LoadLovRowsAsync()
        {
            var rows = new List<LovRow>();
            bool ledgerDataExist = IsLedgerDataExist();

            ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadLovRowsAsync: started for ledger \"{LOV_SelectedLedger?.LedgerName}\" (LedgerId={LOV_SelectedLedger?.LedgerId}), ledgerDataExist={ledgerDataExist}.");

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
                    ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadLovRowsAsync: calling CommonFunctions.FillResponsibilitiesAsync for LedgerId={LOV_SelectedLedger.LedgerId}, CubeId={AppState.Instance.SelectedCube.CubeId}.");
                    await CommonFunctions.FillResponsibilitiesAsync(LOV_SelectedLedger.LedgerId, AppState.Instance.SelectedCube.CubeId, token);
                    ServiceLocator.Logger?.LogDebug("GLLovViewModel.LoadLovRowsAsync: CommonFunctions.FillResponsibilitiesAsync completed.");
                }

                var repository = new DataRepository();

                // a. All SEGMENTs (and their counts from SEGMENT_VALUES)
                // Dispatched to the UI thread since the Segments setter raises PropertyChanged.
                await _dispatcher.InvokeAsync(() =>
                {
                    Segments = repository.GetSegments(LOV_SelectedLedger.CubeId, LOV_SelectedLedger.LedgerId);
                });
                ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadLovRowsAsync: DataRepository.GetSegments returned {Segments?.Count ?? 0} segment(s).");

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
                ServiceLocator.Logger?.LogDebug($"GLLovViewModel.LoadLovRowsAsync: built {rows.Count} LOV row(s) total.");

                // Fill ViewModel collection on UI thread
                await _dispatcher.InvokeAsync(() =>
                {
                    LOVRows.Clear();
                    foreach (var r in rows)
                    {
                        LOVRows.Add(r);
                    }

                    DataLoadedAction?.Invoke();
                });

                if (!ledgerDataExist && HideBusyAsyncAction != null)
                {
                    var task = HideBusyAsyncAction.Invoke();
                    await task;
                }

                ServiceLocator.Logger?.LogDebug("GLLovViewModel.LoadLovRowsAsync: completed successfully.");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("GLLovViewModel.LoadLovRowsAsync: ledger data fetch operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "GLLovViewModel.LoadLovRowsAsync");
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

        private static void LogError(Exception ex, string context = "GLLovViewModel")
        {
            ServiceLocator.Logger?.LogException(ex, context);
        }
    }
}
