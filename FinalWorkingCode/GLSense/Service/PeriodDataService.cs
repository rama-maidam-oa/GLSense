using AddinExpress.MSO;
using GLSense.Helpers;
using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using System;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.ObjectModel;

namespace GLSense.Service
{
    internal class PeriodDataService : IPeriodDataService
    {
        public ObservableCollection<PeriodModel> GetPeriodsForLedger(string ledger, bool allowRemoteFetch = true)
        {
            LogUtility.LogDebug($"PeriodDataService.GetPeriodsForLedger: ledger={ledger}, allowRemoteFetch={allowRemoteFetch}");
            try
            {
                // Validate inputs
                if (string.IsNullOrEmpty(ledger))
                    throw new ArgumentException("Ledger name cannot be null or empty");

                if (AppState.Instance.SelectedCube == null)
                    throw new InvalidOperationException("No cube selected");

                // Get ledger ID from selected cube
                var ledgerId = AppState.Instance.SelectedCube.GetLedgerIdByName(ledger);
                if (!ledgerId.HasValue)
                    throw new ArgumentException($"Ledger '{ledger}' not found in selected cube");

                // Get cube ID
                var cubeId = AppState.Instance.SelectedCube.CubeId;
                var ledgerIdValue = ledgerId.Value;

                if (allowRemoteFetch)
                {
                    // Ensure ledger setup data is loaded before fetching periods
                    EnsureLedgerDataLoaded(cubeId, ledgerIdValue);
                    EnsurePeriodDataLoaded(cubeId, ledgerIdValue);
                }

                // Call database method
                var repo = new DataRepository();
                var result = repo.GetPeriods(cubeId, ledgerIdValue);
                LogUtility.LogDebug($"PeriodDataService.GetPeriodsForLedger: CubeId={cubeId}, LedgerId={ledgerIdValue}, periodCount={result?.Count ?? 0}");
                return result;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"PeriodDataService.GetPeriodsForLedger: failed for ledger '{ledger}'");
                throw new ArgumentException($"Failed to get periods for ledger '{ledger}': {ex.Message}", ex);
            }
        }
        private static void EnsureLedgerDataLoaded(long cubeId, long ledgerId)
        {
            LogUtility.LogDebug($"PeriodDataService.EnsureLedgerDataLoaded: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
                {
                    LogUtility.LogDebug("PeriodDataService.EnsureLedgerDataLoaded: login not completed or no cube selected, skipping.");
                    return;
                }

                if (DataRepository.GetTableItemsCount(cubeId, ledgerId, "SEGMENTS") > 0)
                {
                    LogUtility.LogDebug("PeriodDataService.EnsureLedgerDataLoaded: SEGMENTS already populated, skipping remote fetch.");
                    return;
                }

                LogUtility.LogDebug($"PeriodDataService.EnsureLedgerDataLoaded: fetching ledger setup data via FillResponsibilitiesAsync for CubeId={cubeId}, LedgerId={ledgerId}");
                using var ctsHelper = new CancellationHelper();
                // OISR-23022: run the fetch on the thread pool. Blocking the calling thread directly deadlocked when that
                // thread was Excel's / a window's: the fetch's continuations were posted back to the blocked thread, so the
                // response was never processed and the 5-minute timeout never fired (drilldown hung on ledger-setup-data).
                var token = ctsHelper.GetToken();
                Task.Run(() => CommonFunctions.FillResponsibilitiesAsync(ledgerId, cubeId, token))
                    .GetAwaiter()
                    .GetResult();
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogError("Ledger data loading was cancelled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to load ledger data for UDFs");
            }
        }

        private static void EnsurePeriodDataLoaded(long cubeId, long ledgerId)
        {
            LogUtility.LogDebug($"PeriodDataService.EnsurePeriodDataLoaded: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
                {
                    LogUtility.LogDebug("PeriodDataService.EnsurePeriodDataLoaded: login not completed or no cube selected, skipping.");
                    return;
                }

                if (DataRepository.GetTableItemsCount(cubeId, ledgerId, "PERIODS") > 0)
                {
                    LogUtility.LogDebug("PeriodDataService.EnsurePeriodDataLoaded: PERIODS already populated, skipping remote fetch.");
                    return;
                }

                LogUtility.LogDebug($"PeriodDataService.EnsurePeriodDataLoaded: fetching period setup data via FillResponsibilitiesAsync for CubeId={cubeId}, LedgerId={ledgerId}");
                using var ctsHelper = new CancellationHelper();
                // OISR-23022: run the fetch on the thread pool. Blocking the calling thread directly deadlocked when that
                // thread was Excel's / a window's: the fetch's continuations were posted back to the blocked thread, so the
                // response was never processed and the 5-minute timeout never fired (drilldown hung on ledger-setup-data).
                var token = ctsHelper.GetToken();
                Task.Run(() => CommonFunctions.FillResponsibilitiesAsync(ledgerId, cubeId, token))
                    .GetAwaiter()
                    .GetResult();
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogError("Period data loading was cancelled by the user.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to load period data for UDFs");
            }
        }
    }
}
