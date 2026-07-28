// PeriodDataService.cs in GLSense.Addin.Core
// Port of GLSense\Service\PeriodDataService.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers). Re-pointed vs. the original:
//   - AppState.Instance.* -> GLSense.Addin.Core.AppState.Instance.* (this project's
//     AppState already exposes SelectedCube/SelectedLedger/IsLoginCompleted).
//   - GLSense.Repositories.DataRepository -> GLSense.Addin.Core.Repositories.DataRepository
//     (GetPeriods/GetTableItemsCount both added to that file in this same pass).
//   - GLSense.Helpers.CancellationHelper -> GLSense.Addin.Core.Helpers.CancellationHelper.
//   - GLSense.Utilities.CommonFunctions.FillResponsibilitiesAsync ->
//     GLSense.Addin.Core.Utilities.CommonFunctions.FillResponsibilitiesAsync (already
//     ported in Group B).
//   - LogUtility.* -> ServiceLocator.Logger?.* (Infrastructure.ServiceLocator).
//   - Dropped an unused "using AddinExpress.MSO;" from the original - nothing in this
//     class actually references ADX types, and GLSense.Addin.Core does not (and must not)
//     reference AddinExpress.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using GLSense.Addin.Core.Utilities;
using System;
using System.Collections.ObjectModel;

namespace GLSense.Addin.Core.Services
{
    internal class PeriodDataService : IPeriodDataService
    {
        public ObservableCollection<PeriodModel> GetPeriodsForLedger(string ledger, bool allowRemoteFetch = true)
        {
            ServiceLocator.Logger?.LogDebug($"PeriodDataService.GetPeriodsForLedger started. Ledger={ledger}, AllowRemoteFetch={allowRemoteFetch}");
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
                else
                {
                    ServiceLocator.Logger?.LogDebug($"PeriodDataService.GetPeriodsForLedger: AllowRemoteFetch=false - skipping remote ledger/period data load for CubeId={cubeId}, LedgerId={ledgerIdValue}.");
                }

                // Call database method
                var repo = new DataRepository();
                var periods = repo.GetPeriods(cubeId, ledgerIdValue);
                ServiceLocator.Logger?.LogDebug($"PeriodDataService.GetPeriodsForLedger: returning {periods?.Count ?? 0} period(s) for Ledger={ledger}, CubeId={cubeId}, LedgerId={ledgerIdValue}");
                return periods;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, $"PeriodDataService.GetPeriodsForLedger failed for Ledger={ledger}");
                throw new ArgumentException($"Failed to get periods for ledger '{ledger}': {ex.Message}", ex);
            }
        }

        private static void EnsureLedgerDataLoaded(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsureLedgerDataLoaded: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
                {
                    ServiceLocator.Logger?.LogDebug("PeriodDataService.EnsureLedgerDataLoaded: login not completed or no cube selected - skipping.");
                    return;
                }

                if (DataRepository.GetTableItemsCount(cubeId, ledgerId, "SEGMENTS") > 0)
                {
                    ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsureLedgerDataLoaded: SEGMENTS already populated for CubeId={cubeId}, LedgerId={ledgerId} - skipping remote fetch.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsureLedgerDataLoaded: no SEGMENTS found - fetching ledger setup data for CubeId={cubeId}, LedgerId={ledgerId}.");
                using (var ctsHelper = new CancellationHelper())
                {
                    CommonFunctions.FillResponsibilitiesAsync(ledgerId, cubeId, ctsHelper.GetToken())
                        .GetAwaiter()
                        .GetResult();
                }
                ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsureLedgerDataLoaded: FillResponsibilitiesAsync completed for CubeId={cubeId}, LedgerId={ledgerId}.");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogError("Ledger data loading was cancelled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to load ledger data for UDFs");
            }
        }

        private static void EnsurePeriodDataLoaded(long cubeId, long ledgerId)
        {
            ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsurePeriodDataLoaded: CubeId={cubeId}, LedgerId={ledgerId}");
            try
            {
                if (!AppState.Instance.IsLoginCompleted || AppState.Instance.SelectedCube == null)
                {
                    ServiceLocator.Logger?.LogDebug("PeriodDataService.EnsurePeriodDataLoaded: login not completed or no cube selected - skipping.");
                    return;
                }

                if (DataRepository.GetTableItemsCount(cubeId, ledgerId, "PERIODS") > 0)
                {
                    ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsurePeriodDataLoaded: PERIODS already populated for CubeId={cubeId}, LedgerId={ledgerId} - skipping remote fetch.");
                    return;
                }

                ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsurePeriodDataLoaded: no PERIODS found - fetching period data for CubeId={cubeId}, LedgerId={ledgerId}.");
                using (var ctsHelper = new CancellationHelper())
                {
                    CommonFunctions.FillResponsibilitiesAsync(ledgerId, cubeId, ctsHelper.GetToken())
                        .GetAwaiter()
                        .GetResult();
                }
                ServiceLocator.Logger?.LogDebug($"PeriodDataService.EnsurePeriodDataLoaded: FillResponsibilitiesAsync completed for CubeId={cubeId}, LedgerId={ledgerId}.");
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogError("Period data loading was cancelled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "Failed to load period data for UDFs");
            }
        }
    }
}
