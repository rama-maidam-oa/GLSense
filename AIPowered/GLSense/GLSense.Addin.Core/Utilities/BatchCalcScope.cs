// BatchCalcScope.cs in GLSense.Addin.Core
// Ported from GLSense\Utilities\BatchCalcScope.cs (FinalWorkingCode) verbatim.
// Group F (transitive dependency of BulkRefreshProcess.RefreshSheetAsync/
// RefreshWorkbookAsync, which wrap their formula-recalculation loop in this scope).
// Re-pointed vs. the original: none needed - AppState.Instance.StartBatchCalc already
// exists in this project's AppState (added alongside CalculatedBalances/SingleRefresh/
// ResetFormulas for this same Group F pass).
using GLSense.Addin.Core.Infrastructure;
using System;

namespace GLSense.Addin.Core.Utilities
{
    /// <summary>
    /// Disposable scope for batch calculation mode.
    /// Ensures AppState.Instance.StartBatchCalc is set to true on entry,
    /// and automatically reset to false on exit.
    /// </summary>
    public sealed class BatchCalcScope : IDisposable
    {
        private bool _disposed;

        public BatchCalcScope()
        {
            AppState.Instance.StartBatchCalc = true;
            ServiceLocator.Logger?.LogDebug("BatchCalcScope: entered - AppState.StartBatchCalc=true.");
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                AppState.Instance.StartBatchCalc = false;
                _disposed = true;
                ServiceLocator.Logger?.LogDebug("BatchCalcScope: exited - AppState.StartBatchCalc=false.");
            }
        }
    }
}
