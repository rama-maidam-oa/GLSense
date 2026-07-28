using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GLSense.Utilities
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
            LogUtility.LogDebug("BatchCalcScope: entered (StartBatchCalc=true).");
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                AppState.Instance.StartBatchCalc = false;
                LogUtility.LogDebug("BatchCalcScope: exited (StartBatchCalc=false).");
                _disposed = true;
            }
        }
    }

}
