// CancellationHelper.cs in GLSense.Addin.Core
// Ported verbatim from GLSense\Helpers\CancellationHelper.cs (FinalWorkingCode).
// Dropped the unused "using ControlzEx.Standard;" - nothing in this class needs it,
// and GLSense.Addin.Core does not reference ControlzEx.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Threading;

namespace GLSense.Addin.Core.Helpers
{
#nullable enable
    /// <summary>
    /// Robust, reusable CancellationTokenSource helper for WPF / Excel COM Add-ins.
    /// Features:
    /// - Lazy initialization
    /// - Safe multiple Cancel/Dispose calls
    /// - Throws ObjectDisposedException when trying to get Token after disposal
    /// - Thread-safe
    /// </summary>
    public sealed class CancellationHelper : IDisposable
    {
        private CancellationTokenSource? _cts;
        private readonly object _lock = new();
        private bool _disposed;
        private bool _cancelled;
        private CancellationToken _cancelledToken = new CancellationToken(true);

        /// <summary>
        /// Gets a CancellationToken for the current operation.
        /// Creates the CancellationTokenSource lazily on first call.
        /// Returns a cancelled token if cancellation was already requested.
        /// </summary>
        public CancellationToken GetToken()
        {
            lock (_lock)
            {
                if (_disposed)
                    return _cancelledToken; // Return cancelled token instead of throwing

                // If already cancelled, ensure we return a cancelled token
                if (_cancelled)
                {
                    _cts ??= new CancellationTokenSource();
                    _cts.Cancel();
                    return _cts.Token;
                }

                _cts ??= new CancellationTokenSource();
                return _cts.Token;
            }
        }

        /// <summary>
        /// Returns true if cancellation has been requested.
        /// </summary>
        public bool IsCancellationRequested
        {
            get
            {
                lock (_lock)
                {
                    return _disposed || _cancelled || _cts?.IsCancellationRequested == true;
                }
            }
        }

        public void Cancel()
        {
            lock (_lock)
            {
                if (_disposed || _cancelled) return;

                _cancelled = true;
                _cts ??= new CancellationTokenSource();

                try
                {
                    _cts.Cancel();
                }
                catch (ObjectDisposedException ex)
                {
                    /* Safe to ignore */
                    ServiceLocator.Logger?.LogDebug($"CancellationHelper.Cancel: CTS already disposed (benign race) - {ex.Message}");
                }
            }
        }

        public void CancelAfter(TimeSpan delay)
        {
            lock (_lock)
            {
                ThrowIfDisposed();

                if (_cancelled) return; // Already cancelled

                _cts ??= new CancellationTokenSource();
                try
                {
                    _cts.CancelAfter(delay);
                }
                catch (ObjectDisposedException ex)
                {
                    // If disposed, it's already cancelled anyway
                    ServiceLocator.Logger?.LogDebug($"CancellationHelper.CancelAfter: CTS already disposed (benign race) - {ex.Message}");
                }
            }
        }

        public void Dispose()
        {
            CancellationTokenSource? toDispose = null;

            lock (_lock)
            {
                if (_disposed) return;

                _disposed = true;
                _cancelled = true;
                toDispose = _cts;
                _cts = null;
            }

            try
            {
                toDispose?.Dispose();
            }
            catch (ObjectDisposedException ex)
            {
                // Already disposed, ignore
                ServiceLocator.Logger?.LogDebug($"CancellationHelper.Dispose: CTS already disposed (benign race) - {ex.Message}");
            }
            GC.SuppressFinalize(this);
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(CancellationHelper));
        }

        ~CancellationHelper() => Dispose();
    }
#nullable disable
}
