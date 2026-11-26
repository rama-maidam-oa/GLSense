// ProgressWindowHost.cs in GLSense.Addin.Core
//
// Shows GLWaitWindow progress windows on one dedicated, long-lived STA UI thread instead of the main
// WPF dispatcher. The main WPF dispatcher IS Excel's own thread (WpfAppManager.EnsureApplication runs
// there), so a progress window on it froze - and opened late, behind its loading placeholder - whenever
// the operation it reports on kept Excel busy (refresh, reset, drilldown writes, row hide/unhide...).
// On this thread the window paints, updates its step text and elapsed time, and closes on time no
// matter what Excel is doing. See CLAUDE.md 112/114.
//
// Callers must reach the window only through win.Dispatcher (InvokeAsync/Invoke) - every existing
// caller already does. The thread stays alive after a window closes, so a late
// win.Dispatcher.InvokeAsync(...) (e.g. RequestClose in a finally after Cancel) still runs instead of
// being aborted; it is shut down when the add-in's AppDomain unloads (hot reload / Excel exit).
// Known limit: the window stays owned by Excel's window, so a click on Cancel is handled only when
// Excel's thread pumps messages again (painting and timers are unaffected).

#nullable enable

using System;
using System.Threading;
using System.Windows.Threading;
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Views;

namespace GLSense.Addin.Core.Utilities
{
    internal static class ProgressWindowHost
    {
        private static readonly object _lock = new object();
        private static Dispatcher? _dispatcher;
        private static bool _unloadHooked;
        // Windows whose operation may still talk to them. The thread is stopped only after this reaches 0 and
        // stays 0 for IdleShutdownDelay, so a Reload/Excel exit normally finds no progress thread at all
        // (a thread blocked in Dispatcher.Run can stall AppDomain.Unload).
        private static int _leases;
        private static DispatcherTimer? _idleTimer; // lives on the progress thread
        private static readonly TimeSpan IdleShutdownDelay = TimeSpan.FromSeconds(10);

        /// <summary>
        /// Creates a GLWaitWindow on the progress UI thread, sets the optional title/message, and shows it
        /// (without activating it, so it never waits on a busy Excel). Returns null if it could not be
        /// created - callers already treat a null window as "no progress window".
        /// The window holds the thread until its operation's <paramref name="cts"/> is disposed (every caller
        /// disposes it when the operation ends - after any late win.Dispatcher call, e.g. a RequestClose after
        /// Cancel), or, without a cts, until the window closes.
        /// </summary>
        internal static GLWaitWindow? Show(CancellationHelper? cts, string? title = null, string? message = null)
        {
            bool leased = false;
            try
            {
                Dispatcher dispatcher;
                int generation;
                lock (_lock)
                {
                    dispatcher = GetDispatcher();
                    _leases++;
                    generation = _generation;
                    leased = true;
                }
                // The active workbook's window, read here on the caller's (Excel's) thread: the window is owned by it so
                // it moves with that workbook in Alt+Tab / z-order (CLAUDE.md 124).
                IntPtr excelWindow = IntPtr.Zero;
                try { excelWindow = ServiceLocator.ExcelHandle; } catch (Exception ex) { ServiceLocator.Logger?.LogDebug($"ProgressWindowHost: Excel window handle unavailable: {ex.Message}"); }
                GLWaitWindow win = dispatcher.Invoke(() =>
                {
                    _idleTimer?.Stop();
                    var w = new GLWaitWindow(cts) { ShowActivated = false, ExcelOwnerHandle = excelWindow };
                    if (title != null) w.SetProcessTitle(title);
                    if (message != null) w.SetProcessMessage(message);
                    return w;
                });

                _current = win;
                win.Closed += (_, __) => Interlocked.CompareExchange(ref _current, null, win);

                int released = 0;
                void ReleaseOnce() { if (Interlocked.Exchange(ref released, 1) == 0) ReleaseLease(generation); }
                if (cts != null) cts.OnDisposed(ReleaseOnce);
                else win.Closed += (_, __) => ReleaseOnce();
                leased = false; // owned by the callback now

                // Show asynchronously: the calling (Excel) thread doesn't wait for the first paint, and the
                // caller's own win.Dispatcher.InvokeAsync updates queue up behind it in order.
                _ = dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        win.Show();
                        win.StartMonitoring();
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "ProgressWindowHost.Show: showing the progress window failed");
                    }
                });
                return win;
            }
            catch (Exception ex)
            {
                if (leased) ReleaseLease(_generation);
                ServiceLocator.Logger?.LogException(ex, "ProgressWindowHost.Show: creating the progress window failed");
                return null;
            }
        }

        // The progress window shown last and not yet closed (OISR-23022: ApiHelper reports "Waiting for the server" in it).
        private static GLWaitWindow? _current;

        /// <summary>
        /// Shows <paramref name="message"/> as the step text of the open progress window and on the step line of the
        /// busy overlay that is up (Login, Cube details, Process window...), where there is one. Every server request
        /// reports through here ("Waiting for the server: ...", "Receiving ...", "Response received ..."), so the user
        /// sees the time is the server's, not the caller's last step (OISR-23022: a drilldown showed "Creating balance
        /// dto..." for 13 minutes while ledger-setup-data was pending).
        /// </summary>
        internal static void ReportStep(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;
            var win = _current;
            try { if (win != null) _ = win.Dispatcher.InvokeAsync(() => win.SetProcessMessage(message)); }
            catch (Exception ex) { ServiceLocator.Logger?.LogDebug($"ProgressWindowHost.ReportStep: {ex.Message}"); }
            try { AppOverlay.ReportBusyStep(message); }
            catch (Exception ex) { ServiceLocator.Logger?.LogDebug($"ProgressWindowHost.ReportStep (overlay): {ex.Message}"); }
        }

        // Bumped by Shutdown(): a lease taken before it must not be released against the next thread's count.
        private static int _generation;

        private static void ReleaseLease(int generation)
        {
            Dispatcher? dispatcher;
            lock (_lock)
            {
                if (generation != _generation) return; // that thread was already shut down
                if (--_leases > 0) return;
                dispatcher = _dispatcher;
            }
            // Start (or restart) the idle countdown on the progress thread itself.
            _ = dispatcher?.InvokeAsync(() =>
            {
                if (_idleTimer == null)
                {
                    _idleTimer = new DispatcherTimer(DispatcherPriority.Background) { Interval = IdleShutdownDelay };
                    _idleTimer.Tick += (_, __) =>
                    {
                        _idleTimer!.Stop();
                        lock (_lock)
                        {
                            if (_leases > 0) return; // a new window arrived meanwhile
                            _dispatcher = null;      // the next Show starts a fresh thread
                        }
                        _idleTimer = null;
                        ServiceLocator.Logger?.LogDebug("ProgressWindowHost: progress UI thread idle - stopping it.");
                        Dispatcher.CurrentDispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
                    };
                }
                _idleTimer.Stop();
                _idleTimer.Start();
            });
        }

        private static Dispatcher GetDispatcher()
        {
            lock (_lock)
            {
                if (_dispatcher != null && !_dispatcher.HasShutdownStarted)
                    return _dispatcher;

                Dispatcher? created = null;
                Exception? startError = null;
                using (var ready = new ManualResetEventSlim(false))
                {
                    var thread = new Thread(() =>
                    {
                        try
                        {
                            DpiAwarenessHelper.SetPerMonitorAware(); // same per-thread DPI context the main WPF thread uses
                            created = Dispatcher.CurrentDispatcher;
                            SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext(created));
                        }
                        catch (Exception ex)
                        {
                            startError = ex;
                        }
                        finally
                        {
                            ready.Set();
                        }

                        if (startError != null) return;
                        try
                        {
                            Dispatcher.Run();
                        }
                        catch (Exception ex)
                        {
                            ServiceLocator.Logger?.LogException(ex, "ProgressWindowHost: progress UI thread stopped with an error");
                        }
                    })
                    {
                        IsBackground = true,
                        Name = "GLSense progress UI"
                    };
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    ready.Wait();
                    _thread = thread;
                }

                if (created == null)
                    throw new InvalidOperationException("Progress UI thread failed to start.", startError);

                _dispatcher = created;
                if (!_unloadHooked)
                {
                    // Safety net only - AddinEntry.Shutdown calls Shutdown() explicitly before the unload.
                    AppDomain.CurrentDomain.DomainUnload += (_, __) => Shutdown();
                    _unloadHooked = true;
                }
                ServiceLocator.Logger?.LogDebug("ProgressWindowHost: progress UI thread started.");
                return _dispatcher;
            }
        }

        private static Thread? _thread;

        /// <summary>
        /// Stops the progress UI thread now (add-in shutdown / Reload). A thread blocked in Dispatcher.Run can't be
        /// aborted by AppDomain.Unload, which then times out; this waits up to 2 s for the thread to end.
        /// </summary>
        internal static void Shutdown()
        {
            Dispatcher? dispatcher;
            Thread? thread;
            lock (_lock)
            {
                dispatcher = _dispatcher;
                thread = _thread;
                _dispatcher = null;
                _thread = null;
                _leases = 0;
                _generation++;
            }
            if (dispatcher == null) return;
            try
            {
                if (!dispatcher.HasShutdownStarted)
                    dispatcher.BeginInvokeShutdown(DispatcherPriority.Send);
                bool ended = thread == null || thread.Join(2000);
                ServiceLocator.Logger?.LogDebug($"ProgressWindowHost.Shutdown: progress UI thread {(ended ? "stopped" : "still running after 2 s")}.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ProgressWindowHost.Shutdown");
            }
        }
    }
}
