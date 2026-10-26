// RibbonComboPrimer.cs in GLSense.Addin.Core\Helpers
//
// Workaround for an Excel regression (first seen on 16.0.20527.20088, Current Channel,
// Oct-2026): in a new Excel session a ribbon comboBox ignores text set from code until
// its dropdown has been opened once - by the user, or programmatically. Excel does
// request getText and receives the right value; it just doesn't draw it. Opening and
// closing the dropdown once through UI Automation has the same effect as the user doing
// it, after which text set from code is displayed for the rest of the session.
// Full write-up: docs\office-ribbon-combobox-display-bug.md. Remove once Microsoft
// ships a fix.
//
// Constraints:
//   - UI Automation runs on a worker thread: expanding the control needs Excel's main
//     thread to respond, so calling it from that thread would deadlock.
//   - UIA only sees ribbon controls on the currently selected tab, so this is called
//     right after the Cube dialog closes, when the user is on the GLSense tab.
//   - Each Excel window has its own ribbon, so priming is tracked per window handle.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Threading;

namespace GLSense.Addin.Core.Helpers
{
    internal static class RibbonComboPrimer
    {
        private static readonly HashSet<IntPtr> _primedWindows = new HashSet<IntPtr>();
        private static readonly object _lock = new object();

        /// <summary>
        /// Opens and closes the ribbon comboBox named <paramref name="accessibleName"/>
        /// (its caption without the trailing colon, e.g. "Ledger") once for the Excel
        /// window <paramref name="excelHwnd"/>, then runs <paramref name="then"/> on
        /// <paramref name="dispatcher"/>'s thread. If this window was already primed,
        /// <paramref name="then"/> is not run - there is nothing to re-apply. Never throws.
        /// </summary>
        public static void PrimeOnce(IntPtr excelHwnd, string accessibleName, Dispatcher dispatcher, Action then)
        {
            if (excelHwnd == IntPtr.Zero)
                return;

            lock (_lock)
            {
                if (!_primedWindows.Add(excelHwnd))
                    return;
            }

            var worker = new Thread(() =>
            {
                bool primed = false;
                try
                {
                    var excel = AutomationElement.FromHandle(excelHwnd);
                    var condition = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox),
                        new PropertyCondition(AutomationElement.NameProperty, accessibleName));
                    var combo = excel.FindFirst(TreeScope.Descendants, condition);

                    if (combo == null)
                    {
                        ServiceLocator.Logger?.LogWarn($"RibbonComboPrimer: ribbon combo '{accessibleName}' not found (is its tab showing?) - not primed.");
                    }
                    else
                    {
                        var expandCollapse = (ExpandCollapsePattern)combo.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                        expandCollapse.Expand();
                        Thread.Sleep(150);
                        expandCollapse.Collapse();
                        primed = true;
                        ServiceLocator.Logger?.LogDebug($"RibbonComboPrimer: primed ribbon combo '{accessibleName}' for Excel window 0x{excelHwnd.ToInt64():X}.");
                    }
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogWarn($"RibbonComboPrimer: priming '{accessibleName}' failed - {ex.GetType().Name}: {ex.Message}");
                }

                if (!primed)
                {
                    // Let the next attempt in this window try again.
                    lock (_lock) { _primedWindows.Remove(excelHwnd); }
                    return;
                }

                if (then != null && dispatcher != null && !dispatcher.HasShutdownStarted)
                    dispatcher.BeginInvoke(then);
            })
            {
                IsBackground = true,
                Name = "GLSense RibbonComboPrimer"
            };
            worker.SetApartmentState(ApartmentState.MTA);
            worker.Start();
        }
    }
}
