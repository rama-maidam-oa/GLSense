using GLSense.Addin.Core.Infrastructure;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace GLSense.Addin.Core.Utilities
{
    /// <summary>
    /// Makes a single designated "fill" column in a DataGrid stretch to consume any left-over
    /// width so there is no blank gap after it, while every column (including the fill column
    /// itself, most of the time) stays sized to its actual content instead of a star (*) width.
    ///
    /// Why this exists: a DataGridColumn with Width="*" reports a huge desired width when it is
    /// measured with an unconstrained/infinite available width - which is exactly what happens
    /// on every window that uses SizeToContent="WidthAndHeight" (see DpiAwareWindow-derived windows
    /// across this project). That caused those windows to always grow to their MaxWidth cap
    /// instead of fitting their actual data (e.g. GLServerConfiguration's "Instance Configuration"
    /// grid always opening at its widest allowed size even with only a couple of short rows).
    /// Switching the affected columns to Width="Auto" fixed the over-growth, but on its own
    /// leaves a blank gap to the right of the grid whenever the window ends up wider than the
    /// content needs (MinWidth forcing extra room, or the user dragging the window wider).
    ///
    /// This helper re-adds "fill the remaining space" behavior for one designated column, driven
    /// off actual (always-finite) SizeChanged/Loaded measurements instead of WPF's star-column
    /// layout, so it can never re-trigger the original infinite-measure bug.
    /// </summary>
    public static class DataGridColumnFillHelper
    {
        // Allowance for the vertical scrollbar's track width, only reserved when that
        // scrollbar is actually visible (see IsVerticalScrollBarVisible below) - otherwise
        // this left a permanent blank strip after the fill column even when every row fit
        // without scrolling.
        private const double ScrollBarAllowance = 20d;

        // Floor for the fill column's resolved width - see the "always resolve to a
        // concrete pixel width" comment in Refresh() below for why this replaced the old
        // "leave it at DataGridLength.Auto" fallback.
        private const double MinFillColumnWidth = 80d;

        // Re-entrancy guard, keyed per grid (this helper can be wired to more than one
        // DataGrid at once - e.g. GLSegmentManager's dgLeft AND dgRight). Refresh() sets
        // fillColumn.Width twice (Auto, then a resolved concrete value) and calls
        // grid.UpdateLayout() in between - each of those can itself raise another
        // SizeChanged on the same grid (a column width change can nudge the grid's own
        // rendered size, especially under a SizeToContent-driven window), which would
        // otherwise re-enter Refresh() before the first call finishes and potentially
        // never settle. Guarding against re-entrancy makes this method's worst case "skip
        // one redundant pass" instead of a runaway resize loop - suspected root cause of a
        // "window opens fine, then starts resizing and hangs" report on GLSegmentManager
        // that persisted across three unrelated Grid row/layout rewrites (see CLAUDE.md's
        // GLSegmentManager section, 26.3.4).
        private static readonly HashSet<DataGrid> _refreshing = new();

        /// <summary>
        /// Wires up automatic fill-column behavior for <paramref name="grid"/>. Call once,
        /// typically right after the window's InitializeComponent().
        /// </summary>
        public static void EnableFillColumn(DataGrid grid, DataGridColumn fillColumn)
        {
            if (grid == null || fillColumn == null) return;

            // Both hooks defer the actual Refresh() to a later dispatcher pass
            // (ContextIdle) instead of running it synchronously inside the Loaded/
            // SizeChanged handler itself. Refresh() calls grid.UpdateLayout() (a forced,
            // synchronous re-measure) - running that synchronously from a SizeChanged
            // handler risks it firing WHILE an ancestor is still in the middle of its own
            // layout pass (e.g. DpiAwareWindow.ForceSizeToContentResettle's three UpdateLayout()
            // calls, or the nested DispatcherFrame DpiAwareWindow.PumpDispatcherFrame() pushes
            // right after - which explicitly pumps every operation at Background priority
            // and above). Reported symptom on GLSegmentManager: window opens, the visible
            // gap near the title bar's close button briefly appears (the well-known
            // SizeToContent stale-first-measurement symptom - CLAUDE.md section 1), then
            // DpiAwareWindow's resettle "adjusts the width to close the gap", and immediately
            // hangs/crashes Excel - i.e. the hang coincides exactly with the resettle-and-
            // pump sequence, not with any specific Grid row-structure choice (three
            // unrelated layout rewrites all still hung). Dispatching at ContextIdle -
            // strictly below Background - guarantees Refresh() never runs INSIDE that
            // pump; it simply waits until the dispatcher is genuinely idle, after
            // ForceSizeToContentResettle/PumpDispatcherFrame have both fully returned.
            // Matches the priority GLSegmentManager.xaml.cs's own Window_Loaded already
            // uses for its manual Refresh() calls (DispatcherPriority.ContextIdle) - kept
            // consistent here so every call site defers the same way.
            grid.Loaded += (s, e) =>
                grid.Dispatcher.BeginInvoke(new Action(() => Refresh(grid, fillColumn)),
                    System.Windows.Threading.DispatcherPriority.ContextIdle);
            grid.SizeChanged += (s, e) =>
                grid.Dispatcher.BeginInvoke(new Action(() => Refresh(grid, fillColumn)),
                    System.Windows.Threading.DispatcherPriority.ContextIdle);
        }

        /// <summary>
        /// Recomputes the fill column's width. Safe to call manually after code populates or
        /// replaces a grid's ItemsSource at runtime (e.g. after a data reload), since that can
        /// change every other column's natural (Auto) width without necessarily changing the
        /// grid's own ActualWidth/SizeChanged firing.
        /// </summary>
        public static void Refresh(DataGrid grid, DataGridColumn fillColumn)
        {
            if (grid == null || fillColumn == null) return;
            if (!_refreshing.Add(grid)) return; // already refreshing this grid - see _refreshing's comment above

            // Setting fillColumn.Width = DataGridLength.Auto a few lines below is only ever
            // meant to be a transient, in-method measurement trick - but on a window using
            // SizeToContent="WidthAndHeight" (every DpiAwareWindow-derived window here), WPF's
            // SizeToContent engine measures the window on EVERY layout pass, including the
            // instant this method flips the column to Auto. A DataGridColumn at Auto reports
            // its full natural (unclamped) width - for a long "Description" cell that can be
            // hundreds of pixels wider than the column's eventual resolved width - so for one
            // layout pass the window's desired size balloons, and the window visibly grows to
            // fit it before this method sets the column back to its resolved (smaller) pixel
            // width a few lines later. Reported symptom: switching to a segment with longer
            // Account descriptions made the grid - and then the window - resize, after which
            // the trailing "Is-Summary" column was pushed out of view with a horizontal
            // scrollbar appearing. Freezing SizeToContent to Manual for the duration of this
            // method (mirroring DpiAwareWindow.ForceSizeToContentResettle's own
            // toggle-Manual-then-restore pattern) means the transient Auto width is never
            // measured by the window at all - only the final, already-resolved pixel width is,
            // once SizeToContent is restored at the very end.
            Window window = null;
            SizeToContent originalSizeToContent = SizeToContent.Manual;
            bool restoreSizeToContent = false;

            try
            {
                if (grid.Columns.Count == 0 || grid.ActualWidth <= 0) return;

                window = Window.GetWindow(grid);
                if (window != null && window.SizeToContent != SizeToContent.Manual)
                {
                    originalSizeToContent = window.SizeToContent;
                    window.SizeToContent = SizeToContent.Manual;
                    restoreSizeToContent = true;
                }

                // Captured BEFORE the Auto re-measure below (which unconditionally
                // overwrites fillColumn.Width) so the idempotency check further down is
                // comparing against what the column actually looked like on screen, not
                // against the transient "Auto" this method sets a few lines from now.
                var previousWidth = fillColumn.Width;

                // Let the fill column re-measure to its natural content width first, so
                // "naturalFillWidth" below reflects the actual current data, not a stale
                // fixed pixel value from a previous adjustment.
                fillColumn.Width = DataGridLength.Auto;
                grid.UpdateLayout();

                double othersWidth = 0;
                foreach (var col in grid.Columns)
                {
                    if (ReferenceEquals(col, fillColumn)) continue;
                    if (col.Visibility != Visibility.Visible) continue;
                    othersWidth += col.ActualWidth;
                }

                double rowHeaderWidth =
                    (grid.HeadersVisibility == DataGridHeadersVisibility.All ||
                     grid.HeadersVisibility == DataGridHeadersVisibility.Row)
                        ? grid.RowHeaderActualWidth
                        : 0;

                double naturalFillWidth = fillColumn.ActualWidth;
                double scrollBarAllowance = IsVerticalScrollBarVisible(grid) ? ScrollBarAllowance : 0d;
                double available = grid.ActualWidth - othersWidth - rowHeaderWidth - scrollBarAllowance;

                // Always resolve to a concrete pixel width here - never leave the fill column
                // at DataGridLength.Auto. A virtualizing DataGrid's Auto column keeps
                // re-measuring against whichever rows are CURRENTLY realized, so leaving it
                // Auto (the old "content needs more than available" fallback) meant that
                // scrolling to a row with a longer value re-grew the column - and, under this
                // window's SizeToContent="WidthAndHeight", the whole window - on every scroll,
                // which could squeeze a fixed-width column after it (e.g. GLSegmentManager's
                // "Is-Summary" column, sitting after "Description") out of view. Clamping to a
                // sane floor keeps the column - and the grid's overall layout - stable while
                // scrolling; the grid scrolls horizontally instead if content is genuinely
                // wider than that floor.
                double resolvedWidth = available > naturalFillWidth ? available : Math.Max(available, MinFillColumnWidth);

                // Skip the assignment entirely if it wouldn't meaningfully change anything
                // versus what the column looked like before this call - an idempotent
                // no-op re-assignment is still a DependencyProperty write that can ripple
                // into another layout pass for no benefit. Compares against previousWidth
                // (captured above, before the Auto re-measure), not fillColumn.Width's
                // current value - which is always "Auto" at this point in the method.
                if (previousWidth.IsAuto || Math.Abs(previousWidth.Value - resolvedWidth) > 0.5)
                {
                    fillColumn.Width = new DataGridLength(resolvedWidth);
                }
                else
                {
                    fillColumn.Width = previousWidth;
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "DataGridColumnFillHelper.Refresh");
            }
            finally
            {
                // Restore SizeToContent LAST, after fillColumn has already been set to its
                // final resolved (non-Auto) pixel width above - so the one re-measure this
                // triggers sees only the correctly-clamped layout, never the transient Auto
                // width, and the window settles at its proper size instead of staying
                // pinned at whatever size it happened to be frozen at.
                if (restoreSizeToContent && window != null)
                {
                    window.SizeToContent = originalSizeToContent;
                }
                _refreshing.Remove(grid);
            }
        }

        /// <summary>
        /// Whether the DataGrid's internal vertical ScrollBar is currently rendered (i.e. row
        /// content actually overflows the grid's viewport). Reserving scrollbar width when it
        /// isn't shown left a permanent blank strip to the right of the fill column.
        /// </summary>
        private static bool IsVerticalScrollBarVisible(DependencyObject visual)
        {
            if (visual == null) return false;

            int count = VisualTreeHelper.GetChildrenCount(visual);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(visual, i);

                if (child is ScrollBar scrollBar &&
                    scrollBar.Orientation == Orientation.Vertical &&
                    scrollBar.Visibility == Visibility.Visible)
                {
                    return true;
                }

                if (IsVerticalScrollBarVisible(child))
                    return true;
            }

            return false;
        }
    }
}
