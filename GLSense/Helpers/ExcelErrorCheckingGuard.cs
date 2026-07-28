using GLSense.Utilities;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Helpers
{
    /// <summary>
    /// Captures Excel ErrorCheckingOptions at construction, lets you apply your
    /// preferred values during the session, and restores the original on Dispose.
    /// </summary>
    public sealed class ExcelErrorCheckingGuard : IDisposable
    {
        private readonly Excel.Application _app;
        private readonly (bool BackgroundChecking,
                          bool EvaluateToError,
                          bool TextDate,
                          bool NumberAsText,
                          bool InconsistentFormula,
                          bool OmittedCells,
                          bool UnlockedFormulaCells,
                          bool ListDataValidation,
                          bool EmptyCellReferences) _original;
        private bool _applied;
        private bool _disposed;

        public ExcelErrorCheckingGuard(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            var o = _app.ErrorCheckingOptions;

            _original = (o.BackgroundChecking,
                         o.EvaluateToError,
                         o.TextDate,
                         o.NumberAsText,
                         o.InconsistentFormula,
                         o.OmittedCells,
                         o.UnlockedFormulaCells,
                         o.ListDataValidation,
                         o.EmptyCellReferences);

            LogUtility.LogDebug("ExcelErrorCheckingGuard: captured original ErrorCheckingOptions.");
        }

        /// <summary>
        /// Apply your runtime settings (idempotent).
        /// </summary>
        public void Apply()
        {
            if (_disposed || _applied) return;

            var o = _app.ErrorCheckingOptions;
            o.BackgroundChecking = false;
            o.EvaluateToError = false;
            o.TextDate = false;
            o.NumberAsText = false;
            o.InconsistentFormula = false;
            o.OmittedCells = false;
            o.UnlockedFormulaCells = false;
            o.ListDataValidation = false;
            o.EmptyCellReferences = false;

            _applied = true;
            LogUtility.LogDebug("ExcelErrorCheckingGuard: applied session ErrorCheckingOptions.");
        }

        /// <summary>
        /// Restore the captured values (safe to call multiple times).
        /// </summary>
        public void Restore()
        {
            if (_disposed) return;

            try
            {
                var o = _app.ErrorCheckingOptions;
                o.BackgroundChecking = _original.BackgroundChecking;
                o.EvaluateToError = _original.EvaluateToError;
                o.TextDate = _original.TextDate;
                o.NumberAsText = _original.NumberAsText;
                o.InconsistentFormula = _original.InconsistentFormula;
                o.OmittedCells = _original.OmittedCells;
                o.UnlockedFormulaCells = _original.UnlockedFormulaCells;
                o.ListDataValidation = _original.ListDataValidation;
                o.EmptyCellReferences = _original.EmptyCellReferences;
                LogUtility.LogDebug("ExcelErrorCheckingGuard: restored original ErrorCheckingOptions.");
            }
            catch
            {
                // Excel/COM can throw during shutdown; ignore to keep shutdown clean.
            }

            _applied = false;
        }

        public void Dispose()
        {
            if (_disposed) return;
            Restore();
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }

}
