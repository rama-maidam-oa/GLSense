// SegmentDataService.cs in GLSense.Addin.Core
// Port of GLSense\Service\SegmentDataService.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers). Re-pointed vs. the original: AppState.Instance.* ->
// GLSense.Addin.Core.AppState.Instance.*; GLSense.Repositories.DataRepository ->
// GLSense.Addin.Core.Repositories.DataRepository (GetSegments already existed from
// Group B; GetAllSegmentValues added to that file in this same pass).
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Interfaces;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Repositories;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace GLSense.Addin.Core.Services
{
    internal class SegmentDataService : ISegmentDataService
    {
        public ObservableCollection<SegmentModel> GetSegments(string ledgerName)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetSegments: LedgerName={ledgerName}");

            if (AppState.Instance.SelectedCube == null)
            {
                ServiceLocator.Logger?.LogWarn("SegmentDataService.GetSegments: no cube selected.");
                throw new InvalidOperationException("No cube selected");
            }

            try
            {
                var cubeId = AppState.Instance.SelectedCube.CubeId;
                var ledgerId = GetDefaultLedgerId(ledgerName) ?? AppState.Instance.SelectedLedger.LedgerId;
                var repo = new DataRepository();

                var segments = repo.GetSegments(cubeId, ledgerId);
                ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetSegments: returning {segments?.Count ?? 0} segment(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return segments;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDataService.GetSegments");
                throw;
            }
        }

        public ObservableCollection<SegmentValueModel> GetSegmentValues(string ledgerName)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetSegmentValues: LedgerName={ledgerName}");

            if (AppState.Instance.SelectedCube == null)
            {
                ServiceLocator.Logger?.LogWarn("SegmentDataService.GetSegmentValues: no cube selected.");
                throw new InvalidOperationException("No cube selected");
            }

            try
            {
                var cubeId = AppState.Instance.SelectedCube.CubeId;
                var ledgerId = GetDefaultLedgerId(ledgerName) ?? AppState.Instance.SelectedLedger.LedgerId;

                var values = DataRepository.GetAllSegmentValues(cubeId, ledgerId);
                ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetSegmentValues: returning {values?.Count ?? 0} segment value(s) for CubeId={cubeId}, LedgerId={ledgerId}");
                return values;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDataService.GetSegmentValues");
                throw;
            }
        }

        public string GetSegmentNameBySequence(int sequence, string ledgerName)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetSegmentNameBySequence: Sequence={sequence}, LedgerName={ledgerName}");

            try
            {
                var segments = GetSegments(ledgerName);
                if (segments == null || !segments.Any())
                {
                    ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetSegmentNameBySequence: no segments found for LedgerName={ledgerName}");
                    return null;
                }

                var orderedSegments = segments.OrderBy(s => s.ApplicationColumnName).ToList();
                if (sequence > 0 && sequence <= orderedSegments.Count)
                    return orderedSegments[sequence - 1].SegmentName;

                ServiceLocator.Logger?.LogWarn($"SegmentDataService.GetSegmentNameBySequence: Sequence={sequence} out of range (0-{orderedSegments.Count}) for LedgerName={ledgerName}");
                return null;
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDataService.GetSegmentNameBySequence");
                throw;
            }
        }

        public string ResolveSegmentName(object segmentName, string ledgerName)
        {
            ServiceLocator.Logger?.LogDebug($"SegmentDataService.ResolveSegmentName: SegmentName={segmentName}, LedgerName={ledgerName}");

            if (segmentName == null || segmentName is System.Reflection.Missing)
            {
                ServiceLocator.Logger?.LogDebug("SegmentDataService.ResolveSegmentName: SegmentName is null/Missing - returning null.");
                return null;
            }

            try
            {
                if (IsNumeric(segmentName))
                {
                    int sequence = Convert.ToInt32(segmentName);
                    ServiceLocator.Logger?.LogDebug($"SegmentDataService.ResolveSegmentName: SegmentName is numeric - resolving via Sequence={sequence}.");
                    return GetSegmentNameBySequence(sequence, ledgerName);
                }

                return segmentName.ToString();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "SegmentDataService.ResolveSegmentName");
                throw;
            }
        }

        private static long? GetDefaultLedgerId(string ledgerName)
        {
            var ledgerIdByName = AppState.Instance.SelectedCube.GetLedgerIdByName(ledgerName);
            if (ledgerIdByName.HasValue)
            {
                ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetDefaultLedgerId: resolved LedgerId={ledgerIdByName.Value} by LedgerName='{ledgerName}'.");
                return ledgerIdByName;
            }

            ServiceLocator.Logger?.LogDebug($"SegmentDataService.GetDefaultLedgerId: LedgerName='{ledgerName}' not found in selected cube - falling back to AppState.Instance.SelectedLedger.");
            var ledgerId = AppState.Instance.SelectedLedger.LedgerId;

            return ledgerId;
        }

        private static bool IsNumeric(object value)
        {
            if (value == null) return false;
            return value is sbyte || value is byte || value is short || value is ushort ||
                   value is int || value is uint || value is long || value is ulong ||
                   value is float || value is double || value is decimal;
        }
    }
}
