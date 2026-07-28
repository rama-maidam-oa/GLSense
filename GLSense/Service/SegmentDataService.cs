using GLSense.Interfaces;
using GLSense.Models;
using GLSense.Repositories;
using GLSense.Utilities;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace GLSense.Service
{
    internal class SegmentDataService : ISegmentDataService
    {
        public ObservableCollection<SegmentModel> GetSegments(string ledgerName)
        {
            LogUtility.LogDebug($"SegmentDataService.GetSegments: ledgerName={ledgerName}");

            if (AppState.Instance.SelectedCube == null)
            {
                LogUtility.LogError("SegmentDataService.GetSegments: No cube selected.");
                throw new InvalidOperationException("No cube selected");
            }

            var cubeId = AppState.Instance.SelectedCube.CubeId;
            var ledgerId = GetDefaultLedgerId(ledgerName) ?? AppState.Instance.SelectedLedger.LedgerId;
            var repo = new DataRepository();

            LogUtility.LogDebug($"SegmentDataService.GetSegments: CubeId={cubeId}, LedgerId={ledgerId}");
            return repo.GetSegments(cubeId, ledgerId);
        }

        public ObservableCollection<SegmentValueModel> GetSegmentValues(string ledgerName)
        {
            LogUtility.LogDebug($"SegmentDataService.GetSegmentValues: ledgerName={ledgerName}");

            if (AppState.Instance.SelectedCube == null)
            {
                LogUtility.LogError("SegmentDataService.GetSegmentValues: No cube selected.");
                throw new InvalidOperationException("No cube selected");
            }

            var cubeId = AppState.Instance.SelectedCube.CubeId;
            var ledgerId = GetDefaultLedgerId(ledgerName) ?? AppState.Instance.SelectedLedger.LedgerId;

            LogUtility.LogDebug($"SegmentDataService.GetSegmentValues: CubeId={cubeId}, LedgerId={ledgerId}");
            return DataRepository.GetAllSegmentValues(cubeId, ledgerId);
        }
        public string GetSegmentNameBySequence(int sequence, string ledgerName)
        {
            LogUtility.LogDebug($"SegmentDataService.GetSegmentNameBySequence: sequence={sequence}, ledgerName={ledgerName}");

            var segments = GetSegments(ledgerName);
            if (segments == null || !segments.Any())
            {
                LogUtility.LogDebug("SegmentDataService.GetSegmentNameBySequence: no segments found, returning null.");
                return null;
            }

            var orderedSegments = segments.OrderBy(s => s.ApplicationColumnName).ToList();
            if (sequence > 0 && sequence <= orderedSegments.Count)
                return orderedSegments[sequence - 1].SegmentName;

            LogUtility.LogWarn($"SegmentDataService.GetSegmentNameBySequence: sequence {sequence} out of range (count={orderedSegments.Count}).");
            return null;
        }

        public string ResolveSegmentName(object segmentName, string ledgerName)
        {
            LogUtility.LogDebug($"SegmentDataService.ResolveSegmentName: segmentName={segmentName}, ledgerName={ledgerName}");

            if (segmentName == null || segmentName is System.Reflection.Missing)
                return null;

            if (IsNumeric(segmentName))
            {
                int sequence = Convert.ToInt32(segmentName);
                return GetSegmentNameBySequence(sequence,ledgerName);
            }

            return segmentName.ToString();
        }

        private static long? GetDefaultLedgerId(string ledgerName)
        {
            var ledgerId = AppState.Instance.SelectedCube.GetLedgerIdByName(ledgerName) ?? AppState.Instance.SelectedLedger.LedgerId;

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
