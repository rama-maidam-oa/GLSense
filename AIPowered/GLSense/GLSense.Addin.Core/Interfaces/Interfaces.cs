// Interfaces.cs in GLSense.Addin.Core
// Port of GLSense\Interfaces\Interfaces.cs (FinalWorkingCode) for Group C
// (Segment/Period pickers).
//   - IWarningHost: implemented by all 7 Group C views (GLGetPeriod, GLGetPeriodByYear,
//     GLGetPeriodByDate, GLGetPeriodDetails, GLGetPeriodStartEnd, GLSegmentFunctions,
//     GLDailyRates) so ExcelRefEditControl can walk up the visual tree and surface
//     cell-selection warnings without a direct reference back to the hosting window.
//   - IPeriodDataService / ISegmentDataService: kept internal (assembly-scoped) exactly
//     like the original - Services\DataServiceLocator.cs (this project's re-pointed
//     Service\ServiceLocator.cs) and the 7 ViewModels are all in this same assembly, so
//     internal visibility works the same way it did in the old GLSense project.
using GLSense.Addin.Core.Models;
using System.Collections.ObjectModel;

namespace GLSense.Addin.Core.Interfaces
{
    public interface IWarningHost
    {
        void CellSelectionWarning(string message);
    }

    internal interface IPeriodDataService
    {
        ObservableCollection<PeriodModel> GetPeriodsForLedger(string ledger, bool allowRemoteFetch = true);
    }

    internal interface ISegmentDataService
    {
        ObservableCollection<SegmentModel> GetSegments(string ledgerName);
        ObservableCollection<SegmentValueModel> GetSegmentValues(string ledgerName);
        string GetSegmentNameBySequence(int sequence, string ledgerName);
        string ResolveSegmentName(object segmentName, string ledgerName);
    }
}
