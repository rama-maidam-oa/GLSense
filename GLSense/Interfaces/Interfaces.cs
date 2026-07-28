using GLSense.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace GLSense.Interfaces
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
