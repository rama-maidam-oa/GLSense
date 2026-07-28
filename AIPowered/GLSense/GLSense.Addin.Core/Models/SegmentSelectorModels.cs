// SegmentSelectorModels.cs in GLSense.Addin.Core
// Port of GLSense\Models\AllModels.cs (FinalWorkingCode) - the small POCOs shared by
// GLLovViewModel/SimpleSegmentViewModel/SegmentSelectorViewModel (Group H - Balance
// Configurator pane + LOVs/Roller/Account dialogs port). None of the old classes derived
// from GLSense.Base.NotifyBase, so they are copied verbatim (no re-basing needed) -
// consistent with the convention already established in Models\PeriodModels.cs's header
// comment (implement INotifyPropertyChanged directly where notification is needed; these
// particular models need none).
namespace GLSense.Addin.Core.Models
{
    /// <summary>
    /// One entry in the "right" (selected) grid shared by GLRollerGroups/GLSegmentValues/
    /// (and, out of scope here, GLSegmentRef) - a single value, or a Value1/Value2 range,
    /// tagged with which segment it belongs to.
    /// </summary>
    public class SegmentSelectionModel
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }
        public string Segment { get; set; }
    }

    /// <summary>
    /// Fired by SimpleSegmentViewModel/SegmentSelectorViewModel after paging/filtering
    /// changes so the hosting window's code-behind can scroll both DataGrids back to the
    /// top (GLRollerGroups/GLSegmentValues both subscribe to ScrollToTopRequested).
    /// </summary>
    public class ScrollToTopMessage
    {
        public bool ScrollLeft { get; set; } = true;
        public bool ScrollRight { get; set; } = true;
        public string Trigger { get; set; } = "DataLoaded";
    }

    /// <summary>
    /// One row in GLLOVs' left-hand "Available LOVs" DataGrid - a named list of values
    /// (segment, DB table, or hardcoded) with its item count and category.
    /// </summary>
    public class LovRow
    {
        public string Name { get; set; }          // E.g. "Account", "Activity", ...
        public int ItemsCount { get; set; }       // Number of choices
        public string Category { get; set; }      // "Segment", "Database", "Hardcoded"
    }

    /// <summary>
    /// Marker interface for GLRollerGroups' left DataGrid row model - either a
    /// <see cref="TitleRow"/> (a roller-group header) or a <see cref="SegmentDataRow"/> (a
    /// leaf segment value under the currently-open group).
    /// </summary>
    public interface ISegmentRow
    {
    }

    public class TitleRow : ISegmentRow
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string SummaryFlag { get; set; } = "RG"; // Indicates a title/group-header row
    }

    public class SegmentDataRow : ISegmentRow
    {
        public string SegmentName { get; set; }
        public string SegmentValue { get; set; }
        public string Description { get; set; }
        public string SummaryFlag { get; set; }
        public long SegmentValueSetId { get; set; }
    }
}
