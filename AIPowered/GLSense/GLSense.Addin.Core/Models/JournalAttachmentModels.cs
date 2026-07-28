// JournalAttachmentModels.cs in GLSense.Addin.Core
// Port of GLSense\Models\AllModels.cs (FinalWorkingCode) - just the three small DTOs used
// by the journal-attachment download flow (JournalAttachments.cs):
// JournalAttachments/JrnalAttachRequest/JournalAttachmentRecord. This project has no
// single AllModels.cs - models are split into per-feature files (see BalanceDtoModel.cs/
// PeriodModels.cs/etc.), so these are grouped into their own file instead of appended to
// an existing one. Verbatim port - no logic, just plain DTOs (property names/casing
// preserved exactly since they round-trip through JsonSerializer against the server API).
namespace GLSense.Addin.Core.Models
{
    /// <summary>Request payload for POST .../journal-attachment-files (list attachments for a journal).</summary>
    public class JournalAttachments
    {
        public long cubeId { get; set; }
        public long journalHeaderId { get; set; }
    }

    /// <summary>Request payload for POST .../journal-attachments (download selected attachment files).</summary>
    public class JrnalAttachRequest
    {
        public long cubeId { get; set; }
        public long[] fileIds { get; set; }
    }

    /// <summary>One row of the journal-attachment-files response.</summary>
    public class JournalAttachmentRecord
    {
        public string FILE_ID { get; set; }
        public string FILE_NAME { get; set; }
    }
}
