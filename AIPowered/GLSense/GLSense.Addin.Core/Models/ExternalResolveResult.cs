// ExternalResolveResult.cs in GLSense.Addin.Core
// Ported from GLSense\Models\AllModels.cs (FinalWorkingCode) - split out into its own
// file here since Addin.Core doesn't have an AllModels.cs equivalent yet.
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Addin.Core.Models
{
    public sealed class ExternalResolveResult
    {
        public Excel.Workbook Workbook { get; set; }
        public Excel.Worksheet Worksheet { get; set; }
        public Excel.Range Range { get; set; }
    }
}
