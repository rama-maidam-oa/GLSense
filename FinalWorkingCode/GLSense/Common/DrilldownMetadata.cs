using System.Drawing;

namespace GLSense.Common
{
    public static class DrilldownMetadata
    {
        /// <summary>
        /// Display text from enum Description (fallback to name).
        /// </summary>
        public static string GetDisplay(DrilldownType dd) => dd.GetDescription();

        /// <summary>
        /// Tab color for each drilldown type (System.Drawing.Color).
        /// </summary>
        public static Color GetColor(DrilldownType dd)
        {
            return dd switch
            {
                DrilldownType.BL => Color.FromArgb(192, 255, 192),
                DrilldownType.JL => Color.FromArgb(255, 255, 192),
                DrilldownType.SL => Color.FromArgb(255, 224, 192),
                DrilldownType.BL_JL => Color.FromArgb(255, 255, 192),
                DrilldownType.BL_SL => Color.FromArgb(255, 224, 192),
                DrilldownType.BLDD_SL => Color.FromArgb(255, 224, 192),
                DrilldownType.UF => Color.FromArgb(242, 206, 239),
                DrilldownType.BLDD_UF => Color.FromArgb(242, 206, 239),
                DrilldownType.CM => Color.FromArgb(228, 158, 221),
                _ => Color.LightGray,
            };
        }

        /// <summary>
        /// Excel/Interop-friendly OLE color code.
        /// </summary>
        public static int GetOleColor(DrilldownType dd) =>
            ColorTranslator.ToOle(GetColor(dd));
    }
}
