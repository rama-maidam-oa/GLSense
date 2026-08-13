using GLSense.Common;
using GLSense.Helpers;
using GLSense.Utilities;
using GLSense.Views;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GLSense.Drilldowns
{
    public class DDDatatoWorksheet
    {
        private Excel.Application DD_ExcelApp { get; }
        private readonly string DD_Type;
        private readonly string DD_Parts;
        private static readonly string DD_DummyCol = "DUMMYCOL";
        private readonly object DD_DrillValues;
        private CancellationToken Dd_token { get; }
        private string DD_TableObjname;
        private GLWaitWindow DD_win { get; }
        public DDDatatoWorksheet(Excel.Application xlapp, object drillValues, string dDType, string drilldownParts, CancellationToken token = default, GLWaitWindow win = null)
        {
            DD_ExcelApp = xlapp;
            if (win != null)
            {
                DD_win = win;
            }
            else
            {
                DD_win = null;
            }

            Dd_token = token;

            DD_Type = dDType;
            DD_Parts = drilldownParts;
            DD_DrillValues = drillValues;
        }
        public async Task DD_DatetoWorksheet()
        {
            using (new LogUtility.LogScope($"DD_DatetoWorksheet ({DD_Type})"))
            {
            LogUtility.LogDebug($"DDDatatoWorksheet.DD_DatetoWorksheet started. DD_Type={DD_Type}, DD_Parts={DD_Parts}");
            bool snapshotTaken = false;
            bool originalScreenUpdating = true;
            bool originalDisplayAlerts = true;
            bool originalEnableEvents = true;
            bool originalDisplayStatusBar = true;
            Excel.XlCalculation originalCalculation = Excel.XlCalculation.xlCalculationAutomatic;

            try
            {

                if (DD_ExcelApp != null)
                {
                    snapshotTaken = true;
                    originalScreenUpdating = DD_ExcelApp.ScreenUpdating;
                    originalDisplayAlerts = DD_ExcelApp.DisplayAlerts;
                    originalEnableEvents = DD_ExcelApp.EnableEvents;
                    originalDisplayStatusBar = DD_ExcelApp.DisplayStatusBar;
                    originalCalculation = DD_ExcelApp.Calculation;

                    DD_ExcelApp.ScreenUpdating = false;
                    DD_ExcelApp.DisplayAlerts = false;
                    DD_ExcelApp.EnableEvents = false;
                    DD_ExcelApp.DisplayStatusBar = false;
                    DD_ExcelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                }

                Excel.Workbook ActWrkBook = DD_ExcelApp.ActiveWorkbook;

                DD_TableObjname = "ORB_DD_" + DateTime.Now.ToString("ddMMyyyyhhmm");
                Dictionary<string, string> DataTypeDict = new Dictionary<string, string>();
                Dictionary<string, string> FormatDict = new Dictionary<string, string>();
                Dictionary<string, string> SubTotalsDict = new Dictionary<string, string>();
                List<string> DisplayColumnName = new List<string>();
                List<string> ActualColumnName = new List<string>();

                Dd_token.ThrowIfCancellationRequested();

                

                DrillDownQueryData drillsData = TryDeserializeDrillData(DD_DrillValues);

                if (drillsData == null || drillsData.records.Length == 0)
                {
                    HandleNoRecords(DD_Type, drillsData?.msg ?? "No records exists");
                    return;
                }

                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() =>
                          DD_win.SetProcessMessage("Extracting meta data..."));

                var metadataDict = ExtractMetadata(drillsData, DataTypeDict, FormatDict, SubTotalsDict, DisplayColumnName, ActualColumnName);

                Dd_token.ThrowIfCancellationRequested();

                long lastRow = 5 + drillsData.records.Length;
                int lastColumn = DisplayColumnName.Count > 0 ? DisplayColumnName.Count : -1;

                // Check Excel row limitation before attempting to write data.
                try
                {
                    long maxRows = 1048576;

                    if (lastRow > maxRows)
                    {
                        var msg = $"The number of records ({drillsData.records.Length}) requires {lastRow} rows which exceeds Excel's maximum rows ({maxRows}).";

                        HandleNoRecords(DD_Type, msg);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    // If anything goes wrong determining the sheet limit, log and continue (so existing behavior remains)
                    LogUtility.LogException(ex, "Error checking Excel row limitation.");
                }

                if (lastColumn == -1)
                {
                    // Every other early-exit in this method (no records, row-limit
                    // exceeded, array-copy failure, worksheet-prep failure) tells the user
                    // something via HandleNoRecords - this one used to just return with
                    // nothing logged and nothing shown, which is exactly what made the
                    // underlying metadata-empty bug above so hard to notice: the drilldown
                    // looked like it did nothing at all. Kept as a defensive fallback in
                    // case column detection ever comes back empty for some other reason.
                    LogUtility.LogWarn($"DDDatatoWorksheet.DD_DatetoWorksheet: no columns could be determined from metadata or record keys (DD_Type={DD_Type}) - nothing to write.");
                    HandleNoRecords(DD_Type, "Unable to determine columns for this drilldown. Refer Excel logs for more information.");
                    return;
                }

                object[,] sampleobj = PrepareDataArray(drillsData, DisplayColumnName, ActualColumnName, DataTypeDict);

                if (sampleobj == null)
                {
                    HandleNoRecords(DD_Type, "Error in copying data to array. Refer excel logs for more information.");
                    return;
                }

                Excel.Worksheet ws = PrepareWorksheet(DD_Parts, ActWrkBook);

                if (ws == null)
                {
                    HandleNoRecords(DD_Type, "Error in getting/setting worksheet. Refer excel logs for more information.");
                    return;
                }

                ApplyDataFormats(ws, lastRow, DisplayColumnName, ActualColumnName, DataTypeDict, FormatDict, SubTotalsDict);
                PopulateSheet(ws, sampleobj, lastRow, lastColumn, ActualColumnName);
                ApplyFormatting(ws, lastRow, lastColumn, DD_TableObjname, metadataDict, DD_Type);
                
                FinalizeSheet(ws, lastColumn, DD_Type, DD_Parts);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                if (snapshotTaken && DD_ExcelApp != null)
                {
                    try
                    {
                        DD_ExcelApp.ScreenUpdating = originalScreenUpdating;
                        DD_ExcelApp.DisplayAlerts = originalDisplayAlerts;
                        DD_ExcelApp.EnableEvents = originalEnableEvents;
                        DD_ExcelApp.DisplayStatusBar = originalDisplayStatusBar;
                        DD_ExcelApp.Calculation = originalCalculation;
                    }
                    catch (Exception restoreEx)
                    {
                        LogUtility.LogException(restoreEx, "Error restoring Excel settings after drilldown.");
                    }
                }
            }

            await Task.CompletedTask;
            }
        }
        private void FinalizeSheet(Excel.Worksheet ws, int lastColumn, string ddType, object objStr)
        {
            LogUtility.LogDebug("Applying final formats to drilldown worksheet.");

            // Safely split the incoming object into string array (handles nulls)
            var listValues = objStr?.ToString()?.Split('*') ?? Array.Empty<string>();

            Excel.Range fr;

            try
            {
                Dd_token.ThrowIfCancellationRequested();

                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Applying header format..."));
                // Header block formatting (rows 1–3, up to the last column)
                fr = ws.Range[ws.Cells[1, 1], ws.Cells[3, lastColumn]];
                fr.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(21, 96, 130));

                // A1: Title based on DDType
                var a1 = ws.Range["A1"];
                var ddEnum = DrilldownHelpers.ParseOrDefault(ddType, fallback: DrilldownType.BL);

                // Set display name
                a1.Value = DrilldownMetadata.GetDisplay(ddEnum);

                a1.Font.Size = 12;
                a1.Font.Color = ColorTranslator.ToOle(Color.White);

                Dd_token.ThrowIfCancellationRequested();

                // C1: Timestamp
                var c1 = ws.Range["C1"];
                c1.Value = "Downloaded on : " + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
                c1.Font.Size = 10;
                c1.Font.Color = ColorTranslator.ToOle(Color.White);

                // C2: Time Zone
                var c2 = ws.Range["C2"];
                c2.Value = $"Time Zone: {TimeZoneInfo.Local.DisplayName}";
                c2.Font.Size = 10;
                c2.Font.Color = ColorTranslator.ToOle(Color.White);

                // A2: Reference (ListValues[1] if present)
                var a2 = ws.Range["A2"];
                a2.Value = "Reference : " + (listValues.Length > 1 ? listValues[1] : string.Empty);
                a2.Font.Size = 10;
                a2.Font.Color = ColorTranslator.ToOle(Color.White);

                // A3: Additional text (ListValues[4] if present)
                var a3 = ws.Range["A3"];
                a3.Value = (listValues.Length > 4 ? listValues[4] : string.Empty);
                a3.Font.Size = 10;
                a3.Font.Color = ColorTranslator.ToOle(Color.White);

                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Applying excel sheet tab color..."));

                ws.Tab.Color = DrilldownMetadata.GetOleColor(ddEnum);

                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Modifying the ribbon state..."));

                AddinModule.RibbonHelper.ApplyState("ApplySheetActiveState");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exception in finalizing drilldown sheet");
            }
            finally
            {
                LogUtility.LogDebug("Applying final formats completed.");
            }
        }

        private void ApplyFormatting(
                Excel.Worksheet ws,
                long lastRow,
                int lastColumn,
                string tableObjName,
                Dictionary<string, Dictionary<string, object>> metadataDict,
                string ddType)
        {
            LogUtility.LogDebug("Applying formats to drilldown worksheet like customdrilldown, attachments, etc.");

            try
            {
                Dd_token.ThrowIfCancellationRequested();

                if (!HasValidRange(lastRow, lastColumn) || ws == null)
                {
                    return;
                }

                var listObject = CreateListObject(ws, lastRow, lastColumn, tableObjName);

                if (listObject == null)
                {
                    return;
                }

                var customDrilldownExists = ApplyPerColumnFormatting(
                    ws,
                    lastColumn,
                    metadataDict,
                    ddType,
                    tableObjName);

                if (customDrilldownExists)
                {
                    CreateCustomDrilldownXMLPart(ws.Name, metadataDict);
                }

                ApplyTablePostFormatting(listObject);

                HideInternalColumns(listObject);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private static bool HasValidRange(long lastRow, int lastColumn)
        {
            return lastRow >= 1 && lastColumn >= 1;
        }

        private Excel.ListObject CreateListObject(
            Excel.Worksheet ws,
            long lastRow,
            int lastColumn,
            string tableObjName)
        {
            try
            {
                Dd_token.ThrowIfCancellationRequested();

                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Creating excel data table object..."));

                var srcRange = ws.Range[ws.Cells[5, 1], ws.Cells[lastRow, lastColumn]];

                var listObj = ws.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    srcRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);

                listObj.Name = tableObjName;
                listObj.TableStyle = "TableStyleLight9";

                return listObj;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error creating ListObject.");
                return null;
            }
        }

        private bool ApplyPerColumnFormatting(
            Excel.Worksheet ws,
            int lastColumn,
            Dictionary<string, Dictionary<string, object>> metadataDict,
            string ddType,
            string tableObjName)
        {
            var customDrilldownExists = false;

            for (int columnIndex = 1; columnIndex <= lastColumn; columnIndex++)
            {
                var aliasColName = GetCellTextSafe(ws, 5, columnIndex);
                var actualColName = GetCellTextSafe(ws, 4, columnIndex);

                ApplySubledgerLinkStyling(
                    ws,
                    ddType,
                    aliasColName,
                    actualColName,
                    tableObjName);

                ApplyAttachmentHyperlinks(
                    ws,
                    ddType,
                    aliasColName,
                    actualColName,
                    tableObjName);

                customDrilldownExists |= ApplyCustomDrilldownAndFormula(
                    ws,
                    metadataDict,
                    aliasColName,
                    actualColName,
                    tableObjName);
            }

            return customDrilldownExists;
        }

        private static string GetCellTextSafe(Excel.Worksheet ws, int row, int column)
        {
            try
            {
                var rng = (Excel.Range)ws.Cells[row, column];
                var value = rng.Value2;
                return value != null ? value.ToString() : DD_DummyCol;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"GetCellTextSafe: failed reading cell [{row},{column}] (non-fatal, using dummy column): {ex.Message}");
                return DD_DummyCol;
            }
        }

        private void ApplySubledgerLinkStyling(
            Excel.Worksheet ws,
            string ddType,
            string aliasColName,
            string actualColName,
            string tableObjName)
        {
            Dd_token.ThrowIfCancellationRequested();

            if (!IsJournalOrSubledger(ddType))
            {
                return;
            }

            if (!IsSubledgerColumn(aliasColName, actualColName))
            {
                return;
            }

            try
            {
                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Applying subledger cell format like hyperlink..."));

                var dataBody = ws.ListObjects[tableObjName]
                    .ListColumns[aliasColName]?
                    .DataBodyRange;

                if (dataBody == null)
                {
                    return;
                }

                dataBody.Font.Color = ColorTranslator.ToOle(Color.Blue);
                dataBody.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error styling drilldown column '{aliasColName}'.");
            }
        }

        private static bool IsJournalOrSubledger(string ddType)
        {
            return ddType == "JL" || ddType == "BL_JL";
        }

        private static bool IsSubledgerColumn(string aliasColName, string actualColName)
        {
            return StringEquals(aliasColName, "SUB_DRILL_DOWN") ||
                   StringEquals(aliasColName, "SUBLEDGER_DRILL_DOWN") ||
                   StringEquals(actualColName, "SUB_DRILL_DOWN") ||
                   StringEquals(actualColName, "SUBLEDGER_DRILL_DOWN");
        }

        private static bool StringEquals(string value, string compareTo)
        {
            return string.Equals(value, compareTo, StringComparison.OrdinalIgnoreCase);
        }

        private void ApplyAttachmentHyperlinks(
            Excel.Worksheet ws,
            string ddType,
            string aliasColName,
            string actualColName,
            string tableObjName)
        {
            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Applying attachment hyperlinks..."));

            if (!IsJournalOrSubledger(ddType))
            {
                return;
            }

            if (!IsAttachmentColumn(aliasColName, actualColName))
            {
                return;
            }

            var range = GetColumnDataBodyRange(ws, tableObjName, aliasColName, "attachment links");
            if (range == null)
            {
                return;
            }

            foreach (Excel.Range cell in range)
            {
                AddAttachmentHyperlink(cell);
            }
        }

        private static bool IsAttachmentColumn(string aliasColName, string actualColName)
        {
            return StringEquals(aliasColName, "ATTACHMENT") ||
                   StringEquals(actualColName, "ATTACHMENT");
        }

        private static Excel.Range GetColumnDataBodyRange(
            Excel.Worksheet ws,
            string tableObjName,
            string columnName,
            string logContext)
        {
            Excel.Range dummyRange = null;
            try
            {
                var range = ws.ListObjects[tableObjName]
                    .ListColumns[columnName]?
                    .DataBodyRange;

                if (range != null)
                {
                    LogUtility.LogDebug(
                        $"Applying {logContext} for column {columnName} and for range {range.Address}.");
                }

                return range;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                return dummyRange;
            }
        }

        private void AddAttachmentHyperlink(Excel.Range cell)
        {
            var val = cell.Value2 ?? cell.Value;
            var displayText = val?.ToString();

            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Adding attachment hyperlinks..."));

            if (string.IsNullOrWhiteSpace(displayText))
            {
                return;
            }

            try
            {
                var subAddress = cell.get_Address(
                                Type.Missing,
                                Type.Missing,
                                Excel.XlReferenceStyle.xlA1,
                                false,
                                Type.Missing);

                cell.Hyperlinks.Add(
                    Anchor: cell,
                    Address: string.Empty,
                    SubAddress: subAddress,
                    ScreenTip: "Journals Attachments Link",
                    TextToDisplay: displayText);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(
                    ex,
                    $"Error adding attachment hyperlink to '{cell.Worksheet.Name}'!" +
                    $"{cell.Address[Type.Missing, Type.Missing, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]}");
            }
        }

        private bool ApplyCustomDrilldownAndFormula(
            Excel.Worksheet ws,
            Dictionary<string, Dictionary<string, object>> metadataDict,
            string aliasColName,
            string actualColName,
            string tableObjName)
        {
            if (metadataDict == null ||
                metadataDict.Count == 0 ||
                !metadataDict.ContainsKey(aliasColName))
            {
                return false;
            }

            var dictItem = metadataDict[aliasColName];

            var drilldownConfig = GetDictionaryValue(dictItem, "customDrilldownConfig");
            var customFormula = GetDictionaryValue(dictItem, "customFormula");

            ApplyCustomFormula(
                ws,
                dictItem,
                customFormula,
                aliasColName,
                actualColName,
                tableObjName);

            var hasCustomDrilldown = ApplyCustomDrilldownLinks(
                ws,
                drilldownConfig,
                aliasColName,
                actualColName,
                tableObjName);

            return hasCustomDrilldown;
        }

        private static string GetDictionaryValue(
            Dictionary<string, object> dict,
            string key)
        {
            return dict.ContainsKey(key) ? dict[key].ToString() ?? string.Empty : string.Empty;
        }

        private void ApplyCustomFormula(
            Excel.Worksheet ws,
            Dictionary<string, object> dictItem,
            string customFormula,
            string aliasColName,
            string actualColName,
            string tableObjName)
        {

            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Adding/Inserting custom formulas..."));

            if (string.IsNullOrEmpty(customFormula))
            {
                return;
            }

            var range = GetColumnDataBodyRange(ws, tableObjName, aliasColName, "custom drilldown formula");
            if (range == null)
            {
                return;
            }

            EnsureTextDataTypeUsesGeneral(dictItem, range, actualColName, customFormula);

            try
            {
                range.Value2 = "=" + customFormula;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(
                    ex,
                    $"Error setting custom formula for column {actualColName}. Formula: {customFormula}. " +
                    "Either the formula is incorrect or referenced columns are missing.");
            }
        }

        private static void EnsureTextDataTypeUsesGeneral(
            Dictionary<string, object> dictItem,
            Excel.Range range,
            string actualColName,
            string formula)
        {
            try
            {
                if (dictItem.TryGetValue("dataType", out var dt) &&
                    string.Equals(dt.ToString(), "TEXT", StringComparison.OrdinalIgnoreCase))
                {
                    range.NumberFormat = AppConstants.General;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(
                    ex,
                    $"Error setting column format for custom formula column {actualColName}. Formula: {formula}");
            }
        }

        private bool ApplyCustomDrilldownLinks(
            Excel.Worksheet ws,
            string drilldownConfig,
            string aliasColName,
            string actualColName,
            string tableObjName)
        {
            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.SetProcessMessage("Adding custom drilldown links..."));

            if (string.IsNullOrEmpty(drilldownConfig))
            {
                return false;
            }

            var range = GetCustomDrilldownRange(ws, tableObjName, aliasColName, actualColName);
            if (range == null)
            {
                return true; // config exists even if range failed
            }

            foreach (Excel.Range cell in range)
            {
                AddCustomDrilldownHyperlink(cell);
            }

            return true;
        }

        private Excel.Range GetCustomDrilldownRange(
            Excel.Worksheet ws,
            string tableObjName,
            string aliasColName,
            string actualColName)
        {
            Excel.Range dummyRange = null;
            try
            {
                Dd_token.ThrowIfCancellationRequested();

                var range = ws.ListObjects[tableObjName]
                    .ListColumns[aliasColName]?
                    .DataBodyRange;

                if (range != null)
                {
                    LogUtility.LogDebug(
                        $"Applying custom drilldown links for column {aliasColName} and for range {range.Address}.");
                }

                return range;
            }
            catch
            {
                LogUtility.LogError(
                    $"{MethodBase.GetCurrentMethod().Name}|Issue in setting the column range for custom drilldown configuration. " +
                    $"Column Name [ {actualColName} ] and display name [{aliasColName} ]");
                return dummyRange;
            }
        }

        private static void AddCustomDrilldownHyperlink(Excel.Range cell)
        {
            try
            {
                var val = cell.Value2 ?? cell.Value;
                var displayText = val?.ToString();

                if (string.IsNullOrWhiteSpace(displayText))
                {
                    return;
                }

                var subAddress = cell.get_Address(
                                Type.Missing,
                                Type.Missing,
                                Excel.XlReferenceStyle.xlA1,
                                false,
                                Type.Missing);

                cell.Hyperlinks.Add(
                    Anchor: cell,
                    Address: string.Empty,
                    SubAddress: subAddress,
                    ScreenTip: "Custom Drilldown",
                    TextToDisplay: displayText);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(
                    ex,
                    $"Exception applying hyperlink to range '{cell.Worksheet.Name}'!" +
                    $"{cell.Address[Type.Missing, Type.Missing, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]}");
            }
        }

        private void ApplyTablePostFormatting(Excel.ListObject listObject)
        {
            try
            {
                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() =>
                        DD_win.SetProcessMessage("Formatting excel table object..."));
                var range = listObject.Range;
                range.Columns.AutoFit();
                range.Font.Size = 9;
            }
            catch (Exception ex)
            {
                // ignore per original VB behavior - still log for diagnostics
                LogUtility.LogDebug($"ApplyTablePostFormatting: failed to auto-fit/format table range (ignored, matches legacy behavior): {ex.Message}");
            }
        }

        private void HideInternalColumns(Excel.ListObject listObject)
        {
            foreach (Excel.Range colHdr in listObject.HeaderRowRange)
            {
                Dd_token.ThrowIfCancellationRequested();

                var value = colHdr.Value2 ?? colHdr.Value;
                var text = value?.ToString();

                if (IsInternalHeader(text))
                {
                    colHdr.EntireColumn.Hidden = true;
                }
            }
        }

        private static readonly HashSet<string> InternalHeaders = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
            {
                "DRILL_DOWN1", "DRILLDOWN1", "DRILL_DOWN2", "DRILLDOWN2",
                "DRILL_DOWN3", "DRILLDOWN3", "SUBLEDGER_VIEW"
            };

        private static bool IsInternalHeader(string headerText)
        {
            return !string.IsNullOrEmpty(headerText) &&
                   InternalHeaders.Contains(headerText);
        }


        private void CreateCustomDrilldownXMLPart(
            string sheetName,
            Dictionary<string, Dictionary<string, object>> dict)
        {
            if (dict == null || dict.Count == 0)
                return;

            LogUtility.LogDebug("Saving custom drilldown data to excel.");

            var wb = DD_ExcelApp?.ActiveWorkbook;
            RemoveExistingDrilldownParts(wb, sheetName);
            AddDrilldownPart(wb, sheetName, dict);
        }

        private void RemoveExistingDrilldownParts(Excel.Workbook wb, string sheetName)
        {
            if (wb?.CustomXMLParts == null || wb.CustomXMLParts.Count == 0)
                return;

            try
            {
                Dd_token.ThrowIfCancellationRequested();

                var cxps = wb.CustomXMLParts;

                // CustomXMLParts collections are 1-based
                for (int i = cxps.Count; i >= 1; i--)
                {
                    var part = cxps[i];
                    var xml = part?.XML;

                    if (ContainsDrilldownSheet(xml, sheetName))
                    {
                        part.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private static bool ContainsDrilldownSheet(string xml, string sheetName)
        {
            // Require the DRILLDOWNSHEET marker before falling back to the sheet-name
            // substring check. Without this, a plain substring match on sheetName against
            // the *entire* XML string can false-positive on unrelated CustomXMLParts whose
            // content happens to mention the sheet name - e.g. the CubeId-keyed
            // DRILLDOWNMETADATA part added in GLDrilldownCustomization.xaml.cs stores the raw
            // drilldown-metadata API response, which can very plausibly contain the literal
            // text "Balance Drilldown"/"Journal Drilldown"/etc. (the same names used as sheet
            // names here) - that part has no COLUMNNAME/DRILLDOWNSHEET content at all, so this
            // guard prevents RemoveExistingDrilldownParts from deleting it as collateral damage.
            return !string.IsNullOrEmpty(xml) &&
                   xml.IndexOf("DRILLDOWNSHEET", StringComparison.OrdinalIgnoreCase) >= 0 &&
                   xml.IndexOf(sheetName, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void AddDrilldownPart(
            Excel.Workbook wb,
            string sheetName,
            Dictionary<string, Dictionary<string, object>> dict)
        {
            if (wb == null)
                return;

            try
            {
                var xmlDoc = BuildDrilldownDocument(sheetName, dict);
                var xmlString = xmlDoc?.ToString();

                if (string.IsNullOrWhiteSpace(xmlString))
                    return;

                LogUtility.LogDebug("Custom drill down columns JSON details.");
                LogUtility.LogDebug(xmlString);

                wb.CustomXMLParts.Add(xmlString);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private XDocument BuildDrilldownDocument(
            string sheetName,
            Dictionary<string, Dictionary<string, object>> dict)
        {
            var safeSheetName = CommonFunctions.EscapeXml(sheetName);
            var root = new XElement("DRILLDOWNSHEET", safeSheetName);
            var xmlDoc = new XDocument(root);

            foreach (var dictItem in dict.Values.Where(v => v != null))
            {
                Dd_token.ThrowIfCancellationRequested();
                AddColumnElementIfNeeded(root, dictItem);
            }

            return xmlDoc;
        }

        private static void AddColumnElementIfNeeded(
            XElement root,
            Dictionary<string, object> dictItem)
        {
            var config = GetValueOrEmpty(dictItem, "customDrilldownConfig");
            if (string.IsNullOrWhiteSpace(config))
                return;

            var columnName = CommonFunctions.EscapeXml(GetValueOrEmpty(dictItem, "displayName"));

            var columnElement = new XElement(
                "COLUMNNAME",
                new XAttribute("Name", columnName),
                new XCData(config));

            root.Add(columnElement);
        }

        private static string GetValueOrEmpty(
            Dictionary<string, object> dictItem,
            string key)
        {
            return dictItem.TryGetValue(key, out var value)
                ? value.ToString() ?? string.Empty
                : string.Empty;
        }


        private void PopulateSheet(
            Excel.Worksheet ws,
            object[,] sampleobj,
            long lastRow,
            int lastColumn,
            List<string> actualColumnName)
        {
            LogUtility.LogDebug($"Populating drilldown data to worksheet {ws?.Name}");
            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                  DD_win.SetProcessMessage("Populating sheet with the drilldown data..."));
            try
            {
                if (ws == null) throw new ArgumentNullException(nameof(ws));
                if (lastRow < 5 || lastColumn < 1)
                {
                    LogUtility.LogDebug("Invalid lastRow/lastColumn values; skipping population.");
                    return;
                }

                int loopInt = 0;

                // Data population range (starts at row 5, col 1)
                Excel.Range fr = ws.Range[ws.Cells[5, 1], ws.Cells[lastRow, lastColumn]];
                fr.Value2 = sampleobj ?? throw new ArgumentNullException(nameof(sampleobj)); // Use Value2 for COM interop

                LogUtility.LogDebug($"Data population range {fr.Address}");

                if (actualColumnName != null && actualColumnName.Count > 0)
                {
                    loopInt = 1;
                    foreach (string newstr in actualColumnName)
                    {
                        Dd_token.ThrowIfCancellationRequested();
                        Excel.Range rng = (Excel.Range)ws.Cells[4, loopInt];
                        rng.Value2 = newstr;
                        loopInt++;
                    }

                    // If fewer provided headers than columns, fill remainder from first data row
                    if (loopInt <= lastColumn)
                    {
                        for (int c = loopInt; c <= lastColumn; c++)
                        {
                            Dd_token.ThrowIfCancellationRequested();
                            Excel.Range rng = (Excel.Range)ws.Cells[4, c];
                            Excel.Range rng1 = (Excel.Range)ws.Cells[5, c];
                            rng.Value2 = rng1.Value2;
                        }
                    }

                    // Style header row (row 4) font color to white
                    Excel.Range headerRange = ws.Range[ws.Cells[4, 1], ws.Cells[4, lastColumn]];
                    headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error populating drilldown data");
            }
            finally
            {
                LogUtility.LogDebug("Drilldown data populating completed.");
            }
        }

        private void ApplyDataFormats(
            Excel.Worksheet ws,
            long lastRow,
            List<string> displayColumnName,
            List<string> actualColumnName,
            Dictionary<string, string> dataTypeDict,
            Dictionary<string, string> formatDict,
            Dictionary<string,string> subTotalsDict)
        {
            LogUtility.LogDebug("Applying data formats to excel.");

            try
            {
                if (ws == null) throw new ArgumentNullException(nameof(ws));
                if (displayColumnName == null) throw new ArgumentNullException(nameof(displayColumnName));

                dataTypeDict ??= new Dictionary<string, string>();
                formatDict ??= new Dictionary<string, string>();
                subTotalsDict ??= new Dictionary<string, string>();

                int columnIndex = 0;

                for (int i = 0; i < displayColumnName.Count; i++)
                {
                    Dd_token.ThrowIfCancellationRequested();
                    columnIndex++;

                    var displayKey = displayColumnName[i];
                    // actualColumnName is built in lockstep with displayColumnName (both
                    // FillColumnAndTypeInfo and IncludeMissingRecordKeys append to both lists
                    // together for every column), so the same index gives the matching
                    // actual/raw column name for this position.
                    var actualKey = (actualColumnName != null && i < actualColumnName.Count) ? actualColumnName[i] : null;

                    var dataType = GetDictionaryValueOrEmpty(dataTypeDict, displayKey, actualKey);
                    var format = GetDictionaryValueOrEmpty(formatDict, displayKey, actualKey);
                    var subTotalFunction = GetDictionaryValueOrEmpty(subTotalsDict, displayKey, actualKey);

                    var rng = GetColumnRange(ws, lastRow, columnIndex);

                    ApplyColumnFormat(rng, displayKey, dataType, format);

                    if (!string.IsNullOrWhiteSpace(subTotalFunction))
                    {
                        ApplyColumnSubTotals(rng, displayKey, dataType, format, subTotalFunction);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error populating applying data formats");
            }

            LogUtility.LogDebug("Applying data formats to excel completed.");
        }

        // dataTypeDict/formatDict/subTotalsDict are keyed by whatever string FillColumnAndTypeInfo
        // used as "header" (the metadata's displayName) - but the two column-name lists this method
        // receives can be mismatched against that key for a given column (e.g. metadata gives
        // columnName="BEGIN_BALANCE" but displayName="BEGIN BALANCE", so a lookup using only one of
        // the two names can silently miss the dictionary entry and fall back to "General" format
        // with no error). Checking displayKey first, then falling back to actualKey, means either
        // name resolves correctly as long as ONE of them matches the dictionary's key.
        private static string GetDictionaryValueOrEmpty(
            Dictionary<string, string> dict,
            string displayKey,
            string actualKey)
        {
            if (dict.TryGetValue(displayKey, out var value) && !string.IsNullOrEmpty(value))
            {
                return value;
            }

            if (!string.IsNullOrEmpty(actualKey) && dict.TryGetValue(actualKey, out var actualValue) && !string.IsNullOrEmpty(actualValue))
            {
                return actualValue;
            }

            return string.Empty;
        }

        private static Excel.Range GetColumnRange(
            Excel.Worksheet ws,
            long lastRow,
            int columnIndex)
        {
            var endRow = Math.Max(lastRow, 6);
            return ws.Range[ws.Cells[6, columnIndex], ws.Cells[endRow, columnIndex]];
        }

        private void ApplyColumnFormat(
            Excel.Range range,
            string columnKey,
            string dataType,
            string format)
        {
            if (string.IsNullOrEmpty(dataType))
            {
                range.NumberFormat = "General";
                LogUtility.LogDebug($"Applied format \"General\" for range {range.Address}.");
                return;
            }

            LogUtility.LogDebug($"Applying data formats for column {columnKey} and for range {range.Address}.");
            LogUtility.LogDebug($"Data type is {dataType}.");

            var upperType = dataType.ToUpperInvariant();

            Dd_token.ThrowIfCancellationRequested();

            switch (upperType)
            {

                case "TEXT":
                    if (DD_win != null)
                        _ = DD_win.Dispatcher.InvokeAsync(() =>
                            DD_win.SetProcessMessage("Applying text format..."));
                    ApplySpecificFormat(range, format, "@",
                        "Applying default format as there is an exception in reading format information {0}. Default format is \"TEXT\".",
                        "Applying default format as there is no format information. Default format is \"@\".");
                    break;

                case "DECIMAL":
                    if (DD_win != null)
                        _ = DD_win.Dispatcher.InvokeAsync(() =>
                            DD_win.SetProcessMessage("Applying decimal format..."));
                    ApplySpecificFormat(range, format, "#,##0.00_);[Red](#,##0.00)",
                        "Applying default format as there is an exception in reading format information {0}. Default format is \"#,##0.00_);[Red](#,##0.00)\".",
                        "Applying default format as there is no format information. Default format is \"#,##0.00_);[Red](#,##0.00)\".");
                    break;

                case "INTEGER":
                    if (DD_win != null)
                        _ = DD_win.Dispatcher.InvokeAsync(() =>
                            DD_win.SetProcessMessage("Applying integer format..."));
                    ApplySpecificFormat(range, format, "0_);[Red](0)",
                        "Applying default format as there is an exception in reading format information {0}. Default format is \"0_);[Red](0)\".",
                        "Applying default format as there is no format information. Default format is \"0_);[Red](0)\".");
                    break;

                case "DATE":
                    if (DD_win != null)
                        _ = DD_win.Dispatcher.InvokeAsync(() =>
                            DD_win.SetProcessMessage("Applying date format..."));
                    ApplySpecificFormat(range, format, "m/d/yyyy",
                        "Applying default format as there is an exception in reading format information {0}. Default format is \"m/d/yyyy\".",
                        "Applying default format as there is no format information. Default format is \"m/d/yyyy\".");
                    break;

                default:
                    if (DD_win != null)
                        _ = DD_win.Dispatcher.InvokeAsync(() =>
                          DD_win.SetProcessMessage("Applying default (general) format..."));
                    range.NumberFormat = AppConstants.General;
                    LogUtility.LogDebug($"Unknown data type \"{dataType}\". Applied format \"General\".");
                    break;
            }
        }
        private void ApplyColumnSubTotals(
            Excel.Range range,
            string columnKey,
            string dataType,
            string format,
            string funcName)
        {

            LogUtility.LogDebug($"Applying sub totals for column {columnKey} and range {range?.Address}.");

            Dd_token.ThrowIfCancellationRequested();

            if (DD_win != null)
            {
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                DD_win.SetProcessMessage($"Applying sub totals for column \"{columnKey}\""));
            }

            // Determine cell format
            string funcFormat = GetFormatForDataType(dataType, format);

            ApplySubTotalFunction(range, funcName, funcFormat);

        }
        private static string GetFormatForDataType(string dataType, string overrideFormat)
        {
            const string defaultNumericFormat = "#,##0";
            if (string.IsNullOrWhiteSpace(dataType))
                return defaultNumericFormat;

            var upper = dataType.ToUpperInvariant();
            if ((upper == "DECIMAL" || upper == "INTEGER") && !string.IsNullOrWhiteSpace(overrideFormat))
                return overrideFormat;

            return defaultNumericFormat;
        }

        private static void ApplySubTotalFunction(
        Excel.Range range,
        string funcName,
        string funcFormat)
        {
            if (range == null) return;
            if (string.IsNullOrWhiteSpace(funcName)) return;

            try
            {
                int funcInt = GetSubtotalFunctionCode(funcName);
                if (funcInt == 0)
                {
                    return;
                }

                Excel.Range targetCell = range.Offset[1,0];
                LogUtility.LogDebug($"Applying subtotals in {targetCell.Address} with format {funcFormat}");

                targetCell.NumberFormat = funcFormat;
                string formula = $"=SUBTOTAL({funcInt}, {range.Address})";
                LogUtility.LogDebug($"Sub total formula: {formula}");
                targetCell.Value2 = formula;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exception in adding subtotals formula");
            }
        }

        private static int GetSubtotalFunctionCode(string funcName)
        {
            if (string.IsNullOrWhiteSpace(funcName)) return 0;

            // mapping of lowercase function names to SUBTOTAL codes
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                        {
                        { "none", 0 },
                        { "average", 101 },
                        { "count", 102 },
                        { "counta", 103 },
                        { "max", 104 },
                        { "min", 105 },
                        { "product", 106 },
                        { "stdev.s", 107 },
                        { "stdev.p", 108 },
                        { "sum", 109 },
                        { "var.s", 110 },
                        { "var.p", 111 }
                        };

            return map.TryGetValue(funcName.Trim(), out var code) ? code : 0;
        }


        private static void ApplySpecificFormat(
            Excel.Range range,
            string format,
            string defaultFormat,
            string exceptionMessageTemplate,
            string noFormatMessage)
        {

            if (!string.IsNullOrEmpty(format))
            {
                try
                {
                    range.NumberFormat = format;
                    LogUtility.LogDebug($"Applying format {format}");
                }
                catch (Exception ex)
                {
                    range.NumberFormat = defaultFormat;
                    LogUtility.LogWarn(string.Format(exceptionMessageTemplate, format) + $" Exception: {ex.Message}");
                }
            }
            else
            {
                range.NumberFormat = defaultFormat;
                LogUtility.LogDebug(noFormatMessage);
            }
        }

        private Excel.Worksheet PrepareWorksheet(object objStr, Excel.Workbook actWb)
        {
            LogUtility.LogDebug("Getting or creating worksheet for drilldown data.");
            Excel.Worksheet ws;

            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                      DD_win.SetProcessMessage("Creating or getting worksheet (if exists)..."));

            try
            {
                // Normalize input to string and split by '*'
                var listValues = (objStr?.ToString() ?? string.Empty).Split('*');

                // First token is the sheet name
                string shtName = listValues.Length > 0 ? listValues[0].Trim() : "Sheet1";

                shtName = CommonFunctions.SanitizeSheetName(shtName);

                Dd_token.ThrowIfCancellationRequested();

                // Create or reuse worksheet
                ws = FindWorksheetForDrilldown(actWb, shtName);

                if (ws == null)
                {
                    ws = actWb.Worksheets.Add() as Excel.Worksheet;
                    ws.Name = shtName;

                    LogUtility.LogDebug($"Created worksheet {ws.Name} for drilldown data.");
                }
                else
                {
                    LogUtility.LogDebug($"Got worksheet {ws.Name} for drilldown data.");

                    // Ensure visibility (Visible is an enum, not bool)
                    if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    {
                        ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    }

                    // If there is a table starting with "ORB_DD_", capture its name
                    if (ws.ListObjects != null && ws.ListObjects.Count > 0)
                    {
                        // ListObjects collection is 1-based
                        var lo = ws.ListObjects[1];
                        if (lo != null && lo.Name != null && lo.Name.StartsWith("ORB_DD_", StringComparison.Ordinal))
                        {
                            DD_TableObjname = lo.Name;
                        }
                    }

                    ws.Activate();
                }

                ws.Cells.Clear();
                StoreWorksheetMarker(ws, shtName);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in creating/getting excel sheet.");
                ws = null;
            }

            return ws;
        }

        private Excel.Worksheet FindWorksheetForDrilldown(Excel.Workbook actWb, string shtName)
        {
            if (actWb?.Worksheets == null || string.IsNullOrWhiteSpace(shtName))
            {
                return null;
            }

            try
            {
                foreach (Excel.Worksheet worksheet in actWb.Worksheets)
                {
                    if (worksheet == null)
                    {
                        continue;
                    }

                    try
                    {
                        if (string.Equals(worksheet.Name, shtName, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(GetWorksheetMarkerValue(worksheet), shtName, StringComparison.OrdinalIgnoreCase))
                        {
                            return worksheet;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error while searching for existing drilldown worksheet.");
            }

            return null;
        }

        private static string GetWorksheetMarkerValue(Excel.Worksheet worksheet)
        {
            try
            {
                var markerCell = worksheet.Range[AppConstants.DrilldownSheetMarkerCellAddress];
                var value = markerCell?.Value2 ?? markerCell?.Value;
                return value?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"GetWorksheetMarkerValue: failed reading marker cell on sheet '{worksheet?.Name}' (non-fatal): {ex.Message}");
                return string.Empty;
            }
        }

        private static void StoreWorksheetMarker(Excel.Worksheet worksheet, string sheetName)
        {
            try
            {
                var markerCell = worksheet.Range[AppConstants.DrilldownSheetMarkerCellAddress];
                markerCell.Value2 = sheetName;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error storing drilldown worksheet marker.");
            }
        }
        private object[,] PrepareDataArray(
            DrillDownQueryData drillsData,
            List<string> displayColumnName,
            List<string> actualColumnName,
            Dictionary<string, string> dataTypeDict)
        {
            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                DD_win.SetProcessMessage("Preparing data array..."));
            var rowCount = (drillsData?.records?.Length ?? 0) + 1;
            var colCount = displayColumnName.Count;

            var sampleobj = new object[rowCount, colCount];

            try
            {
                var headerArray = FillHeaderRow(sampleobj, displayColumnName);

                Dictionary<string, int> columnIndexMap =
                    BuildColumnIndexMap(actualColumnName, displayColumnName);

                FillDataRows(sampleobj, headerArray, columnIndexMap, drillsData?.records, dataTypeDict);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in preparing data for excel sheet transfer.");
                sampleobj = null;
            }

            return sampleobj;
        }
        private static Dictionary<string, int> BuildColumnIndexMap(
                        List<string> actualColumnNames,
                        List<string> displayColumnNames)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < displayColumnNames.Count; i++)
            {
                // map display name
                map[displayColumnNames[i]] = i;

                // map actual column name
                if (i < actualColumnNames.Count)
                    map[actualColumnNames[i]] = i;
            }

            return map;
        }
        private string[] FillHeaderRow(
            object[,] sampleobj,
            List<string> displayColumnName)
        {
            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
               DD_win.SetProcessMessage("Creating header rows..."));
            var headerArray = new string[displayColumnName.Count];

            for (int i = 0; i < displayColumnName.Count; i++)
            {
                Dd_token.ThrowIfCancellationRequested();

                var key = displayColumnName[i];
                sampleobj[0, i] = key;
                headerArray[i] = key;
            }

            return headerArray;
        }

        private void FillDataRows(
            object[,] sampleobj,
            string[] headerArray,
            Dictionary<string, int> columnIndexMap,
            Dictionary<string, object>[] records,
            Dictionary<string, string> dataTypeDict)
        {
            if (records == null || records.Length == 0)
            {
                return;
            }

            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                   DD_win.SetProcessMessage("Filling data rows..."));

            for (int row = 0; row < records.Length; row++)
            {
                Dd_token.ThrowIfCancellationRequested();

                var record = records[row];
                if (record == null)
                {
                    continue;
                }

                foreach (var key in record.Keys)
                {
                    if (!columnIndexMap.TryGetValue(key, out int colIndex))
                        continue;

                    var rawValue = record[key];
                    var rawText = rawValue?.ToString() ?? string.Empty;
                    sampleobj[row + 1, colIndex] =
                        GetCellValue(key, rawText, dataTypeDict);
                }
            }
        }

        private object GetCellValue(
            string key,
            string rawValue,
            Dictionary<string, string> dataTypeDict)
        {
            if (IsDateColumn(key, dataTypeDict))
            {
                return ConvertDateOrFallback(rawValue);
            }

            return ProtectExcelFormulaLikeText(rawValue);
        }

        private static bool IsDateColumn(
            string key,
            Dictionary<string, string> dataTypeDict)
        {
            if (dataTypeDict == null)
            {
                return false;
            }

            return dataTypeDict.TryGetValue(key, out var dt) &&
                   string.Equals(dt, "DATE", StringComparison.OrdinalIgnoreCase);
        }

        private object ConvertDateOrFallback(string rawValue)
        {
            if (string.IsNullOrEmpty(rawValue))
            {
                return rawValue;
            }

            try
            {
                return JavaDateConv(rawValue);
            }
            catch
            {
                // Conversion failed: keep original
                return rawValue;
            }
        }

        private static object ProtectExcelFormulaLikeText(object rawValue)
        {
            var text = SafeToString(rawValue);
            var equalsCount = CountEqualsSafely(text, rawValue);

            if (equalsCount > 1)
            {
                // Ensure we write as text for Excel
                return "'" + text;
            }

            return rawValue;
        }

        private static string SafeToString(object value)
        {
            return value?.ToString() ?? string.Empty;
        }

        private static int CountEqualsSafely(string text, object rawValue)
        {
            if (string.IsNullOrWhiteSpace(text) || !text.Contains("="))
            {
                return 0;
            }

            try
            {
                return text.Count(c => c == '=');
            }
            catch (Exception ex)
            {
                LogUtility.LogError(
                    $"There are multiple '=' symbols in string value: {rawValue}. Exception: {ex.Message}");
                return 0;
            }
        }
        private string JavaDateConv(string javaLongStr)
        {
            if (string.IsNullOrWhiteSpace(javaLongStr))
                return string.Empty;

            if (long.TryParse(javaLongStr, out var ms))
                return JavaDateConv(ms);

            // If the value isn't purely numeric, you might want to log or return as-is
            LogUtility.LogException(new FormatException($"Invalid epoch ms: {javaLongStr}"));
            return javaLongStr;
        }

        private string JavaDateConv(long? javaLong)
        {
            // Null or zero → empty string (same as VB)
            if (javaLong == null || javaLong == 0)
                return string.Empty;

            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                  DD_win.SetProcessMessage("Converting java date to excel supported date..."));

            try
            {
                // Epoch (UTC) + milliseconds
                var epochUtc = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                var dateTime = epochUtc.AddMilliseconds(javaLong.Value);

                // Return in "yyyy-MM-dd HH:mm:ss"
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                // VB returned javaLong (implicitly converted to String); we’ll do the same explicitly.
                return javaLong?.ToString() ?? string.Empty;
            }
        }

        private Dictionary<string, Dictionary<string, object>> ExtractMetadata(
                        DrillDownQueryData drillsData,
                        Dictionary<string, string> dataTypeDict,
                        Dictionary<string, string> formatDict,
                        Dictionary<string,string> subTotalsDict,
                        List<string> actualColumnName,
                        List<string> displayColumnName)
        {
            LogUtility.LogDebug("Extracting meta data started.");

            try
            {
                if (DD_win != null)
                    _ = DD_win.Dispatcher.InvokeAsync(() =>
                       DD_win.SetProcessMessage("Extracting meta data..."));

                var metadataSource = ResolveMetadataSource(drillsData);

                // Root cause of "data comes back but nothing writes to Excel" (no error,
                // no log line explaining why): this used to return here whenever
                // metadataSource was empty, which skipped IncludeMissingRecordKeys below -
                // the ONLY thing that populates displayColumnName/actualColumnName when the
                // server sends no metadata for a drilldown (a real, valid response shape -
                // "metadata":[] with real "records" alongside it, seen in practice for at
                // least journal drilldowns). With those lists left empty,
                // DD_DatetoWorksheet's lastColumn ends up -1 and it bails out silently.
                // metadataDict itself is allowed to be empty; what must NOT be skipped is
                // deriving columns from the records' own keys.
                var metadataDict = (metadataSource != null && metadataSource.Length > 0)
                    ? BuildMetadataDictionary(metadataSource)
                    : new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);

                if (metadataDict.Count > 0)
                {
                    FillColumnAndTypeInfo(metadataDict, dataTypeDict, formatDict, subTotalsDict, actualColumnName, displayColumnName);
                }

                IncludeMissingRecordKeys(drillsData?.records, displayColumnName, actualColumnName);

                return metadataDict;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in extracting meta data information.");
                return new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
            }
            finally
            {
                LogUtility.LogDebug("Extracting meta data completed.");
            }
        }

        // Decides whether to use the server-provided drillsData.metadata (default) or the
        // metadata previously saved locally via GLDrilldownCustomization's "Save Locally" button
        // (see Common\DrilldownMetadataXmlStore.cs), based on the "Overwrite drilldown metadata
        // with locally saved" user preference. Everything downstream of this point
        // (BuildMetadataDictionary/FillColumnAndTypeInfo/IncludeMissingRecordKeys) is completely
        // unchanged either way, since both sources produce the same
        // Dictionary&lt;string, object&gt;[] per-column shape.
        private Dictionary<string, object>[] ResolveMetadataSource(DrillDownQueryData drillsData)
        {
            //UserConfig.OverwriteDrilldownMetadata = true; // For testing purposes, force this to true. Remove or comment out in production.

            if (UserConfig.OverwriteDrilldownMetadata)
            {
                var localMetadata = TryGetLocalMetadata();
                if (localMetadata != null && localMetadata.Length > 0)
                {
                    LogUtility.LogDebug($"DDDatatoWorksheet.ExtractMetadata: using locally saved drilldown metadata (DD_Type={DD_Type}).");
                    return localMetadata;
                }

                LogUtility.LogWarn($"DDDatatoWorksheet.ExtractMetadata: 'Overwrite drilldown metadata with locally saved' is enabled, but no local drilldown metadata exists for the selected cube/drilldown type (DD_Type={DD_Type}). Falling back to server-provided metadata.");
            }

            return drillsData?.metadata;
        }

        // Reads the CubeId-keyed CustomXMLPart saved via the "Save Locally" button and pulls out
        // just the array for this drilldown's type (DD_Type, e.g. "BL"/"JL"/"SL"/...). Returns
        // null if the preference is off, no cube is selected, no workbook is active, no local
        // metadata was ever saved for this cube, or the saved payload has no entry for this
        // drilldown type (e.g. UNIFIED/CM on a non-fusion cube).
        private Dictionary<string, object>[] TryGetLocalMetadata()
        {
            try
            {
                var cube = AppState.Instance.SelectedCube;
                if (cube == null)
                {
                    LogUtility.LogWarn("DDDatatoWorksheet.TryGetLocalMetadata: no selected cube, cannot read locally saved drilldown metadata.");
                    return null;
                }

                var wb = DD_ExcelApp?.ActiveWorkbook;
                if (wb == null)
                {
                    LogUtility.LogWarn("DDDatatoWorksheet.TryGetLocalMetadata: no active workbook, cannot read locally saved drilldown metadata.");
                    return null;
                }

                if (!DrilldownMetadataXmlStore.TryRead(wb, cube.CubeId, out string rawJson))
                {
                    return null;
                }

                var ddEnum = DrilldownHelpers.ParseOrDefault(DD_Type, fallback: DrilldownType.BL);
                string recordsKey = GetLocalMetadataRecordsKey(ddEnum);

                return DrilldownMetadataXmlStore.ExtractDrilldownTypeMetadata(rawJson, recordsKey);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DDDatatoWorksheet.TryGetLocalMetadata");
                return null;
            }
        }

        // Maps a DrilldownType to the top-level key used under the drilldown-metadata API
        // response's "records" object. Confirmed against a live non-fusion cube's response on
        // 01-Aug-2026: {"msg":"default","records":{"BALANCE":[...],"JOURNAL":[...],
        // "SUBLEDGER":[...]}} - "UNIFIED" wasn't present for that cube, matching the API's own
        // note that a fourth ("Unified") set of records is only returned for fusion-based cubes.
        // This mapping table is the point of contact for adjusting these exact key names once
        // confirmed against a fusion-based cube's response.
        private static string GetLocalMetadataRecordsKey(DrilldownType ddType)
        {
            return ddType switch
            {
                DrilldownType.BL => "BALANCE",
                DrilldownType.JL => "JOURNAL",
                DrilldownType.SL => "SUBLEDGER",
                DrilldownType.BL_JL => "JOURNAL",
                DrilldownType.BL_SL => "SUBLEDGER",
                DrilldownType.BLDD_SL => "SUBLEDGER",
                DrilldownType.BLDD_UF => "UNIFIED",
                DrilldownType.UF => "UNIFIED",
                DrilldownType.CM => "UNIFIED",
                _ => "BALANCE"
            };
        }
        private Dictionary<string, Dictionary<string, object>> BuildMetadataDictionary(IEnumerable<Dictionary<string, object>> metadata)
        {
            var metadataDict = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);

            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                DD_win.SetProcessMessage("Building metadata dictionary..."));

            foreach (var metaItem in metadata)
            {
                if (metaItem == null)
                    continue;

                if (!metaItem.TryGetValue("displayName", out var displayName))
                    continue;

                if (string.IsNullOrEmpty(displayName.ToString()))
                    continue;

                metadataDict[displayName.ToString()] = metaItem;
            }

            return metadataDict;
        }
        private void FillColumnAndTypeInfo(
                Dictionary<string, Dictionary<string, object>> metadataDict,
                Dictionary<string, string> dataTypeDict,
                Dictionary<string, string> formatDict,
                Dictionary<string, string> subTotalsDict,
                List<string> actualColumnName,
                List<string> displayColumnName)
        {
            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                DD_win.SetProcessMessage("Extracting columns and its data types..."));

            foreach (var kvp in metadataDict)
            {
                var header = kvp.Key;
                var metaItem = kvp.Value;

                AddIfNotEmpty(metaItem, "displayName", displayColumnName);
                AddIfNotEmpty(metaItem, "columnName", actualColumnName);
                AddIfNotEmptyToDict(metaItem, "dataType", dataTypeDict, header);
                AddIfNotEmptyToDict(metaItem, "format", formatDict, header);
                AddIfNotEmptyToDict(metaItem, "subtotalFunction", subTotalsDict, actualColumnName[actualColumnName.Count-1]);
            }
        }

        private void IncludeMissingRecordKeys(
            IEnumerable<Dictionary<string, object>> records,
            List<string> displayColumnName,
            List<string> actualColumnName)
        {
            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() =>
                 DD_win.SetProcessMessage("Adding missing records..."));

            if (records == null)
                return;

            foreach (var record in records)
            {
                Dd_token.ThrowIfCancellationRequested();

                if (record == null)
                    continue;

                foreach (var key in record.Keys)
                {
                    Dd_token.ThrowIfCancellationRequested();

                    if (displayColumnName.Contains(key))
                        continue;

                    displayColumnName.Add(key);
                    actualColumnName.Add(key);
                }
            }
        }

        private void AddIfNotEmpty(
            Dictionary<string, object> metaItem,
            string key,
            List<string> targetList)
        {
            if (metaItem.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value.ToString()))
            {
                Dd_token.ThrowIfCancellationRequested();

                targetList.Add(value.ToString());
            }
        }

        private static void AddIfNotEmptyToDict(
            Dictionary<string, object> metaItem,
            string key,
            Dictionary<string, string> targetDict,
            string dictKey)
        {
            if (metaItem.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value.ToString()))
            {
                targetDict[dictKey] = value.ToString();
            }
        }

        private void HandleNoRecords(string DDType, string msg = "No records found.")
        {
            var ddEnum = DrilldownHelpers.ParseOrDefault(DDType, fallback: DrilldownType.BL);


            // Get display name from metadata (reads Description attribute)
            var title = DrilldownMetadata.GetDisplay(ddEnum);

            if (DD_win != null)
                _ = DD_win.Dispatcher.InvokeAsync(() => DD_win.RequestClose());
            CommonFunctions.GLSenseMessage(
                $"{title}{Environment.NewLine}{msg}",
                MessageBoxIcon.Error,
                MessageBoxButtons.OK);
        }

        private DrillDownQueryData TryDeserializeDrillData(object drillvalues)
        {
            string json = string.Empty;
            try
            {
                if (drillvalues == null)
                    return null;

                Dd_token.ThrowIfCancellationRequested();

                // Case 1: Already a string → use directly
                if (drillvalues is string s)
                {
                    json = s;
                }
                // Case 2: JsonNode (from previous System.Text.Json parsing) → serialize to string
                else if (drillvalues is JsonNode node)
                {
                    json = node.ToJsonString(new JsonSerializerOptions { WriteIndented = false });
                }
                // Case 3: JsonElement (e.g., from JsonDocument) → get raw text
                else if (drillvalues is JsonElement element)
                {
                    json = element.GetRawText();
                }
                // Fallback: ToString() — in case it's some other object containing JSON
                else
                {
                    json = drillvalues.ToString() ?? string.Empty;
                }

                if (string.IsNullOrWhiteSpace(json))
                    return null;

                // Deserialize using System.Text.Json
                return JsonSerializer.Deserialize<DrillDownQueryData>(json, JsonGlobals.Options);
            }
            catch (JsonException jsonEx)
            {
                LogUtility.LogException(jsonEx, "Invalid JSON format for drilldown data");
                LogUtility.LogRawJson("DDDatatoWorksheet.DeserializeDrillValues", json);
                return null;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error deserializing drilldown data");
                LogUtility.LogRawJson("DDDatatoWorksheet.DeserializeDrillValues", json);
                return null;
            }
        }

    }

    public class DrillDownQueryData
    {
        public Dictionary<string, object>[] records { get; set; }
        public Dictionary<string, object>[] metadata { get; set; }
        public string status { get; set; }
        public string msg { get; set; }
    }
}
