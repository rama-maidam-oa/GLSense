// JournalAttachments.cs in GLSense.Addin.Core
// Ported from GLSense\AddinModule.cs (FinalWorkingCode) - the journal-attachment download
// flow: HandleJournalAttachmentHyperlink/ProcessAttachmentResponse/PopulateJournalDictionary/
// ShowAttachmentsDialog/DownloadSelectedAttachments/GLSense_DownloadFile. Deliberately
// deferred out of the earlier Group E "Drilldowns" pass along with the custom-drilldown
// flow (see CustomDrilldown.cs) - both are triggered by SheetFollowHyperlink, whose
// host-side classification/dispatch is a separate, later pass (see AddinEntry.cs
// OnExcelEvent's "SheetFollowHyperlink" case).
//
// PUBLIC ENTRY POINT (host-side wiring depends on this exact signature):
//   public static async Task RunJournalAttachmentFlow(string journalHeaderIdText)
//
// Preconditions the later pass's AddinEntry.OnExcelEvent("SheetFollowHyperlink", ...)
// handler is expected to have already checked, exactly like the old monolith's
// adxExcelAppEvents1_SheetFollowHyperlink did before calling HandleJournalAttachmentHyperlink
// (see AddinModule.cs's IsValidDrilldownSheet, and CustomDrilldown.cs's header comment for
// the sibling IsCustomDrilldownHyperlink check that routes the OTHER kind of hyperlink to
// CustomDrilldown.RunCustomDrilldown instead):
//   - IsValidDrilldownSheet(sht): sht.ListObjects.Count > 0 &&
//     sht.ListObjects[1].Name.StartsWith("ORB_DD_").
//   - !IsCustomDrilldownHyperlink(hyperlink) (no "CUSTOM DRILLDOWN" ScreenTip) and
//     hyperlink.Parent is a Range - i.e. this is a plain ATTACHMENT-column hyperlink (see
//     DDDatatoWorksheet.ApplyAttachmentHyperlinks for how that hyperlink was created).
// journalHeaderIdText is the clicked cell's raw value (old code read
// hyperlinkRange.Value2 directly and truncated it to a long) - the host is expected to
// pass hyperlink.Parent's (a Range) .Value2, stringified, across the AppDomain boundary
// (only strings/primitives cross - see IGLSenseAddin.OnRibbonAction's existing
// convention, which this mirrors).
//
// Re-pointed vs. the original (business logic/URLs/JSON shapes unchanged):
//   - namespace GLSense -> GLSense.Addin.Core; GLSense.Helpers/.Models/.Utilities/.Views ->
//     GLSense.Addin.Core.* equivalents.
//   - AppState.Instance.ExcelApp -> ServiceLocator.ExcelApp.
//   - LogUtility.* (static) -> ServiceLocator.Logger?.* (instance via context).
//   - long journalHeaderId = (long)Math.Truncate(Convert.ToDouble(hyperlinkRange.Value2, ...))
//     -> parsed from the incoming journalHeaderIdText string instead of a live Range.Value2.
//   - SafeInvokeWpf(...) (old AddinModule.cs private helper) -> WpfAppManager.
//     InvokeOnWpfThread(...) directly, using the exact win.CenterInExcel/ModalToExcel/
//     ShowInTaskbar/ShowDialog() idiom AddinEntry.cs's ShowGroupCWindow already established
//     for every other BaseWindow-derived modal dialog in this project - no separate
//     ShowWithOwner/ShowDialogWithOwner call needed (BaseWindow sets the Excel owner
//     automatically via ServiceLocator.ExcelHandle).
//   - CommonMethods.DisableExcelSettings()/EnableExcelSettings() and the CancellationHelper
//     scope, previously owned by the outer adxExcelAppEvents1_SheetFollowHyperlink handler
//     in the old monolith (which wrapped both the custom-drilldown AND journal-attachment
//     branches together), are now owned directly by RunJournalAttachmentFlow - see
//     CustomDrilldown.cs's header comment for the same split rationale.
//   - StrictCertificateValidator.Validate, ApiHelper.ServerAPI, ApiResponseHelper.Parse<T>,
//     JsonGlobals.Options, AppConstants.RestSecure, JournalAttachments/JrnalAttachRequest/
//     JournalAttachmentRecord (Models\JournalAttachmentModels.cs), Views.AttachmentsDialog:
//     all already ported/added alongside this file - used as-is, no changes needed.
using GLSense.Addin.Core.Helpers;
using GLSense.Addin.Core.Infrastructure;
using GLSense.Addin.Core.Models;
using GLSense.Addin.Core.Utilities;
using GLSense.Addin.Core.Views;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
// Alias to disambiguate from this file's own static class (both named "JournalAttachments" -
// the DTO in Models\JournalAttachmentModels.cs, and this flow class), since a plain
// unqualified reference to "JournalAttachments" from inside this class would otherwise bind
// to the enclosing static class itself (which cannot be constructed).
using JournalAttachmentsRequest = GLSense.Addin.Core.Models.JournalAttachments;

namespace GLSense.Addin.Core.Drilldowns
{
    public static class JournalAttachments
    {
        /// <summary>
        /// Entry point called (by a later pass) from AddinEntry.OnExcelEvent's
        /// SheetFollowHyperlink handler once the host has verified IsValidDrilldownSheet
        /// and that the clicked hyperlink is a plain journal-attachment link (not a custom
        /// drilldown link - see CustomDrilldown.RunCustomDrilldown for that branch). See
        /// file header for the exact preconditions this method assumes have already been
        /// checked, and for the exact meaning of journalHeaderIdText.
        /// </summary>
        public static async Task RunJournalAttachmentFlow(string journalHeaderIdText)
        {
            ServiceLocator.Logger?.LogDebug($"JournalAttachments.RunJournalAttachmentFlow started. journalHeaderIdText='{journalHeaderIdText}'.");

            if (!GuardLoginAndExcel())
                return;

            using var ctsHelper = new CancellationHelper();
            CancellationToken token = ctsHelper.GetToken();

            if (!CommonMethods.TryDisableExcelSettings("JournalAttachments.RunJournalAttachmentFlow"))
                return;

            try
            {
                await HandleJournalAttachmentHyperlink(journalHeaderIdText, token);
            }
            catch (OperationCanceledException)
            {
                ServiceLocator.Logger?.LogWarn("Journal attachment operation was canceled by the user.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "JournalAttachments.RunJournalAttachmentFlow");
            }
            finally
            {
                CommonMethods.TryEnableExcelSettings("JournalAttachments.RunJournalAttachmentFlow");
            }
        }

        private static async Task HandleJournalAttachmentHyperlink(string journalHeaderIdText, CancellationToken token)
        {
            ServiceLocator.Logger?.LogDebug($"JournalAttachments.HandleJournalAttachmentHyperlink started. journalHeaderIdText='{journalHeaderIdText}'.");

            if (!double.TryParse(journalHeaderIdText, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedValue))
            {
                ServiceLocator.Logger?.LogWarn($"Unable to parse journal header id from '{journalHeaderIdText}'.");
                return;
            }

            long journalHeaderId = (long)Math.Truncate(parsedValue);

            var jsonObj = new JournalAttachmentsRequest
            {
                cubeId = AppState.Instance.SelectedCube.CubeId,
                journalHeaderId = journalHeaderId
            };

            token.ThrowIfCancellationRequested();

            string apiUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}journal-attachment-files";
            string httpPostText = JsonSerializer.Serialize(jsonObj, JsonGlobals.Options);

            ServiceLocator.Logger?.LogDebug($"JournalAttachments.HandleJournalAttachmentHyperlink: POST {apiUrl}");
            string responseData = await ApiHelper.ServerAPI(apiUrl, "JSON", httpPostText, "POST", token);
            ServiceLocator.Logger?.LogDebug($"JournalAttachments.HandleJournalAttachmentHyperlink: response received from server. Length={responseData?.Length ?? 0}.");

            token.ThrowIfCancellationRequested();

            var attachmentResult = ProcessAttachmentResponse(responseData);
            if (!attachmentResult.success) return;

            PopulateJournalDictionary(attachmentResult.records, token);

            if (AppState.Instance.JournalDictionary.Count > 0)
            {
                ShowAttachmentsDialog();
            }

            await DownloadSelectedAttachments(token);
        }

        private static (bool success, List<JournalAttachmentRecord> records) ProcessAttachmentResponse(string responseData)
        {
            if (string.IsNullOrWhiteSpace(responseData))
            {
                CommonFunctions.GLSenseMessage("Empty response from server.", MessageBoxImage.Error);
                return (false, null);
            }

            var result = ApiResponseHelper.Parse<List<JournalAttachmentRecord>>(responseData, JsonGlobals.Options);
            if (!result.IsSuccess)
            {
                CommonFunctions.GLSenseMessage(result.ErrorMessage, MessageBoxImage.Error);
                return (false, null);
            }

            var records = result.Value;
            if (records == null || records.Count == 0)
            {
                CommonFunctions.GLSenseMessage("No attachment records found.", MessageBoxImage.Information);
                return (false, null);
            }

            return (true, records);
        }

        private static void PopulateJournalDictionary(List<JournalAttachmentRecord> records, CancellationToken token)
        {
            AppState.Instance.AttachIDs = string.Empty;
            AppState.Instance.JournalDictionary.Clear();

            foreach (var record in records)
            {
                token.ThrowIfCancellationRequested();

                if (string.IsNullOrWhiteSpace(record.FILE_ID))
                    continue;

                if (!AppState.Instance.JournalDictionary.ContainsKey(record.FILE_ID))
                {
                    AppState.Instance.JournalDictionary.Add(record.FILE_ID, record.FILE_NAME ?? string.Empty);
                }
            }
        }

        private static void ShowAttachmentsDialog()
        {
            WpfAppManager.InvokeOnWpfThread(() =>
            {
                try
                {
                    var win = new AttachmentsDialog
                    {
                        CenterInExcel = true,
                        ModalToExcel = true,
                        ShowInTaskbar = false
                    };

                    win.ShowDialog();

                    ServiceLocator.Logger?.LogDebug("ShowAttachmentsDialog: Dialog closed.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "ShowAttachmentsDialog: ShowDialog error");
                }
            });
        }

        private static async Task DownloadSelectedAttachments(CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(AppState.Instance.AttachIDs))
                return;

            var jrAttachIDs = AppState.Instance.AttachIDs
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(id => id.Trim())
                .Where(id => long.TryParse(id, out _))
                .Select(long.Parse)
                .ToList();

            if (jrAttachIDs.Count == 0)
                return;

            var downloadRequest = new JrnalAttachRequest
            {
                cubeId = AppState.Instance.SelectedCube.CubeId,
                fileIds = jrAttachIDs.ToArray()
            };

            string downloadUrl = $"{AppState.Instance.LoginUrl}{AppConstants.RestSecure}journal-attachments";
            string downloadPayload = JsonSerializer.Serialize(downloadRequest, JsonGlobals.Options);

            await GLSense_DownloadFile(downloadUrl, downloadPayload, token);
            token.ThrowIfCancellationRequested();
        }

        private static async Task GLSense_DownloadFile(string strURL, string postData, CancellationToken token)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug($"JournalAttachments.GLSense_DownloadFile: POST {strURL}");

                token.ThrowIfCancellationRequested();

                var handler = new HttpClientHandler
                {
                    SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                    ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate
                };

                token.ThrowIfCancellationRequested();

                using var client = new HttpClient(handler);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", AppState.Instance.LoginToken);
                client.Timeout = Timeout.InfiniteTimeSpan;

                HttpContent content = string.IsNullOrWhiteSpace(postData) ? null : new StringContent(postData, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(strURL, content, token);

                token.ThrowIfCancellationRequested();

                if (response.IsSuccessStatusCode)
                {
                    string fileName = response.Content.Headers.ContentDisposition?.FileName?.Trim('"') ?? "DownloadedFile.zip";

                    string downloadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", fileName);

                    using var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None);
                    await response.Content.CopyToAsync(fileStream);

                    ServiceLocator.Logger?.LogDebug($"JournalAttachments.GLSense_DownloadFile: download succeeded, saved to '{downloadPath}'.");
                    CommonFunctions.GLSenseMessage($"Attachment saved to downloads folder as \"{fileName}\"", MessageBoxImage.Information);
                }
                else
                {
                    ServiceLocator.Logger?.LogError($"JournalAttachments.GLSense_DownloadFile: download failed: {response.StatusCode} - {response.ReasonPhrase}");
                }
            }
            catch (HttpRequestException ex)
            {
                ServiceLocator.Logger?.LogException(ex, "JournalAttachments.GLSense_DownloadFile");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "JournalAttachments.GLSense_DownloadFile");
            }
        }

        // ---------------------------------------------------------------------
        // Local helper duplicate (this project's per-file convention - see
        // BalanceHighlighter.cs/RowVisibilityProcessor.cs/RangeRefresher.cs/
        // DrillCellHighlighter.cs/CustomDrilldown.cs for the same idiom).
        // ---------------------------------------------------------------------
        private static bool GuardLoginAndExcel() => AppState.Instance.IsLoginCompleted && ServiceLocator.ExcelApp != null;
    }
}
