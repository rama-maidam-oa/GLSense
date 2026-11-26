// OnlineUpdateCheck.cs in GLSense (host)
//
// Post-login add-in update check (login -> update check -> broadcast): the same steps the Reload Add-in window's
// Online mode runs (GLReloadSourcePicker.LoadOnlineReleasesAsync / BtnFetchAndReload_Click), without any UI, so
// Addin.Core can ask "is there a newer release?" right after login and reload into it on the user's Yes.
using GLSense.Contracts;
using GLSense.Loader.Core;
using GLSense.Shared;
using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace GLSense
{
    internal static class OnlineUpdateCheck
    {
        /// <summary>
        /// The newest released GLSense on the server whose release date is later than the loaded one's, or null when
        /// there is none (including a server with no release index - HTTP 404). Throws on any other failure; the caller
        /// logs it and carries on with login.
        /// "Later release date", not just "not on this machine" (the Reload window's IsSelectable): an older release
        /// that was never downloaded is not an update. A newer release that is already downloaded but not loaded is.
        /// </summary>
        internal static async Task<OnlineReleaseRow> FindNewerReleaseAsync(IGLSenseContext context, string loginUrl, string loginToken, TimeSpan timeout)
        {
            string url = loginUrl.TrimEnd('/') + GLReloadSourcePicker.ReleaseIndexPath;
            string listJson;
            using (var client = GLReloadSourcePicker.CreateHttpClient(timeout))
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", loginToken);
                try
                {
                    using (var response = await GLReloadSourcePicker.GetCheckedAsync(client, url).ConfigureAwait(false))
                        listJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                }
                catch (HttpRequestException ex) when (ex.Message == GLReloadSourcePicker.NotOnServerMessage)
                {
                    context.Logger?.LogDebug($"OnlineUpdateCheck: {url} returned 404 - no updates published.");
                    return null;
                }
            }

            var parsed = new VersionParser().ParseVersionJson(listJson);
            if (!parsed.Success || parsed.AllVersions == null)
                throw new InvalidOperationException($"could not parse the release list: {parsed.ErrorMessage ?? "empty response"}");

            var released = OnlineReleaseFilter.ReleasedOnly(listJson, parsed.AllVersions);
            var local = ReleaseHistoryStore.ReadAll(context.Paths.ReleaseHistoryFile);
            var rows = OnlineReleaseClassifier.Classify(released, local, context.Version, context.ReleaseDate);

            DateTime loaded = OnlineReleaseClassifier.ParseReleaseDateOrMin(context.ReleaseDate);
            return rows
                .Where(r => !r.IsCurrentlyLoaded && !string.IsNullOrWhiteSpace(r.FolderName))
                .Where(r => OnlineReleaseClassifier.ParseReleaseDateOrMin(r.ReleaseDate) > loaded)
                .OrderByDescending(r => OnlineReleaseClassifier.ParseReleaseDateOrMin(r.ReleaseDate))
                .FirstOrDefault();
        }

        /// <summary>
        /// Downloads the release's manifest + zip and catalogs it (an already-extracted release folder is reused, see
        /// UpdateBootstrapper.ExtractAndCatalog). Returns the release to load, or null when it could not be cataloged
        /// or the server returned a different release than the one asked for. Throws on HTTP failures.
        /// </summary>
        internal static async Task<ResolvedRelease> FetchAndCatalogAsync(IGLSenseContext context, string loginUrl, string loginToken, OnlineReleaseRow row)
        {
            // Release zips can be several MB - same 5-minute limit as the Reload window.
            using (var client = GLReloadSourcePicker.CreateHttpClient(TimeSpan.FromMinutes(5)))
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", loginToken);

                string manifestJson;
                using (var response = await GLReloadSourcePicker.GetCheckedAsync(client, GLReloadSourcePicker.BuildVersionUrl(loginUrl, row.FolderName, "manifest")).ConfigureAwait(false))
                    manifestJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                byte[] zipBytes;
                using (var response = await GLReloadSourcePicker.GetCheckedAsync(client, GLReloadSourcePicker.BuildVersionUrl(loginUrl, row.FolderName, "zip")).ConfigureAwait(false))
                    zipBytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                var resolved = new UpdateBootstrapper().ExtractAndCatalog(context, manifestJson, zipBytes, "Online");
                if (resolved == null)
                    return null;

                if (!string.Equals(resolved.Version, row.Version, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(resolved.ReleaseDate, row.ReleaseDate, StringComparison.OrdinalIgnoreCase))
                {
                    context.Logger?.LogWarn($"OnlineUpdateCheck: server returned {resolved.Version} ({resolved.ReleaseDate}) for {row.Version} ({row.ReleaseDate}) - not reloading.");
                    return null;
                }

                return resolved;
            }
        }
    }
}
