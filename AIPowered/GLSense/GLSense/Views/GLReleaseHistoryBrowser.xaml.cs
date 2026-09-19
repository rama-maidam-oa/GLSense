// GLReleaseHistoryBrowser.xaml.cs in GLSense\Views
using GLSense.Loader.Core;
using GLSense.Shared;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace GLSense
{
    public partial class GLReleaseHistoryBrowser : Window
    {
        public ResolvedRelease Chosen { get; private set; }

        /// <summary>
        /// Display-only wrapper around a catalog ReleaseEntry, adding the "is this the
        /// release currently loaded into this Excel session" flag used to highlight/
        /// pre-select the matching row. Deliberately NOT added directly onto ReleaseEntry
        /// itself - that class is round-tripped through JsonSerializer.Serialize by
        /// ReleaseHistoryStore, so a property added there would get persisted into
        /// ReleaseHistory.json for no reason.
        /// </summary>
        private class ReleaseRow
        {
            public ReleaseEntry Entry { get; set; }
            public string Version => Entry.Version;
            public string ReleaseDate => Entry.ReleaseDate;
            public string Source => Entry.Source;
            public string Notes => Entry.Notes;
            public bool IsCurrentlyLoaded { get; set; }
            public string LoadedMarker => IsCurrentlyLoaded ? "●" : string.Empty;
        }

        public GLReleaseHistoryBrowser()
        {
            InitializeComponent();
            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
            LoadEntries();
        }

        private void LoadEntries()
        {
            var paths = GlobalsEx.Context.Paths;
            string activeFolderName = GlobalsEx.Context.ActiveFolderName;

            // Reconcile first (spec section 6) so a stale entry - whose Versions\
            // folder was deleted by disk cleanup, an AppData purge, etc., with no
            // reinstall involved at all - is never shown as selectable.
            var entries = ReleaseHistoryStore.Reconcile(paths.ReleaseHistoryFile, paths.VersionsPath);

            var rows = entries
                .OrderByDescending(e => e.ReleaseDate)
                .Select(e => new ReleaseRow
                {
                    Entry = e,
                    IsCurrentlyLoaded = !string.IsNullOrEmpty(activeFolderName) &&
                        string.Equals(e.FolderName, activeFolderName, System.StringComparison.OrdinalIgnoreCase)
                })
                .ToList();

            GridReleases.ItemsSource = rows;

            if (entries.Count == 0)
            {
                TxtStatus.Text = "No version recorded yet.";
            }
            else
            {
                TxtStatus.Inlines.Clear();
                TxtStatus.Inlines.Add(new System.Windows.Documents.Run("Select a version, then click "));
                TxtStatus.Inlines.Add(new System.Windows.Documents.Run("Load This Version")
                {
                    FontWeight = FontWeights.SemiBold
                });
                TxtStatus.Inlines.Add(new System.Windows.Documents.Run("."));
            }

            var loadedRow = rows.FirstOrDefault(r => r.IsCurrentlyLoaded);
            if (loadedRow != null)
            {
                TxtLoadedStatus.Text = $"Currently loaded: version {loadedRow.Version}, released on {loadedRow.ReleaseDate}.";
                GridReleases.SelectedItem = loadedRow;
                GridReleases.ScrollIntoView(loadedRow);
            }
            else
            {
                // Not necessarily an error - e.g. the running release's folder was
                // reconciled away, or ActiveFolderName isn't populated in this context.
                TxtLoadedStatus.Text = $"Currently loaded: version {GlobalsEx.Context.Version} (not found in the history list).";
            }
        }

        private void GridReleases_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var row = GridReleases.SelectedItem as ReleaseRow;
            BtnLoad.IsEnabled = row != null;
            BtnRemove.IsEnabled = row != null;
        }

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            if (!(GridReleases.SelectedItem is ReleaseRow row)) return;
            var picked = row.Entry;

            // Deliberately no version-gate here (unlike RibReload's Online/Offline
            // path) - loading an older release on purpose is this window's entire
            // reason to exist. See spec section 8.
            Chosen = new ResolvedRelease
            {
                Version = picked.Version,
                ReleaseDate = picked.ReleaseDate,
                FolderName = picked.FolderName
            };

            DialogResult = true;
            Close();
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            if (!(GridReleases.SelectedItem is ReleaseRow row)) return;

            if (row.IsCurrentlyLoaded)
            {
                MessageBox.Show(
                    "The currently loaded version is active and cannot be removed. Load another version first, then remove this one.",
                    "Cannot Remove Active Version",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var confirmation = MessageBox.Show(
                $"Remove version {row.Version} released on {row.ReleaseDate} from this machine?\n\n" +
                "This permanently deletes its local files and history entry. " +
                "The release can be downloaded again later if it is still available.",
                "Remove Version",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (confirmation != MessageBoxResult.Yes) return;

            ReleaseDeleteResult result;
            Exception error;
            try
            {
                var paths = GlobalsEx.Context.Paths;
                result = ReleaseHistoryStore.Delete(
                    paths.ReleaseHistoryFile,
                    paths.VersionsPath,
                    row.Entry.FolderName,
                    GlobalsEx.Context.ActiveFolderName,
                    out error);
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(
                    ex,
                    $"GLReleaseHistoryBrowser: failed to remove version '{row.Version}'.");
                MessageBox.Show(
                    $"The version could not be removed.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    "Remove Version Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            switch (result)
            {
                case ReleaseDeleteResult.Deleted:
                    GlobalsEx.Context?.Logger?.LogDebug(
                        $"GLReleaseHistoryBrowser: removed version '{row.Version}' folder '{row.Entry.FolderName}'.");
                    LoadEntries();
                    break;

                case ReleaseDeleteResult.ActiveRelease:
                    MessageBox.Show(
                        "The selected version is currently active and cannot be removed. Load another version first.",
                        "Cannot Remove Active Version",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    break;

                case ReleaseDeleteResult.FolderDeleteFailed:
                    GlobalsEx.Context?.Logger?.LogException(
                        error,
                        $"GLReleaseHistoryBrowser: could not remove folder '{row.Entry.FolderName}'.");
                    MessageBox.Show(
                        $"The version could not be removed because its files could not be deleted.{Environment.NewLine}{Environment.NewLine}{error?.Message}",
                        "Remove Version Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    break;

                case ReleaseDeleteResult.CatalogWriteFailed:
                    GlobalsEx.Context?.Logger?.LogException(
                        error,
                        $"GLReleaseHistoryBrowser: folder '{row.Entry.FolderName}' was deleted but the history catalog could not be updated.");
                    MessageBox.Show(
                        $"The version files were deleted, but the history catalog could not be updated.{Environment.NewLine}{Environment.NewLine}{error?.Message}",
                        "History Update Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    LoadEntries();
                    break;

                case ReleaseDeleteResult.InvalidFolderName:
                    GlobalsEx.Context?.Logger?.LogWarn(
                        $"GLReleaseHistoryBrowser: refused unsafe release folder '{row.Entry.FolderName}'.");
                    MessageBox.Show(
                        "The selected version has an invalid local folder name and was not removed.",
                        "Remove Version Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    break;

                case ReleaseDeleteResult.NotFound:
                    MessageBox.Show(
                        "That version is no longer present in the local history. The list will be refreshed.",
                        "Version Not Found",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LoadEntries();
                    break;
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
