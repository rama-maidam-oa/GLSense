// GLReleaseHistoryBrowser.xaml.cs in GLSense\Views
using GLSense.Loader.Core;
using GLSense.Shared;
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

            TxtStatus.Text = entries.Count == 0
                ? "No releases recorded yet."
                : "Select a release, then click Load This Release.";

            var loadedRow = rows.FirstOrDefault(r => r.IsCurrentlyLoaded);
            if (loadedRow != null)
            {
                TxtLoadedStatus.Text = $"Currently loaded: version {loadedRow.Version}, released {loadedRow.ReleaseDate}.";
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
            BtnLoad.IsEnabled = GridReleases.SelectedItem is ReleaseRow;
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

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
