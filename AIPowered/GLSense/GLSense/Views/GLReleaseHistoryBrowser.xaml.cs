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

        public GLReleaseHistoryBrowser()
        {
            InitializeComponent();
            LoadEntries();
        }

        private void LoadEntries()
        {
            var paths = GlobalsEx.Context.Paths;

            // Reconcile first (spec section 6) so a stale entry - whose Versions\
            // folder was deleted by disk cleanup, an AppData purge, etc., with no
            // reinstall involved at all - is never shown as selectable.
            var entries = ReleaseHistoryStore.Reconcile(paths.ReleaseHistoryFile, paths.VersionsPath);

            GridReleases.ItemsSource = entries
                .OrderByDescending(e => e.ReleaseDate)
                .ToList();

            TxtStatus.Text = entries.Count == 0
                ? "No releases recorded yet."
                : "Select a release, then click Load This Release.";
        }

        private void GridReleases_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BtnLoad.IsEnabled = GridReleases.SelectedItem is ReleaseEntry;
        }

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            if (!(GridReleases.SelectedItem is ReleaseEntry picked)) return;

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
