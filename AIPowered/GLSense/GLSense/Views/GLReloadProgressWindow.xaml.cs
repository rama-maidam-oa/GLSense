// GLReloadProgressWindow.xaml.cs in GLSense\Views
//
// Host-side busy indicator shown for the duration of AddinModule.ReloadAddinCore()'s
// blocking AppDomain unload/resolve/load work - previously the only feedback during that
// entire operation was a WinForms wait cursor, with no visible indication of what was
// happening or how long it might take, and nothing stopping a stray click/Alt+F4/Escape.
//
// Plain host-side WPF window, same reasoning as GLReloadSourcePicker/GLReleaseHistoryBrowser
// (see CLAUDE.md section 40.4) - can't inherit GLSense.Addin.Core.Views.BaseWindow, since it
// must keep working even while Addin.Core is between instances (that's the whole reason it
// exists). The caller must set Owner via WindowInteropHelper before calling ShowAndRender(),
// same pattern as those two windows.
//
// Deliberately NOT closeable while the reload is in progress: Closing is always cancelled
// unless AllowCloseAndClose() (called by the caller once ReloadAddinCore actually finishes,
// success or failure) has set _allowClose first - covers the title bar's X button, Alt+F4,
// and Escape all in one place, since all three ultimately raise the same Closing event.
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;

namespace GLSense
{
    public partial class GLReloadProgressWindow : Window
    {
        private bool _allowClose;

        public GLReloadProgressWindow(string message)
        {
            InitializeComponent();
            if (!string.IsNullOrWhiteSpace(message)) TxtMessage.Text = message;

            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
            Closing += GLReloadProgressWindow_Closing;
        }

        private void GLReloadProgressWindow_Closing(object sender, CancelEventArgs e)
        {
            if (!_allowClose) e.Cancel = true;
        }

        /// <summary>
        /// Show()'s WPF equivalent of BaseWindow.PumpDispatcherFrame() (see CLAUDE.md
        /// section 1.4d) - Show() alone returns immediately without painting anything, and
        /// the caller is about to run genuinely blocking, non-yielding work (AppDomain
        /// unload/load) on this same thread right afterward, with no further opportunity for
        /// this window to render until that work finishes. Forcing two dispatcher passes here
        /// (Render, then Background) guarantees the window's first frame is actually on
        /// screen before the freeze starts, instead of a blank/unpainted window appearing and
        /// disappearing around a block it never got to render into.
        /// </summary>
        public void ShowAndRender()
        {
            Show();
            Dispatcher.Invoke(DispatcherPriority.Render, new Action(() => { }));
            Dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { }));
        }

        public void AllowCloseAndClose()
        {
            _allowClose = true;
            Close();
        }
    }
}
