// ConfiguratorPaneHost.cs in GLSense.Addin.Core
//
// Group H (Balance Configurator pane) - HWND-reparenting bridge between the host's
// GLConfiguratorPane (a WinForms AddinExpress.XL.ADXExcelTaskPane - host-only, can't be
// referenced from this project) and this project's WPF content.
//
// Why this exists: the old monolith's GLConfiguratorPane directly did
// `_wpfControl = new GLBalanceConfigurator(this); _host = new ElementHost { Child =
// _wpfControl }; this.Controls.Add(_host);` - trivial, because everything ran in one
// AppDomain. In this architecture GLBalanceConfigurator (and its ~3000-line
// GLConfiguratorViewModel) live in Addin.Core's separate, hot-reloadable AppDomain, and a
// WPF FrameworkElement is not MarshalByRefObject - it cannot cross an AppDomain boundary
// by reference, so the host cannot `new` it up or set it as an ElementHost.Child directly.
//
// Instead: this class creates a real, self-contained, borderless top-level Window - the
// exact same "own dedicated WPF thread via WpfAppManager" trick every DpiAwareWindow-derived
// dialog in this migration already uses - whose Content is GLBalanceConfigurator, and
// exposes only its native window handle (an IntPtr; a blittable value that crosses an
// AppDomain boundary in the same process safely, unlike a live object reference) via
// IGLSenseAddin.CreateConfiguratorPaneContent(). The host (GLSense\GLConfiguratorPane.cs)
// then:
//   1. Win32 SetParent's that handle into its own WinForms panel.
//   2. Rewrites its window style bits (WS_POPUP -> WS_CHILD, drops WS_CAPTION/
//      WS_THICKFRAME) so it renders as embedded content instead of a floating window.
//   3. Keeps it in sync via MoveWindow on resize.
//   4. AttachThreadInput's the two threads so keyboard focus/Tab-navigation flows into
//      the reparented window correctly (it's still pumped by Addin.Core's own WPF
//      thread, not the host's/Excel's main STA thread).
//
// This keeps GLBalanceConfigurator/GLConfiguratorViewModel fully hot-reloadable (the
// whole point of the two-AppDomain architecture), at the cost of this one bridge class
// plus the host-side Win32 interop in GLConfiguratorPane.cs - there is no other novel
// infrastructure needed anywhere else in this project for task panes specifically.
using GLSense.Addin.Core.Infrastructure;
using System;
using System.Windows;
using System.Windows.Interop;

namespace GLSense.Addin.Core.Views
{
    public static class ConfiguratorPaneHost
    {
        private static Window _window;
        private static GLBalanceConfigurator _content;
        private static readonly object _lock = new object();

        /// <summary>
        /// Idempotent - if the content already exists (e.g. RibFSG toggled the pane
        /// closed and reopened without a full Excel/AppDomain reload), returns the
        /// existing handle rather than creating a second Window.
        /// </summary>
        public static IntPtr CreateContent()
        {
            ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.CreateContent invoked");
            lock (_lock)
            {
                if (_window != null)
                {
                    ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.CreateContent: existing window found, reusing handle");
                    try
                    {
                        return new WindowInteropHelper(_window).Handle;
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "ConfiguratorPaneHost.CreateContent: existing handle read failed, recreating");
                        _window = null;
                        _content = null;
                    }
                }

                IntPtr handle = IntPtr.Zero;

                Utilities.WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        _content = new GLBalanceConfigurator();
                        _content.OnCloseRequested += OnContentCloseRequested;

                        _window = new Window
                        {
                            Content = _content,
                            WindowStyle = WindowStyle.None,
                            ResizeMode = ResizeMode.NoResize,
                            ShowInTaskbar = false,
                            AllowsTransparency = false,
                            SizeToContent = SizeToContent.Manual,
                            // Off-screen until the host SetParent's + MoveWindow's this
                            // into its panel - avoids a visible flash at the primary
                            // monitor's origin before it's docked into place.
                            Left = -32000,
                            Top = -32000,
                            Width = 600,
                            Height = 400
                        };

                        // Show() (even off-screen) is the well-tested path for this kind
                        // of embedding recipe - it fully pumps the Window through Loaded/
                        // render initialization before the host reparents its HWND, unlike
                        // relying on WindowInteropHelper.EnsureHandle() alone.
                        _window.Show();

                        handle = new WindowInteropHelper(_window).Handle;
                        ServiceLocator.Logger?.LogDebug($"ConfiguratorPaneHost.CreateContent: new window created, handle={handle}");
                    }
                    catch (Exception ex)
                    {
                        ServiceLocator.Logger?.LogException(ex, "ConfiguratorPaneHost.CreateContent");
                        handle = IntPtr.Zero;
                    }
                });

                return handle;
            }
        }

        private static void OnContentCloseRequested()
        {
            try
            {
                ServiceLocator.RibbonController?.HideConfiguratorPane();
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ConfiguratorPaneHost.OnContentCloseRequested");
            }
        }

        /// <summary>Old monolith's GLConfiguratorPane.RelaunchPane().</summary>
        public static void Relaunch()
        {
            ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.Relaunch invoked");
            if (_content == null)
            {
                ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.Relaunch: no existing content, nothing to relaunch");
                return;
            }

            try
            {
                Utilities.WpfAppManager.InvokeOnWpfThread(() =>
                {
                    // Fire-and-forget: ReLoadConfigurator is async (network + UI work);
                    // this call is dispatched fire-and-forget from OnRibbonAction/host
                    // event handlers the same way every other Group C-G ribbon action is.
                    _ = _content.ReLoadConfigurator();
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ConfiguratorPaneHost.Relaunch");
            }
        }

        /// <summary>Old monolith's GLConfiguratorPane.ResetPaneReference().</summary>
        public static void ResetReference()
        {
            ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.ResetReference invoked");
            if (_content == null)
            {
                ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.ResetReference: no existing content, nothing to reset");
                return;
            }

            try
            {
                Utilities.WpfAppManager.InvokeOnWpfThread(() =>
                {
                    GLBalanceConfigurator.ResetCellReference();
                });
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "ConfiguratorPaneHost.ResetReference");
            }
        }

        /// <summary>
        /// Tears down the hosted Window - used at Shutdown/logoff and before a hot-reload
        /// swap so the outgoing AppDomain's Window doesn't linger reparented inside the
        /// host's still-alive task pane.
        /// </summary>
        public static void Close()
        {
            ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.Close invoked");
            lock (_lock)
            {
                if (_window == null)
                {
                    ServiceLocator.Logger?.LogDebug("ConfiguratorPaneHost.Close: no existing window, nothing to close");
                    return;
                }

                try
                {
                    var windowToClose = _window;
                    Utilities.WpfAppManager.InvokeOnWpfThread(() =>
                    {
                        windowToClose.Close();
                    });
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "ConfiguratorPaneHost.Close");
                }
                finally
                {
                    _window = null;
                    _content = null;
                }
            }
        }
    }
}
