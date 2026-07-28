using GLSense.Addin.Core.Infrastructure;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Resources;
using Wpf.Ui.Appearance;

namespace GLSense.Addin.Core.Utilities
{
    public static class WpfUiBootstrapper
    {
        private static bool _initialized;
        private static readonly object _lock = new object();

        public static void Initialize()
        {
            if (_initialized)
                return;

            lock (_lock)
            {
                if (_initialized)
                    return;

                try
                {
                    ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Starting initialization...");

                    // Step 1: Ensure Application exists
                    EnsureApplication();

                    // Step 2: Load all required resources manually
                    LoadAllResourcesManually();

                    _initialized = true;
                    ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Initialization completed successfully.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "WpfUiBootstrapper.Initialize: Initialization failed");
                    throw;
                }
            }
        }

        private static void EnsureApplication()
        {
            if (Application.Current == null)
            {
                ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Creating new Application instance...");
                var app = new Application();

                app.DispatcherUnhandledException += (s, e) =>
                {
                    ServiceLocator.Logger?.LogException(e.Exception, "WpfUiBootstrapper: Unhandled WPF Dispatcher exception (suppressed, UI kept alive)");
                    e.Handled = true;
                };
            }
            else
            {
                ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Using existing Application instance.");
            }
        }

        private static void LoadAllResourcesManually()
        {
            if (Application.Current == null)
                throw new InvalidOperationException("Application.Current is null");

            var app = Application.Current;
            var mergedDictionaries = app.Resources.MergedDictionaries;

            ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Loading all resources manually...");

            // Clear any existing dictionaries to avoid conflicts
            var existingDicts = mergedDictionaries.ToList();
            foreach (var dict in existingDicts)
            {
                mergedDictionaries.Remove(dict);
                ServiceLocator.Logger?.LogDebug($"WpfUiBootstrapper: Removed existing dictionary: {dict.Source}");
            }

            // Step 1: Create and add all required resources directly
            AddRequiredResources(app);

            // Step 2: Try to load WPF-UI resources from pack URIs
            LoadWpfUiFromPackUris(app);

            // Step 3: If still missing, add fallback resources
            AddFallbackResources(app);

            ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Resource loading complete.");
        }

        private static void AddRequiredResources(Application app)
        {
            ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Adding required resources...");

            // Define the accent color (Windows 11 default blue)
            var accentColor = (Color)ColorConverter.ConvertFromString("#0078D7");

            // Create brushes
            var accentBrush = new SolidColorBrush(accentColor);
            var subtleBrush = new SolidColorBrush(Color.FromArgb(255, 240, 240, 240));
            var subtleSecondaryBrush = new SolidColorBrush(Color.FromArgb(255, 225, 225, 225));
            var solidAccentBrush = new SolidColorBrush(accentColor);

            // Freeze brushes for performance
            accentBrush.Freeze();
            subtleBrush.Freeze();
            subtleSecondaryBrush.Freeze();
            solidAccentBrush.Freeze();

            // Add all required resources to Application.Current.Resources
            var resources = app.Resources;

            // System Accent Colors
            AddResourceIfMissing(resources, "SystemAccentColor", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimaryBrush", accentBrush);
            AddResourceIfMissing(resources, "SystemAccentColorSecondary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorTertiary", accentColor);

            // Control Brushes (these are what you're missing)
            AddResourceIfMissing(resources, "ControlBackgroundBrush", new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)));
            AddResourceIfMissing(resources, "ControlSubtleBackgroundBrush", subtleBrush);
            AddResourceIfMissing(resources, "ControlSubtleSecondaryBrush", subtleSecondaryBrush);
            AddResourceIfMissing(resources, "ControlSolidAccentBrush", solidAccentBrush);
            AddResourceIfMissing(resources, "ControlTextBrush", new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)));
            AddResourceIfMissing(resources, "ControlBorderBrush", new SolidColorBrush(Color.FromArgb(255, 200, 200, 200)));

            // Card resources
            AddResourceIfMissing(resources, "CardBackgroundFillColorDefaultBrush", new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)));
            AddResourceIfMissing(resources, "CardStrokeColorDefaultBrush", new SolidColorBrush(Color.FromArgb(255, 220, 220, 220)));

            // Text colors
            AddResourceIfMissing(resources, "TextFillColorPrimaryBrush", new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)));
            AddResourceIfMissing(resources, "TextFillColorSecondaryBrush", new SolidColorBrush(Color.FromArgb(255, 100, 100, 100)));

            // Application background
            AddResourceIfMissing(resources, "ApplicationBackgroundBrush", new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)));

            ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Required resources added.");
        }

        private static void AddResourceIfMissing(ResourceDictionary resources, string key, object value)
        {
            if (!resources.Contains(key))
            {
                resources[key] = value;
                ServiceLocator.Logger?.LogDebug($"WpfUiBootstrapper: Added resource '{key}'");
            }
            else
            {
                ServiceLocator.Logger?.LogDebug($"WpfUiBootstrapper: Resource '{key}' already exists");
            }
        }

        private static void LoadWpfUiFromPackUris(Application app)
        {
            try
            {
                ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Trying to load WPF-UI from pack URIs...");

                var mergedDictionaries = app.Resources.MergedDictionaries;

                // Try to load theme
                try
                {
                    var themeUri = new Uri("pack://application:,,,/Wpf.Ui;component/Resources/Theme/Light.xaml", UriKind.Absolute);
                    var themeDict = new ResourceDictionary { Source = themeUri };
                    mergedDictionaries.Add(themeDict);
                    ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Theme loaded successfully.");
                }
                catch (Exception ex)
                {
                    // Non-fatal - AddFallbackResources (called after this) covers any
                    // resources that fail to load here, but still log the full exception:
                    // if the pack URI is wrong or the Wpf.Ui assembly isn't deployed, every
                    // dialog will silently render with fallback colors instead of the real
                    // theme, and this is the only place that would ever surface why.
                    ServiceLocator.Logger?.LogException(ex, "WpfUiBootstrapper.LoadWpfUiFromPackUris: Failed to load theme");
                }

                // Try to load controls
                try
                {
                    var controlsUri = new Uri("pack://application:,,,/Wpf.Ui;component/Resources/Wpf.Ui.xaml", UriKind.Absolute);
                    var controlsDict = new ResourceDictionary { Source = controlsUri };
                    mergedDictionaries.Add(controlsDict);
                    ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Controls loaded successfully.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "WpfUiBootstrapper.LoadWpfUiFromPackUris: Failed to load controls");
                }
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WpfUiBootstrapper.LoadWpfUiFromPackUris: Pack URI loading failed");
            }
        }

        private static void AddFallbackResources(Application app)
        {
            ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Adding fallback resources...");

            var mergedDictionaries = app.Resources.MergedDictionaries;

            // Create a fallback resource dictionary with all required resources
            var fallbackDict = new ResourceDictionary();

            // Add any missing resources that might still be needed
            var keysToCheck = new[]
            {
                "ControlSubtleSecondaryBrush",
                "ControlSolidAccentBrush",
                "SystemAccentColorPrimary",
                "SystemAccentColorPrimaryBrush",
                "ControlBackgroundBrush",
                "ControlTextBrush",
                "ControlBorderBrush",
                "CardBackgroundFillColorDefaultBrush",
                "CardStrokeColorDefaultBrush",
                "TextFillColorPrimaryBrush",
                "TextFillColorSecondaryBrush",
                "ApplicationBackgroundBrush"
            };

            foreach (var key in keysToCheck)
            {
                // Check if resource exists anywhere
                bool exists = app.Resources.Contains(key);

                if (!exists)
                {
                    // Check merged dictionaries
                    foreach (var dict in mergedDictionaries)
                    {
                        if (dict.Contains(key))
                        {
                            exists = true;
                            break;
                        }
                    }
                }

                if (!exists)
                {
                    // Add fallback based on key
                    object value = key switch
                    {
                        "ControlSubtleSecondaryBrush" => new SolidColorBrush(Color.FromArgb(255, 225, 225, 225)),
                        "ControlSolidAccentBrush" => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D7")),
                        "SystemAccentColorPrimary" => (Color)ColorConverter.ConvertFromString("#0078D7"),
                        "SystemAccentColorPrimaryBrush" => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D7")),
                        "ControlBackgroundBrush" => new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)),
                        "ControlTextBrush" => new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)),
                        "ControlBorderBrush" => new SolidColorBrush(Color.FromArgb(255, 200, 200, 200)),
                        "CardBackgroundFillColorDefaultBrush" => new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)),
                        "CardStrokeColorDefaultBrush" => new SolidColorBrush(Color.FromArgb(255, 220, 220, 220)),
                        "TextFillColorPrimaryBrush" => new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)),
                        "TextFillColorSecondaryBrush" => new SolidColorBrush(Color.FromArgb(255, 100, 100, 100)),
                        "ApplicationBackgroundBrush" => new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)),
                        _ => null
                    };

                    if (value != null)
                    {
                        fallbackDict[key] = value;
                        ServiceLocator.Logger?.LogDebug($"WpfUiBootstrapper: Added fallback resource '{key}'");
                    }
                }
            }

            // Add the fallback dictionary if it has any resources
            if (fallbackDict.Count > 0)
            {
                mergedDictionaries.Add(fallbackDict);
                ServiceLocator.Logger?.LogDebug($"WpfUiBootstrapper: Fallback dictionary added with {fallbackDict.Count} resources");
            }
        }

        public static void SetDarkTheme()
        {
            if (!_initialized) Initialize();

            try
            {
                ApplicationThemeManager.Apply(ApplicationTheme.Dark);
                ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Theme changed to Dark.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WpfUiBootstrapper.SetDarkTheme: Failed to set Dark theme");
            }
        }

        public static void SetLightTheme()
        {
            if (!_initialized) Initialize();

            try
            {
                ApplicationThemeManager.Apply(ApplicationTheme.Light);
                ServiceLocator.Logger?.LogDebug("WpfUiBootstrapper: Theme changed to Light.");
            }
            catch (Exception ex)
            {
                ServiceLocator.Logger?.LogException(ex, "WpfUiBootstrapper.SetLightTheme: Failed to set Light theme");
            }
        }

        public static bool IsInitialized => _initialized;
    }
}