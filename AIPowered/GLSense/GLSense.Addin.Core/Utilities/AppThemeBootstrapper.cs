using GLSense.Addin.Core.Infrastructure;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace GLSense.Addin.Core.Utilities
{
    public static class AppThemeBootstrapper
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
                    ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Starting initialization...");

                    EnsureApplication();
                    LoadAllResourcesManually();

                    _initialized = true;
                    ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Initialization completed successfully.");
                }
                catch (Exception ex)
                {
                    ServiceLocator.Logger?.LogException(ex, "AppThemeBootstrapper.Initialize: Initialization failed");
                    throw;
                }
            }
        }

        private static void EnsureApplication()
        {
            if (Application.Current == null)
            {
                ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Creating new Application instance...");
                var app = new Application();

                app.DispatcherUnhandledException += (s, e) =>
                {
                    ServiceLocator.Logger?.LogException(e.Exception, "AppThemeBootstrapper: Unhandled WPF Dispatcher exception (suppressed, UI kept alive)");
                    e.Handled = true;
                };
            }
            else
            {
                ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Using existing Application instance.");
            }
        }

        private static void LoadAllResourcesManually()
        {
            if (Application.Current == null)
                throw new InvalidOperationException("Application.Current is null");

            var app = Application.Current;
            var mergedDictionaries = app.Resources.MergedDictionaries;

            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Loading all resources manually...");

            var existingDicts = mergedDictionaries.ToList();
            foreach (var dict in existingDicts)
            {
                mergedDictionaries.Remove(dict);
                ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Removed existing dictionary: {dict.Source}");
            }

            AddRequiredResources(app);
            AddFallbackResources(app);

            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Resource loading complete.");
        }

        private static void AddRequiredResources(Application app)
        {
            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Adding required resources...");

            var accentColor = (Color)ColorConverter.ConvertFromString("#0078D7");

            var accentBrush = new SolidColorBrush(accentColor);
            var subtleBrush = new SolidColorBrush(Color.FromArgb(255, 240, 240, 240));
            var subtleSecondaryBrush = new SolidColorBrush(Color.FromArgb(255, 225, 225, 225));
            var solidAccentBrush = new SolidColorBrush(accentColor);

            accentBrush.Freeze();
            subtleBrush.Freeze();
            subtleSecondaryBrush.Freeze();
            solidAccentBrush.Freeze();

            var resources = app.Resources;

            AddResourceIfMissing(resources, "SystemAccentColor", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimaryBrush", accentBrush);
            AddResourceIfMissing(resources, "SystemAccentColorSecondary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorTertiary", accentColor);

            AddResourceIfMissing(resources, "ControlBackgroundBrush", new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)));
            AddResourceIfMissing(resources, "ControlSubtleBackgroundBrush", subtleBrush);
            AddResourceIfMissing(resources, "ControlSubtleSecondaryBrush", subtleSecondaryBrush);
            AddResourceIfMissing(resources, "ControlSolidAccentBrush", solidAccentBrush);
            AddResourceIfMissing(resources, "ControlTextBrush", new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)));
            AddResourceIfMissing(resources, "ControlBorderBrush", new SolidColorBrush(Color.FromArgb(255, 200, 200, 200)));

            AddResourceIfMissing(resources, "CardBackgroundFillColorDefaultBrush", new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)));
            AddResourceIfMissing(resources, "CardStrokeColorDefaultBrush", new SolidColorBrush(Color.FromArgb(255, 220, 220, 220)));

            AddResourceIfMissing(resources, "TextFillColorPrimaryBrush", new SolidColorBrush(Color.FromArgb(255, 0, 0, 0)));
            AddResourceIfMissing(resources, "TextFillColorSecondaryBrush", new SolidColorBrush(Color.FromArgb(255, 100, 100, 100)));

            AddResourceIfMissing(resources, "ApplicationBackgroundBrush", new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)));

            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Required resources added.");
        }

        private static void AddResourceIfMissing(ResourceDictionary resources, string key, object value)
        {
            if (!resources.Contains(key))
            {
                resources[key] = value;
                ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Added resource '{key}'");
            }
        }

        private static void AddFallbackResources(Application app)
        {
            ServiceLocator.Logger?.LogDebug("AppThemeBootstrapper: Adding fallback resources...");

            var mergedDictionaries = app.Resources.MergedDictionaries;
            var fallbackDict = new ResourceDictionary();

            var keysToCheck = new[]
            {
                "ControlSubtleSecondaryBrush", "ControlSolidAccentBrush", "SystemAccentColorPrimary",
                "SystemAccentColorPrimaryBrush", "ControlBackgroundBrush", "ControlTextBrush",
                "ControlBorderBrush", "CardBackgroundFillColorDefaultBrush", "CardStrokeColorDefaultBrush",
                "TextFillColorPrimaryBrush", "TextFillColorSecondaryBrush", "ApplicationBackgroundBrush"
            };

            foreach (var key in keysToCheck)
            {
                bool exists = app.Resources.Contains(key) || mergedDictionaries.Any(d => d.Contains(key));
                if (exists)
                    continue;

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
                    ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Added fallback resource '{key}'");
                }
            }

            if (fallbackDict.Count > 0)
            {
                mergedDictionaries.Add(fallbackDict);
                ServiceLocator.Logger?.LogDebug($"AppThemeBootstrapper: Fallback dictionary added with {fallbackDict.Count} resources");
            }
        }

        public static bool IsInitialized => _initialized;
    }
}
