using ControlzEx.Theming;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;

namespace GLSense.Utilities
{
    public static class MahAppsBootstrapper
    {
        private static bool _initialized;
        private static readonly object _lock = new object();
        private static string _currentAccentHex;
        private static string _currentBaseTheme;

        /// <summary>
        /// Initialize MahApps once during ribbon load
        /// </summary>
        public static void Init(string accentHex, string baseTheme)
        {
            // If already initialized with same values, just return
            if (_initialized && _currentAccentHex == accentHex && _currentBaseTheme == baseTheme)
            {
                return;
            }

            lock (_lock)
            {
                // If initialized with different values, we need to update
                if (_initialized)
                {
                    UpdateTheme(accentHex, baseTheme);
                    return;
                }

                // First time initialization
                _currentAccentHex = accentHex;
                _currentBaseTheme = baseTheme;

                WpfAppManager.EnsureApplication();

                var app = Application.Current;
                if (app == null) return;

                try
                {
                    if (app.Dispatcher.CheckAccess())
                    {
                        InitializeMahApps(app, accentHex, baseTheme);
                        // Validate expected keys and log missing ones
                        ValidateThemeResources(app);
                    }
                    else
                    {
                        app.Dispatcher.Invoke(() =>
                        {
                            InitializeMahApps(app, accentHex, baseTheme);
                            ValidateThemeResources(app);
                        });
                    }

                    _initialized = true;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "MahAppsBootstrapper.Init");
                }
            }
        }

        /// <summary>
        /// Preload all resources to ensure they're ready when windows open
        /// Call this once after Init()
        /// </summary>
        public static void PreloadResources()
        {
            if (!_initialized)
            {
                LogUtility.LogWarn("Cannot preload resources before initialization");
                return;
            }

            try
            {
                var app = Application.Current;
                if (app == null) return;

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        // Force load all merged dictionaries
                        foreach (var dict in app.Resources.MergedDictionaries)
                        {
                            if (dict.Source != null)
                            {
                                var keys = dict.Keys; // Force loading
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "MahAppsBootstrapper.PreloadResources (inner)");
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "MahAppsBootstrapper.PreloadResources");
            }
        }
        /// <summary>
        /// Update theme at runtime (call this if user changes theme in settings)
        /// </summary>
        public static void UpdateTheme(string accentHex, string baseTheme)
        {
            if (!_initialized)
            {
                Init(accentHex, baseTheme);
                return;
            }

            lock (_lock)
            {
                if (_currentAccentHex == accentHex && _currentBaseTheme == baseTheme)
                    return;

                _currentAccentHex = accentHex;
                _currentBaseTheme = baseTheme;

                var app = Application.Current;
                if (app == null) return;

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        // Update accent colors
                        SetAccentColors(app, accentHex);

                        // Update theme brushes
                        DefineEssentialBrushes(app, baseTheme);

                        // Try to update via ThemeManager if available
                        try
                        {
                            var currentTheme = ThemeManager.Current.DetectTheme(app);
                            if (currentTheme != null)
                            {
                                var accentColor = (Color)ColorConverter.ConvertFromString(accentHex);
                                var newTheme = new Theme(
                                    name: $"{baseTheme}.{accentHex.Replace("#", "")}",
                                    displayName: $"{baseTheme} {accentHex}",
                                    baseColorScheme: baseTheme,
                                    colorScheme: accentHex.Replace("#", ""),
                                    primaryAccentColor: accentColor,
                                    showcaseBrush: new SolidColorBrush(accentColor),
                                    isRuntimeGenerated: true,
                                    isHighContrast: false);

                                ThemeManager.Current.ChangeTheme(app, newTheme);
                            }
                        }
                        catch
                        {
                            // ThemeManager update failed, but we already updated colors manually
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "MahAppsBootstrapper.UpdateTheme");
                    }
                });
            }
        }
        private static void InitializeMahApps(Application app, string accentHex, string baseTheme)
        {
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                // Load minimal required resources
                LoadResourceIfMissing(app, $"pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml");
                LoadResourceIfMissing(app, $"pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml");

                var theme = baseTheme.Equals("Dark", StringComparison.OrdinalIgnoreCase)
                    ? "Dark.Blue"
                    : "Light.Blue";

                LoadResourceIfMissing(app, $"pack://application:,,,/MahApps.Metro;component/Styles/Themes/{theme}.xaml");

                // Load GLSense global styles (required for SuggestAppendComboBox and other controls)
                LoadResourceIfMissing(app, $"pack://application:,,,/GLSense;component/Themes/GlobalStyles.xaml");
                LoadResourceIfMissing(app, $"pack://application:,,,/GLSense;component/Themes/Generic.xaml");

                // Set accent colors
                SetAccentColors(app, accentHex);

                // Define essential brushes
                DefineEssentialBrushes(app, baseTheme);

                _initialized = true;
            }
        }

        private static void LoadResourceIfMissing(Application app, string uriString)
        {
            try
            {
                var uri = new Uri(uriString, UriKind.Absolute);

                if (app.Resources.MergedDictionaries.Any(d => d.Source == uri))
                    return;

                var dict = new ResourceDictionary { Source = uri };
                var keys = dict.Keys; // Force loading

                app.Resources.MergedDictionaries.Add(dict);
            }
            catch (Exception ex) 
            {
                LogUtility.LogException(ex, $"Resource not found at : {uriString}");
            }
        }

        private static void SetAccentColors(Application app, string accentHex)
        {
            try
            {
                var baseColor = (Color)ColorConverter.ConvertFromString(accentHex);

                var accent2 = Lighten(baseColor, 0.10);
                var accent3 = Darken(baseColor, 0.18);
                var accent4 = Darken(baseColor, 0.25);
                var accentSelected = Darken(baseColor, 0.08);

                app.Resources["MahApps.Colors.Accent"] = baseColor;
                app.Resources["MahApps.Colors.Accent2"] = accent2;
                app.Resources["MahApps.Colors.Accent3"] = accent3;
                app.Resources["MahApps.Colors.Accent4"] = accent4;
                app.Resources["MahApps.Colors.AccentSelected"] = accentSelected;

                app.Resources["MahApps.Brushes.Accent"] = CreateFrozenBrush(baseColor);
                app.Resources["MahApps.Brushes.Accent2"] = CreateFrozenBrush(accent2);
                app.Resources["MahApps.Brushes.Accent3"] = CreateFrozenBrush(accent3);
                app.Resources["MahApps.Brushes.Accent4"] = CreateFrozenBrush(accent4);
                app.Resources["MahApps.Brushes.AccentSelected"] = CreateFrozenBrush(accentSelected);

                // Ideal foreground
                double luminance = (0.299 * baseColor.R + 0.587 * baseColor.G + 0.114 * baseColor.B) / 255.0;
                Color idealColor = luminance > 0.5 ? Colors.Black : Colors.White;

                app.Resources["MahApps.Colors.IdealForeground"] = idealColor;
                app.Resources["MahApps.Brushes.IdealForeground"] = CreateFrozenBrush(idealColor);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "MahAppsBootstrapper.SetAccentColors");
            }
        }

        private static void DefineEssentialBrushes(Application app, string baseTheme)
        {
            // keep logic simple and split responsibilities to reduce cognitive complexity
            bool isLight = baseTheme.Equals("Light", StringComparison.OrdinalIgnoreCase);
            RegisterCoreBrushes(app, isLight);
            RegisterCompatibilityPairs(app, isLight);
        }

        private static void RegisterCoreBrushes(Application app, bool isLight)
        {
            var essentialBrushes = new Dictionary<string, Color>
            {
                ["MahApps.Brushes.Gray1"] = Color.FromRgb(250, 250, 250),
                ["MahApps.Brushes.Gray2"] = Color.FromRgb(245, 245, 245),
                ["MahApps.Brushes.Gray3"] = Color.FromRgb(238, 238, 238),
                ["MahApps.Brushes.Gray4"] = Color.FromRgb(224, 224, 224),
                ["MahApps.Brushes.Gray5"] = Color.FromRgb(204, 204, 204),
                ["MahApps.Brushes.Gray6"] = Color.FromRgb(189, 189, 189),
                ["MahApps.Brushes.Gray7"] = Color.FromRgb(158, 158, 158),
                ["MahApps.Brushes.Gray8"] = Color.FromRgb(117, 117, 117),
                ["MahApps.Brushes.Gray9"] = Color.FromRgb(97, 97, 97),
                ["MahApps.Brushes.Gray10"] = Color.FromRgb(66, 66, 66),

                ["MahApps.Brushes.ThemeForeground"] = isLight ? Colors.Black : Colors.White,
                ["MahApps.Brushes.ThemeBackground"] = isLight ? Colors.White : Color.FromRgb(30, 30, 30),
                ["MahApps.Brushes.WindowBackground"] = isLight ? Colors.White : Color.FromRgb(30, 30, 30),
                ["MahApps.Brushes.Control.Background"] = isLight ? Colors.White : Color.FromRgb(45, 45, 45),
                ["MahApps.Brushes.Control.Border"] = isLight ? Color.FromRgb(170, 170, 170) : Color.FromRgb(85, 85, 85),

                // HIGHLIGHT BRUSH
                ["MahApps.Brushes.Highlight"] = isLight ? Color.FromRgb(230, 230, 230) : Color.FromRgb(70, 70, 70),
                ["MahApps.Brushes.HighlightText"] = isLight ? Colors.Black : Colors.White,

                // Selection brushes
                ["MahApps.Brushes.Selected"] = isLight ? Color.FromRgb(218, 238, 248) : Color.FromRgb(50, 70, 90),
                ["MahApps.Brushes.SelectedText"] = isLight ? Colors.Black : Colors.White,

                // Add the missing/common keys that some XAML expects
                ["MahApps.Brushes.WindowBorder"] = isLight ? Color.FromRgb(200, 200, 200) : Color.FromRgb(60, 60, 60),
                ["MahApps.Brushes.Control.Foreground"] = isLight ? Colors.Black : Colors.White,
                ["MahApps.Brushes.Text"] = isLight ? Colors.Black : Colors.White,
                ["MahApps.Brushes.DisabledForeground"] = isLight ? Color.FromRgb(142, 142, 142) : Color.FromRgb(105, 105, 105)
            };

            foreach (var brush in essentialBrushes)
            {
                if (!app.Resources.Contains(brush.Key))
                {
                    var solidBrush = CreateFrozenBrush(brush.Value);
                    app.Resources[brush.Key] = solidBrush;
                }

                string colorKey = brush.Key.Replace("Brushes", "Colors");
                if (!app.Resources.Contains(colorKey))
                {
                    app.Resources[colorKey] = brush.Value;
                }
            }
        }

        private static void RegisterCompatibilityPairs(Application app, bool isLight)
        {
            // Ensure some frequently referenced color/brush pairs exist for compatibility
            EnsureResourcePair(app, "MahApps.Colors.AccentForeground", Colors.White);
            EnsureResourcePair(app, "MahApps.Brushes.AccentForeground", Colors.White);

            EnsureResourcePair(app, "MahApps.Brushes.Control.MouseOver", isLight ? Color.FromRgb(245, 245, 245) : Color.FromRgb(55, 55, 55));
            EnsureResourcePair(app, "MahApps.Colors.Highlight", isLight ? Color.FromRgb(230, 230, 230) : Color.FromRgb(70, 70, 70));

            // Provide data-grid specific keys referenced by runtime warnings
            EnsureResourcePair(app, "MahApps.Brushes.DataGrid.Selection.Text.MouseOver", isLight ? Colors.Black : Colors.White);
            EnsureResourcePair(app, "MahApps.Brushes.DataGrid.Selection.Text.Inactive", isLight ? Colors.Black : Colors.White);
            EnsureResourcePair(app, "MahApps.Brushes.DataGrid.Selection.Text", isLight ? Colors.Black : Colors.White);
        }

        private static void EnsureResourcePair(Application app, string key, Color color)
        {
            try
            {
                if (!app.Resources.Contains(key))
                {
                    if (key.Contains(".Brushes."))
                    {
                        var brush = CreateFrozenBrush(color);
                        app.Resources[key] = brush;
                    }
                    else if (key.IndexOf("Brushes", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        var brush = CreateFrozenBrush(color);
                        app.Resources[key] = brush;
                    }
                    else
                    {
                        app.Resources[key] = color;
                    }
                }
                // Also ensure Colors counterpart if missing
                if (key.Contains("Brushes"))
                {
                    var colorKey = key.Replace("Brushes", "Colors");
                    if (!app.Resources.Contains(colorKey))
                        app.Resources[colorKey] = color;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"MahAppsBootstrapper.EnsureResourcePair: {key}");
            }
        }

        private static SolidColorBrush CreateFrozenBrush(Color color)
        {
            var brush = new SolidColorBrush(color);
            if (brush.CanFreeze)
                brush.Freeze();
            return brush;
        }

        private static Color Lighten(Color color, double factor)
        {
            byte r = (byte)Math.Min(255, color.R + (255 - color.R) * factor);
            byte g = (byte)Math.Min(255, color.G + (255 - color.G) * factor);
            byte b = (byte)Math.Min(255, color.B + (255 - color.B) * factor);
            return Color.FromArgb(color.A, r, g, b);
        }

        private static Color Darken(Color color, double factor)
        {
            byte r = (byte)Math.Max(0, color.R * (1 - factor));
            byte g = (byte)Math.Max(0, color.G * (1 - factor));
            byte b = (byte)Math.Max(0, color.B * (1 - factor));
            return Color.FromArgb(color.A, r, g, b);
        }

        // Runtime validator to log missing expected keys
        private static readonly string[] ExpectedKeys = new[]
        {
            "MahApps.Colors.Accent",
            "MahApps.Brushes.Accent",
            "MahApps.Colors.IdealForeground",
            "MahApps.Brushes.IdealForeground",
            "MahApps.Brushes.Highlight",
            "MahApps.Brushes.HighlightText",
            "MahApps.Brushes.Selected",
            "MahApps.Brushes.SelectedText",
            "MahApps.Brushes.Control.Background",
            "MahApps.Brushes.Control.Border",
            "MahApps.Brushes.WindowBackground",
            "MahApps.Brushes.WindowBorder",
            "MahApps.Brushes.Control.Foreground",
            "MahApps.Brushes.Text",
            "MahApps.Brushes.DisabledForeground",
            // Add the data-grid keys to the expected list so missing ones get logged
            "MahApps.Brushes.DataGrid.Selection.Text.MouseOver",
            "MahApps.Brushes.DataGrid.Selection.Text.Inactive",
            "MahApps.Brushes.DataGrid.Selection.Text"
        };

        private static void ValidateThemeResources(Application app)
        {
            if (app == null) return;
            foreach (var key in ExpectedKeys)
            {
                if (!app.Resources.Contains(key))
                {
                    LogUtility.LogWarn($"MahApps resource missing: {key}");
                }
            }
        }
    }
}