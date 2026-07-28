using GLSense.Utilities;
using System;
using System.Reflection;

namespace GLSense.Helpers
{
    internal static class RibbonReflectionHelper
    {
        /// <summary>
        /// Safely attempts to get a ribbon control instance from the AddinModule.
        /// Tries public property, public field, then (as a last resort) non-public field.
        /// Returns null if not found or access is not allowed.
        /// </summary>
        public static object GetRibbonControl(object addinModuleInstance, string controlName)
        {
            if (addinModuleInstance == null || string.IsNullOrWhiteSpace(controlName))
                return null;

            var t = addinModuleInstance.GetType();

            // 1) Try public property
            try
            {
                var prop = t.GetProperty(controlName, BindingFlags.Instance | BindingFlags.Public);
                if (prop != null && prop.CanRead)
                {
                    try
                    {
                        return prop.GetValue(addinModuleInstance, null);
                    }
                    catch (TargetInvocationException ex)
                    {
                        LogUtility.LogDebug($"RibbonReflectionHelper.GetRibbonControl: property getter for '{controlName}' threw ({ex.InnerException?.Message ?? ex.Message}), trying field lookup.");
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"RibbonReflectionHelper.GetRibbonControl: failed reading property '{controlName}' ({ex.Message}), trying field lookup.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"RibbonReflectionHelper.GetRibbonControl: GetProperty('{controlName}') failed ({ex.Message}), trying field lookup.");
            }

            // 2) Try public field
            try
            {
                var field = t.GetField(controlName, BindingFlags.Instance | BindingFlags.Public);
                if (field != null)
                {
                    try
                    {
                        return field.GetValue(addinModuleInstance);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"RibbonReflectionHelper.GetRibbonControl: failed reading public field '{controlName}' ({ex.Message}), trying non-public field lookup.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"RibbonReflectionHelper.GetRibbonControl: GetField('{controlName}') (public) failed ({ex.Message}), trying non-public field lookup.");
            }

            // 3) LAST RESORT: try non-public field (fragile — log and return)
            try
            {
                var nonPublicField = t.GetField(controlName, BindingFlags.Instance | BindingFlags.NonPublic);
                if (nonPublicField != null)
                {
                    try
                    {
                        var val = nonPublicField.GetValue(addinModuleInstance);
                        // optional: only return if not null and type looks like a ribbon control
                        if (val != null)
                            return val;

                        LogUtility.LogWarn($"RibbonReflectionHelper.GetRibbonControl: non-public field '{controlName}' resolved to null.");
                    }
                    catch (FieldAccessException ex)
                    {
                        // Not allowed to access non-public member in this environment
                        LogUtility.LogWarn($"RibbonReflectionHelper.GetRibbonControl: not allowed to access non-public field '{controlName}' ({ex.Message}).");
                        return null;
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogWarn($"RibbonReflectionHelper.GetRibbonControl: failed reading non-public field '{controlName}' ({ex.Message}).");
                    }
                }
                else
                {
                    LogUtility.LogWarn($"RibbonReflectionHelper.GetRibbonControl: control '{controlName}' not found as property, public field, or non-public field on {t.Name}.");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"RibbonReflectionHelper.GetRibbonControl: GetField('{controlName}') (non-public) failed ({ex.Message}).");
            }

            return null;
        }
    }
}
