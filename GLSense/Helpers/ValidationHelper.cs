using GLSense.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GLSense.Helpers
{
    /// <summary>
    /// Centralized validation helper to reduce repeated validation code
    /// </summary>
    public static class ValidationHelper
    {
        /// <summary>
        /// Validates that a string is not null or empty
        /// </summary>
        public static bool ValidateNotEmpty(string value, string fieldName, Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope($"ValidateNotEmpty({fieldName})"))
            {
                try
                {
                    bool isValid = !string.IsNullOrWhiteSpace(value);

                    if (!isValid)
                    {
                        string message = $"{fieldName} cannot be empty.";
                        LogUtility.LogWarn($"Validation failed: {message}");
                        showWarning?.Invoke(message);
                    }
                    else
                    {
                        LogUtility.LogDebug($"Validation passed: {fieldName}");
                    }

                    return isValid;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ValidateNotEmpty: {fieldName}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Validates that a collection is not null or empty
        /// </summary>
        public static bool ValidateCollectionNotEmpty<T>(IEnumerable<T> collection, string fieldName, Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope($"ValidateCollectionNotEmpty({fieldName})"))
            {
                try
                {
                    bool isValid = collection != null && collection.Any();

                    if (!isValid)
                    {
                        string message = $"{fieldName} cannot be empty.";
                        LogUtility.LogWarn($"Validation failed: {message}");
                        showWarning?.Invoke(message);
                    }
                    else
                    {
                        int count = collection.Count();
                        LogUtility.LogDebug($"Validation passed: {fieldName} has {count} items");
                    }

                    return isValid;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ValidateCollectionNotEmpty: {fieldName}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Validates that an object is not null
        /// </summary>
        public static bool ValidateNotNull(object value, string fieldName, Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope($"ValidateNotNull({fieldName})"))
            {
                try
                {
                    bool isValid = value != null;

                    if (!isValid)
                    {
                        string message = $"{fieldName} is required.";
                        LogUtility.LogWarn($"Validation failed: {message}");
                        showWarning?.Invoke(message);
                    }
                    else
                    {
                        LogUtility.LogDebug($"Validation passed: {fieldName} is not null");
                    }

                    return isValid;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ValidateNotNull: {fieldName}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Validates that a value is within a range
        /// </summary>
        public static bool ValidateRange(int value, int min, int max, string fieldName, Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope($"ValidateRange({fieldName})"))
            {
                try
                {
                    bool isValid = value >= min && value <= max;

                    if (!isValid)
                    {
                        string message = $"{fieldName} must be between {min} and {max}.";
                        LogUtility.LogWarn($"Validation failed: {message} (Value: {value})");
                        showWarning?.Invoke(message);
                    }
                    else
                    {
                        LogUtility.LogDebug($"Validation passed: {fieldName} = {value} is within [{min}, {max}]");
                    }

                    return isValid;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, $"ValidateRange: {fieldName}");
                    return false;
                }
            }
        }

        /// <summary>
        /// Validates cube and ledger selection
        /// </summary>
        public static bool ValidateCubeAndLedgerSelection(Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope("ValidateCubeAndLedgerSelection"))
            {
                try
                {
                    bool cubeValid = AppState.Instance.SelectedCube != null;
                    bool ledgerValid = AppState.Instance.SelectedLedger != null;

                    LogUtility.LogDebug($"Cube selected: {cubeValid}, Ledger selected: {ledgerValid}");

                    if (!cubeValid)
                    {
                        string message = "Please select a cube first.";
                        LogUtility.LogWarn(message);
                        showWarning?.Invoke(message);
                        return false;
                    }

                    if (!ledgerValid)
                    {
                        string message = "Please select a ledger first.";
                        LogUtility.LogWarn(message);
                        showWarning?.Invoke(message);
                        return false;
                    }

                    LogUtility.LogDebug("Cube and ledger validation passed");
                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "ValidateCubeAndLedgerSelection");
                    return false;
                }
            }
        }

        /// <summary>
        /// Validates login status
        /// </summary>
        public static bool ValidateLoginStatus(Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope("ValidateLoginStatus"))
            {
                try
                {
                    bool isLoggedIn = AppState.Instance.IsLoginCompleted;

                    LogUtility.LogDebug($"Login completed: {isLoggedIn}");

                    if (!isLoggedIn)
                    {
                        string message = "Please log in first.";
                        LogUtility.LogWarn(message);
                        showWarning?.Invoke(message);
                        return false;
                    }

                    LogUtility.LogDebug("Login validation passed");
                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "ValidateLoginStatus");
                    return false;
                }
            }
        }

        /// <summary>
        /// Validates Excel application availability
        /// </summary>
        public static bool ValidateExcelAvailable(Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope("ValidateExcelAvailable"))
            {
                try
                {
                    bool excelAvailable = AppState.Instance.ExcelApp != null && 
                                         ExcelComHelper.IsExcelAppAlive(AppState.Instance.ExcelApp);

                    LogUtility.LogDebug($"Excel available: {excelAvailable}");

                    if (!excelAvailable)
                    {
                        string message = "Excel application is not available.";
                        LogUtility.LogError(message);
                        showWarning?.Invoke(message);
                        return false;
                    }

                    LogUtility.LogDebug("Excel validation passed");
                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "ValidateExcelAvailable");
                    return false;
                }
            }
        }

        /// <summary>
        /// Validates all prerequisites for operations (Excel, Login, Cube, Ledger)
        /// </summary>
        public static bool ValidateAllPrerequisites(Action<string> showWarning = null)
        {
            using (new LogUtility.LogScope("ValidateAllPrerequisites"))
            {
                try
                {
                    LogUtility.LogDebug("Validating all prerequisites");

                    if (!ValidateExcelAvailable(showWarning))
                        return false;

                    if (!ValidateLoginStatus(showWarning))
                        return false;

                    if (!ValidateCubeAndLedgerSelection(showWarning))
                        return false;

                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionHelper.LogDetailedException(ex, "ValidateAllPrerequisites");
                    return false;
                }
            }
        }
    }
}
