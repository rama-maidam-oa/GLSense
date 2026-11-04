using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace GLSense.Utilities
{
    /// <summary>
    /// OISR-22830: keeps very large texts (API responses, raw JSON) out of the main log file.
    /// Text over <see cref="ThresholdChars"/> is written in full to the add-in's Temp folder
    /// (GLSense_Logs\Temp), and the main log gets only a short preview plus that file's path, so
    /// the debug log stays readable. <see cref="CleanUp"/> runs at add-in start: it creates the
    /// folder if missing and removes this helper's own files older than <see cref="KeepDays"/> days
    /// (other files in the folder, e.g. refresh working files, are left alone).
    /// </summary>
    public static class LargeLogText
    {
        public const int ThresholdChars = 100 * 1024;
        public const int PreviewChars = 1000;
        public const int KeepDays = 7;

        // File names written by NewFilePath: yyyyMMdd_HHmmss_fff_<name>.<ext>
        private static readonly Regex OwnFileName = new Regex(@"^\d{8}_\d{6}_\d{3}_", RegexOptions.Compiled);

        private static string Folder(string tempFolder) =>
            string.IsNullOrWhiteSpace(tempFolder) ? Path.GetTempPath() : tempFolder;

        /// <summary>New, unique file path in the Temp folder (folder created if needed).</summary>
        public static string NewFilePath(string tempFolder, string nameHint, string extension = ".json")
        {
            string folder = Folder(tempFolder);
            Directory.CreateDirectory(folder);
            return Path.Combine(folder, $"{DateTime.Now:yyyyMMdd_HHmmss_fff}_{SafeName(nameHint)}{extension}");
        }

        /// <summary>
        /// Returns <paramref name="text"/> unchanged when it is small. Otherwise saves it to the
        /// Temp folder and returns a preview plus the saved file's path.
        /// </summary>
        public static string ToLogText(string text, string tempFolder, string nameHint)
        {
            if (text == null || text.Length <= ThresholdChars)
                return text;

            string preview = text.Substring(0, PreviewChars);
            try
            {
                string path = NewFilePath(tempFolder, nameHint);
                File.WriteAllText(path, text);
                return $"{preview}... [{text.Length:N0} chars - full text saved to {path}]";
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                return $"{preview}... [{text.Length:N0} chars - truncated, could not save the full text: {ex.Message}]";
            }
        }

        /// <summary>
        /// Creates the Temp folder if missing and deletes this helper's own files in it that are
        /// older than <paramref name="keepDays"/> days.
        /// </summary>
        public static void CleanUp(string tempFolder, int keepDays = KeepDays)
        {
            try
            {
                string folder = Folder(tempFolder);
                Directory.CreateDirectory(folder);

                DateTime cutoff = DateTime.Now.AddDays(-keepDays);
                foreach (string file in Directory.EnumerateFiles(folder))
                {
                    try
                    {
                        if (OwnFileName.IsMatch(Path.GetFileName(file)) && File.GetLastWriteTime(file) < cutoff)
                            File.Delete(file);
                    }
                    catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                    {
                        // In use or locked - leave it for the next start.
                    }
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // Clean-up is best effort and must never stop the add-in from starting.
            }
        }

        private static string SafeName(string hint)
        {
            if (string.IsNullOrWhiteSpace(hint))
                return "text";

            string name = new string(hint.Select(c => char.IsLetterOrDigit(c) || c == '-' || c == '_' ? c : '_').ToArray()).Trim('_');
            if (name.Length > 60)
                name = name.Substring(0, 60);
            return name.Length == 0 ? "text" : name;
        }
    }
}
