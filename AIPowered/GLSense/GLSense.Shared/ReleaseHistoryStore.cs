// ReleaseHistoryStore.cs in GLSense.Shared
//
// Owns all reads/writes of ReleaseHistory.json - the permanent, append-only catalog of
// every Addin.Core release ever adopted on this machine. See
// docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md section 4.
//
// Every read-modify-write cycle (Append, Reconcile) is protected by a named
// cross-process Mutex, so two Excel processes triggering a reload at close to the same
// time can never race a lost update. Every write goes to a temp file first, then
// replaces the real file, so a crash mid-write can never leave a corrupt catalog.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;

namespace GLSense.Shared
{
    public static class ReleaseHistoryStore
    {
        private const string MutexName = "Global\\GLSense_ReleaseHistory_Mutex";
        private static readonly TimeSpan MutexTimeout = TimeSpan.FromSeconds(10);

        public static List<ReleaseEntry> ReadAll(string releaseHistoryFile)
        {
            return WithLock(() => ReadAllUnlocked(releaseHistoryFile));
        }

        /// <summary>Appends one entry. Process-safe (named Mutex) and crash-safe
        /// (atomic write-then-replace).</summary>
        public static void Append(string releaseHistoryFile, ReleaseEntry entry)
        {
            WithLock(() =>
            {
                var entries = ReadAllUnlocked(releaseHistoryFile);
                entries.Add(entry);
                WriteAllUnlocked(releaseHistoryFile, entries);
            });
        }

        /// <summary>Removes every entry whose Versions\{FolderName}\ no longer contains
        /// any .dll file, and returns the surviving list. Cheap - only
        /// Directory.Exists/GetFiles checks, no file content reads. Called (a) when a
        /// reinstall of an already-known release is detected (UpdateBootstrapper), and
        /// (b) every time the Release History browser is opened (Phase C).</summary>
        public static List<ReleaseEntry> Reconcile(string releaseHistoryFile, string versionsPath)
        {
            return WithLock(() =>
            {
                var entries = ReadAllUnlocked(releaseHistoryFile);
                var survivors = entries.Where(e => ReleaseFolderHasDlls(versionsPath, e.FolderName)).ToList();
                if (survivors.Count != entries.Count)
                    WriteAllUnlocked(releaseHistoryFile, survivors);
                return survivors;
            });
        }

        private static T WithLock<T>(Func<T> action)
        {
            using (var mutex = new Mutex(false, MutexName))
            {
                bool acquired;
                try
                {
                    acquired = mutex.WaitOne(MutexTimeout);
                }
                catch (AbandonedMutexException)
                {
                    // A previous holder terminated (e.g. an AppDomain.Unload-aborted thread -
                    // see CLAUDE.md sections 29/37) without releasing the mutex. We now
                    // legitimately own it - the catalog file itself isn't corrupted just
                    // because a prior holder didn't release cleanly (writes are atomic, see
                    // WriteAllUnlocked), so proceed.
                    acquired = true;
                }

                if (!acquired)
                    throw new TimeoutException($"Could not acquire '{MutexName}' within {MutexTimeout} - refusing to read/write ReleaseHistory.json unprotected.");

                try
                {
                    return action();
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
        }

        private static void WithLock(Action action)
        {
            WithLock<object>(() => { action(); return null; });
        }

        /// <summary>Builds the Versions\ folder name for a release:
        /// V{version}_{releaseDateSafe}. Computed exactly once, at extraction time
        /// (UpdateBootstrapper) - every other consumer resolves a release's folder by
        /// reading the stored FolderName from its catalog entry, never by
        /// recomputing this.</summary>
        public static string BuildFolderName(string version, string releaseDate)
        {
            char[] illegal = Path.GetInvalidFileNameChars();
            var safeDate = new string((releaseDate ?? string.Empty)
                .Select(c => illegal.Contains(c) ? '-' : c).ToArray());
            return $"V{version}_{safeDate}";
        }

        private static bool ReleaseFolderHasDlls(string versionsPath, string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName)) return false;
            var folder = Path.Combine(versionsPath, folderName);
            return Directory.Exists(folder) && Directory.GetFiles(folder, "*.dll").Any();
        }

        private static List<ReleaseEntry> ReadAllUnlocked(string releaseHistoryFile)
        {
            if (!File.Exists(releaseHistoryFile)) return new List<ReleaseEntry>();
            var json = File.ReadAllText(releaseHistoryFile);
            if (string.IsNullOrWhiteSpace(json)) return new List<ReleaseEntry>();
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            return JsonSerializer.Deserialize<List<ReleaseEntry>>(json, options) ?? new List<ReleaseEntry>();
        }

        private static void WriteAllUnlocked(string releaseHistoryFile, List<ReleaseEntry> entries)
        {
            var directory = Path.GetDirectoryName(releaseHistoryFile);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(entries, options);

            string tempFile = releaseHistoryFile + ".tmp";
            File.WriteAllText(tempFile, json);

            if (File.Exists(releaseHistoryFile))
                File.Replace(tempFile, releaseHistoryFile, null);
            else
                File.Move(tempFile, releaseHistoryFile);
        }
    }
}
