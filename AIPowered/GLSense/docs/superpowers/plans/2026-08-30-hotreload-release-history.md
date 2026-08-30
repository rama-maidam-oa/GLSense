# Hot-reload Release History Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace `RibReload`'s single "reload from Manifest folder" action with an
Online/Offline picker, and add a permanent, browsable catalog of every Addin.Core
release ever adopted on this machine (`ReleaseHistory.json`) so a tester can jump to any
past release without an MSI reinstall.

**Architecture:** `Versions\` folders are re-keyed by the release's full timestamp (not
version number alone, since version numbers can repeat across distinct releases).
`ReleaseHistoryStore` (new, in `GLSense.Shared`) owns all reads/writes of
`ReleaseHistory.json` behind a named cross-process `Mutex` with atomic
write-then-replace. `UpdateBootstrapper` (in `GLSense.Loader.Core`) is extended to
detect a genuinely first-ever run (catalog absent), reconcile stale entries on a
detected reinstall, and append a catalog entry on every extract. Two new WPF windows
live in the `GLSense` host project (never in `GLSense.Addin.Core`, since that assembly
is exactly what's being replaced): a reload-source picker (Online/Offline) and a
release-history browser. A new `IGLSenseAddin.GetLoginInfo()` cross-domain method lets
the host read login state for Online mode without `AppState` itself crossing the
AppDomain boundary.

**Tech Stack:** .NET Framework 4.8.1, WPF (new to the `GLSense` host project — it is
otherwise plain WinForms/Add-in Express), `System.Text.Json` (already referenced by
`GLSense.Shared`), .NET Remoting (existing AppDomain/`MarshalByRefObject` boundary), no
automated test framework anywhere in this solution.

**Spec:** `docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md`

## Global Constraints

- `ReleaseHistory.json` writes MUST be atomic (write to a `.tmp` file, then
  `File.Replace`/`File.Move`) and MUST be protected by a named system `Mutex`
  (`Global\GLSense_ReleaseHistory_Mutex`) around every read-modify-write cycle.
- `Versions\{folderName}\` folders are permanent once created — nothing in this plan
  deletes or overwrites one except the explicit fresh-install wipe path (spec §5) and
  the idempotent re-extract-of-an-identical-release case (spec §5.2).
- `Versions\` folder naming: `V{version}_{releaseDateSafe}`, where `releaseDateSafe` is
  `releaseDate` with every character in `Path.GetInvalidFileNameChars()` replaced by
  `-`. Computed exactly once, at extraction time — never recomputed elsewhere.
- Every `IGLSenseAddin` addition (this plan adds one: `GetLoginInfo()`) MUST be called
  from the host wrapped in `try`/`catch`, tolerating the call failing entirely — an
  older historical build reloaded via the Release History browser may not implement it.
  `IGLSenseContext` additions (this plan adds `ActiveFolderName`) do NOT need this
  treatment — see spec §9 for why the risk is directional.
- New host-side WPF windows (`GLReloadSourcePicker`, `GLReleaseHistoryBrowser`) live in
  the `GLSense` project, in `namespace GLSense` (flat, matching every other file in that
  project — do not introduce a `GLSense.Views` namespace even though the files live
  under a `Views\` folder), and must not reference `GLSense.Addin.Core`.
- RibReload's Online/Offline gate: reject a candidate release that is not strictly
  newer (by `releaseDate`, parsed as `DateTime`) than `GlobalsEx.Context.ReleaseDate`.
  The Release History browser (Phase C) has **no such gate** — loading an older release
  on purpose is its entire reason to exist.
- No automated test project exists in this solution. Every task's verification step is
  a build (via `msbuild`, run from PowerShell) plus a manual check described in the
  task. If `msbuild`/the Windows SDK is unavailable in whatever environment executes
  this plan, say so explicitly rather than claiming a build passed.

---

## Phase A — Release history data model & fresh-install detection

### Task A1: `ReleaseEntry` + `ReleaseHistoryStore` (new, in `GLSense.Shared`)

**Files:**
- Create: `GLSense.Shared\ReleaseEntry.cs`
- Create: `GLSense.Shared\ReleaseHistoryStore.cs`
- Modify: `GLSense.Shared\GLSense.Shared.csproj` (add two `<Compile>` items)

**Interfaces:**
- Produces: `GLSense.Shared.ReleaseEntry` (`Version`, `ReleaseDate`, `FolderName`,
  `Checksum`, `Notes`, `Source` — all `string`), `GLSense.Shared.ReleaseHistoryStore`
  static methods: `ReadAll(string releaseHistoryFile) : List<ReleaseEntry>`,
  `Append(string releaseHistoryFile, ReleaseEntry entry) : void`,
  `Reconcile(string releaseHistoryFile, string versionsPath) : List<ReleaseEntry>`,
  `BuildFolderName(string version, string releaseDate) : string`.

- [ ] **Step 1: Create `ReleaseEntry.cs`**

```csharp
// ReleaseEntry.cs in GLSense.Shared
using System;

namespace GLSense.Shared
{
    // [Serializable] even though this doesn't currently cross the AppDomain boundary -
    // it's read/written purely host-side and inside UpdateBootstrapper (also host-side,
    // GLSense.Loader.Core) - kept Serializable anyway for consistency with VersionInfo
    // and in case a future call site needs to hand one across.
    [Serializable]
    public class ReleaseEntry
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string FolderName { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }

        /// <summary>"Install" (MSI-seeded first run, or an ordinary local dev-loop
        /// rebuild picked up automatically at Excel startup), "Online", or "Offline".
        /// See docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md
        /// section 4.</summary>
        public string Source { get; set; }
    }
}
```

- [ ] **Step 2: Create `ReleaseHistoryStore.cs`**

```csharp
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
            using (var mutex = new Mutex(false, MutexName))
            {
                bool acquired = false;
                try
                {
                    acquired = mutex.WaitOne(MutexTimeout);
                    return ReadAllUnlocked(releaseHistoryFile);
                }
                finally
                {
                    if (acquired) mutex.ReleaseMutex();
                }
            }
        }

        /// <summary>Appends one entry. Process-safe (named Mutex) and crash-safe
        /// (atomic write-then-replace).</summary>
        public static void Append(string releaseHistoryFile, ReleaseEntry entry)
        {
            using (var mutex = new Mutex(false, MutexName))
            {
                bool acquired = false;
                try
                {
                    acquired = mutex.WaitOne(MutexTimeout);
                    var entries = ReadAllUnlocked(releaseHistoryFile);
                    entries.Add(entry);
                    WriteAllUnlocked(releaseHistoryFile, entries);
                }
                finally
                {
                    if (acquired) mutex.ReleaseMutex();
                }
            }
        }

        /// <summary>Removes every entry whose Versions\{FolderName}\ no longer contains
        /// any .dll file, and returns the surviving list. Cheap - only
        /// Directory.Exists/GetFiles checks, no file content reads. Called (a) when a
        /// reinstall of an already-known release is detected (UpdateBootstrapper), and
        /// (b) every time the Release History browser is opened (Phase C).</summary>
        public static List<ReleaseEntry> Reconcile(string releaseHistoryFile, string versionsPath)
        {
            using (var mutex = new Mutex(false, MutexName))
            {
                bool acquired = false;
                try
                {
                    acquired = mutex.WaitOne(MutexTimeout);
                    var entries = ReadAllUnlocked(releaseHistoryFile);
                    var survivors = entries.Where(e => ReleaseFolderHasDlls(versionsPath, e.FolderName)).ToList();
                    if (survivors.Count != entries.Count)
                        WriteAllUnlocked(releaseHistoryFile, survivors);
                    return survivors;
                }
                finally
                {
                    if (acquired) mutex.ReleaseMutex();
                }
            }
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
```

- [ ] **Step 3: Add both files to `GLSense.Shared.csproj`**

Open `GLSense.Shared\GLSense.Shared.csproj`, find the `<ItemGroup>` containing
`<Compile Include="PathProvider.cs" />` (or similar existing `<Compile>` entries), and
add two sibling lines:

```xml
    <Compile Include="ReleaseEntry.cs" />
    <Compile Include="ReleaseHistoryStore.cs" />
```

- [ ] **Step 4: Build to verify**

Run (PowerShell, from the solution root `D:\SQLLite_Test\GLSense\AIPowered\GLSense`):

```powershell
msbuild GLSense.sln /t:GLSense_Shared /p:Configuration=Debug
```

Expected: `Build succeeded.` If `msbuild` isn't on PATH, locate it via
`vswhere.exe` (usually under `C:\Program Files (x86)\Microsoft Visual
Studio\Installer\vswhere.exe -latest -find **\MSBuild.exe`) and use its full path. If no
MSBuild/Visual Studio toolchain is available in the environment executing this task,
state that plainly instead of claiming the build passed.

- [ ] **Step 5: Commit**

```bash
git add GLSense.Shared/ReleaseEntry.cs GLSense.Shared/ReleaseHistoryStore.cs GLSense.Shared/GLSense.Shared.csproj
git commit -m "Add ReleaseEntry/ReleaseHistoryStore: the release-history catalog data model"
```

---

### Task A2: `IPathProvider.ReleaseHistoryFile`

**Files:**
- Modify: `GLSense.Contracts\IPathProvider.cs`
- Modify: `GLSense.Shared\PathProvider.cs`

**Interfaces:**
- Consumes: none new.
- Produces: `IPathProvider.ReleaseHistoryFile : string` — full path to
  `ReleaseHistory.json`, a sibling of `Manifest\` and `Versions\` (not inside either).

- [ ] **Step 1: Add the property to the interface**

In `GLSense.Contracts\IPathProvider.cs`, add one line alongside the existing
`ManifestFile`/`ManifestDirectory` declarations:

```csharp
        string ManifestFile { get; }
        string ManifestDirectory { get; }
        string ReleaseHistoryFile { get; }
```

- [ ] **Step 2: Implement it in `PathProvider`**

In `GLSense.Shared\PathProvider.cs`, add one line alongside the existing
`ManifestDirectory`/`ManifestFile` properties:

```csharp
        public string ManifestDirectory => Path.Combine(_root, "Manifest");
        public string ManifestFile => Path.Combine(ManifestDirectory, "manifest.json");
        public string ReleaseHistoryFile => Path.Combine(_root, "ReleaseHistory.json");
```

- [ ] **Step 3: Build to verify**

```powershell
msbuild GLSense.sln /t:GLSense_Shared;GLSense_Contracts /p:Configuration=Debug
```

Expected: `Build succeeded.`

- [ ] **Step 4: Commit**

```bash
git add GLSense.Contracts/IPathProvider.cs GLSense.Shared/PathProvider.cs
git commit -m "Add IPathProvider.ReleaseHistoryFile path"
```

---

### Task A3: `ResolvedRelease` + rewritten `UpdateBootstrapper` (folder keying, fresh-install detection, reconciliation, catalog append)

**Files:**
- Create: `GLSense.Loader.Core\ResolvedRelease.cs`
- Modify: `GLSense.Loader.Core\UpdateBootstrapper.cs` (full rewrite of
  `ResolveVersionToLoad`; `ExtractLocalZipAndAdopt` renamed to
  `ExtractManifestZipAndAdopt` with new behavior)
- Modify: `GLSense.Loader.Core\GLSense.Loader.Core.csproj` (add one `<Compile>` item)

**Interfaces:**
- Consumes: `GLSense.Shared.ReleaseHistoryStore` (Task A1),
  `IPathProvider.ReleaseHistoryFile` (Task A2).
- Produces: `GLSense.Loader.Core.ResolvedRelease` (`Version`, `ReleaseDate`,
  `FolderName` — all `string`), `UpdateBootstrapper.ResolveVersionToLoad(IGLSenseContext
  context, string source = "Install") : ResolvedRelease` (was `: string`, and had no
  `source` parameter — every existing call site must be updated, see Task A4/B4).

- [ ] **Step 1: Create `ResolvedRelease.cs`**

```csharp
// ResolvedRelease.cs in GLSense.Loader.Core
namespace GLSense.Loader.Core
{
    /// <summary>
    /// Identifies exactly which Addin.Core release to load: FolderName is the only
    /// thing AddinDomainLoader actually needs (Versions\{FolderName}\ is where the DLLs
    /// live); Version/ReleaseDate are kept for display (GLAbout, log lines, ribbon
    /// messages). See
    /// docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md section 3.1.
    /// </summary>
    public class ResolvedRelease
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string FolderName { get; set; }
    }
}
```

- [ ] **Step 2: Rewrite `UpdateBootstrapper.cs`**

Replace the entire file with:

```csharp
// UpdateBootstrapper.cs in GLSense.Loader.Core
//
// Pre-AppDomain-load bootstrap: decides which release of GLSense.Addin.Core to load,
// gets it onto disk if it isn't there yet, and maintains ReleaseHistory.json - the
// permanent catalog of every release ever adopted on this machine. See
// docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md.
//
// Decision tree:
//   1. ReleaseHistory.json does not exist -> first-ever run on this machine. If
//      Manifest\ has both manifest.json and a zip (the MSI's bundled seed), wipe any
//      stray Versions\ content, extract+catalog it (source "Install"), delete the
//      Manifest folder, done. If no seed is present, fall through to step 2/3 with an
//      empty catalog.
//   2. ReleaseHistory.json exists and Manifest\ has both manifest.json and a zip:
//      compare the manifest's version+releaseDate against every existing catalog
//      entry. An exact match means this is a reinstall of an already-known release -
//      reconcile the catalog (drop entries whose folder no longer exists) before
//      extracting. Either way, extract+catalog normally (the caller's `source`
//      parameter is used as-is - see ResolveVersionToLoad's own doc comment).
//   3. No zip in Manifest\, but Versions\{FolderName}\ for the currently active
//      release already has DLLs -> reuse it, nothing to do.
//   4. Nothing usable anywhere -> return null so the caller can skip loading the
//      AppDomain instead of crashing Excel.
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace GLSense.Loader.Core
{
    public class UpdateBootstrapper
    {
        /// <summary>
        /// Resolves which release to load, extracting/cataloguing a new one if
        /// Manifest\ has a zip waiting. `source` records WHY this call is happening -
        /// "Install" (default) for the automatic Excel-startup path and for an ordinary
        /// local dev-loop rebuild picked up there, or "Online"/"Offline" when this is
        /// called after RibReload's picker window staged a validated release into
        /// Manifest\. Never pass "Online"/"Offline" from the startup path.
        /// </summary>
        public ResolvedRelease ResolveVersionToLoad(IGLSenseContext context, string source = "Install")
        {
            var logger = context.Logger;
            var paths = context.Paths;

            try
            {
                paths.Refresh();

                bool catalogExists = File.Exists(paths.ReleaseHistoryFile);

                if (!catalogExists)
                {
                    logger?.LogDebug("UpdateBootstrapper: ReleaseHistory.json does not exist - treating this as the first-ever run on this machine.");

                    if (Directory.Exists(paths.ManifestDirectory) &&
                        File.Exists(paths.ManifestFile) &&
                        Directory.GetFiles(paths.ManifestDirectory, "*.zip").Any())
                    {
                        if (Directory.Exists(paths.VersionsPath))
                        {
                            logger?.LogDebug($"UpdateBootstrapper: wiping stray Versions\\ content before first-ever seed ('{paths.VersionsPath}').");
                            Directory.Delete(paths.VersionsPath, true);
                        }

                        var seeded = ExtractManifestZipAndAdopt(context, "Install");
                        DeleteManifestFolder(context);
                        return seeded;
                    }

                    logger?.LogDebug("UpdateBootstrapper: no seed manifest+zip found on first-ever run - falling through with an empty catalog.");
                }

                if (Directory.Exists(paths.ManifestDirectory) && File.Exists(paths.ManifestFile))
                {
                    string zipPath = Directory.GetFiles(paths.ManifestDirectory, "*.zip").FirstOrDefault();
                    if (zipPath != null)
                    {
                        string candidateVersion = paths.LatestVersion;
                        string candidateReleaseDate = paths.LatestReleaseDate;

                        var existing = ReleaseHistoryStore.ReadAll(paths.ReleaseHistoryFile);
                        bool isKnownReinstall = existing.Any(e =>
                            string.Equals(e.Version, candidateVersion, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(e.ReleaseDate, candidateReleaseDate, StringComparison.OrdinalIgnoreCase));

                        if (isKnownReinstall)
                        {
                            logger?.LogDebug($"UpdateBootstrapper: manifest in Manifest\\ matches an existing catalog entry (version={candidateVersion}, releaseDate={candidateReleaseDate}) - reconciling before extracting.");
                            ReleaseHistoryStore.Reconcile(paths.ReleaseHistoryFile, paths.VersionsPath);
                        }

                        return ExtractManifestZipAndAdopt(context, source);
                    }
                }

                string localVersion = paths.LatestVersion;
                string localReleaseDate = paths.LatestReleaseDate;
                string folderName = ReleaseHistoryStore.BuildFolderName(localVersion, localReleaseDate);
                string versionFolder = Path.Combine(paths.VersionsPath, folderName);
                bool haveLocalDlls = Directory.Exists(versionFolder) && Directory.GetFiles(versionFolder, "*.dll").Any();

                if (haveLocalDlls)
                {
                    logger?.LogDebug($"UpdateBootstrapper: no zip present, but '{versionFolder}' already has DLLs - using installed version '{localVersion}'.");
                    return new ResolvedRelease { Version = localVersion, ReleaseDate = localReleaseDate, FolderName = folderName };
                }

                logger?.LogError($"UpdateBootstrapper: no zip in '{paths.ManifestDirectory}' and no usable install at '{versionFolder}' - nothing to load.");
                return null;
            }
            catch (Exception ex)
            {
                logger?.LogException(ex, "UpdateBootstrapper.ResolveVersionToLoad");
                return null;
            }
        }

        private ResolvedRelease ExtractManifestZipAndAdopt(IGLSenseContext context, string source)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            string version = paths.LatestVersion;
            string releaseDate = paths.LatestReleaseDate;
            string folderName = ReleaseHistoryStore.BuildFolderName(version, releaseDate);
            string versionFolder = Path.Combine(paths.VersionsPath, folderName);
            string zipPath = Directory.GetFiles(paths.ManifestDirectory, "*.zip").First();

            logger?.LogDebug($"UpdateBootstrapper: extracting '{zipPath}' into '{versionFolder}' (source={source}).");

            if (Directory.Exists(versionFolder))
                Directory.Delete(versionFolder, true);
            Directory.CreateDirectory(versionFolder);

            ZipFile.ExtractToDirectory(zipPath, versionFolder);
            File.Delete(zipPath);

            // Per-version manifest snapshot - a permanent, self-contained record of
            // exactly what this folder is, independent of the transient copy in
            // Manifest\ (which the caller may delete afterward - e.g. the fresh-install
            // path).
            File.Copy(paths.ManifestFile, Path.Combine(versionFolder, "manifest.json"), true);

            var entry = new ReleaseEntry
            {
                Version = version,
                ReleaseDate = releaseDate,
                FolderName = folderName,
                Checksum = paths.LatestChecksum,
                Notes = string.IsNullOrWhiteSpace(paths.LatestNotes) ? "Published by GLSense.Addin.Core" : paths.LatestNotes,
                Source = source
            };
            ReleaseHistoryStore.Append(paths.ReleaseHistoryFile, entry);

            logger?.LogDebug($"UpdateBootstrapper: extracted, catalogued (source={source}), and deleted '{zipPath}'. Adopting '{folderName}'.");

            return new ResolvedRelease { Version = version, ReleaseDate = releaseDate, FolderName = folderName };
        }

        private void DeleteManifestFolder(IGLSenseContext context)
        {
            try
            {
                if (Directory.Exists(context.Paths.ManifestDirectory))
                    Directory.Delete(context.Paths.ManifestDirectory, true);
            }
            catch (Exception ex)
            {
                context.Logger?.LogException(ex, "UpdateBootstrapper.DeleteManifestFolder");
            }
        }
    }
}
```

- [ ] **Step 3: Add `ResolvedRelease.cs` to `GLSense.Loader.Core.csproj`**

Add one line alongside the existing `<Compile>` entries (e.g. next to
`AddinDomainLoader.cs`'s entry):

```xml
    <Compile Include="ResolvedRelease.cs" />
```

Also confirm `GLSense.Loader.Core.csproj` has a `<ProjectReference>` to
`GLSense.Shared.csproj` (it already does — `GLSense.Loader.Core` already depends on
`GLSense.Shared` for `VersionParser` per `CLAUDE.md` section 15). No new reference
needed for `ReleaseHistoryStore`/`ReleaseEntry`.

- [ ] **Step 4: Build to verify**

**Correction found during execution** (`AddinDomainLoader.cs` does not actually call
`ResolveVersionToLoad` at all — only `AddinModule.cs` in the separate `GLSense` host
project does): the command that actually exercises the expected failure must target the
host project, not `GLSense_Loader_Core`:

```powershell
msbuild GLSense.sln /t:GLSense /p:Configuration=Debug
```

Expected: build FAILS at this point with two `CS0029` errors in `GLSense\AddinModule.cs`
(the startup call site and `ReloadAddinCore`'s call site, both still expecting the old
`string`-returning signature) — that's expected and gets fixed in Task A4 (startup site)
and Task B4 (`ReloadAddinCore`). Confirm the failure is specifically these two
`ResolvedRelease`-to-`string` conversion errors, not a syntax error in the files just
written. Separately, `msbuild GLSense.Loader.Core\GLSense.Loader.Core.csproj
/p:Configuration=Debug` (this task's own project, in isolation) should succeed cleanly.

- [ ] **Step 5: Commit**

```bash
git add GLSense.Loader.Core/ResolvedRelease.cs GLSense.Loader.Core/UpdateBootstrapper.cs GLSense.Loader.Core/GLSense.Loader.Core.csproj
git commit -m "Rewrite UpdateBootstrapper: timestamp-based folder keying, fresh-install detection, catalog append"
```

---

### Task A4: Thread `ResolvedRelease`/`ActiveFolderName` through `AddinDomainLoader`, `IGLSenseContext`, and the host's startup call site

**Files:**
- Modify: `GLSense.Contracts\IGLSenseContext.cs`
- Modify: `GLSense\GLSenseContext.cs`
- Modify: `GLSense.Loader.Core\AddinDomainLoader.cs`
- Modify: `GLSense\AddinModule.cs` (startup call site only — `ReloadAddinCore`'s call
  site is handled in Task B4, since it also changes shape there)

**Interfaces:**
- Consumes: `GLSense.Loader.Core.ResolvedRelease` (Task A3).
- Produces: `IGLSenseContext.ActiveFolderName : string { get; set; }`.

- [ ] **Step 1: Add `ActiveFolderName` to the interface**

In `GLSense.Contracts\IGLSenseContext.cs`:

```csharp
        // Version information
        string Version { get; set; }
        string ReleaseDate { get; set; }  // ✅ Added this
        string ActiveFolderName { get; set; } // Versions\{ActiveFolderName}\ is where the currently loaded release's DLLs live - see ResolvedRelease.
        IReadOnlyList<VersionInfo> AllVersions { get; }
```

- [ ] **Step 2: Implement it in `GLSenseContext`**

In `GLSense\GLSenseContext.cs`, alongside the existing `Version`/`ReleaseDate`
auto-properties:

```csharp
        // Version properties - get from PathProvider
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string ActiveFolderName { get; set; }
```

- [ ] **Step 3: Update `AddinDomainLoader.Load()` to use it**

In `GLSense.Loader.Core\AddinDomainLoader.cs`, replace:

```csharp
            string versionFolderName = $"V{_context.Version}";
            string dllPath = Path.Combine(_context.Paths.VersionsPath, versionFolderName);
```

with:

```csharp
            string dllPath = Path.Combine(_context.Paths.VersionsPath, _context.ActiveFolderName);
```

(The `_context.Logger?.LogDebug(...)` line immediately below already references
`dllPath` — leave it as-is.)

- [ ] **Step 4: Update `AddinModule_OnRibbonLoaded`'s call site**

In `GLSense\AddinModule.cs`, find:

```csharp
            GlobalsEx.Context.Logger?.LogDebug("AddinModule_OnRibbonLoaded: resolving version to load via UpdateBootstrapper.");
            string resolvedVersion = new UpdateBootstrapper().ResolveVersionToLoad(GlobalsEx.Context);

            if (string.IsNullOrEmpty(resolvedVersion))
            {
```

Replace with:

```csharp
            GlobalsEx.Context.Logger?.LogDebug("AddinModule_OnRibbonLoaded: resolving version to load via UpdateBootstrapper.");
            var resolved = new UpdateBootstrapper().ResolveVersionToLoad(GlobalsEx.Context);

            if (resolved == null)
            {
```

A few lines further down, find:

```csharp
            GlobalsEx.Context.ReleaseDate = GlobalsEx.Context.Paths?.LatestReleaseDate;
            GlobalsEx.Context.Logger?.LogDebug($"AddinModule_OnRibbonLoaded: resolvedVersion={resolvedVersion}, releaseDate={GlobalsEx.Context.ReleaseDate}");
```

Replace with:

```csharp
            GlobalsEx.Context.Version = resolved.Version;
            GlobalsEx.Context.ReleaseDate = resolved.ReleaseDate;
            GlobalsEx.Context.ActiveFolderName = resolved.FolderName;
            GlobalsEx.Context.Logger?.LogDebug($"AddinModule_OnRibbonLoaded: version={resolved.Version}, releaseDate={resolved.ReleaseDate}, folderName={resolved.FolderName}");
```

(There may be an existing `GlobalsEx.Context.Version = resolvedVersion;` line right
before the `ReleaseDate` line above — remove it, since the block above now sets both.)

Leave `ReloadAddinCore`'s own call site untouched here — Task B4 rewrites that whole
method.

- [ ] **Step 5: Build to verify**

`ReloadAddinCore` elsewhere in `AddinModule.cs` still calls `ResolveVersionToLoad` as if
it returned `string` (untouched by this task — Task B4 rewrites it), so a plain
solution-wide build is still expected to fail at that one remaining call site. To verify
this task's own 4 files are correct in isolation, temporarily comment out the body of
`ReloadAddinCore` (e.g. replace it with `throw new NotImplementedException();`), then run:

```powershell
msbuild GLSense.sln /t:GLSense_Loader_Core;GLSense /p:Configuration=Debug
```

Expected: `Build succeeded.` with `ReloadAddinCore` stubbed this way. Then restore
`ReloadAddinCore`'s real body exactly as it was before committing — the committed state
legitimately still fails to build solution-wide until Task B4 lands.

- [ ] **Step 6: Manual verification (requires a Windows machine with Excel + this
  solution built)**

1. Delete `%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs_New\ReleaseHistory.json` if it
   exists (simulating a machine that predates this feature).
2. Rebuild `GLSense.Addin.Core` (Release, so `post_build.cmd` publishes a fresh
   manifest+zip into `Manifest\`).
3. Launch Excel. Confirm in the logs that `AddinModule_OnRibbonLoaded` resolves a
   release, and that `%LOCALAPPDATA%\...\GLSense_Logs_New\ReleaseHistory.json` now
   exists with exactly one entry, `"source": "Install"`.
4. Confirm `Versions\` now contains a folder named `V{version}_{releaseDateSafe}` (not
   the old `V{version}`-only shape), and that folder contains a `manifest.json`
   snapshot in addition to the DLLs.
5. Confirm the add-in's ribbon loads normally (About window shows the right
   version/build date).

- [ ] **Step 7: Commit**

```bash
git add GLSense.Contracts/IGLSenseContext.cs GLSense/GLSenseContext.cs GLSense.Loader.Core/AddinDomainLoader.cs GLSense/AddinModule.cs
git commit -m "Thread ActiveFolderName through IGLSenseContext/AddinDomainLoader/startup path"
```

---

## Phase B — RibReload Online/Offline picker

### Task B1: `LoginInfo` DTO + `IGLSenseAddin.GetLoginInfo()`

**Files:**
- Create: `GLSense.Contracts\LoginInfo.cs`
- Modify: `GLSense.Contracts\IGLSenseAddin.cs`
- Modify: `GLSense.Addin.Core\AddinEntry.cs`
- Modify: `GLSense.Contracts\GLSense.Contracts.csproj` (add one `<Compile>` item)

**Interfaces:**
- Produces: `GLSense.Contracts.LoginInfo` (`LoginUrl`, `LoginToken` — `string`;
  `IsLoggedIn` — `bool`), `IGLSenseAddin.GetLoginInfo() : LoginInfo`.

- [ ] **Step 1: Create `LoginInfo.cs`**

```csharp
// LoginInfo.cs in GLSense.Contracts
using System;

namespace GLSense.Contracts
{
    // [Serializable] because this crosses the host<->Addin.Core AppDomain boundary as
    // the return value of IGLSenseAddin.GetLoginInfo().
    [Serializable]
    public class LoginInfo
    {
        public string LoginUrl { get; set; }
        public string LoginToken { get; set; }
        public bool IsLoggedIn { get; set; }
    }
}
```

- [ ] **Step 2: Add the method to `IGLSenseAddin`**

In `GLSense.Contracts\IGLSenseAddin.cs`, add before the closing brace of the interface:

```csharp
        /// <summary>
        /// Returns the current login state (LoginUrl/LoginToken/IsLoggedIn) so the
        /// host's RibReload picker (GLReloadSourcePicker, host-side, no dependency on
        /// this project) can build the Online "check for update" request without
        /// AppState (which lives entirely in this project) crossing the AppDomain
        /// boundary directly.
        ///
        /// Added after the very first shipped version of this interface - any caller
        /// MUST wrap this call in try/catch (or otherwise tolerate a
        /// MissingMethodException/RemotingException), because the Release History
        /// browser can reload an OLDER historical build of GLSense.Addin.Core whose own
        /// compiled copy of this interface predates this member. Treat any failure
        /// exactly like "not logged in". See
        /// docs/superpowers/specs/2026-08-30-hotreload-release-history-design.md
        /// section 9.
        /// </summary>
        LoginInfo GetLoginInfo();
```

- [ ] **Step 3: Implement it in `AddinEntry`**

In `GLSense.Addin.Core\AddinEntry.cs`, add a new method (anywhere alongside the other
simple `IGLSenseAddin` implementations, e.g. near `Shutdown()`):

```csharp
        /// <summary>IGLSenseAddin.GetLoginInfo() - see that interface member's own doc
        /// comment.</summary>
        public LoginInfo GetLoginInfo()
        {
            return new LoginInfo
            {
                LoginUrl = AppState.Instance.LoginUrl,
                LoginToken = AppState.Instance.LoginToken,
                IsLoggedIn = AppState.Instance.IsLoggedIn
            };
        }
```

- [ ] **Step 4: Add `LoginInfo.cs` to `GLSense.Contracts.csproj`**

```xml
    <Compile Include="LoginInfo.cs" />
```

- [ ] **Step 5: Build to verify**

```powershell
msbuild GLSense.sln /t:GLSense_Contracts;GLSense_Addin_Core /p:Configuration=Debug
```

Expected: `Build succeeded.`

- [ ] **Step 6: Commit**

```bash
git add GLSense.Contracts/LoginInfo.cs GLSense.Contracts/IGLSenseAddin.cs GLSense.Addin.Core/AddinEntry.cs GLSense.Contracts/GLSense.Contracts.csproj
git commit -m "Add IGLSenseAddin.GetLoginInfo() for the host-side reload picker's Online mode"
```

---

### Task B2: Enable WPF in the `GLSense` host project + create `GLReloadSourcePicker`

**Files:**
- Modify: `GLSense\GLSense.csproj` (add 4 references + 2 items)
- Create: `GLSense\Views\GLReloadSourcePicker.xaml`
- Create: `GLSense\Views\GLReloadSourcePicker.xaml.cs`

**Interfaces:**
- Consumes: `GLSense.Shared.VersionParser`/`VersionParseResult` (existing),
  `GLSense.Contracts.LoginInfo` (Task B1), `GlobalsEx.Addin`/`GlobalsEx.Context`
  (existing, in `namespace GLSense`).
- Produces: `GLSense.GLReloadSourcePicker` — a `Window` with `SelectedSource : string`
  (`"Online"` or `"Offline"`, valid only when `ShowDialog()` returned `true`).

- [ ] **Step 1: Add WPF references and file items to `GLSense.csproj`**

In the `<ItemGroup>` containing the existing `<Reference Include="System.Windows.Forms"
/>` line, add:

```xml
    <Reference Include="WindowsBase" />
    <Reference Include="PresentationCore" />
    <Reference Include="PresentationFramework" />
    <Reference Include="System.Xaml">
      <RequiredTargetFramework>4.0</RequiredTargetFramework>
    </Reference>
```

In the `<ItemGroup>` containing `<Compile Include="AddinModule.cs">`, add:

```xml
    <Compile Include="Views\GLReloadSourcePicker.xaml.cs">
      <DependentUpon>GLReloadSourcePicker.xaml</DependentUpon>
    </Compile>
```

In a new (or existing) `<ItemGroup>`, add the `Page` item (mirroring
`GLSense.Addin.Core.csproj`'s exact pattern for its own windows):

```xml
  <ItemGroup>
    <Page Include="Views\GLReloadSourcePicker.xaml">
      <SubType>Designer</SubType>
      <Generator>MSBuild:Compile</Generator>
    </Page>
  </ItemGroup>
```

Do **not** add `<ProjectTypeGuids>` — see the Global Constraints and spec §7 for why.

- [ ] **Step 2: Create `Views\GLReloadSourcePicker.xaml`**

```xml
<Window x:Class="GLSense.GLReloadSourcePicker"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Reload GLSense Add-in"
        Width="520" Height="440"
        SizeToContent="Manual"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#FFF5F7FA">
    <Window.Resources>
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Foreground" Value="#FF1565C0"/>
            <Setter Property="Margin" Value="0,0,0,6"/>
        </Style>
        <Style x:Key="ActionButton" TargetType="Button">
            <Setter Property="Padding" Value="16,6"/>
            <Setter Property="Margin" Value="6,0,0,0"/>
            <Setter Property="Background" Value="#FF1565C0"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Style="{StaticResource SectionHeader}" Text="Reload GLSense.Addin.Core"/>
        <TextBlock Grid.Row="1" TextWrapping="Wrap" Margin="0,0,0,12" Foreground="#FF555555"
                   Text="This reloads GLSense.Addin.Core.dll from disk without restarting Excel. Any drilldown/refresh/snapshot currently in progress will be interrupted."/>

        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,12">
            <RadioButton x:Name="RbOnline" Content="Online" GroupName="Mode" Margin="0,0,20,0" Checked="Mode_Checked"/>
            <RadioButton x:Name="RbOffline" Content="Offline" GroupName="Mode" Checked="Mode_Checked"/>
        </StackPanel>

        <Grid Grid.Row="3">
            <StackPanel x:Name="OnlinePanel" Visibility="Collapsed">
                <Button x:Name="BtnCheckOnline" Content="Check for Update" HorizontalAlignment="Left"
                        Padding="10,4" Click="BtnCheckOnline_Click"/>
                <ProgressBar x:Name="OnlineProgress" Height="6" Margin="0,10,0,0" IsIndeterminate="True" Visibility="Collapsed"/>
            </StackPanel>

            <StackPanel x:Name="OfflinePanel" Visibility="Collapsed">
                <TextBlock Text="Folder:" Margin="0,0,0,4"/>
                <DockPanel>
                    <Button x:Name="BtnBrowse" Content="Browse..." DockPanel.Dock="Right" Padding="10,4" Click="BtnBrowse_Click"/>
                    <TextBox x:Name="TxtFolder" IsReadOnly="True" Margin="0,0,8,0" Padding="4"/>
                </DockPanel>
            </StackPanel>
        </Grid>

        <Border Grid.Row="4" Background="White" BorderBrush="#FFDDDDDD" BorderThickness="1" Padding="10" Margin="0,12,0,0">
            <TextBlock x:Name="TxtStatus" TextWrapping="Wrap" Text="Select Online or Offline to begin."/>
        </Border>

        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="BtnReload" Content="Reload" Style="{StaticResource ActionButton}" IsEnabled="False" Click="BtnReload_Click"/>
            <Button x:Name="BtnCancel" Content="Cancel" Padding="16,6" Margin="6,0,0,0" Click="BtnCancel_Click"/>
        </StackPanel>
    </Grid>
</Window>
```

- [ ] **Step 3: Create `Views\GLReloadSourcePicker.xaml.cs`**

```csharp
// GLReloadSourcePicker.xaml.cs in GLSense\Views
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows;

namespace GLSense
{
    public partial class GLReloadSourcePicker : Window
    {
        public string SelectedSource { get; private set; }

        private string _candidateManifestPath;
        private string _candidateZipPath;
        private bool _isValidated;

        public GLReloadSourcePicker()
        {
            InitializeComponent();
            InitializeModeAvailability();
        }

        private void InitializeModeAvailability()
        {
            bool onlineAvailable;
            try
            {
                var loginInfo = GlobalsEx.Addin?.GetLoginInfo();
                onlineAvailable = loginInfo != null && loginInfo.IsLoggedIn && !string.IsNullOrWhiteSpace(loginInfo.LoginUrl);
            }
            catch
            {
                // Defensive: an older historical Addin.Core build reloaded via the
                // Release History browser may not implement GetLoginInfo at all.
                onlineAvailable = false;
            }

            RbOnline.IsEnabled = onlineAvailable;
            if (onlineAvailable) RbOnline.IsChecked = true;
            else RbOffline.IsChecked = true;
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (OnlinePanel == null || OfflinePanel == null) return; // fires during InitializeComponent

            bool isOnline = RbOnline.IsChecked == true;
            OnlinePanel.Visibility = isOnline ? Visibility.Visible : Visibility.Collapsed;
            OfflinePanel.Visibility = isOnline ? Visibility.Collapsed : Visibility.Visible;

            ResetValidation();

            if (!isOnline)
            {
                TxtFolder.Text = GetDownloadsFolder();
                ScanFolder(TxtFolder.Text);
            }
        }

        private void ResetValidation()
        {
            _isValidated = false;
            _candidateManifestPath = null;
            _candidateZipPath = null;
            BtnReload.IsEnabled = false;
            TxtStatus.Text = "Select Online or Offline to begin.";
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.SelectedPath = Directory.Exists(TxtFolder.Text) ? TxtFolder.Text : GetDownloadsFolder();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtFolder.Text = dialog.SelectedPath;
                    ScanFolder(dialog.SelectedPath);
                }
            }
        }

        private void ScanFolder(string folder)
        {
            ResetValidation();

            if (!Directory.Exists(folder))
            {
                TxtStatus.Text = $"Folder not found: {folder}";
                return;
            }

            string manifestPath = Directory.GetFiles(folder, "manifest*.json")
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();
            string zipPath = Directory.GetFiles(folder, "v*.zip")
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();

            if (manifestPath == null || zipPath == null)
            {
                TxtStatus.Text = "No manifest.json + zip pair found in this folder.";
                return;
            }

            ValidateCandidate(manifestPath, zipPath);
        }

        private void ValidateCandidate(string manifestPath, string zipPath)
        {
            var parser = new VersionParser();
            var result = parser.ParseVersionFile(manifestPath);

            if (!result.Success)
            {
                TxtStatus.Text = $"Could not parse manifest.json: {result.ErrorMessage}";
                return;
            }

            if (string.IsNullOrWhiteSpace(result.Checksum))
            {
                TxtStatus.Text = "manifest.json has no checksum recorded - cannot verify, refusing to reload.";
                return;
            }

            string actualChecksum = ComputeSha256(zipPath);
            if (!string.Equals(actualChecksum, result.Checksum, StringComparison.OrdinalIgnoreCase))
            {
                TxtStatus.Text = $"Checksum mismatch - the zip may be corrupt or incomplete. Expected {result.Checksum}, got {actualChecksum}.";
                return;
            }

            if (!IsStrictlyNewer(result.ReleaseDate))
            {
                TxtStatus.Text = $"No update available - {result.Version} ({result.ReleaseDate}) is not newer than the currently loaded release ({GlobalsEx.Context?.Version} / {GlobalsEx.Context?.ReleaseDate}).";
                return;
            }

            _candidateManifestPath = manifestPath;
            _candidateZipPath = zipPath;
            _isValidated = true;
            BtnReload.IsEnabled = true;
            TxtStatus.Text = $"Ready to reload: version {result.Version}, released {result.ReleaseDate}.\n{Path.GetFileName(zipPath)} ({new FileInfo(zipPath).Length / 1024} KB)";
        }

        // VersionParseResult.ReleaseDate is already a parsed DateTime (see
        // VersionParser.ParseVersionJson/ParseVersionFile) - takes DateTime directly,
        // does not re-parse a string.
        private bool IsStrictlyNewer(DateTime candidateReleaseDate)
        {
            string baseline = GlobalsEx.Context?.ReleaseDate;
            if (string.IsNullOrWhiteSpace(baseline)) return true; // nothing loaded yet

            if (!DateTime.TryParse(baseline, out var baselineDate)) return true;

            return candidateReleaseDate > baselineDate;
        }

        private void BtnCheckOnline_Click(object sender, RoutedEventArgs e)
        {
            ResetValidation();

            LoginInfo loginInfo;
            try { loginInfo = GlobalsEx.Addin?.GetLoginInfo(); }
            catch { loginInfo = null; }

            if (loginInfo == null || !loginInfo.IsLoggedIn || string.IsNullOrWhiteSpace(loginInfo.LoginUrl))
            {
                TxtStatus.Text = "Not logged in - switch to Offline mode.";
                return;
            }

            OnlineProgress.Visibility = Visibility.Visible;
            BtnCheckOnline.IsEnabled = false;

            try
            {
                string url = loginInfo.LoginUrl.TrimEnd('/') + "/glsense/projectdlls";
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", loginInfo.LoginToken);

                    string responseJson = client.GetStringAsync(url).GetAwaiter().GetResult();

                    var parser = new VersionParser();
                    var result = parser.ParseVersionJson(responseJson);

                    if (!result.Success)
                    {
                        TxtStatus.Text = $"Could not parse server response: {result.ErrorMessage}";
                        return;
                    }

                    if (!IsStrictlyNewer(result.ReleaseDate))
                    {
                        TxtStatus.Text = $"No updates available - server has {result.Version} ({result.ReleaseDate}), which is not newer than the currently loaded release.";
                        return;
                    }

                    string tempZip = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.zip");
                    var zipBytes = client.GetByteArrayAsync(result.DownloadUrl).GetAwaiter().GetResult();
                    File.WriteAllBytes(tempZip, zipBytes);

                    string actualChecksum = ComputeSha256(tempZip);
                    if (!string.Equals(actualChecksum, result.Checksum, StringComparison.OrdinalIgnoreCase))
                    {
                        TxtStatus.Text = $"Downloaded zip failed checksum verification - expected {result.Checksum}, got {actualChecksum}. Not reloading.";
                        File.Delete(tempZip);
                        return;
                    }

                    string tempManifest = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.json");
                    File.WriteAllText(tempManifest, responseJson);

                    _candidateManifestPath = tempManifest;
                    _candidateZipPath = tempZip;
                    _isValidated = true;
                    BtnReload.IsEnabled = true;
                    TxtStatus.Text = $"Ready to reload: version {result.Version}, released {result.ReleaseDate}.";
                }
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Online check failed: {ex.Message}";
            }
            finally
            {
                OnlineProgress.Visibility = Visibility.Collapsed;
                BtnCheckOnline.IsEnabled = true;
            }
        }

        private void BtnReload_Click(object sender, RoutedEventArgs e)
        {
            if (!_isValidated || _candidateManifestPath == null || _candidateZipPath == null) return;

            try
            {
                string manifestDir = GlobalsEx.Context.Paths.ManifestDirectory;
                if (!Directory.Exists(manifestDir)) Directory.CreateDirectory(manifestDir);

                foreach (var oldZip in Directory.GetFiles(manifestDir, "*.zip"))
                    File.Delete(oldZip);

                File.Copy(_candidateZipPath, Path.Combine(manifestDir, Path.GetFileName(_candidateZipPath)), true);
                File.Copy(_candidateManifestPath, GlobalsEx.Context.Paths.ManifestFile, true);

                SelectedSource = RbOnline.IsChecked == true ? "Online" : "Offline";
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                TxtStatus.Text = $"Failed to stage the new release: {ex.Message}";
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private static string ComputeSha256(string filePath)
        {
            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                var hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "");
            }
        }

        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(ref Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr pszPath);

        private static readonly Guid FolderIdDownloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");

        private static string GetDownloadsFolder()
        {
            try
            {
                var guid = FolderIdDownloads;
                if (SHGetKnownFolderPath(ref guid, 0, IntPtr.Zero, out IntPtr pathPtr) == 0)
                {
                    string path = Marshal.PtrToStringUni(pathPtr);
                    Marshal.FreeCoTaskMem(pathPtr);
                    return path;
                }
            }
            catch
            {
                // fall through to the profile-based fallback below
            }
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        }
    }
}
```

- [ ] **Step 4: Build to verify**

```powershell
msbuild GLSense.sln /t:GLSense /p:Configuration=Debug
```

Expected: `Build succeeded.` (`RibReload_OnClick` still calls the pre-existing
`MessageBox`-based flow at this point — Task B4 rewires it to use this new window.)

- [ ] **Step 5: Commit**

```bash
git add GLSense/GLSense.csproj GLSense/Views/GLReloadSourcePicker.xaml GLSense/Views/GLReloadSourcePicker.xaml.cs
git commit -m "Add GLReloadSourcePicker: Online/Offline reload window (not yet wired to RibReload)"
```

---

### Task B3: `VersionParser.ParseVersionJson` sanity check for the Online response shape

**Files:**
- Read-only check: `GLSense.Shared\VersionParser.cs` (no edit expected — this task
  confirms `ParseVersionJson` genuinely parses the same array-of-object shape the
  Online endpoint contract requires, per spec §7.2, before B2's code relies on it)

**Interfaces:**
- Consumes: `GLSense.Shared.VersionParser.ParseVersionJson(string) : VersionParseResult`
  (existing, confirmed present in Task list research: returns `Success`, `Version`,
  `ReleaseDate`, `DownloadUrl`, `Checksum`, `Notes`, `Mandatory`, `ErrorMessage`).

- [ ] **Step 1: Read `ParseVersionJson`'s implementation**

Open `GLSense.Shared\VersionParser.cs` and confirm `ParseVersionJson(string
jsonContent)` deserializes a JSON **array** of objects with `version`/`releaseDate`/
`downloadUrl`/`checksum`/`notes`/`mandatory` keys (case-insensitive) and returns the
entry `GetLatestVersion` selects. This is the exact shape `post_build.cmd` already
writes to the local `manifest.json` (`CLAUDE.md` section 14/17), so the new
`{LoginUrl}/glsense/projectdlls` server endpoint's response body (spec §7.2) must match
it exactly — no new parsing code is written in this plan.

- [ ] **Step 2: If the shape does NOT match** (e.g. `ParseVersionJson` expects
  something other than a bare top-level array), stop and flag this — Task B2's
  `BtnCheckOnline_Click` would need a different call than `ParseVersionJson`, and the
  server-side contract note in spec §7.2 would need updating to match reality. Do not
  guess a fix; this step exists specifically to catch that mismatch before Phase B's
  window code is exercised against a real server.

- [ ] **Step 3: No commit** — this task either confirms the existing contract (nothing
  to commit) or surfaces a mismatch to resolve before proceeding.

---

### Task B4: Rewire `RibReload_OnClick`/`ReloadAddinCore` to use the picker

**Files:**
- Modify: `GLSense\AddinModule.cs`

**Interfaces:**
- Consumes: `GLSense.GLReloadSourcePicker` (Task B2), `GLSense.Loader.Core
  .UpdateBootstrapper.ResolveVersionToLoad(IGLSenseContext, string) : ResolvedRelease`
  (Task A3).
- Produces: `ReloadAddinCore(Func<ResolvedRelease> resolveRelease) : void` (was
  `ReloadAddinCore() : void`, no parameter) — Task C2 also calls this new signature.

- [ ] **Step 1: Replace `RibReload_OnClick`**

Find the existing method (confirmed shape from research: builds a `MessageBox.Show(...)`
confirmation, then calls `ReloadAddinCore()` inside the `_reloadInProgress` guard).
Replace its entire body with:

```csharp
        private void RibReload_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibReload_OnClick fired (pressed={pressed})");
            if (_reloadInProgress) return;

            var picker = new GLReloadSourcePicker();
            bool? pickerResult = picker.ShowDialog();
            if (pickerResult != true) return;

            string source = picker.SelectedSource;

            _reloadInProgress = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            try
            {
                ReloadAddinCore(() => new UpdateBootstrapper().ResolveVersionToLoad(GlobalsEx.Context, source));
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                _reloadInProgress = false;
            }
        }
```

- [ ] **Step 2: Replace `ReloadAddinCore`'s signature and body**

Replace the entire existing `private void ReloadAddinCore()` method with:

```csharp
        private void ReloadAddinCore(Func<ResolvedRelease> resolveRelease)
        {
            try
            {
                GlobalsEx.Context?.Logger?.LogDebug("Reload requested via ribbon (RibReload) or Release History browser.");

                var oldAddin = GlobalsEx.Addin;
                var loader = GlobalsEx.Loader;

                try
                {
                    oldAddin?.Shutdown();
                }
                catch (Exception ex)
                {
                    GlobalsEx.Context?.Logger?.LogException(ex, "ReloadAddinCore: old instance Shutdown failed");
                }

                GlobalsEx.Addin = null;
                loader?.Unload(GlobalsEx.Context);

                ResolvedRelease resolved = resolveRelease();
                if (resolved == null)
                {
                    GlobalsEx.Context?.Logger?.LogError("ReloadAddinCore: could not resolve a release to load.");
                    MessageBox.Show(
                        "Reload failed - no usable add-in version was found. Make sure GLSense.Addin.Core " +
                        "has been rebuilt (its post_build.cmd publishes a zip + manifest.json into the " +
                        "Manifest folder), then try again.",
                        "Reload GLSense Add-in",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                GlobalsEx.Context.Version = resolved.Version;
                GlobalsEx.Context.ReleaseDate = resolved.ReleaseDate;
                GlobalsEx.Context.ActiveFolderName = resolved.FolderName;
                GlobalsEx.Context?.Logger?.LogDebug($"ReloadAddinCore: version={resolved.Version}, releaseDate={resolved.ReleaseDate}, folderName={resolved.FolderName}");

                GlobalsEx.Addin = loader?.Load(GlobalsEx.Context);

                if (GlobalsEx.Addin != null)
                {
                    GlobalsEx.Context?.Logger?.LogDebug("Reload complete - GlobalsEx.Addin re-pointed to a fresh instance.");
                }
                else
                {
                    GlobalsEx.Context?.Logger?.LogError("Reload failed - GlobalsEx.Addin is null after Load(). The add-in is unavailable until Excel is restarted.");
                    MessageBox.Show(
                        "Reload failed - the add-in could not be loaded. Check the logs. " +
                        "Excel will need to be restarted to recover.",
                        "Reload GLSense Add-in",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                GlobalsEx.Context?.Logger?.LogException(ex, "ReloadAddinCore");
                MessageBox.Show(
                    $"Reload failed: {ex.Message}{Environment.NewLine}Excel may need to be restarted to recover.",
                    "Reload GLSense Add-in",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
```

Confirm `using GLSense.Loader.Core;` is already present at the top of `AddinModule.cs`
(it must be, since `UpdateBootstrapper` was already used there) — `ResolvedRelease` is
in the same namespace, so no new `using` is needed.

- [ ] **Step 3: Build to verify**

```powershell
msbuild GLSense.sln /t:GLSense /p:Configuration=Debug
```

Expected: `Build succeeded.`

- [ ] **Step 4: Manual verification**

1. Rebuild `GLSense.Addin.Core` twice in a row (two distinct `manifest.json`+zip drops
   into `Manifest\`, at least a few minutes apart so `releaseDate` differs).
2. Launch Excel once (adopts the first build automatically at startup).
3. Click `RibReload`. Confirm the new window opens; with nothing logged in, "Online" is
   disabled and "Offline" is pre-selected with the Downloads folder shown (empty, so
   status reads "No manifest.json + zip pair found").
4. Manually copy the second build's `manifest.json`+zip from `Manifest\` (or wherever
   `post_build.cmd` just wrote them, before the automatic startup adopt consumed them —
   easiest: copy from a fresh `GLSense.Addin.Core` rebuild into a test folder) into
   Downloads, click Browse to re-scan (or reopen the window), confirm it's found,
   checksum-verified, and shown as "Ready to reload".
5. Click Reload, confirm the add-in reloads without restarting Excel and
   `ReleaseHistory.json` now has a second entry with `"source": "Offline"`.

- [ ] **Step 5: Commit**

```bash
git add GLSense/AddinModule.cs
git commit -m "Rewire RibReload to open the Online/Offline picker instead of a MessageBox confirm"
```

---

## Phase C — Release History browser

### Task C1: `GLReleaseHistoryBrowser` window

**Files:**
- Modify: `GLSense\GLSense.csproj` (add 2 items — WPF references already added in B2)
- Create: `GLSense\Views\GLReleaseHistoryBrowser.xaml`
- Create: `GLSense\Views\GLReleaseHistoryBrowser.xaml.cs`

**Interfaces:**
- Consumes: `GLSense.Shared.ReleaseHistoryStore.Reconcile`/`ReadAll` (Task A1),
  `GLSense.Loader.Core.ResolvedRelease` (Task A3).
- Produces: `GLSense.GLReleaseHistoryBrowser` — a `Window` with `Chosen :
  ResolvedRelease` (`null` unless `ShowDialog()` returned `true`).

- [ ] **Step 1: Add file items to `GLSense.csproj`**

```xml
    <Compile Include="Views\GLReleaseHistoryBrowser.xaml.cs">
      <DependentUpon>GLReleaseHistoryBrowser.xaml</DependentUpon>
    </Compile>
```

```xml
    <Page Include="Views\GLReleaseHistoryBrowser.xaml">
      <SubType>Designer</SubType>
      <Generator>MSBuild:Compile</Generator>
    </Page>
```

- [ ] **Step 2: Create `Views\GLReleaseHistoryBrowser.xaml`**

```xml
<Window x:Class="GLSense.GLReleaseHistoryBrowser"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="GLSense Release History"
        Width="640" Height="420"
        SizeToContent="Manual"
        WindowStartupLocation="CenterScreen"
        Background="#FFF5F7FA">
    <Window.Resources>
        <Style x:Key="ActionButton" TargetType="Button">
            <Setter Property="Padding" Value="16,6"/>
            <Setter Property="Margin" Value="6,0,0,0"/>
            <Setter Property="Background" Value="#FF1565C0"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" FontWeight="SemiBold" FontSize="14" Foreground="#FF1565C0"
                   Margin="0,0,0,10" Text="Every GLSense.Addin.Core release ever adopted on this machine"/>

        <DataGrid x:Name="GridReleases" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True"
                  SelectionMode="Single" SelectionUnit="FullRow" CanUserAddRows="False">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="90"/>
                <DataGridTextColumn Header="Release Date" Binding="{Binding ReleaseDate}" Width="160"/>
                <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="80"/>
                <DataGridTextColumn Header="Notes" Binding="{Binding Notes}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>

        <Border Grid.Row="2" Background="White" BorderBrush="#FFDDDDDD" BorderThickness="1" Padding="10" Margin="0,10,0,0">
            <TextBlock x:Name="TxtStatus" TextWrapping="Wrap" Text="Select a release, then click Load This Release."/>
        </Border>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="BtnLoad" Content="Load This Release" Style="{StaticResource ActionButton}" IsEnabled="False" Click="BtnLoad_Click"/>
            <Button x:Name="BtnCancel" Content="Close" Padding="16,6" Margin="6,0,0,0" Click="BtnCancel_Click"/>
        </StackPanel>
    </Grid>
</Window>
```

- [ ] **Step 3: Create `Views\GLReleaseHistoryBrowser.xaml.cs`**

```csharp
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
```

`GridReleases`'s `SelectionChanged` handler needs wiring in the XAML too — add
`SelectionChanged="GridReleases_SelectionChanged"` to the `<DataGrid ...>` element from
Step 2 (add it as an attribute on that element).

- [ ] **Step 4: Build to verify**

```powershell
msbuild GLSense.sln /t:GLSense /p:Configuration=Debug
```

Expected: `Build succeeded.`

- [ ] **Step 5: Commit**

```bash
git add GLSense/GLSense.csproj GLSense/Views/GLReleaseHistoryBrowser.xaml GLSense/Views/GLReleaseHistoryBrowser.xaml.cs
git commit -m "Add GLReleaseHistoryBrowser: browse and pick any past release (not yet wired to a ribbon button)"
```

---

### Task C2: `RibReleaseHistory` ribbon button

**Files:**
- Modify: `GLSense\AddinModule.Designer.cs`
- Modify: `GLSense\AddinModule.cs`

**Interfaces:**
- Consumes: `GLSense.GLReleaseHistoryBrowser` (Task C1), `ReloadAddinCore(Func
  <ResolvedRelease>)` (Task B4).

- [ ] **Step 1: Declare the control (Designer.cs field)**

Find `public AddinExpress.MSO.ADXRibbonButton RibReload;` in
`GLSense\AddinModule.Designer.cs` and add a sibling declaration immediately after it:

```csharp
        public AddinExpress.MSO.ADXRibbonButton RibReload;
        public AddinExpress.MSO.ADXRibbonButton RibReleaseHistory;
```

Find `this.RibReload = new AddinExpress.MSO.ADXRibbonButton(this.components);` (near the
top of `InitializeComponent()`) and add a sibling line immediately after it:

```csharp
            this.RibReload = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.RibReleaseHistory = new AddinExpress.MSO.ADXRibbonButton(this.components);
```

Find `this.adxRibbonGroup10.Controls.Add(this.RibReload);` and add a sibling line
immediately after it (same group as `RibReload`, so it appears right next to it):

```csharp
            this.adxRibbonGroup10.Controls.Add(this.RibReload);
            this.adxRibbonGroup10.Controls.Add(this.RibReleaseHistory);
```

Find the `// RibReload` property-assignment block (the one ending in `this.RibReload
.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibReload_OnClick);`)
and add a new block immediately after it, using a fresh GUID for `Id` (any valid GUID
string not already used elsewhere in this file — do not reuse `RibReload`'s):

```csharp
            // 
            // RibReleaseHistory
            // 
            this.RibReleaseHistory.Caption = "Release History";
            this.RibReleaseHistory.Id = "adxRibbonButton_9f2e6c1a4b7d4e3f8a1c6d9e2b5f7a3c";
            this.RibReleaseHistory.Image = 40;
            this.RibReleaseHistory.ImageList = this.ImageList_16X16;
            this.RibReleaseHistory.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.RibReleaseHistory.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.RibReleaseHistory.ScreenTip = "Release History";
            this.RibReleaseHistory.SuperTip = "Browse every GLSense.Addin.Core release ever adopted on this machine, and load any of them without restarting Excel.";
            this.RibReleaseHistory.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibReleaseHistory_OnClick);
```

- [ ] **Step 2: Add the click handler in `AddinModule.cs`**

Add a new method immediately after `RibReload_OnClick` in `GLSense\AddinModule.cs`:

```csharp
        private void RibReleaseHistory_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            GlobalsEx.Context?.Logger?.LogDebug($"RibReleaseHistory_OnClick fired (pressed={pressed})");
            if (_reloadInProgress) return;

            var browser = new GLReleaseHistoryBrowser();
            bool? browserResult = browser.ShowDialog();
            if (browserResult != true || browser.Chosen == null) return;

            var chosen = browser.Chosen;

            _reloadInProgress = true;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            try
            {
                ReloadAddinCore(() => chosen);
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                _reloadInProgress = false;
            }
        }
```

- [ ] **Step 3: Build to verify**

```powershell
msbuild GLSense.sln /t:GLSense /p:Configuration=Debug
```

Expected: `Build succeeded.`

- [ ] **Step 4: Manual verification**

1. Launch Excel with at least two `ReleaseHistory.json` entries already present (from
   Phase A/B's manual verification steps).
2. Click the new "Release History" ribbon button. Confirm the grid lists both entries,
   newest first.
3. Select the OLDER entry, click "Load This Release". Confirm it loads successfully
   with **no** "no update available" rejection (this is the behavior that must differ
   from RibReload's gate).
4. Reopen the browser, confirm `GlobalsEx.Context.Version`/`ReleaseDate` (visible via
   `GLAbout`) now reflect the older release, and that `ReleaseHistory.json` was **not**
   modified by this action (still the same two entries, nothing added or reordered).
5. Manually delete one entry's `Versions\{folderName}\` folder on disk, reopen the
   browser, confirm that entry no longer appears (reconciliation) and
   `ReleaseHistory.json` now has one fewer entry.

- [ ] **Step 5: Commit**

```bash
git add GLSense/AddinModule.Designer.cs GLSense/AddinModule.cs
git commit -m "Add RibReleaseHistory ribbon button, wired to the Release History browser"
```

---

## Final self-review notes (from the plan author, not a task)

- **Spec coverage**: §2/§3 folder layout+keying -> Task A3. §4 schema+concurrency ->
  Task A1. §3.1 threading -> Task A4. §5/§6 fresh-install+reconciliation -> Task A3 (§5)
  and Task C1 (§6's "on browser open" trigger) plus Task A3 (§6's "on reinstall match"
  trigger). §7 picker -> Tasks B1-B4. §8 browser -> Tasks C1-C2. §9 discipline -> encoded
  as a Global Constraint and in B1's doc comment. §10 trade-offs -> encoded as B4's wait
  cursor and simply not building any pruning logic anywhere in this plan.
- **Type consistency checked**: `ResolvedRelease` (Version/ReleaseDate/FolderName)
  used identically in Tasks A3, A4, B4, C1, C2. `ReleaseEntry`
  (Version/ReleaseDate/FolderName/Checksum/Notes/Source) used identically in Tasks A1,
  A3, C1. `ReloadAddinCore`'s signature (`Func<ResolvedRelease>` parameter) matches
  between its Task B4 definition and both its Task B4 and Task C2 call sites.
