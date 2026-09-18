# GLReloadSourcePicker Multi-Version Online Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace `GLReloadSourcePicker`'s single-latest-version Online check with a
multi-row list of every server release, letting the user fetch several releases in
one pass while the newest one automatically becomes what's actually loaded.

**Architecture:** A new server-facing list endpoint feeds a pure classification
function (no WPF/HTTP dependencies) that labels each release Loaded/Already
Downloaded/New against the local `ReleaseHistory.json` catalog. A new shared
`UpdateBootstrapper.ExtractAndCatalog` method (factored out of the existing
`ExtractManifestZipAndAdopt`) does the verify-extract-catalog work for both the old
Install/Offline path and the new Online per-row download loop, so there is exactly
one implementation of "turn a manifest+zip into a catalogued release" in the
codebase. The picker's job ends at growing the local catalog - the existing
`UpdateBootstrapper.ResolveVersionToLoad` adoption logic (already shipped, unchanged)
picks up whatever is now newest.

**Tech Stack:** .NET Framework 4.8.1, WPF (host project `GLSense`), `System.Text.Json`
(via `GLSense.Shared.VersionParser`), `System.Net.Http`. No test project exists
anywhere in this solution (verified: zero `[TestMethod]`/`[Fact]`/`[Test]` attributes
in the repo, no `*.Tests.csproj`). Every task's "test" step is therefore a real,
executable PowerShell verification script run against the actual compiled DLLs
(`Add-Type -Path`/`Add-Type -TypeDefinition`, matching the technique already used and
proven earlier in this engagement to verify `ReleaseHistoryStore.BuildFolderName` and
`VersionParser.ParseVersionFile` directly) - not a placeholder, an actual runnable
check. These scripts are throwaway (written to `%TEMP%`, never committed) since there
is no test project to check them into.

**Spec:** `docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md`

## Global Constraints

- AIPowered only - no changes to `FinalWorkingCode` (it has no Online reload mode).
- The `folderName`/`fileName` manifest fields use lowercase `v{version}_{releaseDateSafe}`
  (matches `ReleaseHistoryStore.BuildFolderName`'s existing convention, already fixed
  in this engagement to use lowercase `v`) - never recompute this client-side when a
  manifest already supplies it.
- `"mandatory"` is display-only in this pass - no forced-selection/blocking behavior.
- Every UI-touching call in new code must go through the same dispatcher-safe pattern
  `GLReloadSourcePicker.xaml.cs` already uses (`Dispatcher.CheckAccess()` + synchronous
  `Dispatcher.Invoke`, never `InvokeAsync`) - this VSTO host's WPF window does not
  reliably run under a `DispatcherSynchronizationContext`.
- Endpoint paths (`/glsense/versions`, `/glsense/manifest`, `/glsense/download`) are
  placeholders pending the server team's final routing - defined as named constants so
  they're a one-line change later, not scattered string literals.
- Row dedup key for "is this release already local/loaded" is exactly
  `(Version, ReleaseDate)` string-equality (`OrdinalIgnoreCase`) - the same key
  `ReleaseHistoryStore.Append` and `UpdateBootstrapper.ResolveVersionToLoad` already use.
  Never invent a different comparison.

---

## Task 1: `VersionInfo.cs` - add `FolderName`/`FileName` properties

**Files:**
- Modify: `GLSense.Contracts\VersionInfo.cs`

**Interfaces:**
- Produces: `VersionInfo.FolderName` (string), `VersionInfo.FileName` (string) -
  bind case-insensitively against the JSON keys `folderName`/`fileName` already
  present in every manifest.json this codebase writes (see
  `GLSense.Addin.Core\post_build.cmd`'s STEP 4). Every later task that parses a
  manifest/list response via `VersionParser`/`System.Text.Json` relies on these two
  properties existing.

- [ ] **Step 1: Write the failing verification script**

Save to `%TEMP%\glsense_task1_verify.ps1`:

```powershell
$base = "D:\SQLLite_Test\GLSense\AIPowered\GLSense"
Add-Type -Path "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll"
# System.Text.Json is a NuGet-sourced assembly, not part of the .NET Framework GAC -
# GLSense.Contracts.dll doesn't carry it, so it must be loaded explicitly. Any
# project's bin output that already has it works; GLSense.Shared's is used here
# since it's guaranteed to exist once this repo has been built at all.
Add-Type -Path "$base\GLSense.Shared\bin\Debug\System.Text.Json.dll"

$json = '[{"version":"1.0.0","releaseDate":"2026-01-01T00:00:00","folderName":"v1.0.0_2026-01-01T00-00-00","fileName":"v1.0.0_2026-01-01T00-00-00.zip","checksum":"ABC","notes":"test","mandatory":false}]'
$options = New-Object System.Text.Json.JsonSerializerOptions
$options.PropertyNameCaseInsensitive = $true
# PowerShell 5.1's generic-static-method call syntax
# (Method[GenericArg](args)) doesn't reliably parse a NESTED generic type argument
# like List[VersionInfo] - use the non-generic Deserialize(json, Type, options)
# overload instead, passing the type as a plain value.
$listType = [System.Collections.Generic.List[GLSense.Contracts.VersionInfo]]
$list = [System.Text.Json.JsonSerializer]::Deserialize($json, $listType, $options)

if ($list[0].FolderName -eq "v1.0.0_2026-01-01T00-00-00" -and $list[0].FileName -eq "v1.0.0_2026-01-01T00-00-00.zip") {
    Write-Output "PASS"
} else {
    Write-Output "FAIL: FolderName='$($list[0].FolderName)' FileName='$($list[0].FileName)'"
}
```

- [ ] **Step 2: Run it, confirm it fails**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task1_verify.ps1"
```

Expected: a PowerShell error (`GetProperty`/member-not-found on `FolderName`) or a
line reading `FAIL: FolderName='' FileName=''` - `VersionInfo` doesn't declare these
properties yet, so `System.Text.Json` silently drops the unmapped JSON fields and
they read back as `$null`/empty.

- [ ] **Step 3: Add the properties**

In `GLSense.Contracts\VersionInfo.cs`, add two properties alongside the existing
`DownloadUrl`/`Checksum`/`Notes`/`Mandatory`:

```csharp
        // Manifest schema fields (see manifest.json under the "Manifest" folder).
        // System.Text.Json binds these case-insensitively against the JSON keys
        // "downloadUrl"/"checksum"/"notes"/"mandatory" - no attributes needed.
        public string DownloadUrl { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }
        public bool Mandatory { get; set; }

        // Added for the Online multi-version reload feature - binds against the
        // "folderName"/"fileName" keys every manifest.json already writes (see
        // GLSense.Addin.Core\post_build.cmd STEP 2b/4). FolderName is always
        // lowercase "v{version}_{releaseDateSafe}" - see
        // ReleaseHistoryStore.BuildFolderName, the authoritative generator this
        // mirrors.
        public string FolderName { get; set; }
        public string FileName { get; set; }
```

- [ ] **Step 4: Rebuild `GLSense.Contracts`**

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense.Contracts\GLSense.Contracts.csproj" /p:Configuration=Debug /p:SignAssembly=false /nologo /v:minimal
```

Expected: `Build succeeded.`

- [ ] **Step 5: Rerun the verification script, confirm it passes**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task1_verify.ps1"
```

Expected: `PASS`

- [ ] **Step 6: Commit**

```bash
git add AIPowered/GLSense/GLSense.Contracts/VersionInfo.cs
git commit -m "Add FolderName/FileName to VersionInfo for the Online multi-version reload feature"
```

---

## Task 2: `UpdateBootstrapper.cs` - add `ExtractAndCatalog` (standalone)

**Files:**
- Modify: `GLSense.Loader.Core\UpdateBootstrapper.cs`

**Interfaces:**
- Consumes: `VersionInfo.FolderName`/`FileName` (Task 1),
  `ReleaseHistoryStore.BuildFolderName(string, string)`,
  `ReleaseHistoryStore.Append(string, ReleaseEntry)` (both already exist, unchanged),
  `VersionParser.ParseVersionJson(string)` → `VersionParseResult.AllVersions`
  (`List<VersionInfo>`, already exists, unchanged).
- Produces: `public ResolvedRelease ExtractAndCatalog(IGLSenseContext context, string manifestJson, byte[] zipBytes, string source)` -
  returns `null` (does not throw) on a missing/unparseable manifest entry or a
  checksum mismatch; otherwise extracts into `Versions\{folderName}\`, writes the
  per-version `manifest.json` snapshot, appends to `ReleaseHistory.json`, and returns
  the `ResolvedRelease` for the newly-catalogued release. Task 3 (the
  `ExtractManifestZipAndAdopt` refactor) and Task 7 (the Online download loop) both
  call this exact signature.

- [ ] **Step 1: Write the failing verification script**

Save to `%TEMP%\glsense_task2_verify.ps1` - builds a tiny fixture zip (one text file),
constructs a stub `IGLSenseContext`/`IPathProvider`/`ILogger` via `Add-Type
-TypeDefinition` (no test project exists, so this is a throwaway in-memory fixture,
not a permanent file), and calls `ExtractAndCatalog` directly:

```powershell
$base = "D:\SQLLite_Test\GLSense\AIPowered\GLSense"
Add-Type -Path "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll"
Add-Type -Path "$base\GLSense.Shared\bin\Debug\GLSense.Shared.dll"
Add-Type -Path "$base\GLSense.Loader.Core\bin\Debug\GLSense.Loader.Core.dll"

$scratch = Join-Path $env:TEMP "glsense_task2_scratch"
if (Test-Path $scratch) { Remove-Item $scratch -Recurse -Force }
New-Item -ItemType Directory -Path $scratch | Out-Null
$versionsPath = Join-Path $scratch "Versions"
$releaseHistoryFile = Join-Path $scratch "ReleaseHistory.json"
New-Item -ItemType Directory -Path $versionsPath | Out-Null

# Build a tiny fixture zip containing one file.
$stageDir = Join-Path $scratch "stage"
New-Item -ItemType Directory -Path $stageDir | Out-Null
Set-Content -Path (Join-Path $stageDir "Dummy.txt") -Value "fixture payload"
$zipPath = Join-Path $scratch "fixture.zip"
Compress-Archive -Path "$stageDir\*" -DestinationPath $zipPath -Force
$zipBytes = [System.IO.File]::ReadAllBytes($zipPath)
$checksum = (Get-FileHash -Algorithm SHA256 $zipPath).Hash

$manifestJson = "[{`"version`":`"9.9.9`",`"releaseDate`":`"2026-01-01T00:00:00`",`"folderName`":`"v9.9.9_2026-01-01T00-00-00`",`"fileName`":`"v9.9.9_2026-01-01T00-00-00.zip`",`"checksum`":`"$checksum`",`"notes`":`"fixture`",`"mandatory`":false}]"

Add-Type -ReferencedAssemblies @(
    "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll",
    "System.Runtime.dll"
) -TypeDefinition @"
using System;
using System.Collections.Generic;
using GLSense.Contracts;

public class FakeLogger : ILogger {
    public void LogInfo(string msg) {}
    public void LogWarn(string msg) {}
    public void LogError(string msg, Exception ex = null) {}
    public void LogDebug(string msg) {}
    public void LogException(Exception ex, string context = "") {}
    public void LogRawJson(string context, string rawJson) {}
    public void FlushDebugLogs(string section = "Buffered Logs") {}
    public void LogMethodEntry(string methodName = "") {}
    public void LogMethodExit(string methodName = "") {}
    public IDisposable BeginLogScope(string scopeName) { return null; }
}

public class FakePathProvider : IPathProvider {
    public string VersionsPath { get; set; }
    public string ReleaseHistoryFile { get; set; }
    public string Root { get { throw new NotImplementedException(); } }
    public string Logs { get { throw new NotImplementedException(); } }
    public string Database { get { throw new NotImplementedException(); } }
    public string LoginBrowserPath { get { throw new NotImplementedException(); } }
    public string DrilldownBrowserPath { get { throw new NotImplementedException(); } }
    public string Temp { get { throw new NotImplementedException(); } }
    public string UrlsDirectory { get { throw new NotImplementedException(); } }
    public string Resources { get { throw new NotImplementedException(); } }
    public string ManifestFile { get { throw new NotImplementedException(); } }
    public string ManifestDirectory { get { throw new NotImplementedException(); } }
    public string LatestVersion { get { throw new NotImplementedException(); } }
    public string LatestReleaseDate { get { throw new NotImplementedException(); } }
    public IReadOnlyList<VersionInfo> AllVersions { get { return null; } }
    public string LatestDownloadUrl { get { throw new NotImplementedException(); } }
    public string LatestChecksum { get { throw new NotImplementedException(); } }
    public string LatestNotes { get { throw new NotImplementedException(); } }
    public bool LatestMandatory { get { throw new NotImplementedException(); } }
    public void Ensure() {}
    public void Refresh() {}
    public void WriteManifest(VersionInfo info) {}
}

public class FakeContext : IGLSenseContext {
    public object ExcelApp { get { return null; } }
    public ILogger Logger { get; set; }
    public IPathProvider Paths { get; set; }
    public IRibbonController RibbonController { get { return null; } }
    public IntPtr ExcelHandle { get; set; }
    public bool DebugMode { get; set; }
    public string Version { get; set; }
    public string ReleaseDate { get; set; }
    public string ActiveFolderName { get; set; }
    public IReadOnlyList<VersionInfo> AllVersions { get { return null; } }
    public void SetRibbonController(IRibbonController controller) {}
    public object EdgeAddinInstance { get; set; }
    public object GetEdgeAddinInstance() { return null; }
}
"@

$paths = New-Object FakePathProvider
$paths.VersionsPath = $versionsPath
$paths.ReleaseHistoryFile = $releaseHistoryFile
$context = New-Object FakeContext
$context.Logger = New-Object FakeLogger
$context.Paths = $paths

$bootstrapper = New-Object GLSense.Loader.Core.UpdateBootstrapper
$resolved = $bootstrapper.ExtractAndCatalog($context, $manifestJson, $zipBytes, "Online")

$extractedFile = Join-Path $versionsPath "v9.9.9_2026-01-01T00-00-00\Dummy.txt"
$snapshotManifest = Join-Path $versionsPath "v9.9.9_2026-01-01T00-00-00\manifest.json"
$catalogHasEntry = (Test-Path $releaseHistoryFile) -and ((Get-Content $releaseHistoryFile -Raw) -match "9\.9\.9")

if ($resolved -ne $null -and $resolved.FolderName -eq "v9.9.9_2026-01-01T00-00-00" -and (Test-Path $extractedFile) -and (Test-Path $snapshotManifest) -and $catalogHasEntry) {
    Write-Output "PASS"
} else {
    Write-Output "FAIL: resolved=$resolved extractedExists=$(Test-Path $extractedFile) snapshotExists=$(Test-Path $snapshotManifest) catalogHasEntry=$catalogHasEntry"
}

# Second check: a deliberately wrong checksum must return null and extract nothing.
Remove-Item $versionsPath -Recurse -Force
New-Item -ItemType Directory -Path $versionsPath | Out-Null
Remove-Item $releaseHistoryFile -Force -ErrorAction SilentlyContinue
$badManifestJson = $manifestJson -replace $checksum, "0000000000000000000000000000000000000000000000000000000000000000"
$badResult = $bootstrapper.ExtractAndCatalog($context, $badManifestJson, $zipBytes, "Online")
if ($badResult -eq $null -and -not (Test-Path (Join-Path $versionsPath "v9.9.9_2026-01-01T00-00-00"))) {
    Write-Output "PASS-CHECKSUM-REJECTED"
} else {
    Write-Output "FAIL-CHECKSUM-NOT-REJECTED"
}
```

- [ ] **Step 2: Run it, confirm it fails**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task2_verify.ps1"
```

Expected: a compile/runtime error - `UpdateBootstrapper` has no `ExtractAndCatalog`
member yet.

- [ ] **Step 3: Add `ExtractAndCatalog`**

In `GLSense.Loader.Core\UpdateBootstrapper.cs`, add `using System.Security.Cryptography;`
to the existing `using` block at the top, then add this new public method (placed
right before the existing `private ResolvedRelease ExtractManifestZipAndAdopt(...)`):

```csharp
        /// <summary>
        /// Verifies zipBytes against the manifest's own checksum, extracts into
        /// Versions\{folderName}\, writes manifestJson as that folder's own
        /// manifest.json snapshot, and appends a ReleaseEntry to the catalog. Shared
        /// by ExtractManifestZipAndAdopt (which reads manifestJson/zipBytes from the
        /// Manifest\ folder on disk - see that method) and GLReloadSourcePicker's
        /// Online bulk-download loop (which fetches both over the network - see
        /// docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md).
        /// Returns null (does not throw) on a missing/unparseable version entry or a
        /// checksum mismatch - the caller decides what "one release in a batch
        /// failed" means for its own flow; nothing is extracted or catalogued in
        /// that case.
        /// </summary>
        public ResolvedRelease ExtractAndCatalog(IGLSenseContext context, string manifestJson, byte[] zipBytes, string source)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            var parsed = new VersionParser(logger).ParseVersionJson(manifestJson);
            var info = parsed.AllVersions?.FirstOrDefault();
            if (info == null || string.IsNullOrWhiteSpace(info.Version))
            {
                logger?.LogError("UpdateBootstrapper.ExtractAndCatalog: manifest JSON did not contain a usable version entry.");
                return null;
            }

            string version = info.Version;
            string releaseDate = info.ReleaseDate;
            string folderName = !string.IsNullOrWhiteSpace(info.FolderName)
                ? info.FolderName
                : ReleaseHistoryStore.BuildFolderName(version, releaseDate);
            string checksum = info.Checksum ?? string.Empty;
            string notes = string.IsNullOrWhiteSpace(info.Notes) ? "Published by GLSense.Addin.Core" : info.Notes;

            string actualChecksum;
            using (var sha256 = SHA256.Create())
            {
                actualChecksum = BitConverter.ToString(sha256.ComputeHash(zipBytes)).Replace("-", "");
            }

            if (!string.IsNullOrWhiteSpace(checksum) &&
                !string.Equals(actualChecksum, checksum, StringComparison.OrdinalIgnoreCase))
            {
                logger?.LogError($"UpdateBootstrapper.ExtractAndCatalog: checksum mismatch for '{version}' ({releaseDate}) - expected {checksum}, got {actualChecksum}. Not extracting.");
                return null;
            }

            string versionFolder = Path.Combine(paths.VersionsPath, folderName);
            if (Directory.Exists(versionFolder))
                Directory.Delete(versionFolder, true);
            Directory.CreateDirectory(versionFolder);

            string tempZipPath = Path.Combine(Path.GetTempPath(), $"GLSenseOnline_{Guid.NewGuid():N}.zip");
            try
            {
                File.WriteAllBytes(tempZipPath, zipBytes);
                ZipFile.ExtractToDirectory(tempZipPath, versionFolder);
            }
            finally
            {
                if (File.Exists(tempZipPath))
                    File.Delete(tempZipPath);
            }

            File.WriteAllText(Path.Combine(versionFolder, "manifest.json"), manifestJson);

            var entry = new ReleaseEntry
            {
                Version = version,
                ReleaseDate = releaseDate,
                FolderName = folderName,
                Checksum = checksum,
                Notes = notes,
                Source = source
            };
            ReleaseHistoryStore.Append(paths.ReleaseHistoryFile, entry);

            logger?.LogDebug($"UpdateBootstrapper.ExtractAndCatalog: extracted and catalogued '{folderName}' (source={source}).");

            return new ResolvedRelease { Version = version, ReleaseDate = releaseDate, FolderName = folderName };
        }
```

- [ ] **Step 4: Rebuild `GLSense.Loader.Core`**

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense.Loader.Core\GLSense.Loader.Core.csproj" /p:Configuration=Debug /p:SignAssembly=false /nologo /v:minimal
```

Expected: `Build succeeded.`

- [ ] **Step 5: Rerun the verification script, confirm it passes**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task2_verify.ps1"
```

Expected:
```
PASS
PASS-CHECKSUM-REJECTED
```

- [ ] **Step 6: Commit**

```bash
git add AIPowered/GLSense/GLSense.Loader.Core/UpdateBootstrapper.cs
git commit -m "Add UpdateBootstrapper.ExtractAndCatalog: shared verify-extract-catalog logic"
```

---

## Task 3: Refactor `ExtractManifestZipAndAdopt` to call `ExtractAndCatalog`

**Files:**
- Modify: `GLSense.Loader.Core\UpdateBootstrapper.cs:146-191` (the existing
  `ExtractManifestZipAndAdopt` method body)

**Interfaces:**
- Consumes: `ExtractAndCatalog` (Task 2).
- Produces: no change to `ExtractManifestZipAndAdopt`'s own signature
  (`private ResolvedRelease ExtractManifestZipAndAdopt(IGLSenseContext context, string source)`) -
  every existing caller (`ResolveVersionToLoad`'s two call sites) is unaffected.

**Behavior change accepted as a deliberate, minor hardening, not a regression:** the
existing method never verified the zip's checksum before extracting from the local
`Manifest\` folder (checksum was only ever checked upstream, by
`GLReloadSourcePicker`'s Offline/Online validation before staging). Routing through
`ExtractAndCatalog` adds that verification at this stage too - for real, legitimate
zips built by this repo's own `post_build.cmd` pipeline (whose checksum always
matches, since it's computed from the same file moments after creation), this can
never spuriously fail.

- [ ] **Step 1: Write the failing verification script**

Save to `%TEMP%\glsense_task3_verify.ps1` - this exercises the *public* entry point
(`ResolveVersionToLoad`), not the private method directly, so it's a realistic
end-to-end check of the refactor:

```powershell
$base = "D:\SQLLite_Test\GLSense\AIPowered\GLSense"
Add-Type -Path "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll"
Add-Type -Path "$base\GLSense.Shared\bin\Debug\GLSense.Shared.dll"
Add-Type -Path "$base\GLSense.Loader.Core\bin\Debug\GLSense.Loader.Core.dll"

$scratch = Join-Path $env:TEMP "glsense_task3_scratch"
if (Test-Path $scratch) { Remove-Item $scratch -Recurse -Force }
New-Item -ItemType Directory -Path $scratch | Out-Null
$manifestDir = Join-Path $scratch "Manifest"
$versionsPath = Join-Path $scratch "Versions"
$releaseHistoryFile = Join-Path $scratch "ReleaseHistory.json"
New-Item -ItemType Directory -Path $manifestDir | Out-Null
New-Item -ItemType Directory -Path $versionsPath | Out-Null

$stageDir = Join-Path $scratch "stage"
New-Item -ItemType Directory -Path $stageDir | Out-Null
Set-Content -Path (Join-Path $stageDir "Dummy.txt") -Value "fixture payload"
$zipPath = Join-Path $manifestDir "v8.8.8_2026-02-02T00-00-00.zip"
Compress-Archive -Path "$stageDir\*" -DestinationPath $zipPath -Force
$checksum = (Get-FileHash -Algorithm SHA256 $zipPath).Hash

$manifestFile = Join-Path $manifestDir "manifest.json"
$manifestJson = "[{`"version`":`"8.8.8`",`"releaseDate`":`"2026-02-02T00:00:00`",`"folderName`":`"v8.8.8_2026-02-02T00-00-00`",`"fileName`":`"v8.8.8_2026-02-02T00-00-00.zip`",`"checksum`":`"$checksum`",`"notes`":`"fixture`",`"mandatory`":false}]"
Set-Content -Path $manifestFile -Value $manifestJson

Add-Type -ReferencedAssemblies @(
    "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll"
) -TypeDefinition @"
using System;
using System.Collections.Generic;
using GLSense.Contracts;

public class FakeLogger : ILogger {
    public void LogInfo(string msg) {}
    public void LogWarn(string msg) {}
    public void LogError(string msg, Exception ex = null) {}
    public void LogDebug(string msg) {}
    public void LogException(Exception ex, string context = "") {}
    public void LogRawJson(string context, string rawJson) {}
    public void FlushDebugLogs(string section = "Buffered Logs") {}
    public void LogMethodEntry(string methodName = "") {}
    public void LogMethodExit(string methodName = "") {}
    public IDisposable BeginLogScope(string scopeName) { return null; }
}

public class FakePathProvider : IPathProvider {
    public string ManifestDirectory { get; set; }
    public string ManifestFile { get; set; }
    public string VersionsPath { get; set; }
    public string ReleaseHistoryFile { get; set; }
    public string LatestVersion { get; set; }
    public string LatestReleaseDate { get; set; }
    public string LatestChecksum { get; set; }
    public string LatestNotes { get; set; }
    public bool LatestMandatory { get; set; }
    public IReadOnlyList<VersionInfo> AllVersions { get; set; }
    public string Root { get { throw new NotImplementedException(); } }
    public string Logs { get { throw new NotImplementedException(); } }
    public string Database { get { throw new NotImplementedException(); } }
    public string LoginBrowserPath { get { throw new NotImplementedException(); } }
    public string DrilldownBrowserPath { get { throw new NotImplementedException(); } }
    public string Temp { get { throw new NotImplementedException(); } }
    public string UrlsDirectory { get { throw new NotImplementedException(); } }
    public string Resources { get { throw new NotImplementedException(); } }
    public string LatestDownloadUrl { get { throw new NotImplementedException(); } }
    public void Ensure() {}
    public void Refresh() {}
    public void WriteManifest(VersionInfo info) {}
}

public class FakeContext : IGLSenseContext {
    public object ExcelApp { get { return null; } }
    public ILogger Logger { get; set; }
    public IPathProvider Paths { get; set; }
    public IRibbonController RibbonController { get { return null; } }
    public IntPtr ExcelHandle { get; set; }
    public bool DebugMode { get; set; }
    public string Version { get; set; }
    public string ReleaseDate { get; set; }
    public string ActiveFolderName { get; set; }
    public IReadOnlyList<VersionInfo> AllVersions { get { return null; } }
    public void SetRibbonController(IRibbonController controller) {}
    public object EdgeAddinInstance { get; set; }
    public object GetEdgeAddinInstance() { return null; }
}
"@

$paths = New-Object FakePathProvider
$paths.ManifestDirectory = $manifestDir
$paths.ManifestFile = $manifestFile
$paths.VersionsPath = $versionsPath
$paths.ReleaseHistoryFile = $releaseHistoryFile
$context = New-Object FakeContext
$context.Logger = New-Object FakeLogger
$context.Paths = $paths
# The current, unrefactored ExtractManifestZipAndAdopt reads paths.LatestVersion/
# LatestReleaseDate directly (populated by a real Refresh() in production, which
# this throwaway fixture's Refresh() deliberately doesn't reimplement) - set them
# directly since the test author already knows what $manifestJson above contains.
# After Step 3's refactor, ExtractManifestZipAndAdopt stops reading these at all
# (it reads the manifest FILE's text directly instead) - these two lines become
# inert but harmless once that lands, which is exactly the "identical PASS before
# and after" parity this test is designed to prove.
$paths.LatestVersion = "8.8.8"
$paths.LatestReleaseDate = "2026-02-02T00:00:00"

$bootstrapper = New-Object GLSense.Loader.Core.UpdateBootstrapper
$resolved = $bootstrapper.ResolveVersionToLoad($context, "Install")

$extractedFile = Join-Path $versionsPath "v8.8.8_2026-02-02T00-00-00\Dummy.txt"
$zipStillPresent = Test-Path $zipPath

if ($resolved -ne $null -and $resolved.FolderName -eq "v8.8.8_2026-02-02T00-00-00" -and (Test-Path $extractedFile) -and -not $zipStillPresent) {
    Write-Output "PASS"
} else {
    Write-Output "FAIL: resolved=$resolved extractedExists=$(Test-Path $extractedFile) zipStillPresent=$zipStillPresent"
}
```

- [ ] **Step 2: Run it, confirm it currently passes (baseline)**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task3_verify.ps1"
```

Expected: `PASS` - this is the EXISTING behavior working correctly before the
refactor. Record this as the baseline; the point of Task 3 is that this script keeps
passing identically after the refactor below, proving behavioral parity.

- [ ] **Step 3: Refactor `ExtractManifestZipAndAdopt`**

Replace the existing body (`GLSense.Loader.Core\UpdateBootstrapper.cs:146-191`) with:

```csharp
        private ResolvedRelease ExtractManifestZipAndAdopt(IGLSenseContext context, string source)
        {
            var logger = context.Logger;
            var paths = context.Paths;

            string zipPath = Directory.GetFiles(paths.ManifestDirectory, "*.zip").First();
            string manifestJson = File.ReadAllText(paths.ManifestFile);
            byte[] zipBytes = File.ReadAllBytes(zipPath);

            logger?.LogDebug($"UpdateBootstrapper: extracting '{zipPath}' (source={source}).");

            var resolved = ExtractAndCatalog(context, manifestJson, zipBytes, source);
            if (resolved == null)
            {
                logger?.LogError($"UpdateBootstrapper: ExtractAndCatalog failed for '{zipPath}' - leaving it in place for the next launch to retry.");
                return null;
            }

            // Delete the zip only after the catalog append has genuinely succeeded -
            // if anything above throws, the zip is still there so the next launch can
            // retry the full extract+catalog sequence, instead of being left with DLLs
            // on disk but no catalog entry and no way to retry (the zip already gone).
            File.Delete(zipPath);

            logger?.LogDebug($"UpdateBootstrapper: extracted, catalogued (source={source}), and deleted '{zipPath}'. Adopting '{resolved.FolderName}'.");

            return resolved;
        }
```

- [ ] **Step 4: Rebuild `GLSense.Loader.Core`**

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense.Loader.Core\GLSense.Loader.Core.csproj" /p:Configuration=Debug /p:SignAssembly=false /nologo /v:minimal
```

Expected: `Build succeeded.`

- [ ] **Step 5: Rerun the verification script, confirm identical PASS**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task3_verify.ps1"
```

Expected: `PASS` (same outcome as Step 2's baseline - behavioral parity confirmed).

- [ ] **Step 6: Commit**

```bash
git add AIPowered/GLSense/GLSense.Loader.Core/UpdateBootstrapper.cs
git commit -m "Refactor ExtractManifestZipAndAdopt to share ExtractAndCatalog"
```

---

## Task 4: `OnlineReleaseRow.cs` - new row model

**Files:**
- Create: `GLSense\Views\OnlineReleaseRow.cs`

**Interfaces:**
- Produces: `OnlineReleaseRow` class with `Version`, `ReleaseDate`, `Mandatory`,
  `FolderName`, `FileName`, `Checksum`, `Notes` (plain settable properties),
  `IsCurrentlyLoaded`, `IsAlreadyDownloaded` (plain settable bools),
  `IsSelectable` (computed: `!IsCurrentlyLoaded && !IsAlreadyDownloaded`),
  `StatusText` (computed: "Currently loaded"/"Already Downloaded"/"New"),
  `IsChecked` (bool, raises `PropertyChanged` on change - the DataGrid checkbox
  binding and Task 6/7's button-count logic depend on this). Task 5's classifier
  constructs these; Task 6/7's XAML binds to every property; Task 7's download loop
  reads `FolderName`/`FileName`/`Version`/`ReleaseDate`/`IsChecked`/`IsSelectable`.

- [ ] **Step 1: Write the failing verification script**

Save to `%TEMP%\glsense_task4_verify.ps1`:

```powershell
$base = "D:\SQLLite_Test\GLSense\AIPowered\GLSense"
Add-Type -Path "$base\GLSense\bin\Debug\GLSense.dll"

$row = New-Object GLSense.OnlineReleaseRow
$row.Version = "1.2.3"
$row.IsCurrentlyLoaded = $false
$row.IsAlreadyDownloaded = $false

$eventsRaised = New-Object System.Collections.Generic.List[string]
$handler = [System.ComponentModel.PropertyChangedEventHandler]{
    param($s, $e) $eventsRaised.Add($e.PropertyName)
}
$row.add_PropertyChanged($handler)
$row.IsChecked = $true

if ($row.IsSelectable -eq $true -and $row.StatusText -eq "New" -and $eventsRaised.Contains("IsChecked")) {
    Write-Output "PASS"
} else {
    Write-Output "FAIL: IsSelectable=$($row.IsSelectable) StatusText=$($row.StatusText) events=$($eventsRaised -join ',')"
}
```

- [ ] **Step 2: Run it, confirm it fails**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task4_verify.ps1"
```

Expected: error - `GLSense.OnlineReleaseRow` type does not exist yet.

- [ ] **Step 3: Create the file**

```csharp
// OnlineReleaseRow.cs in GLSense\Views
using System.ComponentModel;

namespace GLSense
{
    /// <summary>
    /// One row in GLReloadSourcePicker's Online-mode release list. Built by
    /// OnlineReleaseClassifier from a server VersionInfo entry plus the local
    /// ReleaseHistory.json catalog - see
    /// docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md.
    /// </summary>
    public class OnlineReleaseRow : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public bool Mandatory { get; set; }
        public string FolderName { get; set; }
        public string FileName { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }

        public bool IsCurrentlyLoaded { get; set; }
        public bool IsAlreadyDownloaded { get; set; }

        /// <summary>Neither currently loaded nor already catalogued locally - the
        /// only state a checkbox can be enabled/checked for.</summary>
        public bool IsSelectable => !IsCurrentlyLoaded && !IsAlreadyDownloaded;

        public string StatusText =>
            IsCurrentlyLoaded ? "Currently loaded" :
            IsAlreadyDownloaded ? "Already Downloaded" :
            "New";

        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (_isChecked == value) return;
                _isChecked = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
```

- [ ] **Step 4: Add the file to the project and rebuild**

In `GLSense\GLSense.csproj:171`, immediately after
`<Compile Include="Views\WindowChromeHelper.cs" />`, add:

```xml
    <Compile Include="Views\OnlineReleaseRow.cs" />
```

(this is an old-style csproj with no implicit file globbing, so new files must be
listed explicitly - `WindowChromeHelper.cs` at line 171 is the existing example of a
plain, non-XAML-backed `Views\*.cs` file added the same way).

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense\GLSense.csproj" /p:Configuration=Debug /p:SignAssembly=false /p:RegisterForComInterop=false /nologo /v:minimal
```

Expected: `Build succeeded.`

- [ ] **Step 5: Rerun the verification script, confirm it passes**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task4_verify.ps1"
```

Expected: `PASS`

- [ ] **Step 6: Commit**

```bash
git add AIPowered/GLSense/GLSense/Views/OnlineReleaseRow.cs AIPowered/GLSense/GLSense/GLSense.csproj
git commit -m "Add OnlineReleaseRow model for the Online multi-version release list"
```

---

## Task 5: `OnlineReleaseClassifier.cs` - pure classification logic

**Files:**
- Create: `GLSense\Views\OnlineReleaseClassifier.cs`

**Interfaces:**
- Consumes: `VersionInfo` (Task 1's `FolderName`/`FileName` additions),
  `ReleaseEntry` (`GLSense.Shared`, existing, unchanged - `Version`/`ReleaseDate`/
  `FolderName`/`Checksum`/`Notes`/`Source`), `OnlineReleaseRow` (Task 4).
- Produces:
  `public static List<OnlineReleaseRow> Classify(IEnumerable<VersionInfo> serverEntries, IReadOnlyList<ReleaseEntry> localEntries, string currentVersion, string currentReleaseDate)` -
  returns one row per server entry with `IsCurrentlyLoaded`/`IsAlreadyDownloaded` set
  and the single newest selectable row pre-checked. Also
  `public static DateTime ParseReleaseDateOrMin(string releaseDate)` - a small shared
  helper Task 7's download loop reuses for ascending-order sorting, so the "what does
  newest/oldest mean" parsing logic exists in exactly one place.

- [ ] **Step 1: Write the failing verification script**

Save to `%TEMP%\glsense_task5_verify.ps1` - covers: Loaded wins over Already
Downloaded when both would technically match (defensive, shouldn't happen in
practice since a release can't be both, but the classifier's `IsCurrentlyLoaded`
check must run first); default selection picks the newest New row; a row with an
unparseable `ReleaseDate` doesn't crash the sort:

```powershell
$base = "D:\SQLLite_Test\GLSense\AIPowered\GLSense"
Add-Type -Path "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll"
Add-Type -Path "$base\GLSense.Shared\bin\Debug\GLSense.Shared.dll"
Add-Type -Path "$base\GLSense\bin\Debug\GLSense.dll"

function New-VersionInfo($version, $releaseDate, $folderName) {
    $vi = New-Object GLSense.Contracts.VersionInfo
    $vi.Version = $version
    $vi.ReleaseDate = $releaseDate
    $vi.FolderName = $folderName
    $vi.FileName = "$folderName.zip"
    return $vi
}

function New-ReleaseEntry($version, $releaseDate, $folderName) {
    $re = New-Object GLSense.Shared.ReleaseEntry
    $re.Version = $version
    $re.ReleaseDate = $releaseDate
    $re.FolderName = $folderName
    return $re
}

$serverEntries = New-Object 'System.Collections.Generic.List[GLSense.Contracts.VersionInfo]'
$serverEntries.Add((New-VersionInfo "1.0.2" "2026-03-03T00:00:00" "v1.0.2_2026-03-03T00-00-00"))  # newest new
$serverEntries.Add((New-VersionInfo "1.0.1" "2026-02-02T00:00:00" "v1.0.1_2026-02-02T00-00-00"))  # older new
$serverEntries.Add((New-VersionInfo "1.0.0" "2026-01-01T00:00:00" "v1.0.0_2026-01-01T00-00-00"))  # currently loaded
$serverEntries.Add((New-VersionInfo "0.9.0" "2025-12-01T00:00:00" "v0.9.0_2025-12-01T00-00-00"))  # already local
$serverEntries.Add((New-VersionInfo "0.8.0" "not-a-date" "v0.8.0_bad"))                            # unparseable date, still "new"

$localEntries = New-Object 'System.Collections.Generic.List[GLSense.Shared.ReleaseEntry]'
$localEntries.Add((New-ReleaseEntry "1.0.0" "2026-01-01T00:00:00" "v1.0.0_2026-01-01T00-00-00"))
$localEntries.Add((New-ReleaseEntry "0.9.0" "2025-12-01T00:00:00" "v0.9.0_2025-12-01T00-00-00"))

$rows = [GLSense.OnlineReleaseClassifier]::Classify($serverEntries, $localEntries, "1.0.0", "2026-01-01T00:00:00")

$loadedRow = $rows | Where-Object { $_.Version -eq "1.0.0" }
$localRow = $rows | Where-Object { $_.Version -eq "0.9.0" }
$newestNewRow = $rows | Where-Object { $_.Version -eq "1.0.2" }
$olderNewRow = $rows | Where-Object { $_.Version -eq "1.0.1" }
$badDateRow = $rows | Where-Object { $_.Version -eq "0.8.0" }

$ok = $true
if ($loadedRow.IsCurrentlyLoaded -ne $true -or $loadedRow.IsSelectable -ne $false) { $ok = $false }
if ($localRow.IsAlreadyDownloaded -ne $true -or $localRow.IsSelectable -ne $false) { $ok = $false }
if ($newestNewRow.IsSelectable -ne $true -or $newestNewRow.IsChecked -ne $true) { $ok = $false }
if ($olderNewRow.IsChecked -ne $false) { $ok = $false }
if ($badDateRow.IsSelectable -ne $true) { $ok = $false }

if ($ok) { Write-Output "PASS" } else {
    Write-Output "FAIL: loaded(loaded=$($loadedRow.IsCurrentlyLoaded),sel=$($loadedRow.IsSelectable)) local(local=$($localRow.IsAlreadyDownloaded),sel=$($localRow.IsSelectable)) newest(sel=$($newestNewRow.IsSelectable),checked=$($newestNewRow.IsChecked)) older(checked=$($olderNewRow.IsChecked)) bad(sel=$($badDateRow.IsSelectable))"
}
```

- [ ] **Step 2: Run it, confirm it fails**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task5_verify.ps1"
```

Expected: error - `GLSense.OnlineReleaseClassifier` type does not exist yet.

- [ ] **Step 3: Create the file**

```csharp
// OnlineReleaseClassifier.cs in GLSense\Views
using GLSense.Contracts;
using GLSense.Shared;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GLSense
{
    /// <summary>
    /// Pure classification logic for GLReloadSourcePicker's Online-mode release
    /// list - no WPF, no HttpClient, so it's directly testable. See
    /// docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md
    /// section 2.
    /// </summary>
    public static class OnlineReleaseClassifier
    {
        public static List<OnlineReleaseRow> Classify(
            IEnumerable<VersionInfo> serverEntries,
            IReadOnlyList<ReleaseEntry> localEntries,
            string currentVersion,
            string currentReleaseDate)
        {
            var rows = new List<OnlineReleaseRow>();
            if (serverEntries == null) return rows;

            localEntries = localEntries ?? new List<ReleaseEntry>();

            foreach (var entry in serverEntries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Version)) continue;

                bool isCurrentlyLoaded =
                    string.Equals(entry.Version, currentVersion, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(entry.ReleaseDate, currentReleaseDate, StringComparison.OrdinalIgnoreCase);

                bool isAlreadyDownloaded = !isCurrentlyLoaded && localEntries.Any(l =>
                    string.Equals(l.Version, entry.Version, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(l.ReleaseDate, entry.ReleaseDate, StringComparison.OrdinalIgnoreCase));

                rows.Add(new OnlineReleaseRow
                {
                    Version = entry.Version,
                    ReleaseDate = entry.ReleaseDate,
                    Mandatory = entry.Mandatory,
                    FolderName = entry.FolderName,
                    FileName = entry.FileName,
                    Checksum = entry.Checksum,
                    Notes = entry.Notes,
                    IsCurrentlyLoaded = isCurrentlyLoaded,
                    IsAlreadyDownloaded = isAlreadyDownloaded
                });
            }

            var newestSelectable = rows
                .Where(r => r.IsSelectable)
                .OrderByDescending(r => ParseReleaseDateOrMin(r.ReleaseDate))
                .FirstOrDefault();

            if (newestSelectable != null)
                newestSelectable.IsChecked = true;

            return rows;
        }

        /// <summary>Shared by Classify's default-selection sort and
        /// GLReloadSourcePicker's ascending download-order sort, so "what counts as
        /// newest/oldest" is defined in exactly one place.</summary>
        public static DateTime ParseReleaseDateOrMin(string releaseDate)
        {
            return DateTime.TryParse(releaseDate, out var dt) ? dt : DateTime.MinValue;
        }
    }
}
```

Note: `VersionInfo` in this task's usage does not yet have a `Mandatory` binding
issue - `Mandatory` already exists on `VersionInfo` (added long before this feature,
see `GLSense.Contracts\VersionInfo.cs`), only `FolderName`/`FileName` were new
(Task 1).

- [ ] **Step 4: Add the file to the project and rebuild**

In `GLSense\GLSense.csproj`, immediately after the
`<Compile Include="Views\OnlineReleaseRow.cs" />` line Task 4 just added, add:

```xml
    <Compile Include="Views\OnlineReleaseClassifier.cs" />
```

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense\GLSense.csproj" /p:Configuration=Debug /p:SignAssembly=false /p:RegisterForComInterop=false /nologo /v:minimal
```

Expected: `Build succeeded.`

- [ ] **Step 5: Rerun the verification script, confirm it passes**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task5_verify.ps1"
```

Expected: `PASS`

- [ ] **Step 6: Commit**

```bash
git add AIPowered/GLSense/GLSense/Views/OnlineReleaseClassifier.cs AIPowered/GLSense/GLSense/GLSense.csproj
git commit -m "Add OnlineReleaseClassifier: pure Loaded/Already-Downloaded/New row logic"
```

---

## Task 6: `GLReloadSourcePicker` - list fetch, grid rendering, row classification, details panel

**Files:**
- Modify: `GLSense\Views\GLReloadSourcePicker.xaml`
- Modify: `GLSense\Views\GLReloadSourcePicker.xaml.cs`

**Interfaces:**
- Consumes: `OnlineReleaseRow` (Task 4), `OnlineReleaseClassifier.Classify` (Task 5),
  `VersionParser.ParseVersionJson(string)` → `VersionParseResult.AllVersions`
  (existing, unchanged), `ReleaseHistoryStore.ReadAll(string)` (existing,
  unchanged), `GlobalsEx.Addin.GetLoginInfo()` → `LoginInfo` (existing, unchanged),
  `GlobalsEx.Context.Version`/`ReleaseDate`/`Paths.ReleaseHistoryFile` (existing,
  unchanged).
- Produces: a working, checkable, detail-viewable release grid populated from a real
  (or fixture) list response. Task 7 adds the download loop that consumes
  `_onlineRows` (the `ObservableCollection<OnlineReleaseRow>` field this task
  creates) and the `GridOnlineReleases`/`BtnFetchAndReload` named elements this
  task's XAML declares.

This task deliberately stops short of the download loop (Task 7) - it's reviewable
and testable on its own: open the window, click "Check for Updates" against a
fixture list response, see correctly-colored/labeled rows, click rows to see
details, check/uncheck boxes. No network write/extraction happens yet.

- [ ] **Step 1: Replace `OnlinePanel`'s content in `GLReloadSourcePicker.xaml`**

Replace this block (the `OnlinePanel` `StackPanel` inside `Grid Grid.Row="3"`):

```xml
            <StackPanel x:Name="OnlinePanel" Visibility="Collapsed">
                <Button x:Name="BtnCheckOnline" Style="{StaticResource SecondaryButton}" Content="Check for Update" HorizontalAlignment="Left"
                        Padding="10,4" Click="BtnCheckOnline_Click"
                        ToolTip="Query the logged-in GLSense server for the latest published release, download it, verify its checksum, and enable Reload if it is newer than what's currently loaded."/>
            </StackPanel>
```

with:

```xml
            <StackPanel x:Name="OnlinePanel" Visibility="Collapsed">
                <Button x:Name="BtnCheckOnline" Style="{StaticResource SecondaryButton}" Content="Check for Updates" HorizontalAlignment="Left"
                        Padding="10,4" Click="BtnCheckOnline_Click"
                        ToolTip="Query the logged-in GLSense server for every release not yet on this machine."/>

                <DataGrid x:Name="GridOnlineReleases" AutoGenerateColumns="False" IsReadOnly="True"
                          SelectionMode="Single" SelectionUnit="FullRow" CanUserAddRows="False"
                          Margin="0,8,0,0" MaxHeight="220"
                          Background="White" RowBackground="White" AlternatingRowBackground="White"
                          BorderBrush="#FFDDDDDD" BorderThickness="1"
                          SelectionChanged="GridOnlineReleases_SelectionChanged"
                          ToolTip="Every release the server knows about. Rows already on this machine (grey) or currently loaded (blue) can't be re-selected here - use Release History for those.">
                    <DataGrid.RowStyle>
                        <Style TargetType="DataGridRow">
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding IsCurrentlyLoaded}" Value="True">
                                    <Setter Property="Background" Value="#FFE3F2FD"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                </DataTrigger>
                                <DataTrigger Binding="{Binding IsAlreadyDownloaded}" Value="True">
                                    <Setter Property="Background" Value="#FFF4F4F4"/>
                                    <Setter Property="Foreground" Value="#FF8A8A8A"/>
                                </DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </DataGrid.RowStyle>
                    <DataGrid.Columns>
                        <DataGridTemplateColumn Header="" Width="34">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <CheckBox IsChecked="{Binding IsChecked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                              IsEnabled="{Binding IsSelectable}" HorizontalAlignment="Center"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="70"/>
                        <DataGridTextColumn Header="Released" Binding="{Binding ReleaseDate}" Width="140"/>
                        <DataGridTextColumn Header="Notes" Binding="{Binding Notes}" Width="*"/>
                        <DataGridTextColumn Header="Status" Binding="{Binding StatusText}" Width="110"/>
                    </DataGrid.Columns>
                </DataGrid>

                <Border x:Name="DetailsPanel" Background="White" BorderBrush="#FFDDDDDD" BorderThickness="1" Padding="10" Margin="0,8,0,0" Visibility="Collapsed">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="100"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                            <RowDefinition/><RowDefinition/><RowDefinition/><RowDefinition/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Version:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailVersion" Grid.Row="0" Grid.Column="1"/>
                        <TextBlock Grid.Row="1" Grid.Column="0" Text="Released:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailReleased" Grid.Row="1" Grid.Column="1"/>
                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Status:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailStatus" Grid.Row="2" Grid.Column="1"/>
                        <TextBlock Grid.Row="3" Grid.Column="0" Text="Notes:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailNotes" Grid.Row="3" Grid.Column="1" TextWrapping="Wrap"/>
                        <TextBlock Grid.Row="4" Grid.Column="0" Text="Mandatory:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailMandatory" Grid.Row="4" Grid.Column="1"/>
                        <TextBlock Grid.Row="5" Grid.Column="0" Text="Folder Name:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailFolderName" Grid.Row="5" Grid.Column="1" TextWrapping="Wrap"/>
                        <TextBlock Grid.Row="6" Grid.Column="0" Text="File Name:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailFileName" Grid.Row="6" Grid.Column="1" TextWrapping="Wrap"/>
                        <TextBlock Grid.Row="7" Grid.Column="0" Text="Checksum:" Foreground="#FF777777"/>
                        <TextBlock x:Name="TxtDetailChecksum" Grid.Row="7" Grid.Column="1" TextWrapping="Wrap"/>
                    </Grid>
                </Border>
            </StackPanel>
```

- [ ] **Step 2: Update the window size and the Online radio's tooltip**

Change the `Window` element's `Width="520" Height="560"` to `Width="660" Height="680"`.

Change `RbOnline`'s `ToolTip` from
`"Check the currently logged-in GLSense server for a newer release and download it automatically. Only available after a successful login this Excel session."`
to
`"Check the currently logged-in GLSense server for every release not yet on this machine. Only available after a successful login this Excel session."`

- [ ] **Step 3: Add the new bottom-row `BtnFetchAndReload` button (visibility toggled with mode)**

In the `Grid.Row="5"` `StackPanel` (the bottom-right button row), add a new button
before the existing `BtnReload`:

```xml
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="BtnFetchAndReload" Content="Fetch Selected &amp; Reload" Style="{StaticResource ActionButton}" IsEnabled="False" Visibility="Collapsed" Click="BtnFetchAndReload_Click"
                    ToolTip="Download every checked release, catalog it, then reload GLSense onto whichever release is now newest - without restarting Excel. Any drilldown/refresh/snapshot in progress will be interrupted."/>
            <Button x:Name="BtnReload" Content="Reload" Style="{StaticResource ActionButton}" IsEnabled="False" Click="BtnReload_Click"
                    ToolTip="Stage the validated release into the Manifest folder and reload GLSense immediately, without restarting Excel. Any drilldown/refresh/snapshot in progress will be interrupted."/>
            <Button x:Name="BtnCancel" Content="Cancel" Style="{StaticResource SecondaryButton}" Click="BtnCancel_Click"
                    ToolTip="Close this window without reloading anything."/>
        </StackPanel>
```

(`BtnFetchAndReload_Click` is implemented in Task 7 - for this task, leave it
unwired by adding a temporary no-op handler so the project compiles; Task 7 replaces
it with the real implementation.)

- [ ] **Step 4: Update `GLReloadSourcePicker.xaml.cs`'s usings and `Mode_Checked`**

Add to the top `using` block:

```csharp
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Controls;
```

Replace `Mode_Checked`'s body:

```csharp
        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (OnlinePanel == null || OfflinePanel == null) return; // fires during InitializeComponent

            bool isOnline = RbOnline.IsChecked == true;
            OnlinePanel.Visibility = isOnline ? Visibility.Visible : Visibility.Collapsed;
            OfflinePanel.Visibility = isOnline ? Visibility.Collapsed : Visibility.Visible;
            BtnFetchAndReload.Visibility = isOnline ? Visibility.Visible : Visibility.Collapsed;
            BtnReload.Visibility = isOnline ? Visibility.Collapsed : Visibility.Visible;

            ResetValidation();

            if (isOnline)
            {
                TxtStatus.Text = "Click \"Check for Updates\" to see every release not yet on this machine.";
                DetailsPanel.Visibility = Visibility.Collapsed;
                UpdateFetchButtonState();
            }
            else
            {
                // Deliberately no default folder and no auto-scan here anymore - this
                // used to default straight to the Downloads folder and immediately
                // scan it the instant the window opened (Downloads is very often the
                // largest, most heavily-populated folder on a machine, so this could
                // visibly delay the window even showing up). Offline mode now starts
                // empty; scanning only ever happens when the user explicitly clicks
                // Browse... (which still conveniently defaults ITS OWN starting folder
                // to Downloads - see BtnBrowse_Click - without scanning anything until
                // a folder is actually chosen).
                TxtStatus.Text = "Click \"Browse...\" to select a folder to scan for a manifest.json + zip pair.";
            }
        }
```

- [ ] **Step 5: Add the `_onlineRows` field and replace `BtnCheckOnline_Click`**

Add a field near the top of the class (next to `_candidateManifestPath` etc.):

```csharp
        private readonly ObservableCollection<OnlineReleaseRow> _onlineRows = new ObservableCollection<OnlineReleaseRow>();
        private const string VersionsListPath = "/glsense/versions";
```

Replace the ENTIRE existing `BtnCheckOnline_Click` method
(`GLReloadSourcePicker.xaml.cs:323-407`) with:

```csharp
        private async void BtnCheckOnline_Click(object sender, RoutedEventArgs e)
        {
            await LoadOnlineReleasesAsync();
        }

        private async Task LoadOnlineReleasesAsync()
        {
            ResetValidation();
            SetBusy(true);
            DetailsPanel.Visibility = Visibility.Collapsed;
            _onlineRows.Clear();
            GridOnlineReleases.ItemsSource = _onlineRows;

            try
            {
                LogStep("Checking login status...");

                LoginInfo loginInfo;
                try { loginInfo = GlobalsEx.Addin?.GetLoginInfo(); }
                catch (Exception ex)
                {
                    loginInfo = null;
                    LogWarning($"Could not read login info from Addin.Core: {ex.Message}");
                }

                if (loginInfo == null || !loginInfo.IsLoggedIn || string.IsNullOrWhiteSpace(loginInfo.LoginUrl))
                {
                    LogWarning("Not logged in - switch to Offline mode.");
                    return;
                }

                string url = loginInfo.LoginUrl.TrimEnd('/') + VersionsListPath;
                LogStep($"Fetching release list from {url}...");

                string listJson;
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", loginInfo.LoginToken);
                    listJson = await client.GetStringAsync(url);
                }

                LogStep("Parsing release list...");
                var parser = new VersionParser();
                var parsedList = parser.ParseVersionJson(listJson);
                if (!parsedList.Success || parsedList.AllVersions == null || parsedList.AllVersions.Count == 0)
                {
                    LogFailure($"Could not parse the release list: {parsedList.ErrorMessage ?? "empty response"}");
                    return;
                }

                var localEntries = ReleaseHistoryStore.ReadAll(GlobalsEx.Context.Paths.ReleaseHistoryFile);
                var rows = OnlineReleaseClassifier.Classify(
                    parsedList.AllVersions,
                    localEntries,
                    GlobalsEx.Context.Version,
                    GlobalsEx.Context.ReleaseDate);

                foreach (var row in rows)
                {
                    row.PropertyChanged += OnlineRow_PropertyChanged;
                    _onlineRows.Add(row);
                }

                int newCount = rows.Count(r => r.IsSelectable);
                if (newCount == 0)
                    LogStep("No new releases available on the server.");
                else
                    LogSuccess($"Found {newCount} release(s) not yet on this machine.");

                UpdateFetchButtonState();
            }
            catch (Exception ex)
            {
                LogFailure($"Failed to load the release list: {ex.Message}", ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void OnlineRow_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(OnlineReleaseRow.IsChecked))
                UpdateFetchButtonState();
        }

        private void UpdateFetchButtonState()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(UpdateFetchButtonState);
                return;
            }

            int count = _onlineRows.Count(r => r.IsChecked);
            BtnFetchAndReload.Content = count > 0 ? $"Fetch Selected & Reload ({count})" : "Fetch Selected & Reload";
            BtnFetchAndReload.IsEnabled = count > 0;
        }

        private void GridOnlineReleases_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var row = GridOnlineReleases.SelectedItem as OnlineReleaseRow;
            if (row == null)
            {
                DetailsPanel.Visibility = Visibility.Collapsed;
                return;
            }

            TxtDetailVersion.Text = row.Version;
            TxtDetailReleased.Text = row.ReleaseDate;
            TxtDetailStatus.Text = row.StatusText;
            TxtDetailNotes.Text = row.Notes;
            TxtDetailMandatory.Text = row.Mandatory ? "Yes" : "No";
            TxtDetailFolderName.Text = row.FolderName;
            TxtDetailFileName.Text = row.FileName;
            TxtDetailChecksum.Text = row.Checksum;
            DetailsPanel.Visibility = Visibility.Visible;
        }

        private void BtnFetchAndReload_Click(object sender, RoutedEventArgs e)
        {
            // Implemented in Task 7 of the implementation plan.
        }
```

(`IsStrictlyNewer`/`PromptNoUpdate` are left untouched by this task - Offline mode
still calls both, at their existing line numbers.)

- [ ] **Step 6: Rebuild and manually verify with a fixture list response**

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense\GLSense.csproj" /p:Configuration=Debug /p:SignAssembly=false /p:RegisterForComInterop=false /nologo /v:minimal
```

Expected: `Build succeeded.` No real server exists yet, so full manual UI
verification (opening the actual window in Excel) happens once Task 7 also lands and
the whole Online flow can be exercised together end-to-end against a temporary
fixture URL swapped into `VersionsListPath`'s base - do not attempt to launch Excel
for this task alone.

- [ ] **Step 7: Commit**

```bash
git add AIPowered/GLSense/GLSense/Views/GLReloadSourcePicker.xaml AIPowered/GLSense/GLSense/Views/GLReloadSourcePicker.xaml.cs
git commit -m "Replace single-version Online check with a classified multi-release list"
```

---

## Task 7: Download loop and dialog completion

**Files:**
- Modify: `GLSense\Views\GLReloadSourcePicker.xaml.cs`

**Interfaces:**
- Consumes: `UpdateBootstrapper.ExtractAndCatalog` (Task 2/3),
  `OnlineReleaseClassifier.ParseReleaseDateOrMin` (Task 5), `_onlineRows`/
  `BtnFetchAndReload`/`LogStep`/`LogSuccess`/`LogFailure`/`SetBusy` (Task 6 and
  pre-existing).
- Produces: the real `BtnFetchAndReload_Click` implementation - fetches
  manifest+zip per checked row, calls `ExtractAndCatalog`, sets `SelectedSource`,
  closes the dialog on at least one success. This is the last piece; after this
  task, `AddinModule.RibReload_OnClick`'s existing, unchanged call to
  `new UpdateBootstrapper().ResolveVersionToLoad(GlobalsEx.Context, source)` adopts
  whatever is now newest with no further code changes anywhere.

- [ ] **Step 1: Write the failing verification script**

This exercises the extraction/cataloging half of the loop directly (the HTTP calls
themselves can't be fixture-tested without a running server or a swapped-in base
URL, which is a manual, not scripted, verification - see Step 4 below). Save to
`%TEMP%\glsense_task7_verify.ps1` - simulates "the loop already has manifestJson and
zipBytes for two checked rows, one with a good checksum and one with a bad one" and
confirms exactly one gets catalogued while the batch doesn't abort:

```powershell
$base = "D:\SQLLite_Test\GLSense\AIPowered\GLSense"
Add-Type -Path "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll"
Add-Type -Path "$base\GLSense.Shared\bin\Debug\GLSense.Shared.dll"
Add-Type -Path "$base\GLSense.Loader.Core\bin\Debug\GLSense.Loader.Core.dll"

$scratch = Join-Path $env:TEMP "glsense_task7_scratch"
if (Test-Path $scratch) { Remove-Item $scratch -Recurse -Force }
New-Item -ItemType Directory -Path $scratch | Out-Null
$versionsPath = Join-Path $scratch "Versions"
$releaseHistoryFile = Join-Path $scratch "ReleaseHistory.json"
New-Item -ItemType Directory -Path $versionsPath | Out-Null

function New-FixtureZip($version) {
    $stageDir = Join-Path $scratch "stage_$version"
    New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
    Set-Content -Path (Join-Path $stageDir "Dummy.txt") -Value "fixture $version"
    $zipPath = Join-Path $scratch "$version.zip"
    Compress-Archive -Path "$stageDir\*" -DestinationPath $zipPath -Force
    return $zipPath
}

$goodZip = New-FixtureZip "7.0.0"
$goodChecksum = (Get-FileHash -Algorithm SHA256 $goodZip).Hash
$goodManifest = "[{`"version`":`"7.0.0`",`"releaseDate`":`"2026-04-01T00:00:00`",`"folderName`":`"v7.0.0_2026-04-01T00-00-00`",`"fileName`":`"v7.0.0_2026-04-01T00-00-00.zip`",`"checksum`":`"$goodChecksum`",`"notes`":`"fixture`",`"mandatory`":false}]"
$goodBytes = [System.IO.File]::ReadAllBytes($goodZip)

$badZip = New-FixtureZip "7.0.1"
$badManifest = "[{`"version`":`"7.0.1`",`"releaseDate`":`"2026-04-02T00:00:00`",`"folderName`":`"v7.0.1_2026-04-02T00-00-00`",`"fileName`":`"v7.0.1_2026-04-02T00-00-00.zip`",`"checksum`":`"0000000000000000000000000000000000000000000000000000000000000000`",`"notes`":`"fixture`",`"mandatory`":false}]"
$badBytes = [System.IO.File]::ReadAllBytes($badZip)

Add-Type -ReferencedAssemblies @(
    "$base\GLSense.Contracts\bin\Debug\GLSense.Contracts.dll",
    "System.Runtime.dll"
) -TypeDefinition @"
using System;
using System.Collections.Generic;
using GLSense.Contracts;

public class FakeLogger : ILogger {
    public void LogInfo(string msg) {}
    public void LogWarn(string msg) {}
    public void LogError(string msg, Exception ex = null) {}
    public void LogDebug(string msg) {}
    public void LogException(Exception ex, string context = "") {}
    public void LogRawJson(string context, string rawJson) {}
    public void FlushDebugLogs(string section = "Buffered Logs") {}
    public void LogMethodEntry(string methodName = "") {}
    public void LogMethodExit(string methodName = "") {}
    public IDisposable BeginLogScope(string scopeName) { return null; }
}

public class FakePathProvider : IPathProvider {
    public string VersionsPath { get; set; }
    public string ReleaseHistoryFile { get; set; }
    public string Root { get { throw new NotImplementedException(); } }
    public string Logs { get { throw new NotImplementedException(); } }
    public string Database { get { throw new NotImplementedException(); } }
    public string LoginBrowserPath { get { throw new NotImplementedException(); } }
    public string DrilldownBrowserPath { get { throw new NotImplementedException(); } }
    public string Temp { get { throw new NotImplementedException(); } }
    public string UrlsDirectory { get { throw new NotImplementedException(); } }
    public string Resources { get { throw new NotImplementedException(); } }
    public string ManifestFile { get { throw new NotImplementedException(); } }
    public string ManifestDirectory { get { throw new NotImplementedException(); } }
    public string LatestVersion { get { throw new NotImplementedException(); } }
    public string LatestReleaseDate { get { throw new NotImplementedException(); } }
    public IReadOnlyList<VersionInfo> AllVersions { get { return null; } }
    public string LatestDownloadUrl { get { throw new NotImplementedException(); } }
    public string LatestChecksum { get { throw new NotImplementedException(); } }
    public string LatestNotes { get { throw new NotImplementedException(); } }
    public bool LatestMandatory { get { throw new NotImplementedException(); } }
    public void Ensure() {}
    public void Refresh() {}
    public void WriteManifest(VersionInfo info) {}
}

public class FakeContext : IGLSenseContext {
    public object ExcelApp { get { return null; } }
    public ILogger Logger { get; set; }
    public IPathProvider Paths { get; set; }
    public IRibbonController RibbonController { get { return null; } }
    public IntPtr ExcelHandle { get; set; }
    public bool DebugMode { get; set; }
    public string Version { get; set; }
    public string ReleaseDate { get; set; }
    public string ActiveFolderName { get; set; }
    public IReadOnlyList<VersionInfo> AllVersions { get { return null; } }
    public void SetRibbonController(IRibbonController controller) {}
    public object EdgeAddinInstance { get; set; }
    public object GetEdgeAddinInstance() { return null; }
}
"@

$paths = New-Object FakePathProvider
$paths.VersionsPath = $versionsPath
$paths.ReleaseHistoryFile = $releaseHistoryFile
$context = New-Object FakeContext
$context.Logger = New-Object FakeLogger
$context.Paths = $paths

$bootstrapper = New-Object GLSense.Loader.Core.UpdateBootstrapper

# Simulates BtnFetchAndReload_Click's per-row loop body: one good, one bad, batch
# continues past the failure.
$succeeded = 0
foreach ($pair in @(@{m=$goodManifest; z=$goodBytes}, @{m=$badManifest; z=$badBytes})) {
    $resolved = $bootstrapper.ExtractAndCatalog($context, $pair.m, $pair.z, "Online")
    if ($resolved -ne $null) { $succeeded++ }
}

$goodExtracted = Test-Path (Join-Path $versionsPath "v7.0.0_2026-04-01T00-00-00\Dummy.txt")
$badExtracted = Test-Path (Join-Path $versionsPath "v7.0.1_2026-04-02T00-00-00")
$catalog = Get-Content $releaseHistoryFile -Raw

if ($succeeded -eq 1 -and $goodExtracted -and -not $badExtracted -and ($catalog -match "7\.0\.0") -and -not ($catalog -match "7\.0\.1")) {
    Write-Output "PASS"
} else {
    Write-Output "FAIL: succeeded=$succeeded goodExtracted=$goodExtracted badExtracted=$badExtracted catalog=$catalog"
}
```

- [ ] **Step 2: Run it, confirm it currently fails or is inapplicable**

```powershell
powershell -NoProfile -File "$env:TEMP\glsense_task7_verify.ps1"
```

Expected: this actually already PASSES at this point, since it only exercises
`ExtractAndCatalog` (already implemented in Task 2/3) directly, not the not-yet-written
`BtnFetchAndReload_Click`. That's fine and expected - this script validates the exact
per-row batch semantics (`continue` past one failure, catalog only the successful
one) that `BtnFetchAndReload_Click`'s real loop must reproduce. Treat this as the
acceptance check for Step 3's implementation: if this script doesn't pass, the loop
logic below is wrong before it's ever wired to real HTTP calls.

- [ ] **Step 3: Implement `BtnFetchAndReload_Click`**

Add to the top `using` block (if not already present from Task 6):

```csharp
using GLSense.Loader.Core;
```

Add two more named constants next to `VersionsListPath` (added in Task 6):

```csharp
        private const string ManifestByFolderPath = "/glsense/manifest";
        private const string DownloadPath = "/glsense/download";
```

Replace the placeholder `BtnFetchAndReload_Click` from Task 6 with:

```csharp
        private async void BtnFetchAndReload_Click(object sender, RoutedEventArgs e)
        {
            var checkedRows = _onlineRows
                .Where(r => r.IsSelectable && r.IsChecked)
                .OrderBy(r => OnlineReleaseClassifier.ParseReleaseDateOrMin(r.ReleaseDate))
                .ToList();

            if (checkedRows.Count == 0) return;

            SetBusy(true);
            int succeeded = 0;

            try
            {
                LoginInfo loginInfo;
                try { loginInfo = GlobalsEx.Addin?.GetLoginInfo(); }
                catch (Exception ex)
                {
                    LogFailure($"Could not read login info from Addin.Core: {ex.Message}", ex);
                    return;
                }

                if (loginInfo == null || !loginInfo.IsLoggedIn || string.IsNullOrWhiteSpace(loginInfo.LoginUrl))
                {
                    LogWarning("Not logged in - switch to Offline mode.");
                    return;
                }

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", loginInfo.LoginToken);

                    foreach (var row in checkedRows)
                    {
                        try
                        {
                            LogStep($"Fetching manifest for {row.Version} ({row.ReleaseDate})...");
                            string manifestUrl = loginInfo.LoginUrl.TrimEnd('/') + ManifestByFolderPath + "?folderName=" + Uri.EscapeDataString(row.FolderName);
                            string manifestJson = await client.GetStringAsync(manifestUrl);

                            LogStep($"Downloading {row.FileName}...");
                            string zipUrl = loginInfo.LoginUrl.TrimEnd('/') + DownloadPath + "?file=" + Uri.EscapeDataString(row.FileName);
                            byte[] zipBytes = await client.GetByteArrayAsync(zipUrl);

                            var resolved = new UpdateBootstrapper().ExtractAndCatalog(GlobalsEx.Context, manifestJson, zipBytes, "Online");
                            if (resolved == null)
                            {
                                LogFailure($"Failed to catalog {row.Version} ({row.ReleaseDate}) - checksum mismatch or invalid manifest.");
                                continue;
                            }

                            LogSuccess($"Cataloged {row.Version}.");
                            succeeded++;
                        }
                        catch (Exception ex)
                        {
                            LogFailure($"Failed to fetch {row.Version} ({row.ReleaseDate}): {ex.Message}", ex);
                        }
                    }
                }

                if (succeeded == 0)
                {
                    LogFailure("No releases were fetched successfully - nothing to reload.");
                    return;
                }

                CompleteOnlineFetch(succeeded, checkedRows.Count);
            }
            finally
            {
                SetBusy(false);
            }
        }

        // DialogResult/Close are Window members with the same UI-thread affinity as
        // any other WPF DependencyObject/Window API - setting DialogResult or calling
        // Close() off the UI thread throws InvalidOperationException. By this point in
        // BtnFetchAndReload_Click, several `await`s (the per-row manifest/zip fetches)
        // have already run, and this VSTO host's WPF window doesn't reliably resume
        // continuations back onto the UI thread - so this closing sequence needs the
        // same unconditional dispatcher guard as AddOnlineRows (see Task 6's fix round
        // for the identical class of bug caught there).
        private void CompleteOnlineFetch(int succeeded, int totalChecked)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => CompleteOnlineFetch(succeeded, totalChecked));
                return;
            }

            SelectedSource = "Online";
            GlobalsEx.Context?.Logger?.LogDebug($"GLReloadSourcePicker: fetched {succeeded} of {totalChecked} selected release(s) - proceeding with reload.");
            DialogResult = true;
            Close();
        }
```

- [ ] **Step 4: Rebuild, then manually verify the full Online flow against a temporary fixture URL**

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "D:\SQLLite_Test\GLSense\AIPowered\GLSense\GLSense\GLSense.csproj" /p:Configuration=Debug /p:SignAssembly=false /p:RegisterForComInterop=false /nologo /v:minimal
```

Expected: `Build succeeded.`

Since no real server exists yet (confirmed "under process" per the design spec),
full manual verification requires a temporary stand-in for the three endpoints. Do
NOT resurrect a local HTTP host for this (the previous `GLSense.LocalUpdateHost`
approach was explicitly abandoned after real connection failures - see this
project's own history). Instead, temporarily point `VersionsListPath`/
`ManifestByFolderPath`/`DownloadPath` at a local `HttpListener`-free stand-in by
running this once, by hand, in a debugger session or a temporary console harness
that serves three fixture files (`versions.json`, a per-release manifest, and a
`.zip`) over `http://localhost:PORT/...` for the duration of a manual test only -
revert the constants back to the real `/glsense/...` paths before committing
anything. Confirm:

1. Currently-loaded row renders with the `#E3F2FD` tint, no checkbox, "Currently
   loaded" status.
2. An already-catalogued local release renders grey, no checkbox, "Already
   Downloaded" status.
3. A genuinely new release is checkable; the newest such release is pre-checked;
   `BtnFetchAndReload`'s label/enabled state track the checked count live.
4. Clicking any row (including Loaded/Already-Downloaded ones) populates the
   details panel with all 8 fields.
5. Checking two new releases and clicking Fetch catalogs both (confirm
   `ReleaseHistory.json` gains two entries and `Versions\` gains two folders).
6. Making one selected row's manifest/zip fetch fail (e.g. point its fixture at a
   404) while another succeeds - confirm the successful one still catalogs, the
   failure is logged, and the dialog still closes/reloads.
7. Every checked row failing - confirm the dialog does not close.

- [ ] **Step 5: Commit**

```bash
git add AIPowered/GLSense/GLSense/Views/GLReloadSourcePicker.xaml.cs
git commit -m "Add the Online multi-version download loop and dialog completion"
```

---

## Self-Review Notes

- **Spec coverage:** API contract (Task 6/7's three named path constants + HTTP
  calls), row classification (Task 5), shared extract+catalog (Task 2/3), download
  loop with best-effort per-row failure handling (Task 7), UI/XAML changes (Task 6),
  threading discipline (every new method follows the existing `SetBusy`/`AppendLine`
  dispatcher-safe pattern - no new mechanism introduced), `OnlineReleaseRow` model
  (Task 4). `PromptNoUpdate`/`IsStrictlyNewer` explicitly left untouched (still used
  by Offline mode) - matches the spec's Non-goals.
- **Type consistency check performed:** `ExtractAndCatalog` returns `ResolvedRelease`
  (not `ReleaseEntry`, correcting a loose signature in the spec's own prose) to stay
  consistent with `ExtractManifestZipAndAdopt`'s existing return type - both Task 2's
  declaration and Task 7's call site use this consistently. `OnlineReleaseRow`'s
  property names (`FolderName`, `FileName`, `IsChecked`, `IsSelectable`,
  `IsCurrentlyLoaded`, `IsAlreadyDownloaded`, `StatusText`) are identical across
  Task 4 (declaration), Task 5 (classifier construction), Task 6 (XAML bindings), and
  Task 7 (loop's `row.FolderName`/`row.FileName`/`row.Version`/`row.ReleaseDate`
  reads).
- **No placeholders:** every task's code block is complete, runnable C#/XAML/
  PowerShell - no "TBD"/"add error handling"/"similar to Task N" shortcuts.
