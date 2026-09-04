# AddinCore Colocated Storage Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move `Manifest\`, `Versions\`, and `ReleaseHistory.json` out of
`%LocalAppData%\ORBIT\Excel_Logs\GLSense_Logs_New\` and into a new `AddinCore\`
folder colocated with `GLSense.dll` itself (`GLSense\bin\{Config}\AddinCore\`
in dev, the installer's product folder in production), so a future
installer's uninstall - which removes its own product folder wholesale - takes
this state with it, leaving no trash behind. `Logs\`/`Database\`/`Temp\`/
`BrowserLogs\`/`Resources\` stay exactly where they are today.

**Architecture:** `GLSense.Shared\PathProvider.cs` gains a second root
(`_installRoot`, an `AddinCore` subfolder of a configurable base path) used
only by `ManifestDirectory`/`ManifestFile`/`VersionsPath`/`ReleaseHistoryFile`;
every other property keeps deriving from the existing `_root`
(Excel_Logs-based), unchanged. `GLSenseContext.cs` configures that base path
once, from `GLSense.dll`'s own on-disk location. Every downstream consumer
(`UpdateBootstrapper`, `AddinDomainLoader`, `ReleaseHistoryStore`,
`GLReloadSourcePicker`, `GLReleaseHistoryBrowser`) already goes through
`IPathProvider` exclusively, so this is the only runtime code change needed.
On the build side, `GLSense.Addin.Core\post_build.cmd` now writes its
zip+manifest.json to one uniform `SetupFiles\{Config}\Manifest\` staging
folder (was: Debug->AppData, Release->a separate project folder);
`GLSense\post_build.cmd` copies from there into its own
`bin\{Config}\AddinCore\Manifest\`; a new build-order-only `ProjectReference`
from `GLSense.csproj` to `GLSense.Addin.Core.csproj` guarantees Addin.Core (and
its post-build) finishes first, in every build entry point - standalone,
whole-solution, or (primarily) `GLSense.Build.csproj` alone.

**Tech Stack:** .NET Framework 4.8.1, classic (non-SDK) MSBuild `.csproj`
projects, batch (`.cmd`) post-build scripts, `System.Text.Json`, .NET Remoting
(`MarshalByRefObject`) across an AppDomain boundary. No automated test project
exists in this solution; verification is via real MSBuild builds and a small
throwaway `csc.exe`-compiled console harness for the one piece of pure,
testable logic (`PathProvider`'s path resolution) - this matches the
established verification convention used throughout this codebase's history
(CLAUDE.md sections 36-44: "build-verified" via real rebuilds, not unit
tests).

**Spec:** `docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md`

## Global Constraints

- AIPowered only. Do not touch `FinalWorkingCode` - it has no AppDomain/hot-reload
  architecture for this to apply to.
- No migration of existing data at the old AppData location. Old
  `Manifest\`/`Versions\`/`ReleaseHistory.json` under `Excel_Logs` become dead,
  unread folders after this ships - do not write cleanup code for them.
- `Logs\`, `Database\`, `Temp\`, `BrowserLogs\`, `Resources\`, `UrlsDirectory`
  in `PathProvider.cs` must remain byte-for-byte unchanged (still derived from
  the existing `_root`).
- New container folder name is `AddinCore` (PascalCase), holding `Manifest\`,
  `Versions\`, and `ReleaseHistory.json` - same sub-structure as today, just
  relocated.
- `GLSense.Addin.Core\SetupFiles\{Config}\` is transient build output only -
  gitignored, never committed, regenerated every build.
- There is exactly ONE copy destination for the zip+manifest:
  `GLSense\bin\{Config}\AddinCore\Manifest\`.
- `GLSense.Build.csproj` must be a correct, complete, single build entry point
  - building it alone (never Addin.Core and GLSense separately) must produce
  the full correct end state.
- All verification builds in this plan use `-p:Configuration=Debug
  -p:SignAssembly=false` (Debug mode skips Authenticode signing entirely per
  `sign_file.cmd`'s own early-exit for non-Release; `SignAssembly=false` works
  around a pre-existing, unrelated missing-.pfx-passphrase issue in this
  sandbox - confirmed via a real build in this session, not assumed). MSBuild
  is at `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\
  Current\Bin\MSBuild.exe`. Invoke its flags with a single leading dash
  (`-p:`, `-nologo`, `-v:minimal`) instead of `/p:` etc. when run through Git
  Bash - a leading `/` gets mistaken for a POSIX path and mangled by MSYS path
  conversion (confirmed in this session: `/nologo` became
  `C:/Program Files/Git/nologo`).
- `csc.exe` (classic .NET Framework C# compiler, needed for the Task 1
  harness) is at `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`.

---

### Task 1: Split `PathProvider`'s roots and configure it from `GLSense.dll`'s own location

**Files:**
- Modify: `GLSense.Shared\PathProvider.cs`
- Modify: `GLSense\GLSenseContext.cs`
- Scratch (not committed): a `csc`-compiled verification harness, written to
  the scratchpad directory.

**Interfaces:**
- Produces: `PathProvider.ConfigureInstallRoot(string installRoot)` (new
  public static method) - must be called before the `PathProvider` instance
  whose paths matter is constructed. `IPathProvider.ManifestDirectory` /
  `.VersionsPath` / `.ReleaseHistoryFile` now resolve to
  `Path.Combine(installRoot, "AddinCore", ...)` instead of the old
  `Excel_Logs`-based root, where `installRoot` is whatever was last passed to
  `ConfigureInstallRoot` (or the historical `Excel_Logs` root if never
  called). `IPathProvider.ManifestFile` is unchanged in shape (still
  `Path.Combine(ManifestDirectory, "manifest.json")`) - only where
  `ManifestDirectory` itself points changes.
- Consumes: nothing new from other tasks - this task is self-contained and
  is a prerequisite for none of the others (they only depend on the build
  scripts / project files, not on this C# change directly), but should land
  first since it's the riskiest, most logic-heavy piece.

- [ ] **Step 1: Build `GLSense.Contracts` and `GLSense.Shared` in Debug, to get baseline DLLs**

Run:
```bash
cd "D:/SQLLite_Test/GLSense/AIPowered/GLSense"
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" "GLSense.Contracts/GLSense.Contracts.csproj" -p:Configuration=Debug -p:SignAssembly=false -nologo -v:minimal
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" "GLSense.Shared/GLSense.Shared.csproj" -p:Configuration=Debug -p:SignAssembly=false -nologo -v:minimal
```
Expected: both builds print `... -> ...\bin\Debug\GLSense.Contracts.dll` /
`...\GLSense.Shared.dll` with no errors.

- [ ] **Step 2: Write the verification harness (will fail to compile against the CURRENT, unmodified code)**

Create `<scratchpad>/PathProviderHarness.cs` (use the scratchpad directory
path from your own environment, e.g.
`C:\Users\RAMAMA~1\AppData\Local\Temp\claude\...\scratchpad\PathProviderHarness.cs`):

```csharp
using GLSense.Shared;
using System;
using System.IO;

class Harness
{
    static int Main()
    {
        string fakeInstallDir = Path.Combine(Path.GetTempPath(), "GLSense_PathProviderHarness_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(fakeInstallDir);
        bool ok = true;
        try
        {
            PathProvider.ConfigureInstallRoot(fakeInstallDir);

            // First instance
            var p1 = new PathProvider();
            string expectedManifestDir = Path.Combine(fakeInstallDir, "AddinCore", "Manifest");
            string expectedVersionsPath = Path.Combine(fakeInstallDir, "AddinCore", "Versions");
            string expectedReleaseHistoryFile = Path.Combine(fakeInstallDir, "AddinCore", "ReleaseHistory.json");
            string expectedManifestFile = Path.Combine(expectedManifestDir, "manifest.json");

            ok &= Check("ManifestDirectory", p1.ManifestDirectory, expectedManifestDir);
            ok &= Check("VersionsPath", p1.VersionsPath, expectedVersionsPath);
            ok &= Check("ReleaseHistoryFile", p1.ReleaseHistoryFile, expectedReleaseHistoryFile);
            ok &= Check("ManifestFile", p1.ManifestFile, expectedManifestFile);

            // Logs/Database/Root must be UNCHANGED - still under the old Excel_Logs root.
            string expectedLogsRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ORBIT", "Excel_Logs", "GLSense_Logs_New");
            ok &= Check("Root", p1.Root, expectedLogsRoot);
            ok &= Check("Logs", p1.Logs, Path.Combine(expectedLogsRoot, "Logs"));
            ok &= Check("Database", p1.Database, Path.Combine(expectedLogsRoot, "Database"));

            // A SECOND instance, constructed after the same ConfigureInstallRoot call,
            // must resolve identically - proving the override is a static, process-wide
            // setting (mirrors how PathProvider.Instance's separate lazy singleton must
            // still pick up the same install root).
            var p2 = new PathProvider();
            ok &= Check("Second instance ManifestDirectory", p2.ManifestDirectory, expectedManifestDir);

            Console.WriteLine(ok ? "ALL CHECKS PASSED" : "CHECKS FAILED");
            return ok ? 0 : 1;
        }
        finally
        {
            try { Directory.Delete(fakeInstallDir, true); } catch { }
        }
    }

    static bool Check(string name, string actual, string expected)
    {
        bool match = string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        Console.WriteLine((match ? "PASS " : "FAIL ") + name + ": actual='" + actual + "' expected='" + expected + "'");
        return match;
    }
}
```

- [ ] **Step 3: Compile the harness and confirm it fails (ConfigureInstallRoot doesn't exist yet)**

Run (adjust `<scratchpad>` to your actual scratchpad path; `GLSense.Shared.dll`'s
folder already contains `GLSense.Contracts.dll`, `System.Text.Json.dll`, and
every other dependency copied there by Step 1's build):
```bash
cd "D:/SQLLite_Test/GLSense/AIPowered/GLSense/GLSense.Shared/bin/Debug"
cp "<scratchpad>/PathProviderHarness.cs" .
"/c/Windows/Microsoft.NET/Framework64/v4.0.30319/csc.exe" /nologo /reference:GLSense.Shared.dll /reference:GLSense.Contracts.dll /out:PathProviderHarness.exe PathProviderHarness.cs
```
Expected: FAIL - `error CS0117: 'PathProvider' does not contain a definition
for 'ConfigureInstallRoot'`.

- [ ] **Step 4: Implement `PathProvider.cs`'s split-root logic**

In `GLSense.Shared\PathProvider.cs`:

Add a static override field right after the existing static fields
(after `private static List<VersionInfo> _allVersions = new();`):

```csharp
        // Base folder that the "AddinCore" hot-reload state (Manifest/Versions/
        // ReleaseHistory.json) is colocated under - normally the folder
        // GLSense.dll itself is running from, set once via ConfigureInstallRoot
        // (see GLSenseContext's constructor). Static (not per-instance) so every
        // PathProvider ever constructed - including PathProvider.Instance's
        // separate lazy singleton - shares the same configured value. Falls back
        // to the historical Excel_Logs-based root if never configured (e.g. a
        // PathProvider constructed in a test harness with no GLSenseContext).
        private static string _installRootOverride;

        /// <summary>
        /// Sets the folder that Manifest/Versions/ReleaseHistory.json are
        /// colocated under (an "AddinCore" subfolder of it) - normally the
        /// folder GLSense.dll itself is running from, so a future installer's
        /// uninstall (which removes that whole folder) takes this state with
        /// it too, instead of leaving it behind under the separate Excel_Logs
        /// tree. Call this before constructing the PathProvider whose paths
        /// matter - safe to call more than once (last call wins, and applies
        /// to every PathProvider constructed afterward, since this is a
        /// static field shared process-wide, not an instance field).
        /// </summary>
        public static void ConfigureInstallRoot(string installRoot)
        {
            _installRootOverride = installRoot;
        }
```

Add the `_installRoot` instance field next to `_root`/`_basePath`:

```csharp
        private readonly string _root;
        private readonly string _basePath;
        private readonly string _installRoot;
```

In the constructor, right after `_root = Path.Combine(_basePath, "GLSense_Logs_New");`, add:

```csharp
            _installRoot = Path.Combine(_installRootOverride ?? _root, "AddinCore");
```

Change these three properties to derive from `_installRoot` instead of `_root`
(everything else - `Root`, `UrlsDirectory`, `Logs`, `Database`, `Temp`,
`LoginBrowserPath`, `DrilldownBrowserPath`, `Resources` - stays exactly as-is):

```csharp
        public string VersionsPath => Path.Combine(_installRoot, "Versions");
```
```csharp
        // "Manifest" (not "Version") since this folder/file is the update-tracking
        // record (releaseDate/version/downloadUrl/etc.), distinct from "Versions" (plural)
        // which holds the actual hot-reloadable DLL payloads. Colocated with GLSense.dll's
        // own folder (via _installRoot), NOT the Excel_Logs tree - see
        // docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md.
        public string ManifestDirectory => Path.Combine(_installRoot, "Manifest");
        public string ManifestFile => Path.Combine(ManifestDirectory, "manifest.json");
        public string ReleaseHistoryFile => Path.Combine(_installRoot, "ReleaseHistory.json");
```

- [ ] **Step 5: Implement `GLSenseContext.cs`'s configuration call**

In `GLSense\GLSenseContext.cs`, add `using System.IO;` to the usings (it
currently has `GLSense.Contracts`, `GLSense.Shared`, `System`,
`System.Collections.Generic`, `System.Reflection` - no `System.IO`).

Change the constructor from:

```csharp
        public GLSenseContext(object app)
        {
            ExcelApp = app ?? throw new ArgumentNullException(nameof(app));

            // Initialize PathProvider
            Paths = new PathProvider();
            Paths.Ensure();
```

to:

```csharp
        public GLSenseContext(object app)
        {
            ExcelApp = app ?? throw new ArgumentNullException(nameof(app));

            // The AddinCore hot-reload state (Manifest/Versions/ReleaseHistory.json)
            // is colocated with GLSense.dll's own folder so a future installer's
            // uninstall (which removes this whole folder) takes that state with it
            // too, instead of leaving it behind under the separate Excel_Logs tree -
            // see docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md.
            // Deliberately Assembly.GetExecutingAssembly() (this class's own assembly,
            // GLSense.dll) rather than AppDomain.CurrentDomain.BaseDirectory, which can
            // resolve to Excel's own directory rather than the add-in's when hosted via COM.
            string glsenseAssemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            PathProvider.ConfigureInstallRoot(glsenseAssemblyDir);

            // Initialize PathProvider
            Paths = new PathProvider();
            Paths.Ensure();
```

- [ ] **Step 6: Rebuild `GLSense.Contracts` and `GLSense.Shared`**

Run the same two build commands from Step 1 again.
Expected: both succeed with no errors.

- [ ] **Step 7: Recompile and run the harness, confirm it now passes**

```bash
cd "D:/SQLLite_Test/GLSense/AIPowered/GLSense/GLSense.Shared/bin/Debug"
cp "<scratchpad>/PathProviderHarness.cs" .
"/c/Windows/Microsoft.NET/Framework64/v4.0.30319/csc.exe" /nologo /reference:GLSense.Shared.dll /reference:GLSense.Contracts.dll /out:PathProviderHarness.exe PathProviderHarness.cs
./PathProviderHarness.exe
```
Expected: every `Check` line prints `PASS`, final line is `ALL CHECKS
PASSED`, exit code 0.

Note: this run creates real folders under
`%LocalAppData%\ORBIT\Excel_Logs\GLSense_Logs_New\` (Logs/Database/Temp/
BrowserLogs/Resources) as a side effect of `PathProvider`'s constructor
calling `Ensure()` - this is pre-existing behavior, unrelated to this
change, and non-destructive (only creates folders, never deletes anything).

- [ ] **Step 8: Commit**

```bash
git add GLSense.Shared/PathProvider.cs GLSense/GLSenseContext.cs
git commit -m "Colocate AddinCore Manifest/Versions/ReleaseHistory under the GLSense assembly folder

Splits PathProvider's single root into two: the existing Excel_Logs-based
root (Logs/Database/Temp/BrowserLogs/Resources, unchanged) and a new
install-root-based AddinCore folder (Manifest/Versions/ReleaseHistory.json),
configured from GLSense.dll's own on-disk location. A future installer's
uninstall, which removes its own product folder, now takes this state with
it instead of leaving it behind under the separate Excel_Logs tree.

See docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md"
```

---

### Task 2: `GLSense.Addin.Core\post_build.cmd` writes to a single `SetupFiles\{Config}\Manifest\` folder

**Files:**
- Modify: `GLSense.Addin.Core\post_build.cmd`
- Modify: `.gitignore`

**Interfaces:**
- Consumes: nothing from Task 1 (build scripts don't depend on the C# change).
- Produces: `GLSense.Addin.Core\SetupFiles\{Config}\Manifest\manifest.json` and
  `v{version}.zip` - Task 4 depends on this exact path shape to know what to
  `xcopy` from.

- [ ] **Step 1: Add the `.gitignore` entry**

Add this line to `.gitignore` (create the section if none exists for build
output; place near any existing `bin/`/`obj/` ignore entries):

```
GLSense.Addin.Core/SetupFiles/
```

- [ ] **Step 2: Replace the header comment describing the Debug/Release split**

In `GLSense.Addin.Core\post_build.cmd`, replace lines 33-58 (the
`REM ====...` header block starting with "FOLDER-ONLY testing flow" and
ending with "not subject to that same caveat.") with:

```bat
REM ============================================================================
REM Both Debug and Release write a fresh zip + manifest.json into this
REM project's own SetupFiles\{Config}\Manifest\ folder - transient build
REM output (gitignored, not committed, regenerated every build). This is a
REM staging/hand-off point only: GLSense\post_build.cmd (the host project)
REM copies from here into its own bin\{Config}\AddinCore\Manifest\, which is
REM what PathProvider.ManifestDirectory now resolves to at runtime (colocated
REM with GLSense.dll itself, not GLSense_Logs_New) - see
REM docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md.
REM
REM Because GLSense.csproj has a build-order-only ProjectReference to this
REM project (Private=False - see GLSense.csproj), a build of GLSense.Build,
REM GLSense.sln, or GLSense.csproj alone always runs this project's
REM post_build.cmd (refreshing SetupFiles) before GLSense's own post_build.cmd
REM tries to copy from it. A standalone build of ONLY this project (e.g. the
REM Reload button's fast Addin.Core-only iteration loop) still refreshes
REM SetupFiles - GLReloadSourcePicker's Offline mode can browse directly to
REM SetupFiles\{Config}\Manifest\ without needing GLSense to be rebuilt too.
REM ============================================================================
```

- [ ] **Step 3: Replace the Debug/Release branch with one uniform destination**

Replace lines 114-126 (from the `REM Debug -> local %LOCALAPPDATA%...` comment
through `echo Manifest Output Dir: %MANIFEST_DIR%`) with:

```bat
REM One destination for both configurations - see the header comment above.
set MANIFEST_DIR=%PROJECT_DIR%\SetupFiles\%CONFIG%\Manifest
set OUT_ZIP=%MANIFEST_DIR%\v%FILE_VERSION%.zip
set OUT_MANIFEST=%MANIFEST_DIR%\manifest.json

echo Manifest Output Dir: %MANIFEST_DIR%
```

- [ ] **Step 4: Simplify the now-stale "deliberate override" comment**

Replace the comment block right before `if not exist "%MANIFEST_DIR%" mkdir
"%MANIFEST_DIR%"` (currently: "Manifest folder is created here on purpose ...
not covered by that rule at all.") with:

```bat
REM SetupFiles is this project's own folder, not GLSense_Logs_New - creating
REM it here is ordinary build output, not subject to section 14.5's
REM "build shouldn't create GLSense_Logs_New folders" rule at all.
```

- [ ] **Step 5: Build `GLSense.Addin.Core` and confirm the new output location**

```bash
cd "D:/SQLLite_Test/GLSense/AIPowered/GLSense"
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" "GLSense.Addin.Core/GLSense.Addin.Core.csproj" -p:Configuration=Debug -p:SignAssembly=false -nologo -v:minimal
ls -la "GLSense.Addin.Core/SetupFiles/Debug/Manifest/"
```
Expected: build succeeds; the `ls` shows `manifest.json` and a
`v{version}.zip` file with a fresh timestamp (matches the build just run).

- [ ] **Step 6: Confirm the old AppData location is no longer being written to**

```bash
ls -la "/c/Users/RamaMaidam_2fc/AppData/Local/ORBIT/Excel_Logs/GLSense_Logs_New/Manifest/"
```
Expected: the files there (from earlier baseline builds, before this task's
change) keep their OLD timestamp - this build did not touch them. (It's fine
if this folder still exists with stale content - per the spec's Non-goals,
no migration/cleanup of the old location is in scope.)

- [ ] **Step 7: Commit**

```bash
git add GLSense.Addin.Core/post_build.cmd .gitignore
git commit -m "Addin.Core post_build.cmd: write manifest+zip to SetupFiles, not AppData

Both Debug and Release now write to one uniform SetupFiles\{Config}\Manifest\
staging folder instead of the old Debug->AppData / Release->Manifests split.
GLSense's own post_build.cmd will copy from here into its bin\AddinCore\
Manifest\ - see docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md"
```

---

### Task 3: Build-order-only `ProjectReference` from `GLSense.csproj` to `GLSense.Addin.Core.csproj`

**Files:**
- Modify: `GLSense\GLSense.csproj`

**Interfaces:**
- Consumes: nothing directly, but Task 4's `xcopy` step depends on this
  reference existing so that `GLSense.Addin.Core`'s post-build (and therefore
  its `SetupFiles` output) has always already run by the time `GLSense`'s own
  post-build fires.
- Produces: a guarantee that building `GLSense.csproj` (standalone, via the
  solution, or via `GLSense.Build.csproj`) always builds `GLSense.Addin.Core`
  first, without copying `GLSense.Addin.Core.dll` into `GLSense`'s own bin
  output.

- [ ] **Step 1: Add the reference**

In `GLSense\GLSense.csproj`, inside the existing `<ItemGroup>` that already
contains the `ProjectReference`s to `GLSense.Contracts`/`GLSense.Loader.Core`/
`GLSense.Shared` (currently ending at `</ItemGroup>` right after the
`GLSense.Shared` reference), add a new entry:

```xml
    <!-- Not used by any code here - GLSense.dll never calls into
         GLSense.Addin.Core directly (that only ever happens via a child
         AppDomain, loaded at runtime from a Versions\ hot-reload folder).
         Added purely so MSBuild always builds GLSense.Addin.Core - and runs
         its post_build.cmd, populating SetupFiles - before this project's
         own post_build.cmd tries to copy from SetupFiles. Private=False so
         GLSense.Addin.Core.dll is NOT copied into this project's own bin
         output - unlike the GLSense.Loader.Core reference below (which
         intentionally IS copied, since AddinDomainLoader needs it alongside
         whatever gets zipped), Addin.Core.dll must only ever be loaded from a
         Versions\ folder, never sit directly next to GLSense.dll. See
         docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md. -->
    <ProjectReference Include="..\GLSense.Addin.Core\GLSense.Addin.Core.csproj">
      <Project>{76F5002B-6A27-4B68-ADEE-EB673A363228}</Project>
      <Name>GLSense.Addin.Core</Name>
      <Private>False</Private>
    </ProjectReference>
```

- [ ] **Step 2: Build `GLSense.csproj` standalone and confirm ordering + no DLL copy**

```bash
cd "D:/SQLLite_Test/GLSense/AIPowered/GLSense"
rm -f "GLSense.Addin.Core/bin/Debug/GLSense.Addin.Core.dll"
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" "GLSense/GLSense.csproj" -p:Configuration=Debug -p:SignAssembly=false -nologo -v:minimal
ls -la "GLSense.Addin.Core/bin/Debug/GLSense.Addin.Core.dll"
ls "GLSense/bin/Debug/" | grep -i "Addin.Core" || echo "CONFIRMED: not copied into GLSense's bin output"
```
Expected: the first `ls` shows `GLSense.Addin.Core.dll` was rebuilt (proving
MSBuild built it as part of building `GLSense.csproj`, even though it was
deleted first); the `grep` finds nothing, and the `echo` fallback prints -
confirming `Private=False` prevented the copy.

- [ ] **Step 3: Commit**

```bash
git add GLSense/GLSense.csproj
git commit -m "GLSense.csproj: add build-order-only reference to GLSense.Addin.Core

Private=False - guarantees Addin.Core (and its post-build, which populates
SetupFiles) always finishes before GLSense's own post-build runs, without
copying GLSense.Addin.Core.dll into GLSense's bin output. Needed so building
GLSense.csproj (or GLSense.Build.csproj) alone is always sufficient - see
docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md"
```

---

### Task 4: `GLSense\post_build.cmd` copies the manifest+zip into its own bin output

**Files:**
- Modify: `GLSense\post_build.cmd`

**Interfaces:**
- Consumes: `GLSense.Addin.Core\SetupFiles\{Config}\Manifest\` (Task 2's
  output), and relies on Task 3's `ProjectReference` for build ordering.
- Produces: `GLSense\bin\{Config}\AddinCore\Manifest\manifest.json` +
  `v{version}.zip` - the exact location Task 1's `PathProvider.
  ManifestDirectory` now resolves to at runtime.

- [ ] **Step 1: Add the copy step**

In `GLSense\post_build.cmd`, add this block right after the existing three
`call "%SOLUTION_DIR%\sign_file.cmd" ...` lines (before the closing `echo
========================================` / `echo post_build.cmd (GLSense
host) completed` footer):

```bat
echo ========================================
echo Copying AddinCore manifest+zip from GLSense.Addin.Core's SetupFiles
echo ========================================

set ADDINCORE_SOURCE=%SOLUTION_DIR%\GLSense.Addin.Core\SetupFiles\%CONFIG%\Manifest
set ADDINCORE_DEST=%TARGET_DIR%\AddinCore\Manifest

if not exist "%ADDINCORE_SOURCE%" (
    echo WARNING: %ADDINCORE_SOURCE% not found - GLSense.Addin.Core may not have
    echo been built in this configuration yet. Skipping AddinCore manifest copy.
    echo Build GLSense.Build.csproj instead of GLSense.csproj alone to avoid this.
    goto :SkipAddinCoreCopy
)

if not exist "%ADDINCORE_DEST%" mkdir "%ADDINCORE_DEST%"
xcopy /Y /I "%ADDINCORE_SOURCE%\*" "%ADDINCORE_DEST%\"

echo Copied to: %ADDINCORE_DEST%

:SkipAddinCoreCopy
```

- [ ] **Step 2: Build `GLSense.csproj` and confirm the final layout**

```bash
cd "D:/SQLLite_Test/GLSense/AIPowered/GLSense"
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" "GLSense/GLSense.csproj" -p:Configuration=Debug -p:SignAssembly=false -nologo -v:minimal
ls -la "GLSense/bin/Debug/AddinCore/Manifest/"
diff <(md5sum "GLSense/bin/Debug/AddinCore/Manifest/manifest.json" | cut -d' ' -f1) <(md5sum "GLSense.Addin.Core/SetupFiles/Debug/Manifest/manifest.json" | cut -d' ' -f1) && echo "CONFIRMED: identical content"
```
Expected: `GLSense\bin\Debug\AddinCore\Manifest\` contains `manifest.json` and
a `v{version}.zip`; the `diff`+`echo` confirms the copied `manifest.json` is
byte-identical to Addin.Core's `SetupFiles` source.

- [ ] **Step 3: Commit**

```bash
git add GLSense/post_build.cmd
git commit -m "GLSense post_build.cmd: copy AddinCore manifest+zip into own bin output

Copies from GLSense.Addin.Core's SetupFiles staging folder into
bin\{Config}\AddinCore\Manifest\ - the location PathProvider.ManifestDirectory
now resolves to at runtime. See
docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md"
```

---

### Task 5: Fix `GLSense.Build\post_build.cmd`'s summary and verify the single build entry point end to end

**Files:**
- Modify: `GLSense.Build\post_build.cmd`

**Interfaces:**
- Consumes: everything from Tasks 1-4. This is the end-to-end integration
  check for the whole plan.
- Produces: nothing further downstream - this is the final task.

- [ ] **Step 1: Fix the summary printer**

In `GLSense.Build\post_build.cmd`, replace the `if /I "%CONFIG%"=="Release"
(...) else (...)` block (lines 37-45: currently reports the old AppData path
for Release and incorrectly claims Debug publishes no manifest zip at all -
both are stale even before this plan's change) with:

```bat
echo.
echo AddinCore hot-reload state (colocated with GLSense.dll - see
echo docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md):
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\Manifest\manifest.json
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\Manifest\v*.zip
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\Versions\   (populated at runtime by UpdateBootstrapper)
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\ReleaseHistory.json   (populated at runtime by UpdateBootstrapper)
```

- [ ] **Step 2: Clean build of `GLSense.Build.csproj` ALONE, in Debug - the primary end-to-end check**

```bash
cd "D:/SQLLite_Test/GLSense/AIPowered/GLSense"
rm -rf GLSense/bin GLSense.Addin.Core/bin GLSense.Addin.Core/SetupFiles GLSense.Contracts/bin GLSense.Shared/bin GLSense.Loader.Core/bin GLSense.Build/bin
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" "GLSense.Build/GLSense.Build.csproj" -p:Configuration=Debug -p:SignAssembly=false -nologo -v:minimal
```
Expected: succeeds with no errors, and the tail of the output shows the new
summary block from Step 1 with real, resolved paths (not the old AppData
line).

- [ ] **Step 3: Confirm the full resulting directory shape matches the spec**

```bash
find "GLSense/bin/Debug/AddinCore" -type f
```
Expected: exactly
`GLSense/bin/Debug/AddinCore/Manifest/manifest.json` and
`GLSense/bin/Debug/AddinCore/Manifest/v{version}.zip` (no `Versions\` or
`ReleaseHistory.json` yet - those are only created at runtime by
`UpdateBootstrapper`, not at build time, matching `PathProvider.Ensure()`'s
existing behavior of creating `VersionsPath` empty and `ReleaseHistoryFile`
not being pre-created at all).

- [ ] **Step 4: Clean build of `GLSense.Build.csproj` ALONE, in Release - confirm no config-specific regression**

```bash
rm -rf GLSense/bin GLSense.Addin.Core/bin GLSense.Addin.Core/SetupFiles GLSense.Contracts/bin GLSense.Shared/bin GLSense.Loader.Core/bin GLSense.Build/bin
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe" "GLSense.Build/GLSense.Build.csproj" -p:Configuration=Release -p:SignAssembly=false -nologo -v:minimal
find "GLSense/bin/Release/AddinCore" -type f
```
Expected: same shape as Step 3, under `bin\Release\` instead of `bin\Debug\`.
(Signing itself is skipped via `SignAssembly=false`/`sign_file.cmd`'s own
`signtool`-not-found fallback in this sandbox - that's an existing,
unrelated environment limitation, not something this task fixes.)

- [ ] **Step 5: Commit**

```bash
git add GLSense.Build/post_build.cmd
git commit -m "GLSense.Build post_build.cmd: report the real AddinCore output paths

Summary printer was already stale before this change (claimed Debug builds
publish no manifest zip at all, which was never true). Now reports the
actual colocated bin\AddinCore\ location for both configurations.

See docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md"
```

---

### Task 6: Document the change in `CLAUDE.md`

**Files:**
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: the finished state of Tasks 1-5 (this task only writes prose
  describing what was built).

- [ ] **Step 1: Add a new numbered section**

`CLAUDE.md` is this project's living fix/decision log (sections 1-44,
newest last). Append a new section after the last one (currently ends after
section 44's "44.9 GLSense host tooltips looked plain..." block, right before
the "## Deployment note" footer) with:

```markdown
## 45. AddinCore hot-reload state colocated with the GLSense assembly, not AppData (AIPowered only)

`Manifest\`, `Versions\`, and `ReleaseHistory.json` (sections 14-40's
hot-reload/release-history subsystem) moved from
`%LocalAppData%\ORBIT\Excel_Logs\GLSense_Logs_New\` into a new `AddinCore\`
folder colocated with `GLSense.dll` itself (`GLSense\bin\{Config}\AddinCore\`
in dev). Motivation: the planned installer (modeled on
`FinalWorkingCode\GLSense`, installing to `%LocalAppData%\ORBIT\
{Manufacturer}\{Product}\`) removes its own product folder on uninstall, but
has no reason to know about the separate `Excel_Logs` tree - left as-is, this
state would become permanent trash after every uninstall, and could collide
with a fresh reinstall. `Logs\`/`Database\`/`Temp\`/`BrowserLogs\`/
`Resources\` deliberately stay under `Excel_Logs`, unchanged - that's
diagnostic/runtime data meant to persist independently of install/uninstall.

**Runtime**: `GLSense.Shared\PathProvider.cs` now derives
`ManifestDirectory`/`VersionsPath`/`ReleaseHistoryFile` from a second root,
`_installRoot` (an `AddinCore` subfolder of a configurable base path), set via
a new static `PathProvider.ConfigureInstallRoot(string)`.
`GLSense\GLSenseContext.cs`'s constructor calls it with
`Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)` - evaluated
from code inside `GLSense.dll` itself, so it reliably resolves to
`GLSense.dll`'s own folder regardless of Excel's COM-hosting quirks (this is
deliberately NOT `AppDomain.CurrentDomain.BaseDirectory`, which can resolve to
Excel's own directory instead). Every downstream consumer
(`UpdateBootstrapper`, `AddinDomainLoader`, `ReleaseHistoryStore`,
`GLReloadSourcePicker`, `GLReleaseHistoryBrowser`) already went through
`IPathProvider` exclusively - verified by direct code inspection - so this was
the only runtime code change needed.

**Build-time**: `GLSense.Addin.Core\post_build.cmd` now writes its
zip+manifest.json to one uniform `SetupFiles\{Config}\Manifest\` staging
folder (was: Debug->AppData Manifest, Release->this project's own
`Manifests\` folder) - gitignored, transient, regenerated every build.
`GLSense\post_build.cmd` copies from there into its own
`bin\{Config}\AddinCore\Manifest\`, guarded to skip with a warning (not fail)
if Addin.Core hasn't been built yet in that configuration. A new
`Private=False` `ProjectReference` from `GLSense.csproj` to
`GLSense.Addin.Core.csproj` (build-order only, no DLL copy - mirroring the
existing Addin.Core->Loader.Core reference's *ordering* purpose, though that
one deliberately DOES copy its DLL, unlike this one) guarantees Addin.Core's
post-build always finishes first - in a standalone `GLSense.csproj` build, a
whole-solution build, or (the intended single entry point going forward)
`GLSense.Build.csproj` alone.

**Accepted dev-loop trade-off**: rebuilding only `GLSense.Addin.Core` no
longer immediately refreshes the auto-adopt location `UpdateBootstrapper`
reads on next launch/Reload - that now requires `GLSense` (or
`GLSense.Build`) to also rebuild, so its post-build can copy from Addin.Core's
`SetupFiles`. `GLReloadSourcePicker`'s Offline mode (browse-to-a-folder)
remains available for fast Addin.Core-only iteration, pointed directly at
`GLSense.Addin.Core\SetupFiles\{Config}\Manifest\`.

**No migration** of old data at the AppData location - no installer exists
yet, so there's no real installed user base to migrate. Old
`Manifest\`/`Versions\`/`ReleaseHistory.json` under `Excel_Logs` on a dev
machine simply become dead, unread folders after this shipped; delete them by
hand if desired.

Full design: `docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md`.
Full implementation plan: `docs/superpowers/plans/2026-09-04-addincore-colocated-storage.md`.

- **Status**: implemented and build-verified (`GLSense.Build.csproj` built
  clean, standalone, in both Debug and Release, with `-p:SignAssembly=false`
  as a verification-only override for this sandbox's pre-existing missing-
  .pfx-passphrase issue - unrelated to this change). Not yet tested live in
  Excel (no installer exists yet to test uninstall against directly).
```

- [ ] **Step 2: Commit**

```bash
git add CLAUDE.md
git commit -m "Document AddinCore colocated storage change in CLAUDE.md (section 45)"
```
