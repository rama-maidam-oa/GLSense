# AddinCore hot-reload state: colocate with the GLSense assembly (AIPowered only)

## Problem

`Manifest\`, `Versions\`, and `ReleaseHistory.json` (the hot-reload/release-history
subsystem from CLAUDE.md sections 14-40) currently live under
`%LocalAppData%\ORBIT\Excel_Logs\GLSense_Logs_New\`, alongside `Logs\`/`Database\`/
`Temp\`/`BrowserLogs\`/`Resources\`.

No installer exists yet, but the planned one (modeled on
`FinalWorkingCode\GLSense`) will install GLSense to
`%LocalAppData%\ORBIT\{Manufacturer}\{Product}\` - a folder *sibling* to
`Excel_Logs`, not inside it. An uninstaller removes the folder it installed to;
it has no reason to know about, or touch, the separate `Excel_Logs` tree. Left
as-is, every build/run leaves `Manifest\`, `Versions\`, and `ReleaseHistory.json`
behind permanently after uninstall - dead weight that can also collide with a
fresh reinstall (e.g. a stale `ReleaseHistory.json` referencing folders that no
longer exist, or a version-numbering collision against a reused version number,
per CLAUDE.md 14.2/40's discussion of repeated version numbers).

## Goal

Move `Manifest\`, `Versions\`, and `ReleaseHistory.json` into a new `AddinCore\`
folder that sits *next to* `GLSense.dll` itself - `GLSense\bin\{Debug|Release}\
AddinCore\` in dev, and wherever the installer places `GLSense.dll` in
production. An uninstaller that removes the install folder wholesale then takes
this state with it automatically, with zero installer-specific cleanup logic
needed.

`Logs\`, `Database\`, `Temp\`, `BrowserLogs\`, `Resources\`, and the
`ORBIT_URLS.xml` file stay exactly where they are today, under `Excel_Logs\
GLSense_Logs_New\` - unchanged, still runtime/diagnostic data that's meant to
persist independently of install/uninstall cycles.

## Non-goals

- No migration of any existing data at the old AppData location. There is no
  installer yet and therefore no real installed user base to migrate; old
  `Manifest\`/`Versions\`/`ReleaseHistory.json` under `Excel_Logs` simply become
  dead, unread folders after this ships. A developer can delete them by hand;
  nothing in this change reads or writes them going forward.
- No installer project is being created here. This change only prepares the
  runtime/build-time layout so that a future installer's uninstall step is
  correct by construction (remove the install folder = remove everything).
- FinalWorkingCode is out of scope - it has no AppDomain/hot-reload
  architecture for this to apply to.

## Design

### 1. Runtime path resolution

`GLSense.Shared\PathProvider.cs` currently derives every path from one `_root`
(`%LocalAppData%\ORBIT\Excel_Logs\GLSense_Logs_New`). A second root,
`_installRoot`, is added - used **only** by `ManifestDirectory`, `ManifestFile`,
`VersionsPath`, and `ReleaseHistoryFile`. Every other member (`Root`, `Logs`,
`Database`, `Temp`, `LoginBrowserPath`, `DrilldownBrowserPath`, `Resources`,
`UrlsDirectory`) keeps deriving from `_root`, byte-for-byte unchanged.

```
_installRoot = Path.Combine(<configured install root>, "AddinCore")

ManifestDirectory   = _installRoot\Manifest
ManifestFile        = ManifestDirectory\manifest.json      (unchanged - derives from ManifestDirectory)
VersionsPath        = _installRoot\Versions
ReleaseHistoryFile  = _installRoot\ReleaseHistory.json
```

`_installRoot`'s base value is provided via a new static method,
`PathProvider.ConfigureInstallRoot(string installRoot)`, backed by a private
static field. The instance constructor reads
`_installRoot = _installRootOverride ?? _root` (i.e. if never configured, it
falls back to the *old* Excel_Logs-based location - see the "Known limitation"
note below on why this fallback exists rather than throwing).

`GLSense\GLSenseContext.cs`'s constructor calls
`PathProvider.ConfigureInstallRoot(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location))`
before constructing `Paths = new PathProvider()`. `GLSenseContext` is a class
compiled into `GLSense.dll` itself, so `Assembly.GetExecutingAssembly()` called
from its own code reliably resolves to `GLSense.dll`'s own on-disk location -
this is deliberately *not* `AppDomain.CurrentDomain.BaseDirectory`, which can
resolve to Excel's own directory rather than the add-in's when hosted via COM.

**Why this is the only runtime code change needed:** `UpdateBootstrapper.cs`,
`AddinDomainLoader.cs` (via `IGLSenseContext.ActiveFolderName` +
`Paths.VersionsPath`), `ReleaseHistoryStore.cs`, `GLReloadSourcePicker.xaml.cs`,
and `GLReleaseHistoryBrowser.xaml.cs` were all verified (by direct code
inspection, not assumption) to go through `IPathProvider`/`GlobalsEx.Context.
Paths` exclusively - no hardcoded `%LocalAppData%` paths exist anywhere else in
the codebase. Redirecting the properties at the source redirects every
consumer transparently.

**Known limitation, accepted as-is:** `PathProvider.Instance` (a separate,
lazily-constructed static singleton used only by `GLSense.Shared\Logger.cs` as
a fallback for `LatestVersion`/`LatestReleaseDate` in early-startup log lines)
is a different object from `GLSenseContext.Paths`. If it's ever constructed
before `ConfigureInstallRoot` runs (only possible during very early startup,
before `GLSenseContext` exists), it falls back to reading a manifest.json from
the *old* Excel_Logs location - which, after this ships, is normally
empty/missing, so it would show `CreateDefaultManifestFile()`'s hardcoded
default (`11.1.0`) rather than the real current version, in that fallback log
line only. This is a narrow, cosmetic, log-only edge case with no functional
impact - not corrected here.

### 2. Build-time: `GLSense.Addin.Core\post_build.cmd`

The existing Debug-writes-to-AppData / Release-writes-to-`%PROJECT_DIR%\
Manifests` branch is replaced with one uniform destination for both
configurations:

```
GLSense.Addin.Core\SetupFiles\{Config}\Manifest\manifest.json
GLSense.Addin.Core\SetupFiles\{Config}\Manifest\v{FILE_VERSION}.zip
```

This folder is transient build output (added to `.gitignore`, not committed) -
purely a hand-off point to the next step. Everything else about how the zip is
built (staged copy of `CORE_BIN_DIR` excluding `*.pdb`, SHA256 checksum,
`ReleaseNotes.txt`-sourced notes, hand-rolled JSON) is unchanged.

### 3. Build-time: `GLSense\post_build.cmd` (host)

A new step copies from Addin.Core's `SetupFiles` into the host's own bin
output:

```
xcopy /Y /I "<repo>\GLSense.Addin.Core\SetupFiles\%CONFIG%\Manifest\*" "%TARGET_DIR%\AddinCore\Manifest\"
```

i.e. the final destination is `GLSense\bin\{Config}\AddinCore\Manifest\` -
exactly what `PathProvider.ManifestDirectory` now resolves to at runtime. If
the source doesn't exist (e.g. GLSense built standalone, Addin.Core not yet
built in that configuration), this step logs a warning and continues rather
than failing the build.

### 4. Build ordering

`GLSense.csproj` has no `ProjectReference` to `GLSense.Addin.Core.csproj`
today, so nothing guarantees Addin.Core (and its post-build) completes before
GLSense's own post-build runs. A `ProjectReference` from `GLSense.csproj` to
`GLSense.Addin.Core.csproj` is added with `<Private>False</Private>` (no DLL
copy - purely a build-order dependency), mirroring the identical pattern
already used between `GLSense.Addin.Core.csproj` and
`GLSense.Loader.Core.csproj` (CLAUDE.md section 16).

**This is also what makes `GLSense.Build.csproj` a correct single build
entry point**, per explicit request - the developer should never need to
build Addin.Core and GLSense separately for this flow to work. MSBuild
resolves a `ProjectReference` transitively: `GLSense.Build.csproj` already
references both `GLSense.csproj` and `GLSense.Addin.Core.csproj`, and a
`ProjectReference` guarantees the referenced project's full `Build` target
(compile *and* `PostBuildEvent`) completes before the referencing project's
own compile starts. So once `GLSense.csproj` -> `GLSense.Addin.Core.csproj` is
added, building `GLSense.Build.csproj` alone deterministically runs:
`Contracts -> Shared -> Loader.Core -> Addin.Core (+ its post-build,
populating SetupFiles) -> GLSense (+ its post-build, copying SetupFiles into
its own bin\AddinCore\Manifest)`, with no manual "build this project first"
steps. No further `.csproj` changes are needed for this beyond the reference
already listed above - `GLSense.Build.csproj` itself needs no new
`ProjectReference`, since it already references both leaves of the graph.

### 5. Accepted dev-loop trade-off

Today, rebuilding *only* `GLSense.Addin.Core` immediately refreshes the
AppData Manifest folder that `UpdateBootstrapper` auto-adopts from on next
launch/Reload. After this change, that automatic (no-picker) auto-adopt path
only refreshes once `GLSense` itself has been rebuilt (so its post-build can
copy from Addin.Core's `SetupFiles`) - i.e. a solution or `GLSense.Build`
rebuild, not an Addin.Core-only rebuild.

**Confirmed acceptable, per explicit decision:** this trade-off stands as
designed. The Reload ribbon button's Offline mode (`GLReloadSourcePicker`,
browse-to-a-folder) remains available for fast Addin.Core-only iteration
without a full rebuild - the developer points it directly at
`GLSense.Addin.Core\SetupFiles\{Config}\Manifest\`.

## Directory shape (end state)

```
GLSense\bin\{Debug|Release}\
    GLSense.dll
    ... (host bin output, unchanged)
    AddinCore\
        Manifest\
            manifest.json
            v{version}.zip
        Versions\
            V{version}_{releaseDateSafe}\
                GLSense.Addin.Core.dll
                GLSense.Loader.Core.dll
                ... (extracted release payload)
        ReleaseHistory.json

%LocalAppData%\ORBIT\Excel_Logs\GLSense_Logs_New\      (unchanged)
    Logs\
    Database\
    Temp\
    BrowserLogs\
        Login\
        Drilldown\
    Resources\
```

## Files touched

- `GLSense.Shared\PathProvider.cs` - add `_installRoot`, `ConfigureInstallRoot`,
  redirect `ManifestDirectory`/`VersionsPath`/`ReleaseHistoryFile`.
- `GLSense\GLSenseContext.cs` - call `ConfigureInstallRoot` before constructing
  `Paths`.
- `GLSense.Addin.Core\post_build.cmd` - replace the Debug/Release branch with a
  single `SetupFiles\{Config}\Manifest\` destination.
- `GLSense\post_build.cmd` - add the `xcopy` step from Addin.Core's
  `SetupFiles` into `%TARGET_DIR%\AddinCore\Manifest\`.
- `GLSense.csproj` - add the `Private=False` `ProjectReference` to
  `GLSense.Addin.Core.csproj`.
- `.gitignore` - add `GLSense.Addin.Core\SetupFiles\`.
- `GLSense.Build\post_build.cmd` - its build-summary printer currently echoes
  the old AppData path as the "Release publish folder"; update it to report
  the real end state instead (`GLSense\bin\{Config}\AddinCore\Manifest\`) -
  since `GLSense.Build` is now the intended single build entry point, this
  summary should be trustworthy, not cosmetic-only.

No changes needed to: `GLSense.Loader.Core\UpdateBootstrapper.cs`,
`GLSense.Loader.Core\AddinDomainLoader.cs`, `GLSense.Shared\
ReleaseHistoryStore.cs`, `GLSense.Contracts\IPathProvider.cs` (no new members),
`GLSense\Views\GLReloadSourcePicker.xaml.cs`, `GLSense\Views\
GLReleaseHistoryBrowser.xaml.cs` - all already consume paths exclusively
through `IPathProvider`.

## Testing / verification

No installer exists to test uninstall against directly. Verification is:
1. Clean build of **`GLSense.Build.csproj` alone** (not the whole solution,
   not GLSense/Addin.Core individually) in both Debug and Release; confirm
   `GLSense\bin\{Config}\AddinCore\Manifest\manifest.json` + `.zip` are
   produced. This is the primary check, since `GLSense.Build` is meant to be
   the single entry point going forward.
2. Confirm `%LocalAppData%\ORBIT\Excel_Logs\GLSense_Logs_New\` gets no new
   `Manifest\`/`Versions\`/`ReleaseHistory.json` writes after this change
   (only `Logs\`/`Database\`/`Temp\`/`BrowserLogs\`/`Resources\` remain
   written there).
3. Launch Excel; confirm `UpdateBootstrapper` auto-adopts from the new
   location (`Versions\` and `ReleaseHistory.json` appear under `GLSense\bin\
   {Config}\AddinCore\`, not under Excel_Logs).
4. `GLReleaseHistoryBrowser`/`GLReloadSourcePicker` still function correctly
   reading from the new location (no code change needed there, but confirm
   at runtime).
5. Standalone build of `GLSense.Addin.Core` only, without `GLSense`: confirm
   it still succeeds (SetupFiles gets refreshed) and `GLSense`'s post-build
   warning path is exercised harmlessly if `GLSense` is then built separately
   before Addin.Core.

This is AIPowered-only; no changes to FinalWorkingCode.
