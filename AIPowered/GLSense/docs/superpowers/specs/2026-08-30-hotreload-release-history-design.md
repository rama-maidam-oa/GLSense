# Hot-reload Release History (Online/Offline picker + rollback catalog) — Design

Status: design approved by user in conversation; not yet implemented.
Scope: AIPowered only (`GLSense` host project + `GLSense.Addin.Core` + `GLSense.Loader.Core` +
`GLSense.Contracts` + `GLSense.Shared`). Not ported to FinalWorkingCode (no equivalent
architecture exists there).

## 1. Motivation

Today, testing a different build of `GLSense.Addin.Core` means either using `RibReload`
(reloads whatever's sitting in the local `Manifest\` folder) or a full MSI
uninstall/reinstall. Neither lets a tester quickly flip between several *specific past*
builds to compare behavior or bisect a regression. This feature adds:

- **A permanent, append-only catalog** (`ReleaseHistory.json`) of every Addin.Core build
  ever adopted on this machine, keyed by an exact release timestamp (not just a version
  number, since version numbers can stay frozen across multiple distinct releases).
- **A redesigned `RibReload`** that offers Online (check a server endpoint for a newer
  build) and Offline (pick a local folder containing a manifest+zip pair) — both adopt
  through the existing `Manifest\` + `UpdateBootstrapper` pipeline, unchanged.
- **A new ribbon button** that opens a browser over `ReleaseHistory.json`, letting the
  user jump straight to *any* past release — including deliberately going backward,
  which is the whole point of the catalog.

## 2. Folder / file layout

Under `%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs_New\`:

- `Manifest\` — **unchanged in shape**, still the transient staging folder
  (`manifest.json` + a `.zip`) that `UpdateBootstrapper` consumes and clears. Still gets
  written to by `GLSense.Addin.Core\post_build.cmd` on every local build, by the new
  Online download step, and by the new Offline copy step.
- `Versions\{folderName}\` — **folder-naming changes** (see §3). Never deleted or
  overwritten once created — this is what makes rollback possible.
- `ReleaseHistory.json` — **new**, sits at this level (a sibling of `Manifest\` and
  `Versions\`, not inside either). The master catalog. See §4 for schema.

## 3. `Versions\` folder keying (corrected from version-number-only)

**Problem this fixes**: the current scheme names folders `V{version}` only. Two different
releases can share the same version number (confirmed as expected in this project — see
`CLAUDE.md` §14.2's own note that the version is frozen for stretches and `releaseDate` is
the real signal). Keying by version alone means a same-day (or any later) rebuild with an
unchanged version number would silently overwrite and destroy an earlier release's files —
fatal for a rollback feature.

**Fix**: fold the release's exact timestamp into the folder name:

```
V{version}_{releaseDateSafe}
e.g. V11.1.0_2026-08-30T14-32-05
```

`releaseDateSafe` = the release's `releaseDate` string (already written in the
`yyyy-MM-ddTHH:mm:ss` format established in `CLAUDE.md` §19.2) with `:` and any other
Windows-illegal path characters (`/ \ * ? " < >  |`) replaced with `-`.

Each such folder, in addition to the extracted DLLs, gets a **copy of that release's own
`manifest.json`** placed inside it (`Versions\{folderName}\manifest.json`) — a
self-contained, permanent record of exactly what that folder is, independent of the
transient copy that passed through `Manifest\`.

**The folder name is computed exactly once**, at the moment a release is extracted
(`UpdateBootstrapper.ExtractLocalZipAndAdopt`), and is then written verbatim into that
release's `ReleaseHistory.json` entry (§4) as the `folderName` field. No other code path
is allowed to recompute it from `version`+`releaseDate` — every other consumer resolves a
release's folder by reading the catalog entry's stored `folderName`, never by re-deriving
it. This avoids the exact same staleness risk described in §4: if the sanitization rule
ever changes later, old entries stay resolvable because the resolved value was frozen at
write time, not recomputed.

### 3.1 Threading the resolved folder through the existing load path

`UpdateBootstrapper.ResolveVersionToLoad(IGLSenseContext)` currently returns a bare
`string` (the version). It changes to return a small result:

```csharp
public sealed class ResolvedRelease
{
    public string Version { get; set; }
    public string ReleaseDate { get; set; }
    public string FolderName { get; set; }   // Versions\{FolderName}\ is where the DLLs live
}
```

- `AddinDomainLoader.Load()` builds `ApplicationBase`/`PrivateBinPath` from
  `Path.Combine(paths.VersionsPath, resolved.FolderName)` instead of
  `Path.Combine(paths.VersionsPath, $"V{context.Version}")`.
- `IGLSenseContext` gains a new property, `ActiveFolderName` (alongside the existing
  `Version`/`ReleaseDate`), set every time a load/reload succeeds
  (`AddinModule_OnRibbonLoaded` and `ReloadAddinCore`). `Version`/`ReleaseDate` are kept
  as-is for display purposes (GLAbout, log lines) — `ActiveFolderName` is the only thing
  actually used to locate the folder on disk going forward.
- Doc-comment added to `PathProvider.LatestVersion`/`LatestReleaseDate`: once rollback is
  a supported operation, these represent "the *currently adopted* version," not
  necessarily the highest/most-recent version ever seen on this machine — a deliberate
  rollback can leave them holding an older value than some catalog entry.

## 4. `ReleaseHistory.json` — schema and storage

Flat JSON array, one object per adopted release, **appended to forever** (only ever
pruned by reconciliation removing entries whose files are genuinely gone — §6). Chosen
over SQLite: this is a small, linearly-scanned, append-mostly list — a plain JSON file is
human-readable (important for a feature whose whole purpose is manual testing/inspection),
reuses the existing `manifest.json` convention/parser, and needs no new dependency.

```json
[
  {
    "version": "11.1.0",
    "releaseDate": "2026-08-30T14:32:05",
    "folderName": "V11.1.0_2026-08-30T14-32-05",
    "checksum": "ABCD1234...",
    "notes": "Published by GLSense.Addin.Core",
    "source": "Install"
  }
]
```

- `version` / `releaseDate` — raw, human-readable; also what "is this newer than what I
  have" comparisons run against (RibReload's Online/Offline gate — §7).
- `folderName` — the frozen, literal pointer into `Versions\`. Never recomputed.
- `checksum` — the zip's SHA256 at the time it was adopted (carried over from the source
  manifest.json), kept for audit/debugging even though the zip itself no longer exists
  once extracted.
- `notes` — free text, defaults to `"Published by GLSense.Addin.Core"`; not tied to any
  specific build-script name.
- `source` — `"Install"` (from the MSI-seeded first-run manifest), `"Online"`, or
  `"Offline"`. Never `"History"` — picking a past release via the new browser (§8) reads
  the catalog, it never writes to it.

### 4.1 Concurrency and corruption safety

Two real risks identified during review, both fixed the same way every read-modify-write
of this file happens (reconciliation prune, and appending a new entry):

- **Crash mid-write** must never corrupt the file, since a corrupt catalog would break
  reading *all* history, not just lose the newest entry. Fix: serialize to a temp file in
  the same folder, then replace the real file (`File.Replace`, or delete+move) — never
  write in place.
- **Concurrent writers** (a user with more than one Excel process open, each triggering a
  reload around the same time) could otherwise race a classic read-modify-write lost
  update. Fix: acquire a named system `Mutex` (`Global\GLSense_ReleaseHistory_Mutex`)
  around every read-modify-write cycle against this file, released promptly.

## 5. Fresh-install detection & seeding (corrected from "Manifest folder exists")

**Problem this fixes**: the original idea ("Manifest folder's existence signals a fresh
install") collides with the fact that *every* ordinary Online/Offline adopt also
populates that same folder transiently. Both cases would look identical at the next
`OnRibbonLoaded`, risking an ordinary update wrongly triggering a wipe-and-rebuild.

**Fix**: the real signal is **whether `ReleaseHistory.json` exists at all.**

On `AddinModule_OnRibbonLoaded`:

1. **`ReleaseHistory.json` does not exist** → first-ever run on this machine (fresh
   install, or a full AppData wipe).
   - If `Manifest\manifest.json` + a `.zip` are present (the MSI's bundled seed): delete
     the entire `Versions\` directory recursively if it exists (untrusted leftover state,
     since there is no catalog to say what's actually valid), then call
     `UpdateBootstrapper.ResolveVersionToLoad`/`ExtractLocalZipAndAdopt` exactly as the
     ordinary path does — **this is not a separate extraction implementation**. The only
     thing genuinely special about the fresh-install case is the wipe-first step and the
     fact that the catalog append (§4) is creating the very first entry rather than
     appending to existing ones; the extraction mechanics, folder-name computation, and
     catalog-append logic are identical code in both cases. After a successful adopt,
     delete the `Manifest\` folder entirely (per the original ask — a persistently-present
     Manifest folder reads as ongoing infrastructure rather than a one-time install
     artifact). Note: this doesn't mean the folder can never exist again —
     `PathProvider.Ensure()`/an ordinary Online/Offline adopt is free to recreate it
     later; this deletion is only cleaning up *this specific seed instance*.
   - If no seed files are present (edge case — deleted before first launch): fall back to
     today's existing lazy-default-manifest behavior; `ReleaseHistory.json` simply isn't
     created until the first real adopt happens.
2. **`ReleaseHistory.json` already exists** → an ordinary run.
   - If `Manifest\manifest.json` is currently present, read its `version`+`releaseDate`
     and compare against every existing catalog entry:
     - **Exact match found** → this is the "same build reinstalled" case (the original
       point 9 scenario). Run reconciliation now (§6) before proceeding — a matching
       entry might get pruned right here if its folder is missing, in which case the
       zip sitting in `Manifest\` gets extracted fresh and re-added as a new entry.
     - **No match** → a genuinely new/different release arriving normally. No wipe of
       anything — proceed through the ordinary adopt path, appending one new entry.
   - If no `manifest.json` is present in `Manifest\`, nothing special happens —
     `UpdateBootstrapper` resolves from the already-installed `Versions\` folder exactly
     as it does today.

## 6. Reconciliation (self-healing the catalog)

Removes any `ReleaseHistory.json` entry whose `Versions\{folderName}\` no longer contains
DLL files (deleted by disk cleanup, an IT-imposed AppData purge, manual tidying, etc.).
Cheap — just `Directory.Exists` + `GetFiles("*.dll").Any()` per entry, no I/O-heavy work.

Runs in **two** places (the original design only had the first, which misses the far more
common case of files disappearing with no reinstall involved at all):

1. During the reinstall-matching-a-known-release case in §5.2.
2. Every time the new Release History browser (§8) is **opened** — so a stale entry is
   never even shown as selectable, not just cleaned up opportunistically later.

## 7. RibReload: Online / Offline picker (carried over from earlier design rounds)

Lives in `GLSense\Views\GLReloadSourcePicker.xaml`/`.xaml.cs` — the host project, with no
dependency on `GLSense.Addin.Core` (this window must work regardless of whether Addin.Core
is currently loaded, mid-unload, or has never successfully loaded at all — that's the
whole reason it can't live inside Addin.Core). `GLSense.csproj` gains WPF references
(`WindowsBase`/`PresentationCore`/`PresentationFramework`/`System.Xaml`) and `Page`/
`Compile` items for the new file, mirroring `GLSense.Addin.Core.csproj`'s existing recipe
(no `ProjectTypeGuids` change, to avoid disturbing Add-in Express's own project-system
expectations for `AddinModule.cs`). Visual style is a small set of inline WPF styles that
duplicate the app's look (accent color, button/corner-radius conventions) rather than a
shared resource — a copy, not a code dependency.

### 7.1 New cross-domain contract surface

`AppState` (which holds `LoginUrl`/`LoginToken`/`IsLoggedIn`) lives entirely inside
`GLSense.Addin.Core` and isn't exposed across the AppDomain boundary today (only a
mirrored `RibbonController.IsLoggedIn` bool exists host-side). Online mode needs the
actual URL and token, so:

- `GLSense.Contracts.LoginInfo` — `[Serializable]`, `{ string LoginUrl; string
  LoginToken; bool IsLoggedIn; }`.
- `IGLSenseAddin.GetLoginInfo()` → returns it; implemented in `AddinEntry.cs` as a direct
  read of `AppState.Instance`'s three fields.

**Standing discipline this establishes** (see §9 — the most important non-obvious finding
from review): every `IGLSenseAddin` member added from this point forward must be called
defensively (tolerate `MissingMethodException`/`RemotingException`), because the Release
History browser (§8) can load an *older* build whose compiled `IGLSenseAddin`
implementation predates that member. `GetLoginInfo()`'s own call site wraps the call in
try/catch, treating any failure the same as "not logged in" (Online disabled, Offline
auto-selected) — not just a null check.

### 7.2 Window behavior

- Two radio buttons, "Online"/"Offline". On open: `GlobalsEx.Addin?.GetLoginInfo()`
  (wrapped per §7.1) — null/not-logged-in/exception → Online disabled, Offline
  auto-selected; logged in → both enabled, Online selected by default.
- **Online**: `GET {LoginUrl.TrimEnd('/')}/glsense/projectdlls`, header `Authorization:
  Bearer {LoginToken}` (matches `ApiHelper.cs`'s existing pattern). **This endpoint does
  not exist yet** — out of scope for this client-side work, but its contract is fixed
  here so server-side implementation has an exact target: response body is the same JSON
  shape as the local `manifest.json` (`[{version, releaseDate, downloadUrl, checksum,
  notes, mandatory}]`), parsed via the existing `GLSense.Shared.VersionParser
  .ParseVersionJson`. Baseline comparison uses `GlobalsEx.Context.Version`/`ReleaseDate`
  (the values threaded through per §3.1) — **not** any project's compiled assembly
  version. Not strictly newer (by `releaseDate`, per `CLAUDE.md` §14.2's own noted design
  intent) → "No updates available", stop. Strictly newer → download the response's
  `downloadUrl` (same auth header) to a temp file, verify its SHA256 against the
  response's `checksum` (mismatch aborts, nothing is touched), then proceed to §7.3.
- **Offline**: `FolderBrowserDialog`, default folder = the real Downloads folder
  (`SHGetKnownFolderPath` — .NET Framework 4.8.1 has no `SpecialFolder.Downloads`).
  Scans top-level for `manifest*.json` / `v*.zip` (handles browser `(1)`-suffix
  duplicates by picking the newest `LastWriteTime` per pattern), parses via
  `VersionParser.ParseVersionFile`, checksum-verifies, applies the same strictly-newer
  gate, then proceeds to §7.3.
- **Cancel**: `DialogResult = false`, closes immediately, zero side effects.

### 7.3 Shared "adopt" step

Delete any pre-existing `*.zip` in `Manifest\`, copy the new zip in as-is, copy the new
manifest in renamed to exactly `manifest.json` (the literal name `UpdateBootstrapper`
requires), set `DialogResult = true`, close. `RibReload_OnClick` then proceeds exactly as
it does today (`_reloadInProgress` guard, call `ReloadAddinCore()`) — **no changes to
`UpdateBootstrapper.cs`'s consumption logic itself**, beyond the folder-keying fix in §3.

`RibReload` stays enabled regardless of login state (the window itself gates Online vs.
Offline availability, not the ribbon button).

## 8. New "Release History" ribbon button

New control, `RibReleaseHistory`, placed near `RibReload`. Opens
`GLSense\Views\GLReleaseHistoryBrowser.xaml` (also host-side, same no-Addin.Core-
dependency reasoning as §7).

- Runs reconciliation (§6) first, then lists every remaining `ReleaseHistory.json` entry
  — Version / Release Date / Source / Notes columns, newest-first by default.
- User selects a row, clicks "Load This Release". **No version gate at all** — unlike
  RibReload's Online/Offline path, deliberately loading an *older* release is the entire
  point of this feature (backporting/comparison testing). The only check performed is
  that the entry's `Versions\{folderName}\` still has DLL files (already guaranteed by
  reconciliation having just run).
- On confirm: reuses the exact same teardown-and-load machinery as `ReloadAddinCore()`
  (shutdown old instance, null the pointer, unload the AppDomain, load fresh) — but the
  "which release to load" step is parameterized instead of always going through
  `UpdateBootstrapper.ResolveVersionToLoad()`. Concretely, `ReloadAddinCore()` is
  refactored to accept a small resolver delegate/result so both call sites share one
  teardown/load path:
  - RibReload's path: resolver = `new UpdateBootstrapper().ResolveVersionToLoad(...)`.
  - Release History's path: resolver = a `ResolvedRelease` built directly from the
    user's picked catalog entry (`Version`/`ReleaseDate`/`FolderName` copied straight
    across) — no `UpdateBootstrapper`/`Manifest\` involvement at all, since nothing needs
    downloading or extracting.
- Picking an entry **never writes to `ReleaseHistory.json`** — purely a read/select
  operation. Only §7's Online/Offline adopt path ever appends a new entry.
- No checksum re-verification here (nothing was just downloaded) and no zip involved —
  the DLLs are already sitting on disk from when that entry was originally adopted.

## 9. Standing engineering discipline (not a one-time fix)

Because any historical build can be reloaded via §8, and the host (`GLSense.dll`) is
never itself reloaded, **every future addition to `IGLSenseAddin` must be called
defensively** by the host — an older loaded instance may genuinely not implement a member
added after it was built, and that surfaces as a runtime exception, not a compile error.
This should be written into `CLAUDE.md` once implemented, as a permanent rule for anyone
touching `IGLSenseAddin` going forward, not just documented here.

This risk is **directional** and does not apply symmetrically to `IGLSenseContext`
(also new: `ActiveFolderName`, §3.1). Calls flow host→Addin.Core for `IGLSenseAddin`
(the host, always on the current interface, calling into a possibly-older
implementation — genuine crash risk), but `IGLSenseContext` flows the other way
(Addin.Core reading a context object the host constructed against its own,
always-current `IGLSenseContext`). An older Addin.Core build's own compiled copy of that
interface simply doesn't declare newer members at all, so its code was never written to
reference them — adding to `IGLSenseContext` is inert for old builds, not a crash risk.
Only `IGLSenseAddin` additions need the defensive-call treatment.

## 10. Explicitly accepted trade-offs (not fixed, by design)

- **Unbounded disk growth**: `Versions\` folders are permanent by design; with frequent
  same-day releases this will accumulate real disk usage over months. No auto-pruning —
  that would defeat the feature's purpose. A manual cleanup utility is a plausible future
  addition, out of scope now.
- **No progress UI on reload today, and this feature will make the lack more noticeable**
  (`ReloadAddinCore()` blocks Excel's UI thread synchronously for the unload+reload,
  which can take a few seconds per the existing `UnloadTimeoutSeconds = 5` note). Since
  this feature specifically encourages much more frequent reloading, add a simple
  `Cursor.Current = Cursors.WaitCursor` bracket around the reload call in both
  `RibReload_OnClick` and the Release History browser's confirm handler — no new window,
  just existing WinForms cursor feedback, proportionate to the problem.

## 11. Sequencing

Implement in this order (B and C both depend on A's data model existing first):

- **A**: `Versions\` folder re-keying (§3), `ReleaseHistory.json` schema/storage (§4),
  fresh-install detection + reconciliation (§5, §6). Touches `UpdateBootstrapper.cs`,
  `AddinDomainLoader.cs`, `IGLSenseContext.cs`/`GLSenseContext.cs`, `AddinModule.cs`
  (startup + `ReloadAddinCore`).
- **B**: RibReload's Online/Offline picker (§7), including the new `IGLSenseAddin
  .GetLoginInfo()` contract member and its defensive call site.
- **C**: The Release History browser button/window (§8), including the
  `ReloadAddinCore()` parameterization needed to share teardown/load logic with B.

## 12. Manual verification plan (no automated WPF test harness exists in this codebase)

- Fresh-install path: delete `ReleaseHistory.json` (simulating a clean machine), seed
  `Manifest\` by hand, launch Excel, confirm the catalog is created with one `"Install"`
  entry and `Manifest\` is deleted afterward.
- Ordinary update path (existing dev loop): rebuild `GLSense.Addin.Core`, use RibReload's
  Offline mode against the local `Manifest\` folder, confirm a second catalog entry
  appears and the old entry/folder is untouched.
- Same-day double release: rebuild twice in one day, confirm two distinct
  `Versions\` folders exist (not one overwritten) and both remain individually loadable
  via the Release History browser.
- Browser duplicate-suffix case: copy a manifest+zip pair into Downloads twice (simulating
  `manifest (1).json`/`v11.2.0 (1).zip`), confirm Offline mode picks the newest by
  `LastWriteTime`.
- Corrupted zip: truncate a zip's bytes, confirm checksum verification blocks reload with
  a clear error, no files touched.
- Reconciliation: manually delete a `Versions\{folderName}\` folder still referenced by
  an entry, open the Release History browser, confirm that entry no longer appears.
- Rollback: use the Release History browser to load an older entry than the currently
  active one, confirm it loads with no version-gate rejection.
- Concurrent-process write: trigger a reload from two Excel processes at close to the
  same time, confirm both entries end up in `ReleaseHistory.json` (no lost update).
