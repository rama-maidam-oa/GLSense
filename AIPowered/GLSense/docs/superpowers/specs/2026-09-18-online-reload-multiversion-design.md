# GLReloadSourcePicker: multi-version Online mode (AIPowered only)

## Problem

`GLReloadSourcePicker`'s Online mode (`BtnCheckOnline_Click`) only ever checks
for a single "latest" release: it calls `{LoginUrl}/glsense/projectdlls`,
compares the one version it gets back against what's currently loaded, and -
if newer - downloads and stages exactly that one release. This was built
before the server had any real release history of its own.

The server side is now being built to maintain a consolidated,
`ReleaseHistory.json`-equivalent catalog (`versions.json`) that grows by one
entry per push. Multiple releases can accumulate between one user's Excel
sessions - someone might reload the same day a fix ships, or wait several
days and by then have 3-4 releases sitting unclaimed on the server. Today's
single-latest-only check has no way to show that, and no way to let a user
grab more than one release in a single pass.

## Goal

Turn Online mode from "check for one update" into "browse every server
release not yet on this machine, fetch whichever ones you want, and the
newest one becomes what's actually loaded right now." Already-local and
currently-loaded releases are shown (for context and for viewing their
details) but can't be re-selected - `GLReleaseHistoryBrowser` remains the only
place that loads something other than the newest release.

See the confirmed visual reference (mockup) discussed in this design session
for exact layout/colors: white theme matching `GLReloadSourcePicker.xaml`'s
and `GLReleaseHistoryBrowser.xaml`'s existing look (`#2E86AB` primary,
`#E3F2FD` "currently loaded" tint reused from `GLReleaseHistoryBrowser`'s
`IsCurrentlyLoaded` row trigger, light-grey "Already Downloaded" rows, no
minimize/maximize per section 44.4).

## Non-goals

- Server-side implementation is out of scope for this repo - the API
  contract below is what the client needs; the backend team owns building
  it. Endpoint paths shown are placeholders pending their final routing.
- No pagination/filtering of the server's release list. If/when the
  server's catalog grows large enough that showing every release becomes
  unwieldy, that's a follow-up - not addressed here (YAGNI for an
  early-stage rollout).
- `Offline` mode and `GLReleaseHistoryBrowser` are untouched. Offline's
  existing `IsStrictlyNewer`/checksum-verification code is reused as-is,
  not modified.
- `"mandatory"` (already a manifest.json field, schema-only until now) is
  surfaced in this pass purely as a display badge on the row. No
  forced-selection/blocking-Cancel behavior is added for a mandatory
  release - that would be new product behavior nobody has asked for yet.
  If mandatory-update enforcement is wanted later, it's a separate,
  explicitly-scoped follow-up.
- No changes to how `Versions\`/`ReleaseHistory.json` are colocated with
  `GLSense.dll` (see the 2026-09-04 design) - this feature is a consumer of
  that layout, not a change to it.

## Design

### 1. API contract

All three calls reuse the existing authenticated pattern already used by
`BtnCheckOnline_Click` today (`Bearer {loginInfo.LoginToken}` against
`loginInfo.LoginUrl`).

| Call | Method | Purpose | Response shape |
|---|---|---|---|
| **List** | `GET {LoginUrl}/glsense/versions` | Every release the server knows about. | JSON array, one full manifest entry per release - **identical schema to the local `manifest.json`/`ReleaseHistory.json` shape**: `version`, `releaseDate`, `release`, `folderName`, `fileName`, `checksum`, `notes`, `mandatory`. Confirmed by the user: this is the server-side consolidated equivalent of what `post_build.cmd` writes locally, not a trimmed summary. |
| **Manifest by release** | `GET {LoginUrl}/glsense/manifest?folderName={folderName}` | Fetch the authoritative single-release manifest.json right before download. | Single-entry array, same shape as above - parseable by the existing `VersionParser.ParseVersionJson` unchanged. |
| **Zip download** | `GET {LoginUrl}/glsense/download?file={fileName}` | The release's zip payload. | Binary (zip bytes). |

`folderName` is the identifier used for the manifest-by-release call - it
already embeds version+timestamp (`ReleaseHistoryStore.BuildFolderName`), so
it's unambiguous and needs no composite version+releaseDate query with
colons to URL-encode.

**Why re-fetch the manifest per selection when the list already has full
data:** the manifest-by-release response is written byte-for-byte into
`Versions\{folderName}\manifest.json` as the per-version snapshot (matching
what `UpdateBootstrapper.ExtractManifestZipAndAdopt` already does for every
other source). Re-fetching keeps that snapshot authoritative and avoids
inventing a new "reconstruct a manifest.json file from one row of the list
array" code path - the existing `VersionParser`/extraction pipeline is reused
completely unchanged.

### 2. Row classification and default selection

For each server entry, compare `(Version, ReleaseDate)` against two things,
using the *exact* dedup key already used elsewhere in this codebase
(`ReleaseHistoryStore.Append`'s duplicate check,
`UpdateBootstrapper.ResolveVersionToLoad`'s `isKnownReinstall` check) - no new
comparison logic:

1. `GlobalsEx.Context.Version` / `GlobalsEx.Context.ReleaseDate` (what's
   actually loaded right now) → **Loaded** state.
2. `ReleaseHistoryStore.ReadAll(paths.ReleaseHistoryFile)` (everything ever
   catalogued on this machine) → **Already Downloaded** state.
3. Neither → **New** state, checkbox enabled.

Rendering (row states, matching the confirmed mockup):

- **Loaded**: `Background = #FFE3F2FD` (WPF ARGB for the exact light-blue tint
  `GLReleaseHistoryBrowser`'s `IsCurrentlyLoaded` `DataTrigger` already uses),
  `FontWeight = SemiBold`, no checkbox, badge "Currently loaded".
- **Already Downloaded**: light grey (`#F4F4F4`-family), muted text, no
  checkbox, badge "Already Downloaded".
- **New**: default row styling, checkbox enabled, badge "New". The single
  newest **New** row (by `ReleaseDate`) is pre-checked when the list loads;
  every other **New** row starts unchecked but can be checked too.

Clicking *any* row (regardless of state) populates a details panel below the
grid with its full manifest fields (version, released, status, notes,
mandatory, folder name, file name, checksum) - this is read-only viewing, not
a selection action, and is independent of the checkbox.

### 3. Shared extract+catalog method (replaces the old stage-into-Manifest flow)

New method on `UpdateBootstrapper` (`GLSense.Loader.Core`):

```csharp
public ReleaseEntry ExtractAndCatalog(IGLSenseContext context, string manifestJson, byte[] zipBytes, string source)
```

Behavior (lifted directly from `ExtractManifestZipAndAdopt`'s existing body,
generalized to take the manifest/zip as in-memory values instead of reading
them from `paths.ManifestDirectory`):

1. Parse `manifestJson` via `VersionParser.ParseVersionJson` (unchanged).
2. Verify `SHA256(zipBytes)` against the parsed checksum - **abort this one
   release, don't throw** (return `null` / raise a caller-visible failure) if
   it doesn't match, exactly like `ValidateCandidateAsync`'s existing
   Offline-mode checksum check.
3. Extract into `Versions\{folderName}\` (`folderName` from the *parsed
   manifest*, never recomputed client-side).
4. Write `manifestJson` verbatim into that folder as `manifest.json` (the
   per-version snapshot, matching existing behavior).
5. Append a `ReleaseEntry` to `ReleaseHistory.json` via
   `ReleaseHistoryStore.Append` (already idempotent on
   `(Version, ReleaseDate, FolderName)` - safe to call even if this exact
   release somehow gets fetched twice).
6. Return the appended `ReleaseEntry`.

**This replaces `BtnCheckOnline_Click`'s current approach entirely** (write a
temp manifest+zip, copy into `paths.ManifestDirectory`, close the dialog, let
`UpdateBootstrapper.ResolveVersionToLoad`'s branch 2 do the extraction on the
caller's side). The old single-button Online mode becomes redundant once the
list handles "only one new release" as its normal, unremarkable case - it is
removed, not kept alongside the new flow. This also means
`GLReloadSourcePicker.xaml.cs`'s `IsStrictlyNewer`/`PromptNoUpdate` methods
lose their Online-mode caller; they remain, used only by Offline mode, which
is unchanged.

### 4. Download loop (on the action button)

For each **checked** row, in ascending `ReleaseDate` order (so the step log
reads as "building up to the newest," though order has no functional
effect):

1. `LogStep("Fetching manifest for {version} ({releaseDate})...")`
2. `GET` the manifest-by-release endpoint for that row's `folderName`.
3. `LogStep("Downloading {fileName}...")`
4. `GET` the zip endpoint for that manifest's `fileName`.
5. Call `UpdateBootstrapper.ExtractAndCatalog(context, manifestJson, zipBytes, "Online")`.
6. On success: `LogSuccess("Cataloged {version}.")`. On failure at any step
   (network error, parse error, checksum mismatch): `LogFailure(...)` for
   *this row only* and continue to the next checked row - **best-effort,
   never aborts the whole batch** (same non-fatal philosophy as every other
   loop in this codebase, e.g. `ReleaseHistoryStore.Reconcile`,
   `AddinDomainLoader.Unload`'s bounded wait).

After the loop: if at least one row was successfully catalogued, set
`SelectedSource = "Online"`, `DialogResult = true`, close. **No explicit
"pick the latest and stage it" step is needed** - nothing is written to
`Manifest\`, so when the caller's existing
`ReloadAddinCore → UpdateBootstrapper.ResolveVersionToLoad("Online")` runs
next, it lands on the already-existing branch 3 ("no zip in Manifest, use the
catalog's most recent entry with DLLs on disk") and adopts whichever release
is now newest across the *whole* catalog - not necessarily "newest fetched
this run," which is correct: if the user fetches two older-than-today
releases and the actual latest was already local before this operation
started, that already-local latest should still win.

If every checked row failed, the dialog does **not** close (`DialogResult`
stays unset) - same as today's `ValidateCandidateAsync` failure path leaving
`BtnReload` disabled.

### 5. UI/XAML changes (`GLReloadSourcePicker.xaml` / `.xaml.cs`)

- `OnlinePanel` keeps a button (renamed, e.g. `BtnCheckOnline` → "Check for
  Updates") above an initially-empty `DataGrid` bound to a new
  `ObservableCollection<OnlineReleaseRow>`. The list is fetched **only** on
  that button's click, never automatically when the user switches to the
  Online radio button - consistent with section 72's existing precedent
  (Offline mode's folder scan was deliberately changed to require an
  explicit Browse click rather than auto-scanning on mode-switch, to avoid
  slow/networked work running without explicit user intent). Clicking the
  button again re-fetches and replaces the grid's contents, which also
  serves as the "refresh" action - no separate refresh control is needed.
- `DataGrid.RowStyle` gets two `DataTrigger`s mirroring
  `GLReleaseHistoryBrowser.xaml`'s existing `IsCurrentlyLoaded` pattern
  exactly: one for `IsCurrentlyLoaded` (the `#FFE3F2FD` tint), one for
  `IsAlreadyDownloaded` (light grey). Neither trigger touches the checkbox
  column's `IsEnabled` - that's bound directly to `IsSelectable` on the row
  model.
- A details `Border`/`Grid` below the DataGrid (label/value pairs) populated
  on `SelectionChanged`, matching the mockup's details panel.
- The action button's `Content` is bound to show a live count, e.g.
  `Content="{Binding CheckedCount, StringFormat='Fetch Selected & Reload ({0})'}"`,
  and `IsEnabled` requires `CheckedCount > 0`.
- If the list has zero **New** rows (nothing to fetch), the summary area
  shows "No new releases available on the server." and the action button
  stays disabled - **no `MessageBox` for this case** (replaces
  `PromptNoUpdate`'s old modal for the Online path; Offline mode keeps its
  own `PromptNoUpdate` call unchanged, since it's still a single-candidate
  flow).
- Window sizing: the current `520x560` is too small for a grid + summary +
  details + buttons; grow to roughly `GLReleaseHistoryBrowser`'s width family
  (`~660` wide) and taller (`~640-680`). Exact numbers are a polish detail
  settled during implementation, not architecturally load-bearing.

New small model class, `OnlineReleaseRow` (`GLSense` host project, next to
`GLReloadSourcePicker.xaml.cs` - no need for `GLSense.Contracts`/`GLSense.Shared`
placement since nothing outside this window consumes it):

```csharp
public class OnlineReleaseRow : INotifyPropertyChanged
{
    public string Version { get; set; }
    public string ReleaseDate { get; set; }
    public bool Mandatory { get; set; }
    public string FolderName { get; set; }
    public string FileName { get; set; }
    public string Checksum { get; set; }
    public string Notes { get; set; }
    public bool IsCurrentlyLoaded { get; set; }
    public bool IsAlreadyDownloaded { get; set; }
    public bool IsSelectable => !IsCurrentlyLoaded && !IsAlreadyDownloaded;
    public bool IsChecked { get; set; } // raises PropertyChanged; drives CheckedCount
}
```

### 6. Threading

The list fetch and the per-row download loop are all `async`/`await` chains
against `HttpClient`, in the same host window that already has the
documented WPF-dispatcher-thread-loss problem (see the reference memory:
this VSTO add-in's WPF window doesn't reliably run under a
`DispatcherSynchronizationContext`, so a continuation after `await` can
resume off the UI thread). Every UI-touching call in the new code paths
(checkbox state, row collection updates, `LogStep`/`LogSuccess`/`LogFailure`,
button `IsEnabled`) goes through the **same already-dispatcher-safe**
`SetBusy`/`AppendLine`-style pattern this file already uses (`CheckAccess()`
+ synchronous `Dispatcher.Invoke`, never `InvokeAsync`) - no new
thread-safety mechanism needed, just consistent reuse of what's already
there.

## Files touched

- `GLSense\Views\GLReloadSourcePicker.xaml` - replace `OnlinePanel`'s button
  with the DataGrid + details panel; window size bump.
- `GLSense\Views\GLReloadSourcePicker.xaml.cs` - replace `BtnCheckOnline_Click`
  with the list-fetch + row-classification + download-loop logic described
  above; remove the Online-mode caller of `PromptNoUpdate` (Offline's stays).
- `GLSense\Views\OnlineReleaseRow.cs` (new) - the row model above.
- `GLSense.Loader.Core\UpdateBootstrapper.cs` - add `ExtractAndCatalog(...)`,
  refactored out of the existing body of `ExtractManifestZipAndAdopt` so both
  the Install/Offline path and this new Online path share one
  extract-verify-catalog implementation.

No changes needed to: `ReleaseHistoryStore.cs` (already idempotent, already
has the exact dedup key needed), `GLReleaseHistoryBrowser.xaml(.cs)`,
`PathProvider`/`IPathProvider` (no new fields), `VersionInfo.cs`/
`VersionParser.cs` (existing schema already covers every field the new
endpoints return).

## Testing / verification

No real server implementation exists yet (explicitly "under process" per the
requester). Verification for this pass is necessarily against
hand-constructed fixture responses, not a live endpoint:

1. Unit-level: feed `UpdateBootstrapper.ExtractAndCatalog` a known-good
   manifest+zip pair and confirm it extracts/catalogs identically to what
   `ExtractManifestZipAndAdopt` already does today for the Install/Offline
   path - this is a refactor of existing, working logic, so behavioral
   parity (not new behavior) is the bar.
2. A local, temporary fixture (a static JSON file + a locally-built zip,
   *not* a resurrected local HTTP host - the previous local-host approach
   was explicitly abandoned in CLAUDE.md section 17 after real connection
   failures) can stand in for the three endpoints during development by
   temporarily pointing the three URLs at `file://` paths or an in-process
   fake handler, swapped back to the real `HttpClient` calls before merging.
3. Manual UI verification once fixtures are in place: currently-loaded row
   renders with the `#E3F2FD` tint and no checkbox; an already-catalogued
   local release renders grey with no checkbox; a genuinely new release is
   checkable and pre-checked when it's the newest; unchecking/checking
   updates the action button's count and enabled state; clicking any row
   (including Loaded/Already-Downloaded ones) populates the details panel.
4. Multi-select path: check 2+ new releases, run the fetch, confirm
   `ReleaseHistory.json` gains one entry per successfully fetched release and
   `Versions\` gains one folder per successfully fetched release.
5. Partial-failure path: simulate one selected row's checksum failing (or
   its zip endpoint erroring) while another selected row succeeds - confirm
   the successful one still gets catalogued, the failed one is reported via
   `LogFailure`, and the dialog still closes/reloads onto whatever is now
   newest.
6. Full-failure path: every checked row fails - confirm the dialog does not
   close and no reload is triggered.
7. Once a real server implementation exists, re-run the above against it
   end-to-end before considering this feature complete - the fixture-based
   testing above validates the client logic, not the real network contract.

This is AIPowered-only; no changes to FinalWorkingCode (no Online reload
mode exists there to extend).
