# Online multi-version reload: follow-up checklist (once the server exists)

This is a standalone checklist, separate from the implementation plan and the design
spec (both already fully executed and committed - see below). Everything in this file
assumes the client-side work is done; nothing here should require touching
`GLReloadSourcePicker`'s architecture again unless testing turns up a real bug.

**Read first, for context:**
- Design spec: `docs/superpowers/specs/2026-09-18-online-reload-multiversion-design.md`
- Implementation plan (7 tasks + fix history): `docs/superpowers/plans/2026-09-18-online-reload-multiversion.md`
- CLAUDE.md sections 75-76 (schema rework + this feature's own fix-log entry)

**Current state:** code-complete, build-verified via real MSBuild rebuilds and
PowerShell fixture scripts against compiled DLLs. **Never exercised against a real
server or inside a live Excel session.** The three endpoints below don't exist yet.

---

## 1. Confirm the actual API contract matches what the client already assumes

The client was built against these three endpoints, defined as placeholder path
constants in `GLReloadSourcePicker.xaml.cs` (`VersionsListPath`, `ManifestByFolderPath`,
`DownloadPath`) - update the constants if the real paths differ, nothing else should
need to change if the *shapes* below hold:

| Endpoint | Client assumption |
|---|---|
| `GET {LoginUrl}/glsense/versions` | Returns a JSON array, one full manifest entry per server release - same schema as the local `manifest.json`: `version`, `releaseDate`, `release`, `folderName`, `fileName`, `checksum`, `notes`, `mandatory`. |
| `GET {LoginUrl}/glsense/manifest?folderName={folderName}` | Returns a single-entry array in the exact same shape, for exactly the requested `folderName`. |
| `GET {LoginUrl}/glsense/download?file={fileName}` | Returns the zip binary for that release. |

**Three assumptions are load-bearing and were only ever informally noted in code
comments - get these confirmed as real contract terms with the backend team, not just
assumed:**

- [ ] **`releaseDate` format.** Must be exactly `yyyy-MM-ddTHH:mm:ss` (no timezone
  suffix, no space separator) - identical to what `post_build.cmd` writes locally. The
  client now parses this via `DateTime.TryParse` in two places
  (`OnlineReleaseClassifier.ParseReleaseDateOrMin`, `UpdateBootstrapper`'s catalog
  "newest" sort) with a `DateTime.MinValue` fallback - a format mismatch won't crash
  anything, but it WILL silently break "which release is newest" and "is this release
  already local" (the dedup key compares `ReleaseDate` as an exact string).
- [ ] **`folderName` must be a single safe path segment.** The client validates this
  (`UpdateBootstrapper.ExtractAndCatalog` rejects anything where
  `Path.GetFileName(folderName) != folderName`, or that's exactly `"."`/`".."`, or that
  matches the currently-loaded release's folder) - but a server that returns something
  unexpected here will just have every fetch for that release silently rejected and
  logged, not crash. Confirm the server always emits the same
  `v{version}_{releaseDateSafe}` shape `ReleaseHistoryStore.BuildFolderName` produces
  locally, purely for consistency (not required for correctness, since the client
  trusts whatever safe value the manifest supplies).
- [ ] **The manifest-by-folder endpoint must return exactly the requested release.**
  The client validates this too (compares the extracted `resolved.Version`/
  `ReleaseDate` against the row that was actually clicked, logs a failure and skips
  it on mismatch) - but confirm the server's query-parameter filtering actually works
  as intended rather than relying on the client's validation as the only safety net.

## 2. Manual UI verification (the spec's own testing section, now actually runnable)

Set up a temporary way to serve the three endpoints (a real dev/staging server, or a
short-lived local stub) and run through:

- [ ] Currently-loaded release renders with the pale-blue tint (`#E3F2FD`), no
  checkbox, "Currently loaded" status.
- [ ] An already-catalogued local release (check `ReleaseHistory.json` on the test
  machine first) renders grey, no checkbox, "Already Downloaded" status.
- [ ] A genuinely new release is checkable; the single newest such release is
  pre-checked when the list loads; the action button's label/count and enabled state
  track the checked count live as boxes are (un)checked.
- [ ] Clicking any row - including Loaded/Already-Downloaded ones - populates the
  details panel with all 8 fields (Version, Released, Status, Notes, Mandatory,
  Folder Name, File Name, Checksum).
- [ ] Checking 2+ new releases and clicking "Fetch Selected & Reload": confirm
  `ReleaseHistory.json` gains one entry per successfully fetched release, `Versions\`
  gains one folder per successfully fetched release, and the add-in ends up loaded
  onto whichever is now the newest release across the WHOLE catalog (not necessarily
  the newest one just fetched, if an even-newer release was already local beforehand).
- [ ] **Partial-failure path**: make one selected row's manifest/zip fetch fail (point
  it at a 404, or a deliberately wrong checksum) while another selected row succeeds -
  confirm the successful one still gets catalogued, the failure is logged via
  `LogFailure`, and the dialog still closes/reloads onto whatever is now newest.
  Confirm the failed row's extraction genuinely left nothing behind (no partial
  `Versions\` folder for it).
- [ ] **Full-failure path**: every checked row fails - confirm the dialog does NOT
  close and no reload is triggered.
- [ ] **Double-click "Fetch Selected & Reload" during an in-flight download.** This is
  now guarded (`_fetchInProgress` + the button disabling itself via `SetBusy`) but has
  never been exercised live - confirm the second click is genuinely a no-op, not just
  "doesn't crash."
- [ ] **Close the dialog via the title-bar X (or Alt+F4) while a download is
  in-flight.** This is now guarded (`_isClosed` flag checked in `CompleteOnlineFetch`,
  plus an outer try/catch around the whole handler) but has never been exercised live -
  confirm the in-flight loop finishes quietly with no unhandled exception and no
  attempt to touch the closed window.
- [ ] **Window layout with the grid at full height AND the details panel open at the
  same time.** A best-effort mitigation was applied (`GridOnlineReleases.MaxHeight`
  220→160, `Window.Height` 680→760) but was never visually confirmed - check the
  "Fetch Selected & Reload"/"Cancel" buttons are never clipped or pushed off the
  bottom of the window in this state.

## 3. Known, deliberately-deferred minor items (fix opportunistically, not blocking)

None of these need to happen before server integration testing - listed here so they
don't get lost, and can be picked up in the same pass as whatever live-testing fixes
turn up.

**Fixed already** (commit `d85b6ec`, applied directly after review confirmed they were
worth doing - see CLAUDE.md section 76 for the reasoning):
- [x] `BtnReload_Click`'s dead `"Online"` ternary arm - simplified to a plain
  `"Offline"`.
- [x] Stale `_onlineRows` surviving an Online → Offline → Online mode switch - now
  cleared in `Mode_Checked`'s Online branch.
- [x] `UpdateBootstrapper.ExtractAndCatalog` running synchronously on whichever thread
  it's invoked from - `BtnFetchAndReload_Click`'s call site now wraps it in
  `await Task.Run(...)`.
- [x] `LogFailure`'s misleading "Failed to fetch" wording for post-fetch cataloging
  failures - reworded to "Failed to fetch or catalog".
- [x] Fully-qualified `System.Collections.Generic.List<>` in `AddOnlineRows`'s
  signature - added the missing `using` instead.

**Deliberately left as-is** (not bugs, not worth the churn):
- [ ] Two independent SHA256-hex implementations coexist in
  `GLReloadSourcePicker.xaml.cs` (`ComputeSha256`, file-based, used by Offline mode)
  and `UpdateBootstrapper.ExtractAndCatalog` (byte-array-based) - different call
  shapes in different classes; consolidating would mean introducing a shared utility
  purely for DRY's sake with no real benefit.
- [ ] `_onlineRows`' `PropertyChanged` subscriptions are never explicitly unsubscribed
  on `.Clear()` - rows don't root the window and nothing leaks, so unsubscribe logic
  would be handling a problem that doesn't exist.
- [ ] `OnlineReleaseClassifier.Classify` silently drops any server entry with a
  null/blank `Version` rather than surfacing it as, say, an error row - reasonable for
  now, but worth knowing a malformed manifest entry just vanishes from the list.
- [ ] The fallback-to-`ReleaseHistoryStore.BuildFolderName` path in `ExtractAndCatalog`
  (used only when a manifest omits `folderName` entirely) has never been exercised by
  any test - the verification scripts written during implementation only covered the
  manifest-supplies-`folderName` branch. Low risk since this repo's own manifests
  always include it, but worth a fixture test if the server's manifests might ever
  omit it.

## 4. If any of the above surfaces a real bug

Follow this project's own established convention: fix it, then add a new numbered
section to `CLAUDE.md` (after section 76) documenting what broke and why, matching the
style of every other entry in that file. Don't silently patch without a paper trail -
this exact feature already went through two rounds of "a Minor finding was calibrated
too leniently and had to be escalated later" during its own implementation (see CLAUDE.md
section 76's own bug-log for both instances) precisely because compounding risk across
several small gaps is easy to miss without writing it down.
