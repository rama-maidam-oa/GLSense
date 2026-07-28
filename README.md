# GLSense

This repo tracks two parallel implementations of the GLSense Excel add-in during the
transition from the legacy architecture to the new multi-project one.

## Structure

- **`FinalWorkingCode/`** - the current release-track implementation. Single `GLSense`
  project, `DpiAwareWindow`-based UI, static `LogUtility`/`AppPaths` utilities. This is
  what gets released to users today.
- **`AIPowered/`** - the in-progress rewrite. Multi-project architecture (`GLSense` host +
  `GLSense.Addin.Core` class library + `GLSense.Loader.Core`, hot-reload capable),
  `BaseWindow`-based UI, `ServiceLocator`-based services. Being developed in parallel;
  most features are ported to both codebases as they're built or fixed.

Both live on a single `main` branch as top-level folders rather than on separate branches,
since they're structurally unrelated codebases (not a fork of the same tree) that will
never be merged back into each other. Keeping them in one branch lets a single commit
represent "ported change X to both codebases," which is how most day-to-day work here
actually happens.

## Plan

`FinalWorkingCode` remains the release version until `AIPowered` reaches full feature
parity and has been validated. Once that happens, `FinalWorkingCode` will be retired
(its folder removed in a commit) - full history stays available via `git log` regardless.

## Notes

- Build artifacts (`bin/`, `obj/`, `.vs/`, `packages/`) and runtime logs are gitignored.
- `AIPowered/bkpup1/` and `AIPowered/bkpup2/` (manual whole-solution backups taken before/
  during the rewrite) and `FinalWorkingCode/MigrationBackup/` are intentionally not
  tracked - git history now serves that purpose.
- See each codebase's own `CLAUDE.md` (inside `FinalWorkingCode/GLSense/` and
  `AIPowered/GLSense/`) for the detailed running log of fixes and design decisions.
