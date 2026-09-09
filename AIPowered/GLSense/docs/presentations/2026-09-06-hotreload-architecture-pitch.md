# "Reload, Don't Reinstall" — leadership architecture pitch

**Link**: https://claude.ai/code/artifact/0b8df48e-a423-4891-b36f-26197e3cb1ea

**Status**: published, not yet shared publicly for colleague comment as of 2026-09-06 —
share it via the artifact page's own Share menu before circulating.

## What it is

A 9-slide executive deck making the case to standardize on the AIPowered hot-reload
architecture (branded in the deck as **GLSense Modular**) as GLSense's delivery model
going forward, replacing FinalWorkingCode's single-installer model (branded as
**GLSense Classic**). The internal `AIPowered`/`FinalWorkingCode` folder names are
deliberately not used anywhere in the deck itself, since leadership doesn't recognize
them.

## Slides

1. Cover
2. The cost today — why any change to GLSense Classic means a full MSI reinstall on
   every desk
3. The architecture — host-shell / AppDomain-boundary split diagram
   (`GLSense.dll` vs `GLSense.Addin.Core.dll`)
4. The delivery path — side-by-side flow diagram: 6-step reinstall pipeline vs.
   4-step "adopted automatically" pipeline
5. The safety net — Release History timeline diagram (folders keyed by version *and*
   timestamp, since version numbers can repeat)
6. Why it doesn't break — the two fail-safe design rules (unreachable update source →
   silent fallback; older build loaded from history → defensive interface calls)
7. The trade-off, named — honest accounting of the added engineering discipline this
   requires
8. Side-by-side business comparison table
9. The ask — three concrete approval asks

## Next steps

- Share the artifact link via its own Share menu (private by default).
- Once colleagues leave comments on it, they can be triaged/addressed directly on the
  artifact (a thread must be sent to Claude before it can be read/replied to
  programmatically).
- See `CLAUDE.md` section 47 for the same reference, kept alongside the technical fix
  log.
