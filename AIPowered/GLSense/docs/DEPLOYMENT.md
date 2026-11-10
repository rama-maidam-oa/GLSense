# GLSense / XLEdge (AIPowered) – deploying to ORBIT

How to deploy AIPowered GLSense and XLEdge to the orbitbootapp branches (first done for 11.1.3 on 05-Nov-2026). Follow the steps in order.

A deployment isn't just an MSI copy any more. Each add-in has two parts on the server:

- **The MSI**, for new installs. It contains the first core (zip + `manifest.json`).
- **The update files**, for users who already have the add-in (Online Reload). These are served by `OrbitAddonsController`
  (`/rest/secure/orbit-addons/...?type=glsense|xledge`).

## Release types

| Type | When | What goes to orbitbootapp |
|---|---|---|
| **Full release** | The shell changed (`GLSense.dll`/`XLEdge.dll`, loader, contracts, installer), or this is a new product version | Signed MSI + `orbit-artifacts.xml` version and date + update files |
| **DLL-only release** | Only the core changed (`*.Addin.Core`) | Update files only: the new release folder, with its `manifest.json` entry **appended** to the history file. Also the `orbit-artifacts.xml` date (step 13). The MSI stays as it is. |

## Folder layout on the server (orbitbootapp)

```
orbit-app/ORBIT_HOME/download-artifacts/
├── orbit-artifacts.xml                      <version> and <release_date> per product
└── addons/
    ├── glsense/
    │   ├── OrbitGLSense.msi
    │   └── updates/
    │       ├── glsenseManifests.json        full history: one entry per release, appended
    │       └── <folderName>/                e.g. v11.1.3_2026-11-05T16-10-00
    │           ├── manifest.json
    │           └── <fileName>               e.g. v11.1.3_2026-11-05T16-10-00.zip
    └── xledge/
        ├── OrbitXLEdge.msi
        └── updates/
            ├── xledgeManifests.json
            └── <folderName>/
                ├── manifest.json
                └── <fileName>
```

- The history file name is `<type>Manifests.json`, with this exact case. The server only serves releases listed in it.
- `<folderName>` and `<fileName>` must match the release's `manifest.json` exactly.
- Each history entry is that release's `manifest.json` entry, copied as is.
- **`release` flag:** builds write `"release": true`. Online Reload shows only entries whose `release` is true, `"TRUE"`, blank or missing. When several fix builds are published before users get one, set the earlier entries to `false` in the history file, so only the latest is offered.
- `ORBIT_HOME` is gitignored, so new files must be added with `git add -f`.

## Repos and paths

| | GLSense | XLEdge |
|---|---|---|
| Repo / branch | `D:\SQLLite_Test\GLSense`, branch `11.1.3` | `D:\SQLLite_Test\XLEdge\BusinessLayers\11.1.3`, branch `11.1.3` |
| Solution | `AIPowered\GLSense\GLSense.sln` | `AIPowered\XLEdge\XLEdge.sln` |
| Release text | `GLSense.Addin.Core\ReleaseNotes.txt` line 1 | `XLEdge.Addin.Core\ReleaseNotes.txt` line 1 |
| Manifest + zip | `GLSense.Addin.Core\SetupFiles\Release\Manifest\` | `XLEdge.Addin.Core\SetupFiles\Release\Manifest\` |
| MSI | `GLSense\OrbitGLSenseSetups\Release\OrbitGLSense.msi` | `XLEdge\OrbitXLEdgeSetups\Release\OrbitXLEdge.msi` |
| Signing script | `AIPowered\GLSense\sign_file.cmd` | `AIPowered\XLEdge\sign_file.cmd` (gitignored; copy of GLSense's) |
| Archive | `D:\Orbit Excel Add-ons\Setup Files\GLSense\<version>\` | `D:\Orbit Excel Add-ons\Setup Files\XLEdge\<version>\` |

orbitbootapp `development` = `D:\projects\deveelopmet\orbitbootapp`.

## Steps

### A. Prepare (GLSense and XLEdge repos)

1. List the OISR tickets in this release for each add-in.
2. **Release text:** put the tickets on line 1 of each `ReleaseNotes.txt`, as `OISR-NNNNN: <short summary>`, separated by `; `. This becomes the manifest's `notes`, which every customer sees in Version History and the Online list (OISR-22872). So **don't paste the Jira title**: write a short, plain summary of what changed for the user, with no customer names or codes (e.g. "AVX"), no version prefixes ("11.1.1 -") and no internal terms (AppDomain, Addin.Core).
3. Version and commit date:
   - Check the version: `GLSenseSharedVersion.cs` / `XLEdgeSharedVersion.cs`.
   - XLEdge: set `XLEdgeAppConstants.DefaultCommitDate` to the release day.
   - AIPowered GLSense has no commit-date constant (it shows the build's `releaseDate`).
4. Commit and push both branches. The HEAD at build time is the build commit you quote in Jira.

### B. Build (Release)

**One-time per machine (GLSense only):** GLSense's `GLSense.Contracts` / `GLSense.Shared` / `GLSense.Loader.Core` use password-protected `.pfx` strong-name keys. A build that can't read them fails with `MSB3325 ... VS_KEY_xxxxxxxx`.
- Fix it once, in an **administrator** PowerShell: `sn -d VS_KEY_xxxxxxxx`, then `sn -i "<project>\<project>.pfx" VS_KEY_xxxxxxxx`, one key at a time.
- `sn.exe` is in `C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\`, and each import asks for that key's `.pfx` password.
- Key containers are machine-wide on this PC (`sn -m`), so "Object already exists" means a stale container must be deleted first.

XLEdge's keys are password-free `.snk` files. XLEdge's `sign_file.cmd` must be present at `AIPowered\XLEdge\` (a gitignored copy of GLSense's). Without it, a Release build fails with `XLEDGE0002`.

5. Build each solution with `devenv.com <sln> /Rebuild Release`, run from PowerShell (from Bash it exits without building). The Release build:
   - signs the core DLLs before zipping;
   - writes the manifest and zip, plus `InstallerPayload\Orbit<Name>.zip`;
   - signs the shell DLLs in the setup project's pre-build step;
   - builds the MSI.
6. Verify each build:
   - The MSI has today's timestamp.
   - `manifest.json` has the right `version`, `"release": true`, and the OISR text in `notes`.
   - The shell DLLs and core DLL are signed (`Get-AuthenticodeSignature` → Valid).
   - If a DigiCert timestamp call failed during the build, rebuild. A failed signing call doesn't stop the build.
   - **MSI codes (full release):** compare each new MSI against the one it replaces in orbitbootapp. Read `ProductVersion`, `ProductCode` and `UpgradeCode` with the `WindowsInstaller.Installer` COM object. Check that:
     - the **UpgradeCode is the same** (GLSense `{D4B153A7-1EE7-49BB-94C8-912EC0D469F0}`, XLEdge `{D090F1E0-CA8E-4529-B78D-981652FFF708}`);
     - the **ProductCode is different**;
     - the **ProductVersion is higher**.

     Otherwise the new MSI installs next to the old add-in instead of replacing it. On 05-Nov-2026 the recreated GLSense installer had the wrong UpgradeCode, fixed in commit `e4aee8e`.

### C. Sign and archive (full release only)

7. Sign each MSI with `sign_file.cmd "<msi>" Release`. Check it with `Get-AuthenticodeSignature`: Valid, signer Orbit Analytics.
8. Copy the MSI, `manifest.json` and zip to the archive folder for the version.

### D. orbitbootapp (`development`)

9. Run `git -C <orbitbootapp> pull --ff-only`.
10. Full release: copy the signed MSIs over `addons/glsense/OrbitGLSense.msi` and `addons/xledge/OrbitXLEdge.msi`.
11. For each add-in, create `addons/<type>/updates/<folderName>/` and copy that build's `manifest.json` and zip into it.
12. **Append** that build's manifest entry to `addons/<type>/updates/<type>Manifests.json`. Create the file as a one-entry array the first time. Never remove older entries. To hide an earlier fix build, set its `release` to `false`.
13. **Always** update `orbit-artifacts.xml`, keeping the file's encoding:
    - Set `<release_date>` (DDMMMYYYY, e.g. `05NOV2026`) in the GLSENSE / XLEDGE product block of every add-in you released.
    - On a full release, also set `<version>`.
14. Commit only these files, with `git add -f`. The commit message is:
    - line 1: all OISR keys;
    - then one line per ticket, `OISR-NNNNN <Jira title>`.
15. Push, then confirm with `git show --stat` and `git branch -r --contains HEAD`.

### E. Verify and hand off

16. On a test machine:
    - A fresh MSI install loads the add-in.
    - On an existing install: log in, open Reload, choose Online. The new release is listed with its OISR notes, and it downloads and loads.
17. In Jira, for each ticket:
    - Add a comment with the release date and the build commit (one commit, the version branch only).
    - Set the fix version.
    - Move it to **QA Review**. The assignee stays the developer.
