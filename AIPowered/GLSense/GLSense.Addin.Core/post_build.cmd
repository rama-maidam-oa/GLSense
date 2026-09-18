@echo off
echo ========================================
echo Starting post_build.cmd
echo ========================================

REM Get the TargetDir parameter (this is the output folder)
set TARGET_DIR=%1
REM Remove quotes if present
set TARGET_DIR=%TARGET_DIR:"=%
REM Remove trailing backslash if present
if "%TARGET_DIR:~-1%"=="\" set TARGET_DIR=%TARGET_DIR:~0,-1%

echo Target Directory: %TARGET_DIR%

REM Get the project directory (where this script is located)
set PROJECT_DIR=%~dp0
set PROJECT_DIR=%PROJECT_DIR:~0,-1%
echo Project Directory: %PROJECT_DIR%

REM Get the solution directory
for %%F in ("%PROJECT_DIR%") do set SOLUTION_DIR=%%~dpF
set SOLUTION_DIR=%SOLUTION_DIR:~0,-1%
echo Solution Directory: %SOLUTION_DIR%

REM Determine configuration from TargetDir path (assuming standard bin\{Config}\ structure)
for %%F in ("%TARGET_DIR%") do set CONFIG=%%~nxF
echo Configuration: %CONFIG%

set CORE_BIN_DIR=%TARGET_DIR%

echo Core Bin Dir: %CORE_BIN_DIR%

REM ============================================================================
REM Both Debug and Release write a fresh zip + manifest.json into this
REM project's own SetupFiles\{Config}\Manifest\ folder - transient build
REM output (gitignored, not committed, regenerated every build). This is a
REM staging/hand-off point only: GLSense\post_build.cmd (the host project)
REM copies from here into its own bin\{Config}\AddinCore\Manifest\, which is
REM what PathProvider.ManifestDirectory now resolves to at runtime (colocated
REM with GLSense.dll itself, not GLSense_Logs) - see
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

echo ========================================
echo STEP 1: Sign this project's deliverables (Release only)
echo ========================================

REM Sign this project's own DLL here, once, right after it's compiled and
REM BEFORE it gets zipped up below - the zip must contain a signed DLL.
REM GLSense.Contracts.dll/GLSense.Shared.dll/GLSense.Loader.Core.dll also sit
REM in CORE_BIN_DIR (copied in via ProjectReference), but they are NOT signed
REM here - they were already signed once, in their own project's post_build.cmd,
REM before MSBuild copied them here. Re-signing those copies would just waste
REM a signing operation. See sign_file.cmd's own header comment for the full
REM reasoning.
call "%SOLUTION_DIR%\sign_file.cmd" "%CORE_BIN_DIR%\GLSense.Addin.Core.dll" "%CONFIG%"

REM x86\SQLite.Interop.dll / x64\SQLite.Interop.dll already carry a valid
REM vendor (System.Data.SQLite) signature - confirmed via
REM Get-AuthenticodeSignature, left alone. e_sqlite3.dll (both arches) ships
REM genuinely UNSIGNED from its NuGet package (sqlitepclraw.lib.e_sqlite3) -
REM confirmed the same way - and it's a native DLL sitting inside the
REM client-facing zip, so it gets signed here. sqlite_postbuild.cmd (chained
REM before this script in GLSense.Addin.Core.csproj's PostBuildEvent) has
REM already copied both files into CORE_BIN_DIR\x86\ and CORE_BIN_DIR\x64\ by
REM the time this runs.
call "%SOLUTION_DIR%\sign_file.cmd" "%CORE_BIN_DIR%\x86\e_sqlite3.dll" "%CONFIG%"
call "%SOLUTION_DIR%\sign_file.cmd" "%CORE_BIN_DIR%\x64\e_sqlite3.dll" "%CONFIG%"

echo ========================================
echo STEP 2: Resolve version from the compiled DLL
echo ========================================

REM Read the version from the just-built DLL's file-version metadata instead of a
REM hardcoded literal. Single source of truth is GLSenseSharedVersion.cs (solution
REM root), compiled into every assembly via AssemblyFileVersion - read back out here
REM so this script never has to be hand-edited when the version changes.
powershell -NoProfile -Command "$v = (Get-Item '%CORE_BIN_DIR%\GLSense.Addin.Core.dll').VersionInfo; '{0}.{1}.{2}' -f $v.FileMajorPart, $v.FileMinorPart, $v.FileBuildPart" > "%TEMP%\glsense_version.tmp"
set /p FILE_VERSION=<"%TEMP%\glsense_version.tmp"
del "%TEMP%\glsense_version.tmp" >nul 2>&1

if "%FILE_VERSION%"=="" (
    echo WARNING: Could not read version from GLSense.Addin.Core.dll - falling back to 0.0.0
    set FILE_VERSION=0.0.0
)

echo Version: %FILE_VERSION%

echo ========================================
echo STEP 2b: Resolve release date, release folder name, and zip file name
echo ========================================

REM Computed here - before the zip is built in STEP 3 - because the zip's own
REM filename (FILE_NAME, below) now embeds this same timestamp (see STEP 3's
REM comment for why). Used to just be computed later, in STEP 4, when it was
REM only needed for the manifest.json body itself.
REM Local time, not UTC (GLAbout displays this as the build date - it should read
REM like the machine's own clock, not a Z-suffixed UTC timestamp that would be a
REM different wall-clock time for whoever is looking at it).
powershell -NoProfile -Command "[DateTime]::Now.ToString('yyyy-MM-ddTHH:mm:ss')" > "%TEMP%\glsense_releasedate.tmp"
set /p RELEASE_DATE=<"%TEMP%\glsense_releasedate.tmp"
del "%TEMP%\glsense_releasedate.tmp" >nul 2>&1

REM Same algorithm as GLSense.Shared\ReleaseHistoryStore.BuildFolderName -
REM replace every character in the release date that's invalid in a Windows
REM file/folder name (just the two colons in "HH:mm:ss" for this date format)
REM with '-'. Computed once and reused for both FOLDER_NAME and FILE_NAME below.
powershell -NoProfile -Command "$illegal = [System.IO.Path]::GetInvalidFileNameChars(); -join ('%RELEASE_DATE%'.ToCharArray() | ForEach-Object { if ($illegal -contains $_) { '-' } else { $_ } })" > "%TEMP%\glsense_safedate.tmp"
set /p SAFE_DATE=<"%TEMP%\glsense_safedate.tmp"
del "%TEMP%\glsense_safedate.tmp" >nul 2>&1

REM FOLDER_NAME: recorded in manifest.json purely as metadata/traceability - the
REM folder a release is actually extracted into once adopted is still computed
REM independently, at extraction time, by ReleaseHistoryStore.BuildFolderName
REM (UpdateBootstrapper never reads this field back) - kept here so the folder
REM name a release WILL get is visible directly from its own manifest.json
REM without cross-referencing ReleaseHistory.json. Same lowercase
REM "v{version}_{safeDate}" shape as BuildFolderName itself, so the two can
REM never render as different-looking values for the identical
REM version+releaseDate pair. Lowercase "v" (not "V") specifically so this
REM matches FILE_NAME's own "v{version}_{safeDate}.zip" prefix below.
set FOLDER_NAME=v%FILE_VERSION%_%SAFE_DATE%

REM FILE_NAME: the zip's own filename. Used to be just "v{version}.zip" - only
REM safely unique because the version rarely changes between builds; two builds
REM at the same version on the same day would otherwise silently overwrite one
REM another with no way to tell them apart later. Timestamp-suffixed the same
REM way FOLDER_NAME is, for the identical uniqueness reason. MUST keep the
REM lowercase "v" prefix: GLReloadSourcePicker.xaml.cs's Offline folder scan
REM searches for "v*.zip" specifically. Confirmed via a full grep of every
REM *.zip reference in the solution before making this change: every other
REM reader (UpdateBootstrapper included) only ever globs "*.zip" and never
REM compares against a specific literal filename, so this prefix is the only
REM real naming constraint anywhere in the codebase.
set FILE_NAME=v%FILE_VERSION%_%SAFE_DATE%.zip

echo Release Date: %RELEASE_DATE%
echo Folder Name: %FOLDER_NAME%
echo File Name: %FILE_NAME%

echo ========================================
echo STEP 3: Build the zip (Addin.Core's bin output, minus *.pdb)
echo ========================================

if not exist "%CORE_BIN_DIR%" (
    echo ERROR: Core bin directory not found: %CORE_BIN_DIR% - cannot publish.
    goto :SkipManifestPublish
)

REM One destination for both configurations - see the header comment above.
set MANIFEST_DIR=%PROJECT_DIR%\SetupFiles\%CONFIG%\Manifest
set OUT_ZIP=%MANIFEST_DIR%\%FILE_NAME%
set OUT_MANIFEST=%MANIFEST_DIR%\manifest.json

echo Manifest Output Dir: %MANIFEST_DIR%

REM Mirror CORE_BIN_DIR into a small staging folder first, excluding *.pdb
REM (debug symbols aren't needed to run the add-in, just dead weight in the
REM zip). Compress-Archive has no clean way to exclude a pattern while zipping
REM a whole folder without flattening the x86\/x64\/de\/runtimes\ subfolder
REM structure, so a filtered copy first is the reliable option. The ORIGINAL
REM %CORE_BIN_DIR%\*.pdb files are untouched - only this zip excludes them, so
REM local Visual Studio debugging still works exactly as before.
REM
REM CORE_BIN_DIR already includes GLSense.Loader.Core.dll automatically (see
REM the ProjectReference in GLSense.Addin.Core.csproj) - AddinDomainLoader
REM .Load()'s CreateInstanceAndUnwrap(RemoteLoader) call needs that DLL
REM resolvable from wherever this zip ends up extracted.
set ZIP_STAGE_DIR=%TEMP%\GLSense_ZipStage_%RANDOM%
set ZIP_EXCLUDE_LIST=%TEMP%\glsense_zip_exclude.txt

if exist "%ZIP_STAGE_DIR%" rmdir /S /Q "%ZIP_STAGE_DIR%"
mkdir "%ZIP_STAGE_DIR%"
echo .pdb> "%ZIP_EXCLUDE_LIST%"

xcopy /Y /E /I /EXCLUDE:%ZIP_EXCLUDE_LIST% "%CORE_BIN_DIR%\*" "%ZIP_STAGE_DIR%\" 2>&1

REM SetupFiles is this project's own folder, not GLSense_Logs - creating
REM it here is ordinary build output, not subject to section 14.5's
REM "build shouldn't create GLSense_Logs folders" rule at all.
if not exist "%MANIFEST_DIR%" mkdir "%MANIFEST_DIR%"

REM Delete any zip already sitting here before creating this build's zip.
REM Previously unnecessary - the filename was version-only (v{version}.zip), so
REM a same-version rebuild just overwrote it in place. Now that FILE_NAME embeds
REM a timestamp (see STEP 2b), every build produces a differently-named zip, so
REM without this cleanup, MANIFEST_DIR would accumulate one zip per build.
REM That's a real correctness problem, not just clutter: UpdateBootstrapper and
REM GLReloadSourcePicker's Offline scan both resolve "the" zip via a bare
REM Directory.GetFiles(dir, "*.zip").First()/FirstOrDefault() wildcard with no
REM way to prefer the newest - with two zips present, whichever one Windows
REM happens to enumerate first could get extracted/staged instead of the one
REM this build actually produced. GLSense\post_build.cmd (the host project)
REM has the matching cleanup on its own copy destination, for the same reason -
REM see that file's own comment.
for %%Z in ("%MANIFEST_DIR%\*.zip") do del /Q "%%Z" 2>nul

powershell -NoProfile -Command "Compress-Archive -Path '%ZIP_STAGE_DIR%\*' -DestinationPath '%OUT_ZIP%' -Force"

rmdir /S /Q "%ZIP_STAGE_DIR%" 2>nul
del "%ZIP_EXCLUDE_LIST%" >nul 2>&1

if not exist "%OUT_ZIP%" (
    echo WARNING: Failed to create %OUT_ZIP% - nothing for UpdateBootstrapper to extract.
    goto :SkipManifestPublish
)

echo ========================================
echo STEP 3b: Release only - stage a fixed-name copy for the Setup project
echo ========================================

REM OrbitGLSense.vdproj's File System editor needs a SourcePath that never
REM changes release to release - a versioned filename like v11.1.2.zip
REM otherwise forces a manual vdproj edit (and rebuild-the-installer-project
REM step) every single release. InstallerPayload is a SEPARATE folder from
REM Manifest\ above, read ONLY by the .vdproj at installer-build time.
REM GLSense\post_build.cmd's own xcopy (which feeds the live
REM bin\%CONFIG%\AddinCore\Manifest\ folder that PathProvider/UpdateBootstrapper
REM actually read at runtime, for both the Debug dev-loop and a real install)
REM only ever copies from SetupFiles\%CONFIG%\Manifest\, never from here - so
REM this can never introduce a second, ambiguous *.zip into that live folder.
REM Release only: Debug has no installer to feed.
REM INSTALLER_PAYLOAD_DIR is set OUTSIDE the if-block on purpose: cmd.exe
REM expands a parenthesized block's %VAR% references at parse time, before any
REM statement inside that same block (like "set") has run - setting it inside
REM the block below would make every %INSTALLER_PAYLOAD_DIR% reference in that
REM same block resolve to blank instead of the path just assigned.
set INSTALLER_PAYLOAD_DIR=%PROJECT_DIR%\SetupFiles\Release\InstallerPayload
if /I "%CONFIG%"=="Release" (
    if not exist "%INSTALLER_PAYLOAD_DIR%" mkdir "%INSTALLER_PAYLOAD_DIR%"
    copy /Y "%OUT_ZIP%" "%INSTALLER_PAYLOAD_DIR%\OrbitGLSense.zip" >nul
    echo Published: %INSTALLER_PAYLOAD_DIR%\OrbitGLSense.zip
)

echo ========================================
echo STEP 4: Write manifest.json alongside the zip
echo ========================================

powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '%OUT_ZIP%').Hash" > "%TEMP%\glsense_checksum.tmp"
set /p ZIP_CHECKSUM=<"%TEMP%\glsense_checksum.tmp"
del "%TEMP%\glsense_checksum.tmp" >nul 2>&1

REM RELEASE_DATE/FOLDER_NAME/FILE_NAME were already resolved in STEP 2b, before
REM the zip was built - reused here as-is so every field in this manifest
REM describes the exact same release/timestamp, computed exactly once.
REM
REM Overwritten on every build (not "seed if missing" like PathProvider's own
REM CreateDefaultManifestFile) - the whole point is that a fresh manifest.json
REM + zip sitting together in SetupFiles gives GLSense\post_build.cmd something
REM current to copy into bin\AddinCore\Manifest\, which is what triggers
REM UpdateBootstrapper's extract-on-launch path. No per-Configuration branching
REM here - both Debug and Release follow the same SetupFiles path (see the
REM header comment above).
REM
REM "downloadUrl" was removed entirely (rather than left as an empty string) -
REM it was never read by anything for this LOCAL manifest (confirmed via a full
REM grep: the only real reader of VersionInfo.DownloadUrl is
REM GLReloadSourcePicker.xaml.cs's ONLINE flow, which parses a completely
REM separate JSON payload fetched live from {LoginUrl}/glsense/projectdlls, not
REM this file) - the zip is already sitting right next to this manifest, so
REM there was never anywhere for a local download URL to point.
REM
REM "release" (new): always written as false for now - a placeholder for a
REM future "was this a deliberately cut, non-dev release" flag; no consumer
REM reads it yet.
REM
REM "folderName"/"fileName" (new): FOLDER_NAME/FILE_NAME are the exact same
REM values already resolved in STEP 2b - see that step's comments for what
REM they mean and why they're safe to add.
REM
REM Notes text: this manifest.json "notes" field is what GLReleaseHistoryBrowser
REM (GLSense\Views\GLReleaseHistoryBrowser.xaml) shows in its "Notes" column for
REM this release forever after, so a generic "Published by post_build.cmd" string
REM isn't useful for telling releases apart later. ReleaseNotes.txt (this project's
REM own folder, included as a project item - see GLSense.Addin.Core.csproj - so it's
REM directly editable from Solution Explorer/IDE, and committed to source control
REM alongside whatever code change it describes) lets a developer put a real note -
REM typically an OISR ticket reference - on its first line before building; only that
REM first line is read (batch `set /p` reads one line). This is used the same way for
REM BOTH Debug and Release builds - no per-Configuration branching. Leave the first
REM line blank (or delete the file) to fall back to the generic message below.
REM Remember to update/clear ReleaseNotes.txt's first line before your NEXT unrelated
REM build, or that same note will get baked into that release too.
REM The whole "read line 1, trim it, fall back to a default if blank/missing" job is
REM done in ONE PowerShell call (piped through a temp file, then a plain `set /p` -
REM matching the FILE_VERSION/ZIP_CHECKSUM/RELEASE_DATE pattern already used above)
REM rather than in batch, after direct testing turned up three separate, compounding
REM cmd.exe pitfalls with doing this in batch: (1) `set /p VAR=<file` does not
REM reliably stop at line 1 for an LF-only text file (the kind most editors/Write
REM tools produce) - with no CR before the first LF it silently reads the ENTIRE file
REM into the variable instead; (2) `set VAR=` (no value) UNDEFINES the variable
REM rather than setting it to an empty string, so `set /p` reading a blank line can
REM leave it fully undefined, not merely empty; (3) `%VAR:search=replace%` substring
REM syntax on an undefined (not just empty) variable does not reliably expand to
REM empty - it can leave mangled literal text behind. Doing the trim + blank-check +
REM default entirely in PowerShell means MANIFEST_NOTES is always set from a single,
REM already-resolved, non-empty, single-line value - none of the three pitfalls above
REM can occur regardless of how ReleaseNotes.txt was saved or edited.
REM Notes always come from ReleaseNotes.txt regardless of Configuration (Debug/Release)
REM - no per-Config branching here anymore. The only fallback is for when the file is
REM missing or its first line is blank.
set NOTES_FILE=%PROJECT_DIR%\ReleaseNotes.txt
set NOTES_DEFAULT=Published by the GLSense build process
powershell -NoProfile -Command "$n=''; if (Test-Path -LiteralPath '%NOTES_FILE%') { $f = Get-Content -LiteralPath '%NOTES_FILE%' -TotalCount 1 -ErrorAction SilentlyContinue; if ($f) { $n = $f.Trim() } }; if ([string]::IsNullOrWhiteSpace($n)) { $n = '%NOTES_DEFAULT%' }; $n" > "%TEMP%\glsense_notes.tmp"
set /p MANIFEST_NOTES=<"%TEMP%\glsense_notes.tmp"
del "%TEMP%\glsense_notes.tmp" >nul 2>&1
REM Strip characters that would break the hand-rolled JSON below.
set MANIFEST_NOTES=%MANIFEST_NOTES:"=%
(
    echo [
    echo   {
    echo     "version": "%FILE_VERSION%",
    echo     "releaseDate": "%RELEASE_DATE%",
    echo     "release": false,
    echo     "folderName": "%FOLDER_NAME%",
    echo     "fileName": "%FILE_NAME%",
    echo     "checksum": "%ZIP_CHECKSUM%",
    echo     "notes": "%MANIFEST_NOTES%",
    echo     "mandatory": false
    echo   }
    echo ]
) > "%OUT_MANIFEST%"

echo Published: %OUT_ZIP%
echo Published: %OUT_MANIFEST%

:SkipManifestPublish

echo ========================================
echo Deployment completed
echo ========================================
