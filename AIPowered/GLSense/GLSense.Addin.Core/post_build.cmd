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
echo STEP 3: Build the zip (Addin.Core's bin output, minus *.pdb)
echo ========================================

if not exist "%CORE_BIN_DIR%" (
    echo ERROR: Core bin directory not found: %CORE_BIN_DIR% - cannot publish.
    goto :SkipManifestPublish
)

REM One destination for both configurations - see the header comment above.
set MANIFEST_DIR=%PROJECT_DIR%\SetupFiles\%CONFIG%\Manifest
set OUT_ZIP=%MANIFEST_DIR%\v%FILE_VERSION%.zip
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

REM SetupFiles is this project's own folder, not GLSense_Logs_New - creating
REM it here is ordinary build output, not subject to section 14.5's
REM "build shouldn't create GLSense_Logs_New folders" rule at all.
if not exist "%MANIFEST_DIR%" mkdir "%MANIFEST_DIR%"

powershell -NoProfile -Command "Compress-Archive -Path '%ZIP_STAGE_DIR%\*' -DestinationPath '%OUT_ZIP%' -Force"

rmdir /S /Q "%ZIP_STAGE_DIR%" 2>nul
del "%ZIP_EXCLUDE_LIST%" >nul 2>&1

if not exist "%OUT_ZIP%" (
    echo WARNING: Failed to create %OUT_ZIP% - nothing for UpdateBootstrapper to extract.
    goto :SkipManifestPublish
)

echo ========================================
echo STEP 4: Write manifest.json alongside the zip
echo ========================================

powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '%OUT_ZIP%').Hash" > "%TEMP%\glsense_checksum.tmp"
set /p ZIP_CHECKSUM=<"%TEMP%\glsense_checksum.tmp"
del "%TEMP%\glsense_checksum.tmp" >nul 2>&1

REM Local time, not UTC (GLAbout displays this as the build date - it should read like
REM the machine's own clock, not a Z-suffixed UTC timestamp that would be a different
REM wall-clock time for whoever is looking at it).
powershell -NoProfile -Command "[DateTime]::Now.ToString('yyyy-MM-ddTHH:mm:ss')" > "%TEMP%\glsense_releasedate.tmp"
set /p RELEASE_DATE=<"%TEMP%\glsense_releasedate.tmp"
del "%TEMP%\glsense_releasedate.tmp" >nul 2>&1

REM Overwritten on every build (not "seed if missing" like PathProvider's own
REM CreateDefaultManifestFile) - the whole point is that a fresh manifest.json
REM + zip sitting together in SetupFiles gives GLSense\post_build.cmd something
REM current to copy into bin\AddinCore\Manifest\, which is what triggers
REM UpdateBootstrapper's extract-on-launch path. No per-Configuration branching
REM here - both Debug and Release follow the same SetupFiles path (see the
REM header comment above). downloadUrl is left empty - nothing downloads this
REM locally, the zip is already sitting right next to the manifest.
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
    echo     "downloadUrl": "",
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
