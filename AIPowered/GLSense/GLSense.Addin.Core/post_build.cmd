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
REM FOLDER-ONLY testing flow (see CLAUDE.md section 17) - the local-host/HTTP
REM simulation was removed after it kept failing to connect in practice (an
REM easy-to-forget extra process to keep running). This script drops a fresh
REM zip + manifest.json into ONE of two places depending on Configuration:
REM
REM   - Debug:   %LOCALAPPDATA%\...\GLSense_Logs_New\Manifest\ (unchanged) -
REM              GLSense.Loader.Core\UpdateBootstrapper.cs sees both sitting
REM              there together on the next Excel launch (or a manual
REM              "Reload Add-in" click - see AddinModule.ReloadAddinCore) and
REM              extracts the zip into Versions\V{version}\, deleting the zip
REM              afterwards - no network involved at all. This is the local
REM              hot-reload dev-loop flow and is unaffected by the Release
REM              branch below.
REM   - Release: "%PROJECT_DIR%\Manifests\" (this project's own folder, NOT
REM              GLSense_Logs_New) - the main GLSense (host) project references
REM              this fixed, source-relative path to pull the zip+manifest.json
REM              into the shipped MSI. Created if missing.
REM
REM The Debug branch deliberately overrides section 14.5's "build shouldn't
REM create deployment folders, PathProvider.cs owns that" rule, specifically
REM for the Manifest folder, specifically for local testing - once online
REM updates return, this build-time folder write goes away. The Release branch
REM is a different, permanent mechanism (packaging input for the MSI), not a
REM testing convenience, so it is not subject to that same caveat.
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

REM Debug -> local %LOCALAPPDATA% Manifest folder (hot-reload test flow, unchanged).
REM Release -> this project's own "Manifests" folder (MSI packaging input - see
REM the header comment above). Comparison is case-insensitive since MSBuild's
REM $(Configuration) casing can vary by how it was invoked.
if /I "%CONFIG%"=="Release" (
    set MANIFEST_DIR=%PROJECT_DIR%\Manifests
) else (
    set MANIFEST_DIR=%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs_New\Manifest
)
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

REM Manifest folder is created here on purpose (see the header comment above).
REM For Debug this is a deliberate, temporary override of the "build doesn't
REM create GLSense_Logs_New folders" rule, scoped to local testing only. For
REM Release this is just a normal project-relative output folder, not covered
REM by that rule at all.
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
REM + zip sitting together triggers UpdateBootstrapper's extract-on-launch path
REM every time (Debug branch) / a fresh packaging input exists for the MSI
REM (Release branch). downloadUrl is left empty - nothing downloads this
REM locally, the zip is already sitting right next to the manifest.
if /I "%CONFIG%"=="Release" (
    set MANIFEST_NOTES=Published by GLSense.Addin.Core post_build.cmd (Release - MSI packaging input)
) else (
    set MANIFEST_NOTES=Published by GLSense.Addin.Core post_build.cmd (folder-only test flow)
)
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
