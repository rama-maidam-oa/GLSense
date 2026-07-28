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
REM easy-to-forget extra process to keep running). This script now drops a
REM fresh zip + manifest.json directly into the LOCAL Manifest folder
REM (%LOCALAPPDATA%\...\GLSense_Logs_New\Manifest\) on every build.
REM GLSense.Loader.Core\UpdateBootstrapper.cs sees both sitting there together
REM on the next Excel launch (or a manual "Reload Add-in" click - see
REM AddinModule.ReloadAddinCore) and extracts the zip into Versions\V{version}\,
REM deleting the zip afterwards - no network involved at all.
REM
REM This deliberately overrides section 14.5's "build shouldn't create
REM deployment folders, PathProvider.cs owns that" rule, specifically for the
REM Manifest folder, specifically for this local testing setup - once this is
REM confirmed working, the plan is to return to real-time/online updates
REM (a real server, no build-time folder writes into GLSense_Logs_New at all).
REM ============================================================================

echo ========================================
echo STEP 1: Resolve version from the compiled DLL
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
echo STEP 2: Build the zip (Addin.Core's bin output, minus *.pdb)
echo ========================================

if not exist "%CORE_BIN_DIR%" (
    echo ERROR: Core bin directory not found: %CORE_BIN_DIR% - cannot publish.
    goto :SkipManifestPublish
)

set MANIFEST_DIR=%LOCALAPPDATA%\ORBIT\Excel_Logs\GLSense_Logs_New\Manifest
set OUT_ZIP=%MANIFEST_DIR%\v%FILE_VERSION%.zip
set OUT_MANIFEST=%MANIFEST_DIR%\manifest.json

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

REM Manifest folder is created here on purpose (see the header comment above -
REM this is a deliberate, temporary override of the "build doesn't create
REM GLSense_Logs_New folders" rule, scoped to this local testing setup only).
if not exist "%MANIFEST_DIR%" mkdir "%MANIFEST_DIR%"

powershell -NoProfile -Command "Compress-Archive -Path '%ZIP_STAGE_DIR%\*' -DestinationPath '%OUT_ZIP%' -Force"

rmdir /S /Q "%ZIP_STAGE_DIR%" 2>nul
del "%ZIP_EXCLUDE_LIST%" >nul 2>&1

if not exist "%OUT_ZIP%" (
    echo WARNING: Failed to create %OUT_ZIP% - nothing for UpdateBootstrapper to extract.
    goto :SkipManifestPublish
)

echo ========================================
echo STEP 3: Write manifest.json alongside the zip
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
REM every time. downloadUrl is left empty - nothing downloads this locally, the
REM zip is already sitting right next to the manifest.
(
    echo [
    echo   {
    echo     "version": "%FILE_VERSION%",
    echo     "releaseDate": "%RELEASE_DATE%",
    echo     "downloadUrl": "",
    echo     "checksum": "%ZIP_CHECKSUM%",
    echo     "notes": "Published by GLSense.Addin.Core post_build.cmd (folder-only test flow)",
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
