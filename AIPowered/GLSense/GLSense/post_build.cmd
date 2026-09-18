@echo off
echo ========================================
echo Starting post_build.cmd (GLSense host)
echo ========================================

REM Get the TargetDir parameter (this is the output folder)
set TARGET_DIR=%1
set TARGET_DIR=%TARGET_DIR:"=%
if "%TARGET_DIR:~-1%"=="\" set TARGET_DIR=%TARGET_DIR:~0,-1%

REM Get the project directory (where this script is located)
set PROJECT_DIR=%~dp0
set PROJECT_DIR=%PROJECT_DIR:~0,-1%

REM Get the solution directory (this project's parent folder)
for %%F in ("%PROJECT_DIR%") do set SOLUTION_DIR=%%~dpF
set SOLUTION_DIR=%SOLUTION_DIR:~0,-1%

REM Determine configuration from TargetDir path (bin\{Config}\ structure)
for %%F in ("%TARGET_DIR%") do set CONFIG=%%~nxF

echo Target Directory: %TARGET_DIR%
echo Configuration: %CONFIG%

REM This project's own outputs (GLSense.dll, adxloader.GLSense.dll,
REM adxloader64.GLSense.dll - the actual files registered in Excel's COM
REM registry and loaded at every Excel startup) are DELIBERATELY NOT signed
REM here anymore (see CLAUDE.md section 41 for the original back-and-forth
REM that led to signing them from this script, and section 53 for why that
REM was reversed). Signing these 3 files now happens once, in the separate
REM HOST add-in installer project, as part of building the MSI - NOT on
REM every dev rebuild here. That installer project's own build step must
REM sign the actual DLL/EXE bytes it packages (not merely the resulting
REM .msi wrapper) - an MSI's own Authenticode signature does NOT propagate
REM to the individual files it extracts, so signing only the .msi would
REM silently reintroduce the exact Add-in-Express-refuses-to-load /
REM AV-flags-unsigned-code-in-a-trusted-process problem section 41 already
REM solved once.
REM
REM GLSense.Contracts.dll/GLSense.Shared.dll/GLSense.Loader.Core.dll also sit
REM in this output folder (copied in via ProjectReference), and still sign
REM themselves in their own project's post_build.cmd, on every dev Release
REM rebuild - untouched by this change. PARKED, explicitly unresolved: see
REM CLAUDE.md section 53 - whether those 3 libraries should also move to
REM installer-time signing is a real open question, deliberately not decided
REM here. Raise it once the HOST installer project actually exists; don't
REM assume either answer in the meantime.

echo ========================================
echo Copying AddinCore manifest+zip from GLSense.Addin.Core's SetupFiles
echo ========================================

set ADDINCORE_SOURCE=%SOLUTION_DIR%\GLSense.Addin.Core\SetupFiles\%CONFIG%\Manifest
set ADDINCORE_DEST=%TARGET_DIR%\AddinCore\Manifest

if not exist "%ADDINCORE_SOURCE%" (
    echo WARNING: %ADDINCORE_SOURCE% not found - GLSense.Addin.Core may not have
    echo been built in this configuration yet. Skipping AddinCore manifest copy.
    echo Rebuild GLSense.Addin.Core ^(its post_build.cmd populates SetupFiles^) to fix this.
    goto :SkipAddinCoreCopy
)

if not exist "%ADDINCORE_DEST%" mkdir "%ADDINCORE_DEST%"

REM Delete any zip already sitting here before copying this build's fresh one
REM in. GLSense.Addin.Core's post_build.cmd now names its zip with a timestamp
REM (v{version}_{releaseDateSafe}.zip - see that script's STEP 2b), not a
REM stable version-only name, so a plain xcopy (which never deletes anything in
REM the destination) would leave every previous build's differently-named zip
REM sitting here alongside the new one. That's a real correctness problem, not
REM just clutter: UpdateBootstrapper resolves "the" zip in this exact folder
REM via a bare Directory.GetFiles(dir, "*.zip").First()/FirstOrDefault()
REM wildcard with no way to prefer the newest - with two zips present,
REM whichever one Windows happens to enumerate first could get extracted
REM instead of the one this build actually produced. Source-side cleanup in
REM GLSense.Addin.Core's own post_build.cmd already guarantees %ADDINCORE_SOURCE%
REM itself never holds more than one zip, but that alone doesn't help - xcopy
REM only ADDS/overwrites, it never removes something already in the destination
REM that isn't in the source.
for %%Z in ("%ADDINCORE_DEST%\*.zip") do del /Q "%%Z" 2>nul

xcopy /Y /I "%ADDINCORE_SOURCE%\*" "%ADDINCORE_DEST%\"

echo Copied to: %ADDINCORE_DEST%

:SkipAddinCoreCopy

echo ========================================
echo post_build.cmd (GLSense host) completed
echo ========================================
