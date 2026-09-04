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

REM This project's own outputs are signed here: the host COM add-in DLL
REM itself, plus the two Add-in Express loader stubs (the actual files
REM registered in Excel's COM registry and loaded at every Excel startup -
REM the most important files to sign for SmartScreen/AV trust, since an MSI
REM installer's own signature never propagates to the individual files it
REM installs - see CLAUDE.md section 40 for the full back-and-forth history
REM on this decision).
REM
REM Plain (non-FORCE) calls: sign only if not already validly signed, same
REM as every other project's post_build.cmd - skips a wasted signing
REM operation on a rebuild that didn't change the bytes. NOTE this does NOT
REM protect against the specific expired-cert failure mode documented in
REM CLAUDE.md section 40 (signtool verify /pa can keep passing on a
REM timestamped signature even after the signing cert itself expires, while
REM Add-in Express's own loader-trust check does not appear to honor that
REM timestamp the same way) - if that recurs, pass "FORCE" as a 3rd argument
REM to each call below (sign_file.cmd already supports it) to always
REM re-sign with the current cert regardless of what's already there.
REM
REM GLSense.Contracts.dll/GLSense.Shared.dll/GLSense.Loader.Core.dll also sit
REM in this output folder (copied in via ProjectReference), but they are NOT
REM signed here - they were already signed once, in their own project's
REM post_build.cmd, before MSBuild copied them here. Re-signing those copies
REM would just waste a signing operation.
call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\GLSense.dll" "%CONFIG%"
call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\adxloader.GLSense.dll" "%CONFIG%"
call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\adxloader64.GLSense.dll" "%CONFIG%"

echo ========================================
echo Copying AddinCore manifest+zip from GLSense.Addin.Core's SetupFiles
echo ========================================

set ADDINCORE_SOURCE=%SOLUTION_DIR%\GLSense.Addin.Core\SetupFiles\%CONFIG%\Manifest
set ADDINCORE_DEST=%TARGET_DIR%\AddinCore\Manifest

if not exist "%ADDINCORE_SOURCE%" (
    echo WARNING: %ADDINCORE_SOURCE% not found - GLSense.Addin.Core may not have
    echo been built in this configuration yet. Skipping AddinCore manifest copy.
    echo Build GLSense.Build.csproj instead of GLSense.csproj alone to avoid this.
    goto :SkipAddinCoreCopy
)

if not exist "%ADDINCORE_DEST%" mkdir "%ADDINCORE_DEST%"
xcopy /Y /I "%ADDINCORE_SOURCE%\*" "%ADDINCORE_DEST%\"

echo Copied to: %ADDINCORE_DEST%

:SkipAddinCoreCopy

echo ========================================
echo post_build.cmd (GLSense host) completed
echo ========================================
