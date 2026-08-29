@echo off
echo ========================================
echo Starting post_build.cmd (GLSense.Loader.Core)
echo ========================================

REM ============================================================================
REM This script no longer xcopies this project's DLL into
REM %LOCALAPPDATA%\...\GLSense_Logs_New\Versions\vX\ directly - that
REM direct-copy-to-Versions approach was removed everywhere (see
REM GLSense.Addin.Core\post_build.cmd's own comment header and CLAUDE.md
REM section 16): the ONLY way DLLs reach Versions\vX\ now is through
REM GLSense.Loader.Core\UpdateBootstrapper.cs extracting a zip, driven entirely
REM by manifest.json.
REM
REM What this script DOES do: sign this project's own DLL, once, right after
REM it's compiled (Release builds only). GLSense.Addin.Core references this
REM project directly (a ProjectReference added purely for deployment - see
REM CLAUDE.md section 16), so MSBuild copies this already-signed DLL into
REM Addin.Core's own bin output automatically; it must never be re-signed at
REM that copy. See sign_file.cmd's own header comment for the full reasoning.
REM ============================================================================

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

call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\GLSense.Loader.Core.dll" "%CONFIG%"

echo ========================================
echo Deployment completed
echo ========================================
