@echo off
echo ========================================
echo Starting post_build.cmd (GLSense.Contracts)
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

REM Sign this project's own DLL here, once, right after it's compiled - every
REM downstream project (GLSense, GLSense.Addin.Core) that copies this DLL into
REM its own output folder inherits an already-signed copy for free, so it must
REM never be re-signed at those other locations. See sign_file.cmd's own
REM header comment for the full reasoning.
call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\GLSense.Contracts.dll" "%CONFIG%"

echo ========================================
echo post_build.cmd (GLSense.Contracts) completed
echo ========================================
