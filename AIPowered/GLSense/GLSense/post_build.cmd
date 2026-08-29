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

REM Only this project's own outputs are signed here: the host COM add-in DLL
REM itself, plus the two Add-in Express loader stubs (the actual files
REM registered in Excel's COM registry and loaded at every Excel startup -
REM the most important files to sign for SmartScreen/AV trust).
REM
REM GLSense.Contracts.dll/GLSense.Shared.dll/GLSense.Loader.Core.dll also sit
REM in this output folder (copied in via ProjectReference), but they are NOT
REM signed here - they were already signed once, in their own project's
REM post_build.cmd, before MSBuild copied them here. Re-signing those copies
REM would just waste a signing operation. See sign_file.cmd's own header
REM comment for the full reasoning.
call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\GLSense.dll" "%CONFIG%"
call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\adxloader.GLSense.dll" "%CONFIG%"
call "%SOLUTION_DIR%\sign_file.cmd" "%TARGET_DIR%\adxloader64.GLSense.dll" "%CONFIG%"

echo ========================================
echo post_build.cmd (GLSense host) completed
echo ========================================
