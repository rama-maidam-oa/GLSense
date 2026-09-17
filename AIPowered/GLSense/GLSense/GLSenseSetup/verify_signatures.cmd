@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM verify_signatures.cmd - TEMPORARY diagnostic, not a permanent part of the
REM build process. Added specifically to answer one question: does
REM adxpatch.exe (this project's own PostBuildEvent, which runs right before
REM this script) modify GLSense.dll/adxloader.GLSense.dll/adxloader64.GLSense.dll
REM after PreBuildEvent's sign_file.cmd has already signed them?
REM
REM Called from OrbitGLSense.vdproj's PostBuildEvent as:
REM   adxpatch.exe ... && call "$(ProjectDir)verify_signatures.cmd"
REM i.e. AFTER adxpatch.exe runs - so every line this prints reflects the
REM POST-adxpatch state, for direct comparison in the same build log against
REM sign_file.cmd's own "[sign_file][DEBUG] ... post-sign_file.cmd SHA256 ..."
REM lines logged earlier (PRE-adxpatch, during PreBuildEvent).
REM
REM If the SHA256 for a given file is IDENTICAL in both places: adxpatch.exe
REM does not touch that file's bytes, and PreBuildEvent's signing is genuinely
REM what ships. If it DIFFERS, or the signature status below reads anything
REM other than "Valid": adxpatch.exe (or something else in this PostBuildEvent
REM step) invalidated the signature after the fact - see CLAUDE.md for the
REM investigation this diagnostic was added to answer, and remove this script
REM (and its PostBuildEvent wiring) once that question is settled either way.
REM ============================================================================

set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"

echo ========================================
echo POST-ADXPATCH SIGNATURE VERIFICATION (%DATE% %TIME%)
echo ========================================

call :CheckOne "%PROJECT_DIR%\..\bin\Release\GLSense.dll"
call :CheckOne "%PROJECT_DIR%\..\bin\Release\adxloader.GLSense.dll"
call :CheckOne "%PROJECT_DIR%\..\bin\Release\adxloader64.GLSense.dll"

echo ========================================
echo POST-ADXPATCH SIGNATURE VERIFICATION COMPLETE
echo ========================================

exit /b 0

:CheckOne
set "CHECK_FILE=%~1"
if not exist "%CHECK_FILE%" (
    echo [verify-post-adxpatch] %CHECK_FILE% : FILE NOT FOUND
    exit /b 0
)

for /f "usebackq delims=" %%H in (`powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 -LiteralPath '%CHECK_FILE%').Hash"`) do set "CHECK_HASH=%%H"
for /f "usebackq delims=" %%S in (`powershell -NoProfile -Command "(Get-AuthenticodeSignature -LiteralPath '%CHECK_FILE%').Status"`) do set "CHECK_STATUS=%%S"

echo [verify-post-adxpatch] %DATE% %TIME% - %CHECK_FILE% : Status=%CHECK_STATUS% SHA256=%CHECK_HASH%
exit /b 0
