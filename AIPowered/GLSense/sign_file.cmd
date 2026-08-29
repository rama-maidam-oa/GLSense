@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM sign_file.cmd - shared Authenticode signing helper (DigiCert Keylocker)
REM ============================================================================
REM Called from each project's own post_build.cmd, right after that project's
REM own first-party DLL/EXE is compiled - NEVER from a downstream project that
REM merely references/copies that file. An Authenticode signature is embedded
REM in the file's own bytes, so a copy of an already-signed file is already
REM signed; re-running signtool on every copy of the same DLL across multiple
REM output folders would just waste signing operations for nothing.
REM
REM Cert/signtool details sourced from the known-working reference script:
REM   %LOCALAPPDATA%\ORBIT\Excel_Logs\sign_publish.txt
REM
REM Usage:
REM   call "<SolutionDir>sign_file.cmd" "<full path to file>" "<build config>"
REM
REM Behavior:
REM   - Debug (or any non-Release) config: prints a message and exits without
REM     touching the file at all.
REM   - Release config, file already carries a valid signature (signtool
REM     verify /pa succeeds): prints a "skipping" message and exits without
REM     re-signing - this is what avoids burning a signing operation on a
REM     rebuild that didn't actually change the file's bytes.
REM   - Release config, file is unsigned (or was just recompiled with new
REM     bytes, so its old signature no longer applies): prints the file name/
REM     location, then signs it.
REM ============================================================================

set "TARGET_FILE=%~1"
set "BUILD_CONFIG=%~2"

if "%TARGET_FILE%"=="" (
    echo [sign_file] ERROR: no file path was passed to sign_file.cmd.
    exit /b 1
)

if /I not "%BUILD_CONFIG%"=="Release" (
    echo [sign_file] DEBUG mode - no signing needed for "%TARGET_FILE%".
    exit /b 0
)

if not exist "%TARGET_FILE%" (
    echo [sign_file] WARNING: file not found, skipping - "%TARGET_FILE%"
    exit /b 0
)

REM ===== Locate signtool.exe (latest installed Windows SDK) =====
set "SIGNTOOL="
for %%v in (10.0.26100.0 10.0.22621.0 10.0.22000.0 10.0.19041.0 10.0.17763.0 10.0.16299.0) do (
    if exist "C:\Program Files (x86)\Windows Kits\10\bin\%%v\x64\signtool.exe" (
        set "SIGNTOOL=C:\Program Files (x86)\Windows Kits\10\bin\%%v\x64\signtool.exe"
        goto :FoundSignTool
    )
)
if exist "C:\Program Files (x86)\Windows Kits\8.1\bin\x64\signtool.exe" (
    set "SIGNTOOL=C:\Program Files (x86)\Windows Kits\8.1\bin\x64\signtool.exe"
)

:FoundSignTool
if "%SIGNTOOL%"=="" (
    echo [sign_file] ERROR: signtool.exe not found - install the Windows SDK. Skipping "%TARGET_FILE%".
    exit /b 1
)

REM ===== DigiCert Keylocker signing details (same cert as sign_publish.txt) =====
set "CERT_KSP=DigiCert Signing Manager KSP"
set "CERT_KEY_CONTAINER=key_1272019481"
set "CERT_FILE=C:\Program Files\DigiCert\DigiCert Keylocker Tools\orbit_analytics_inc_1272019481_New.p7b"
set "CERT_PASSWORD=ZWCnQMhPFKbi"
set "CERT_SHA1=699E762289B0C51419206A0CA3A2591E6BBC1FD2"
set "TIMESTAMP_URL=http://timestamp.digicert.com"

REM ===== Skip if the file already carries a valid signature =====
"%SIGNTOOL%" verify /pa /q "%TARGET_FILE%" >nul 2>&1
if !errorlevel! equ 0 (
    echo [sign_file] RELEASE mode - already signed, skipping: "%TARGET_FILE%"
    exit /b 0
)

echo [sign_file] RELEASE mode - signing: "%TARGET_FILE%"

"%SIGNTOOL%" sign /csp "%CERT_KSP%" /kc "%CERT_KEY_CONTAINER%" /f "%CERT_FILE%" /p "%CERT_PASSWORD%" /sha1 "%CERT_SHA1%" /tr "%TIMESTAMP_URL%" /td SHA256 /fd SHA256 "%TARGET_FILE%"

if !errorlevel! neq 0 (
    echo [sign_file] ERROR: signing failed for "%TARGET_FILE%"
    exit /b 1
)

echo [sign_file] SUCCESS: signed "%TARGET_FILE%"
exit /b 0
