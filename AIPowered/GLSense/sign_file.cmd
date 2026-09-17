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
REM   call "<SolutionDir>sign_file.cmd" "<full path to file>" "<build config>" ["FORCE"]
REM
REM Behavior:
REM   - Debug (or any non-Release) config: prints a message and exits without
REM     touching the file at all.
REM   - Release config, file already carries a valid signature AND that
REM     signature's own certificate has not passed its expiry date, and
REM     FORCE was not passed: prints a "skipping" message and exits without
REM     re-signing - this is what avoids burning a signing operation on a
REM     rebuild that didn't actually change the file's bytes.
REM   - Release config, file is unsigned, OR its existing certificate has
REM     expired, OR FORCE was passed: prints the file name/location, then
REM     (re-)signs it with whatever certificate is currently configured
REM     below.
REM
REM Why "signed" alone isn't a good enough skip condition: `signtool verify
REM /pa` can keep reporting success on an already-signed file purely because
REM of its RFC3161 timestamp countersignature (/tr .../ /td SHA256 below),
REM even after the signing certificate itself has passed its own expiry
REM date - that's the whole point of timestamping, and it's how Windows
REM itself treats the signature. But at least one real consumer of these
REM files does NOT treat it that way: Add-in Express's own loader-trust
REM check appears to look at the leaf certificate's expiry directly, and
REM refused to load the add-in once GLSense.dll/adxloader*.dll's embedded
REM certificate had expired - confirmed by real-world experience - even
REM though `verify /pa` still passed on those exact files. So the skip
REM check here does two things, both must hold to skip:
REM   1. `signtool verify /pa` succeeds (signature chain/timestamp is
REM      intact - i.e. it's genuinely signed, not just claiming to be).
REM   2. The embedded certificate's own NotAfter date (read via PowerShell's
REM      Get-AuthenticodeSignature, not the timestamp) is still in the
REM      future.
REM If either check fails, the file gets (re-)signed - so a rebuild after
REM the cert has expired self-heals instead of silently keeping a
REM signature some consumers won't trust, with no need to remember to pass
REM FORCE by hand.
REM
REM Optional 3rd argument "FORCE": always (re-)signs unconditionally,
REM skipping both checks above entirely. Kept as an explicit escape hatch
REM (e.g. to force a fresh signature for some other reason) but no longer
REM needed just to handle cert expiry - the two-part check above already
REM catches that automatically. No current caller passes it.
REM ============================================================================

set "TARGET_FILE=%~1"
set "BUILD_CONFIG=%~2"
set "FORCE_MODE=%~3"

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

REM ===== Skip only if already signed AND that signature's cert hasn't expired (unless FORCE) =====
if /I "%FORCE_MODE%"=="FORCE" goto :DoSign

"%SIGNTOOL%" verify /pa /q "%TARGET_FILE%" >nul 2>&1
if !errorlevel! neq 0 (
    echo [sign_file] RELEASE mode - not currently ^(validly^) signed - signing: "%TARGET_FILE%"
    goto :DoSign
)

REM Signature chain/timestamp checks out, but that alone doesn't mean the
REM embedded certificate itself hasn't expired since it was signed (see the
REM header comment above) - read the leaf certificate's own NotAfter date
REM directly and only skip re-signing if it's still in the future.
set "CERT_STATE="
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "try { $sig = Get-AuthenticodeSignature -LiteralPath '%TARGET_FILE%'; if ($sig.SignerCertificate -and $sig.SignerCertificate.NotAfter -gt (Get-Date)) { 'VALID' } else { 'EXPIRED' } } catch { 'EXPIRED' }"`) do set "CERT_STATE=%%D"

if /I "%CERT_STATE%"=="VALID" (
    echo [sign_file] RELEASE mode - already signed and certificate still valid, skipping: "%TARGET_FILE%"
    exit /b 0
)

echo [sign_file] RELEASE mode - existing signature's certificate has expired ^(or could not be read^) - re-signing: "%TARGET_FILE%"

:DoSign
"%SIGNTOOL%" sign /csp "%CERT_KSP%" /kc "%CERT_KEY_CONTAINER%" /f "%CERT_FILE%" /p "%CERT_PASSWORD%" /sha1 "%CERT_SHA1%" /tr "%TIMESTAMP_URL%" /td SHA256 /fd SHA256 "%TARGET_FILE%"

if !errorlevel! neq 0 (
    echo [sign_file] ERROR: signing failed for "%TARGET_FILE%"
    exit /b 1
)

echo [sign_file] SUCCESS: signed "%TARGET_FILE%"
exit /b 0
