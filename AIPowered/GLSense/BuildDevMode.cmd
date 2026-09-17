@echo off
REM ============================================================================
REM BuildDevMode.cmd - launch Visual Studio with Release-config signing
REM disabled, WITHOUT touching your permanent Windows environment.
REM ============================================================================
REM Double-click this file (or run it from a terminal) whenever you want to
REM Build/Rebuild this solution in Release config repeatedly during local dev
REM - e.g. testing a Release-only bug - without burning real, metered DigiCert
REM Keylocker signing operations on every throwaway build.
REM
REM How it works: this script sets GLSENSE_SKIP_SIGNING=1 only in its OWN
REM process's environment, then launches devenv.exe (Visual Studio) as a
REM CHILD of this process. Windows child processes inherit their parent's
REM environment variables, so the Visual Studio instance this script opens -
REM and everything it builds - sees GLSENSE_SKIP_SIGNING as set. Nothing is
REM written to your user/system environment (no `setx`, no registry) - the
REM variable exists only for as long as this one VS instance stays open.
REM
REM Deliberately NOT a `setx`-based on/off switch: a persistent environment
REM variable is easy to forget about, and forgetting to turn it back off
REM would silently leave every future Release build unsigned indefinitely -
REM including one you eventually ship or hand to someone else. With this
REM launcher, the safe (signed) behavior is what you get by default just by
REM opening Visual Studio normally (Start Menu, double-clicking GLSense.sln,
REM Recent Projects) - there's nothing to remember to undo. Only a build
REM launched from a VS instance THIS script itself opened is ever unsigned.
REM
REM Every skipped signing operation still leaves a "_DEV_UNSIGNED_BUILD.txt"
REM marker file next to the unsigned output (see sign_file.cmd), and
REM GLSense.Build's post_build.cmd prints a loud warning at the end of the
REM whole solution build if any such marker is found - so even a build made
REM this way is hard to mistake for a properly signed one later.
REM
REM When you're done with a batch of dev-mode rebuilds: just close this VS
REM instance and reopen the solution normally - the next build will sign as
REM usual, automatically, with nothing to unset.
REM ============================================================================

set "PROJECT_DIR=%~dp0"
set "SLN_PATH=%PROJECT_DIR%GLSense.sln"

if not exist "%SLN_PATH%" (
    echo [BuildDevMode] ERROR: could not find "%SLN_PATH%".
    pause
    exit /b 1
)

set "GLSENSE_SKIP_SIGNING=1"

echo ============================================================
echo  GLSense DEV MODE - signing is DISABLED in this VS instance
echo  (GLSENSE_SKIP_SIGNING is set only for this process and its
echo  children - closing this VS window returns to normal, signed
echo  behavior the next time you open the solution)
echo ============================================================

REM ===== Locate devenv.exe via vswhere (handles any VS 2022 SKU/edition) =====
set "DEVENV="
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if exist "%VSWHERE%" (
    for /f "usebackq tokens=*" %%P in (`"%VSWHERE%" -latest -prerelease -products * -requires Microsoft.Component.MSBuild -property productPath`) do set "DEVENV=%%P"
)

if "%DEVENV%"=="" (
    echo [BuildDevMode] WARNING: could not locate devenv.exe via vswhere - falling back to PATH.
    set "DEVENV=devenv.exe"
)

start "" "%DEVENV%" "%SLN_PATH%"
