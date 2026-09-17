@echo off
setlocal enabledelayedexpansion
echo ========================================================================
echo GLSense.Build: solution build finished
echo ========================================================================

REM Get the TargetDir parameter (this project's own output folder)
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

echo Configuration: %CONFIG%
echo.
echo By the time this ran, MSBuild had already built (in dependency order):
echo   GLSense.Contracts -^> GLSense.Shared -^> GLSense.Loader.Core -^> GLSense + GLSense.Addin.Core
echo Each of those projects' own post_build.cmd already ran too - scroll up
echo for their [sign_file] lines to confirm what was (or wasn't) signed.
echo.
echo Host add-in output (UNSIGNED - see CLAUDE.md section 53; these 3 files
echo are now signed once by the HOST add-in installer project instead):
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\GLSense.dll
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\adxloader.GLSense.dll
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\adxloader64.GLSense.dll
echo.
echo Addin.Core output:
echo   %SOLUTION_DIR%\GLSense.Addin.Core\bin\%CONFIG%\GLSense.Addin.Core.dll

echo.
echo AddinCore hot-reload state (colocated with GLSense.dll - see
echo docs/superpowers/specs/2026-09-04-addincore-colocated-storage-design.md):
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\Manifest\manifest.json
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\Manifest\v*.zip
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\Versions\   (populated at runtime by UpdateBootstrapper)
echo   %SOLUTION_DIR%\GLSense\bin\%CONFIG%\AddinCore\ReleaseHistory.json   (populated at runtime by UpdateBootstrapper)

REM ===== Dev-mode signing check (see sign_file.cmd / CLAUDE.md section 55) =====
REM Runs last, after every project's own post_build.cmd, so this is the final
REM thing printed for the whole solution build - the one place a marker left
REM by GLSENSE_SKIP_SIGNING can't be missed by someone scrolling straight to
REM the bottom of the Output window.
set "DEV_MODE_WARNING_FOUND=0"
for %%D in (
    "%SOLUTION_DIR%\GLSense.Contracts\bin\%CONFIG%"
    "%SOLUTION_DIR%\GLSense.Shared\bin\%CONFIG%"
    "%SOLUTION_DIR%\GLSense.Loader.Core\bin\%CONFIG%"
    "%SOLUTION_DIR%\GLSense.Addin.Core\bin\%CONFIG%"
    "%SOLUTION_DIR%\GLSense.Addin.Core\bin\%CONFIG%\x86"
    "%SOLUTION_DIR%\GLSense.Addin.Core\bin\%CONFIG%\x64"
) do (
    if exist "%%~D\_DEV_UNSIGNED_BUILD.txt" (
        set "DEV_MODE_WARNING_FOUND=1"
        echo   %%~D\_DEV_UNSIGNED_BUILD.txt
    )
)

if "%DEV_MODE_WARNING_FOUND%"=="1" (
    echo.
    echo ****************************************************************
    echo *** WARNING: THIS BUILD CONTAINS UNSIGNED DEV-MODE OUTPUT!    ***
    echo *** GLSENSE_SKIP_SIGNING was set when at least one file above ***
    echo *** was built - see the _DEV_UNSIGNED_BUILD.txt marker^(s^)     ***
    echo *** listed above for exactly which folder^(s^). Do NOT ship,   ***
    echo *** install, or hand off this build until it has been        ***
    echo *** rebuilt with GLSENSE_SKIP_SIGNING unset ^(see CLAUDE.md    ***
    echo *** section 55^).                                             ***
    echo ****************************************************************
)

echo ========================================================================
