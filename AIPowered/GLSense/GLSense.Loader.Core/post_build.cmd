@echo off
echo ========================================
echo Starting post_build.cmd
echo ========================================

REM ============================================================================
REM This script deliberately does nothing anymore. It used to xcopy this
REM project's own DLL directly into %LOCALAPPDATA%\...\GLSense_Logs_New\
REM Versions\vX\ - that direct-copy-to-Versions approach was removed everywhere
REM (see GLSense.Addin.Core\post_build.cmd's own comment header and CLAUDE.md
REM section 16): the ONLY way DLLs reach Versions\vX\ now is through
REM GLSense.Loader.Core\UpdateBootstrapper.cs extracting a zip, driven entirely
REM by manifest.json.
REM
REM GLSense.Loader.Core.dll still needs to end up in that zip (AddinDomainLoader
REM .Load()'s CreateInstanceAndUnwrap(RemoteLoader) call requires it to be
REM resolvable from the new AppDomain's ApplicationBase) - that's now handled by
REM GLSense.Addin.Core referencing this project directly (a new ProjectReference,
REM added purely for this deployment reason, not because Addin.Core's code
REM actually uses anything from here), so MSBuild copies this DLL into Addin.Core
REM 's own bin output automatically, and GLSense.Addin.Core\post_build.cmd zips
REM that whole folder and publishes it to the local update host.
REM
REM This PostBuildEvent/script is left in place (rather than removed from the
REM .csproj) in case a Loader.Core-specific post-build step is ever needed again
REM - if you're looking for where DLLs get deployed, that's Addin.Core's
REM post_build.cmd now, not this file.
REM ============================================================================

echo GLSense.Loader.Core post_build.cmd: no deployment action - see comment header.

echo ========================================
echo Deployment completed
echo ========================================
