@echo off
REM If a target directory is passed in, use that; otherwise default to script location
IF "%~1"=="" (
    SET OUTDIR=%~dp0
) ELSE (
    SET OUTDIR=%~1
)

SET NUGET_CACHE=%USERPROFILE%\.nuget\packages

echo =========================================

REM Rename app.config to GLSense.dll.config
IF EXIST "%OUTDIR%app.config" (
    echo Renaming app.config to GLSense.dll.config
    COPY /Y "%OUTDIR%app.config" "%OUTDIR%GLSense.dll.config"
)

REM Copy application manifest
IF EXIST "%~dp0GLSense.app.manifest" (
    echo Copying GLSense.app.manifest to %OUTDIR%
    COPY /Y "%~dp0GLSense.app.manifest" "%OUTDIR%GLSense.app.manifest"
) ELSE (
    echo WARNING: GLSense.app.manifest not found in %~dp0
)

echo Copying SQLite native DLLs...
echo OUTDIR=%OUTDIR%
echo NUGET_CACHE=%NUGET_CACHE%
echo SQLITE_VERSION=%SQLITE_VERSION%

echo Detecting packages...
REM SQLite package version (from csproj)
SET SQLITE_VERSION=1.0.119
SET SQLITE_STUB=%NUGET_CACHE%\stub.system.data.sqlite.core.netframework\%SQLITE_VERSION%

echo SQLITE_STUB=%SQLITE_STUB%

REM e_sqlite3 package version (from cache scan)
SET ESQLITE_VERSION=2.1.11
SET ESQLITE_PACKAGE=%NUGET_CACHE%\sqlitepclraw.lib.e_sqlite3\%ESQLITE_VERSION%

echo ESQLITE_PACKAGE=%ESQLITE_PACKAGE%

echo Creating target folders if missing...
REM Create x86 and x64 directories if they don't exist
IF NOT EXIST "%OUTDIR%x86\" mkdir "%OUTDIR%x86"
IF NOT EXIST "%OUTDIR%x64\" mkdir "%OUTDIR%x64"

echo --- Copying x86 files ---
set SRC_X86_INTEROP=%SQLITE_STUB%\build\net46\x86\SQLite.Interop.dll
echo Looking for %SRC_X86_INTEROP%
IF EXIST "%SRC_X86_INTEROP%" (
    echo Copying x86 SQLite.Interop.dll to %OUTDIR%x86%\
    COPY /Y "%SRC_X86_INTEROP%" "%OUTDIR%x86\"
) ELSE (
    echo WARNING: x86 SQLite.Interop.dll not found in %SRC_X86_INTEROP%
)

set SRC_X86_ESQL=%ESQLITE_PACKAGE%\runtimes\win-x86\native\e_sqlite3.dll
echo Looking for %SRC_X86_ESQL%
IF EXIST "%SRC_X86_ESQL%" (
    echo Copying x86 e_sqlite3.dll to %OUTDIR%x86%\
    COPY /Y "%SRC_X86_ESQL%" "%OUTDIR%x86\"
) ELSE (
    echo WARNING: x86 e_sqlite3.dll not found in %SRC_X86_ESQL%
)

echo --- Copying x64 files ---
set SRC_X64_INTEROP=%SQLITE_STUB%\build\net46\x64\SQLite.Interop.dll
echo Looking for %SRC_X64_INTEROP%
IF EXIST "%SRC_X64_INTEROP%" (
    echo Copying x64 SQLite.Interop.dll to %OUTDIR%x64%\
    COPY /Y "%SRC_X64_INTEROP%" "%OUTDIR%x64\"
) ELSE (
    echo WARNING: x64 SQLite.Interop.dll not found in %SRC_X64_INTEROP%
)

set SRC_X64_ESQL=%ESQLITE_PACKAGE%\runtimes\win-x64\native\e_sqlite3.dll
echo Looking for %SRC_X64_ESQL%
IF EXIST "%SRC_X64_ESQL%" (
    echo Copying x64 e_sqlite3.dll to %OUTDIR%x64%\
    COPY /Y "%SRC_X64_ESQL%" "%OUTDIR%x64\"
) ELSE (
    echo WARNING: x64 e_sqlite3.dll not found in %SRC_X64_ESQL%
)

echo SQLite native DLL copy complete