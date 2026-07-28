// GLSenseSharedVersion.cs - solution root
//
// SINGLE SOURCE OF TRUTH for the GLSense product version number.
//
// Linked (not copied) into every project in the solution (GLSense,
// GLSense.Addin.Core, GLSense.Shared, GLSense.Contracts, GLSense.Loader.Core) via
// a <Compile Include="..\GLSenseSharedVersion.cs"> entry in each .csproj. Because
// it's a linked file, all 5 assemblies compile this exact same source - there is
// only ONE place to edit when bumping the version.
//
// To release a new version: change the two version strings below, rebuild the
// solution. Every project's compiled AssemblyVersion/AssemblyFileVersion updates
// together, and both post_build.cmd scripts read the new version back out of the
// freshly-built DLL (via PowerShell reflection) instead of using a hardcoded
// batch-file literal - so nothing else needs to change: not the post-build
// scripts, not manifest.json's seeded content, not the deployment folder naming.
//
// Do NOT add [assembly: AssemblyVersion(...)]/[assembly: AssemblyFileVersion(...)]
// back into any individual project's own Properties\AssemblyInfo.cs - that would
// cause a duplicate-attribute compile error against this shared file.
using System.Reflection;

[assembly: AssemblyVersion("11.1.0.0")]
[assembly: AssemblyFileVersion("11.1.0.0")]
