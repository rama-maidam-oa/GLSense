using System.Reflection;
using System.Runtime.InteropServices;

// This assembly is never shipped - GLSense.Build exists purely to sequence
// the real build (see GLSense.Build.csproj's own comment). No functional
// code, no version, nothing here matters at runtime.
[assembly: AssemblyTitle("GLSense.Build")]
[assembly: AssemblyDescription("Build orchestration helper - not shipped")]
[assembly: ComVisible(false)]
