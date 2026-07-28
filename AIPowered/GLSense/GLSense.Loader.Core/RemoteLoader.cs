//GLSense.Loader.Core/RemoteLoader.cs
using GLSense.Contracts;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace GLSense.Loader.Core
{
    public class RemoteLoader : MarshalByRefObject
    {
        private static string _resolverPath;
        private static ILogger _resolverLogger;
        private static readonly object _lock = new object();
        private static bool _resolverRegistered;

        public IGLSenseAddin Create(string folder, IGLSenseContext context)
        {
            lock (_lock)
            {
                _resolverPath = folder;
                _resolverLogger = context.Logger;

                if (!_resolverRegistered)
                {
                    AppDomain.CurrentDomain.AssemblyResolve += ResolveAssemblyInDomain;
                    _resolverRegistered = true;
                }
            }

            var logger = context.Logger;
            var dllPath = Path.Combine(folder, "GLSense.Addin.Core.dll");

            try
            {
                logger?.LogDebug($"RemoteLoader.Create: loading add-in assembly from '{dllPath}'.");

                if (!File.Exists(dllPath))
                    throw new FileNotFoundException($"Could not find add-in assembly: {dllPath}");

                var asm = Assembly.LoadFrom(dllPath);
                logger?.LogDebug($"RemoteLoader.Create: loaded assembly '{asm.FullName}'.");

                var type = asm.GetTypes()
                    .FirstOrDefault(t => typeof(IGLSenseAddin).IsAssignableFrom(t) && !t.IsAbstract && !t.IsInterface);

                if (type == null)
                    throw new InvalidOperationException("Add-in entry type not found in GLSense.Addin.Core.dll");

                var instance = (IGLSenseAddin)Activator.CreateInstance(type);
                logger?.LogDebug($"RemoteLoader.Create: instantiated add-in entry type '{type.FullName}'.");
                return instance;
            }
            catch (Exception ex)
            {
                logger?.LogError($"RemoteLoader.Create: failed to create add-in instance from '{dllPath}'.", ex);
                throw;
            }
        }

        private static Assembly ResolveAssemblyInDomain(object sender, ResolveEventArgs args)
        {
            string path;
            ILogger logger;

            lock (_lock)
            {
                path = _resolverPath;
                logger = _resolverLogger;
            }

            if (string.IsNullOrWhiteSpace(path))
                return null;

            AssemblyName requestedAssembly;
            try
            {
                requestedAssembly = new AssemblyName(args.Name);
            }
            catch (Exception ex)
            {
                logger?.LogError($"AssemblyResolve: invalid assembly name '{args.Name}'.", ex);
                return null;
            }

            var requestedName = requestedAssembly.Name;

            if (string.IsNullOrWhiteSpace(requestedName))
                return null;

            if (requestedName.EndsWith(".resources", StringComparison.OrdinalIgnoreCase) ||
                requestedName.EndsWith(".g.resources", StringComparison.OrdinalIgnoreCase))
            {
                logger?.LogDebug($"AssemblyResolve: ignoring resource assembly request: {args.Name}");
                return null;
            }

            if (IsFrameworkAssembly(requestedName))
            {
                logger?.LogDebug($"AssemblyResolve: ignoring framework assembly request: {args.Name}");
                return null;
            }

            var alreadyLoaded = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a =>
                {
                    try
                    {
                        return string.Equals(a.GetName().Name, requestedName, StringComparison.OrdinalIgnoreCase);
                    }
                    catch (Exception ex)
                    {
                        logger?.LogDebug($"AssemblyResolve: could not read name of a loaded assembly while checking for '{requestedName}'. {ex.Message}");
                        return false;
                    }
                });

            if (alreadyLoaded != null)
            {
                logger?.LogDebug($"AssemblyResolve: already loaded: {alreadyLoaded.FullName}");
                return alreadyLoaded;
            }

            var dllPath = Path.Combine(path, requestedName + ".dll");
            if (File.Exists(dllPath))
            {
                try
                {
                    return Assembly.LoadFrom(dllPath);
                }
                catch (Exception ex)
                {
                    logger?.LogError($"AssemblyResolve: failed loading '{requestedName}' from '{dllPath}'.", ex);
                    return null;
                }
            }

            var exePath = Path.Combine(path, requestedName + ".exe");
            if (File.Exists(exePath))
            {
                try
                {
                    return Assembly.LoadFrom(exePath);
                }
                catch (Exception ex)
                {
                    logger?.LogError($"AssemblyResolve: failed loading '{requestedName}' from '{exePath}'.", ex);
                    return null;
                }
            }

            logger?.LogWarn($"AssemblyResolve: could not resolve '{args.Name}' from '{path}'");
            return null;
        }

        private static bool IsFrameworkAssembly(string assemblyName)
        {
            return assemblyName.StartsWith("System", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.StartsWith("Microsoft", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.Equals("mscorlib", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.Equals("netstandard", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.Equals("WindowsBase", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.Equals("PresentationCore", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.Equals("PresentationFramework", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.Equals("Accessibility", StringComparison.OrdinalIgnoreCase) ||
                   assemblyName.StartsWith("UIAutomation", StringComparison.OrdinalIgnoreCase);
        }

        public override object InitializeLifetimeService()
        {
            return null;
        }
    }
}
