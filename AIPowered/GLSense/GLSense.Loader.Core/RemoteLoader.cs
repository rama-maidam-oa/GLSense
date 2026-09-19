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
        private static string _sharedDependenciesPath;
        private static ILogger _resolverLogger;
        private static readonly object _lock = new object();
        private static bool _resolverRegistered;

        public IGLSenseAddin Create(string folder, string sharedDependenciesPath, IGLSenseContext context)
        {
            lock (_lock)
            {
                _resolverPath = folder;
                _sharedDependenciesPath = sharedDependenciesPath;
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
                return new RemoteAddinProxy(instance);
            }
            catch (Exception ex)
            {
                logger?.LogError(BuildFailureMessage(
                    $"RemoteLoader.Create: failed to create add-in instance from '{dllPath}'.",
                    ex));
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

            string sharedPath;
            lock (_lock)
            {
                sharedPath = _sharedDependenciesPath;
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

            var candidatePaths = new[]
            {
                Path.Combine(path, requestedName + ".dll"),
                Path.Combine(sharedPath ?? string.Empty, requestedName + ".dll")
            };

            foreach (var dllPath in candidatePaths)
            {
                if (!File.Exists(dllPath))
                    continue;

                try
                {
                    return Assembly.LoadFrom(dllPath);
                }
                catch (Exception ex)
                {
                    logger?.LogError(BuildFailureMessage(
                        $"AssemblyResolve: failed loading '{requestedName}' from '{dllPath}'.",
                        ex));
                    return null;
                }
            }

            var exePaths = new[]
            {
                Path.Combine(path, requestedName + ".exe"),
                Path.Combine(sharedPath ?? string.Empty, requestedName + ".exe")
            };

            foreach (var exePath in exePaths)
            {
                if (!File.Exists(exePath))
                    continue;

                try
                {
                    return Assembly.LoadFrom(exePath);
                }
                catch (Exception ex)
                {
                    logger?.LogError(BuildFailureMessage(
                        $"AssemblyResolve: failed loading '{requestedName}' from '{exePath}'.",
                        ex));
                    return null;
                }
            }

            logger?.LogWarn($"AssemblyResolve: could not resolve '{args.Name}' from release '{path}' or shared dependencies '{sharedPath}'");
            return null;
        }

        private static string BuildFailureMessage(string message, Exception exception)
        {
            if (exception == null)
                return message;

            return $"{message} {exception.GetType().FullName}: {exception.Message}{Environment.NewLine}{exception.StackTrace}";
        }

        public sealed class RemoteAddinProxy : MarshalByRefObject, IGLSenseAddin
        {
            private readonly IGLSenseAddin _inner;

            public RemoteAddinProxy(IGLSenseAddin inner)
            {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            }

            public void Initialize(IGLSenseContext context)
            {
                _inner.Initialize(context);
            }

            public void OnRibbonAction(string action, object parameter)
            {
                _inner.OnRibbonAction(action, parameter);
            }

            public bool OnExcelEvent(string eventName, object[] args)
            {
                return _inner.OnExcelEvent(eventName, args);
            }

            public object ExecuteUdf(string functionName, object[] args)
            {
                return _inner.ExecuteUdf(functionName, args);
            }

            public void Shutdown()
            {
                _inner.Shutdown();
            }

            public IntPtr CreateConfiguratorPaneContent()
            {
                return _inner.CreateConfiguratorPaneContent();
            }

            public void RelaunchConfiguratorPane(bool showBusyOverlay = true)
            {
                _inner.RelaunchConfiguratorPane(showBusyOverlay);
            }

            public void ResetConfiguratorPaneReference()
            {
                _inner.ResetConfiguratorPaneReference();
            }

            public bool HasSavedConfigurationSelected()
            {
                return _inner.HasSavedConfigurationSelected();
            }

            public void CloseConfiguratorPaneContent()
            {
                _inner.CloseConfiguratorPaneContent();
            }

            public LoginInfo GetLoginInfo()
            {
                return _inner.GetLoginInfo();
            }

            public string[] GetOpenWindowTitles()
            {
                return _inner.GetOpenWindowTitles();
            }

            public override object InitializeLifetimeService()
            {
                return null;
            }
        }

        public override object InitializeLifetimeService()
        {
            return null;
        }
    }
}
