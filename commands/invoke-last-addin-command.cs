using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutocadBallet.InvokeLastAddinCommand))]

namespace AutocadBallet
{
    public class InvokeLastAddinCommand
    {
#if ACAD2017
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2017\autocad-ballet.dll";
#elif ACAD2018
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2018\autocad-ballet.dll";
#elif ACAD2019
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2019\autocad-ballet.dll";
#elif ACAD2020
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2020\autocad-ballet.dll";
#elif ACAD2021
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2021\autocad-ballet.dll";
#elif ACAD2022
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2022\autocad-ballet.dll";
#elif ACAD2023
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2023\autocad-ballet.dll";
#elif ACAD2024
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2024\autocad-ballet.dll";
#elif ACAD2025
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2025\autocad-ballet.dll";
#elif ACAD2026
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2026\autocad-ballet.dll";
#else
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands\bin\2026\autocad-ballet.dll"; // Default
#endif

        private const string FolderName = "autocad-ballet";
        private const string RuntimeFolderName = "runtime";
        private const string ConfigFileName = "InvokeAddinCommand-last-dll-path";
        private const string LastCommandFileName = "InvokeAddinCommand-history";
        private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName, RuntimeFolderName);
        private static readonly string ConfigFilePath = Path.Combine(ConfigFolderPath, ConfigFileName);
        private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

        // Static dictionary to track assemblies loaded in this AppDomain
        private static Dictionary<string, WeakReference> loadedAssemblyCache = new Dictionary<string, WeakReference>();

        // Stronger cache for recently used assemblies to prevent GC
        private static Dictionary<string, Assembly> strongAssemblyCache = new Dictionary<string, Assembly>();

        // Instance dictionary for dependency resolution within a single invocation
        private Dictionary<string, Assembly> sessionAssemblies = new Dictionary<string, Assembly>();

        // Command info class
        private class CommandInfo
        {
            public string CommandName { get; }
            public string TypeName { get; }
            public string FullTypeName { get; }
            public string MethodName { get; }

            public CommandInfo(string commandName, string typeName, string fullTypeName, string methodName)
            {
                CommandName = commandName;
                TypeName = typeName;
                FullTypeName = fullTypeName;
                MethodName = methodName;
            }
        }

        [CommandMethod("invoke-last-addin-command", CommandFlags.Session | CommandFlags.UsePickSet)]
        public void Execute()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument?.Editor;
            if (ed == null) return;

            try
            {
                if (!File.Exists(LastCommandFilePath))
                {
                    ed.WriteMessage("\nNo previous command found. Run a command using InvokeAddinCommand first.");
                    return;
                }

                string lastCommandName = GetLastCommand();

                if (string.IsNullOrEmpty(lastCommandName))
                {
                    ed.WriteMessage("\nNo command history found.");
                    return;
                }

                ed.WriteMessage($"\nInvoking last command: {lastCommandName}");

                // Resolve the path
                string dllPath = Environment.ExpandEnvironmentVariables(TargetDllPath);

                if (!File.Exists(dllPath))
                {
                    ed.WriteMessage($"\nTarget DLL not found: {dllPath}");
                    ed.WriteMessage($"\nPlease ensure autocad-ballet.dll exists in the expected location.");
                    return;
                }

                // Load assembly using same mechanism as invoke-addin-command
                Assembly assembly = LoadAssemblyWithCaching(dllPath, ed);
                if (assembly == null) return;

                // Find the command in the assembly
                var commands = ExtractCommandsFromAssembly(assembly);
                var targetCommand = commands.FirstOrDefault(c =>
                    c.CommandName.Equals(lastCommandName, StringComparison.OrdinalIgnoreCase));

                if (targetCommand == null)
                {
                    ed.WriteMessage($"\nCommand '{lastCommandName}' not found in the assembly.");
                    return;
                }

                // Invoke the command method directly
                InvokeCommandMethod(targetCommand, assembly, ed);

                // Use safe messaging that handles document closure scenarios
                SafeWriteMessage(ed, $"\nCommand completed successfully.");
            }
            catch (System.Exception ex)
            {
                SafeWriteMessage(ed, $"\n[InvokeLastAddin] Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    SafeWriteMessage(ed, $"\n  Inner: {ex.InnerException.Message}");
                }
                SafeWriteMessage(ed, $"\n  Stack: {ex.StackTrace}");
            }
            finally
            {
                // Clear session assemblies
                sessionAssemblies.Clear();
            }
        }

        private string GetLastCommand()
        {
            try
            {
                string[] lines = File.ReadAllLines(LastCommandFilePath);

                for (int i = lines.Length - 1; i >= 0; i--)
                {
                    if (!string.IsNullOrWhiteSpace(lines[i]))
                    {
                        return lines[i].Trim();
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private Assembly LoadAssemblyWithCaching(string dllPath, Editor ed)
        {
            try
            {
                // Get file info for cache management
                var fileInfo = new FileInfo(dllPath);
                string cacheKey = $"{dllPath}|{fileInfo.LastWriteTimeUtc.Ticks}|{fileInfo.Length}";

                Assembly assembly = null;

                // Check strong cache first
                if (strongAssemblyCache.ContainsKey(cacheKey))
                {
                    assembly = strongAssemblyCache[cacheKey];
                    ed.WriteMessage("\nUsing strongly cached assembly (no file changes detected).");
                }
                // Check weak cache
                else if (loadedAssemblyCache.ContainsKey(cacheKey))
                {
                    var weakRef = loadedAssemblyCache[cacheKey];
                    if (weakRef.IsAlive)
                    {
                        assembly = weakRef.Target as Assembly;
                        // Move to strong cache for future use
                        strongAssemblyCache[cacheKey] = assembly;
                        ed.WriteMessage("\nUsing cached assembly (no file changes detected).");
                    }
                }

                // If not cached or cache is invalid, load the assembly
                if (assembly == null)
                {
                    // Clear old cache entries for this file
                    var strongKeysToRemove = strongAssemblyCache.Keys
                        .Where(k => k.StartsWith(dllPath + "|"))
                        .ToList();
                    foreach (var key in strongKeysToRemove)
                    {
                        strongAssemblyCache.Remove(key);
                    }

                    // Limit strong cache size to prevent memory leaks
                    if (strongAssemblyCache.Count > 10)
                    {
                        var oldestKeys = strongAssemblyCache.Keys.Take(strongAssemblyCache.Count - 5).ToList();
                        foreach (var key in oldestKeys)
                        {
                            strongAssemblyCache.Remove(key);
                        }
                    }

                    var keysToRemove = loadedAssemblyCache.Keys
                        .Where(k => k.StartsWith(dllPath + "|"))
                        .ToList();
                    foreach (var key in keysToRemove)
                    {
                        loadedAssemblyCache.Remove(key);
                    }

                    // Register the assembly resolve event handler for dependencies
                    AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                    try
                    {
                        // Load without triggering AutoCAD registration
                        try
                        {
                            assembly = LoadAssemblyWithoutRegistration(dllPath);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acadEx) when (acadEx.ErrorStatus.ToString().Contains("eDuplicateKey"))
                        {
                            // This happens when commands are already registered - not a fatal error
                            ed.WriteMessage("\nWarning: Commands already registered, reusing existing assembly.");
                            // Try to find the assembly in already loaded assemblies
                            assembly = AppDomain.CurrentDomain.GetAssemblies()
                                .FirstOrDefault(a => a.Location.Equals(dllPath, StringComparison.OrdinalIgnoreCase));
                            if (assembly == null) throw;
                        }

                        // Cache the assembly with a weak reference
                        loadedAssemblyCache[cacheKey] = new WeakReference(assembly);
                        strongAssemblyCache[cacheKey] = assembly;
                        ed.WriteMessage("\nLoaded fresh assembly from disk.");
                    }
                    finally
                    {
                        AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
                    }
                }

                return assembly;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[LoadAssembly] Error: {ex.Message}");
                return null;
            }
        }

        private Assembly LoadAssemblyWithoutRegistration(string assemblyPath)
        {
            byte[] assemblyBytes = File.ReadAllBytes(assemblyPath);

            // Store in session for dependency resolution
            string assemblyName = Path.GetFileNameWithoutExtension(assemblyPath);

            // First, load all potential dependencies in the same directory
            string directory = Path.GetDirectoryName(assemblyPath);
            foreach (string dllFile in Directory.GetFiles(directory, "*.dll"))
            {
                if (dllFile != assemblyPath)
                {
                    string dllName = Path.GetFileNameWithoutExtension(dllFile);
                    if (!sessionAssemblies.ContainsKey(dllName))
                    {
                        try
                        {
                            byte[] dllBytes = File.ReadAllBytes(dllFile);
                            Assembly dllAssembly = Assembly.Load(dllBytes);
                            sessionAssemblies[dllName] = dllAssembly;
                        }
                        catch (BadImageFormatException)
                        {
                            // Skip native DLLs or incompatible assemblies
                            continue;
                        }
                        catch (System.Exception)
                        {
                            // Skip assemblies that fail to load
                            continue;
                        }
                    }
                }
            }

            // Now load the main assembly
            Assembly assembly = null;

            // Try to load in a way that bypasses AutoCAD's automatic processing
            try
            {
                assembly = Assembly.Load(assemblyBytes);
            }
            catch
            {
                try
                {
                    assembly = Assembly.LoadFrom(assemblyPath);
                }
                catch
                {
                    throw new InvalidOperationException($"Failed to load assembly: {assemblyPath}");
                }
            }

            sessionAssemblies[assemblyName] = assembly;
            return assembly;
        }

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // Parse the assembly name
            AssemblyName assemblyName = new AssemblyName(args.Name);
            string shortName = assemblyName.Name;

            // Check session assemblies first
            if (sessionAssemblies.ContainsKey(shortName))
            {
                return sessionAssemblies[shortName];
            }

            // Check cached assemblies
            foreach (var kvp in loadedAssemblyCache)
            {
                if (kvp.Value.IsAlive)
                {
                    var asm = kvp.Value.Target as Assembly;
                    if (asm != null && asm.GetName().Name == shortName)
                    {
                        return asm;
                    }
                }
            }

            // Try to load from the target directory
            string dllPath = Environment.ExpandEnvironmentVariables(TargetDllPath);
            string directory = Path.GetDirectoryName(dllPath);
            string assemblyPath = Path.Combine(directory, shortName + ".dll");

            if (File.Exists(assemblyPath))
            {
                try
                {
                    byte[] bytes = File.ReadAllBytes(assemblyPath);
                    Assembly assembly = Assembly.Load(bytes);
                    sessionAssemblies[shortName] = assembly;
                    return assembly;
                }
                catch (System.Exception)
                {
                    return null;
                }
            }

            return null;
        }

        private List<CommandInfo> ExtractCommandsFromAssembly(Assembly assembly)
        {
            var list = new List<CommandInfo>();

            try
            {
                foreach (var type in assembly.GetTypes())
                {
                    var methods = type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);
                    foreach (var method in methods)
                    {
                        var attributes = method.GetCustomAttributes(false);
                        foreach (var attr in attributes)
                        {
                            // Check for CommandMethodAttribute
                            if (attr.GetType().FullName == "Autodesk.AutoCAD.Runtime.CommandMethodAttribute")
                            {
                                string commandName = TryGetCommandName(attr) ?? method.Name;
                                list.Add(new CommandInfo(
                                    commandName,
                                    type.Name,
                                    type.FullName ?? type.Name,
                                    method.Name));
                            }
                        }
                    }
                }
            }
            catch (ReflectionTypeLoadException ex)
            {
                // Handle partial loading
                foreach (var type in ex.Types.Where(t => t != null))
                {
                    try
                    {
                        var methods = type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);
                        foreach (var method in methods)
                        {
                            var attributes = method.GetCustomAttributes(false);
                            foreach (var attr in attributes)
                            {
                                if (attr.GetType().FullName == "Autodesk.AutoCAD.Runtime.CommandMethodAttribute")
                                {
                                    string commandName = TryGetCommandName(attr) ?? method.Name;
                                    list.Add(new CommandInfo(
                                        commandName,
                                        type.Name,
                                        type.FullName ?? type.Name,
                                        method.Name));
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Skip types that can't be processed
                        continue;
                    }
                }
            }

            return list.OrderBy(c => c.CommandName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private string TryGetCommandName(object cmdAttr)
        {
            var type = cmdAttr.GetType();

            // Try GlobalName property (the actual command name in AutoCAD)
            var prop = type.GetProperty("GlobalName", BindingFlags.Public | BindingFlags.Instance);
            if (prop != null)
            {
                var val = prop.GetValue(cmdAttr) as string;
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            // Fallback to any string property
            foreach (var p in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (p.PropertyType == typeof(string))
                {
                    var v = p.GetValue(cmdAttr) as string;
                    if (!string.IsNullOrWhiteSpace(v))
                        return v;
                }
            }

            return null;
        }

        private void InvokeCommandMethod(CommandInfo info, Assembly assembly, Editor ed)
        {
            try
            {
                ed.WriteMessage($"\nExecuting command: {info.CommandName}");

                // Find the type and method
                Type targetType = assembly.GetType(info.FullTypeName);
                if (targetType == null)
                {
                    throw new InvalidOperationException($"Type '{info.FullTypeName}' not found in assembly");
                }

                MethodInfo targetMethod = targetType.GetMethod(info.MethodName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);
                if (targetMethod == null)
                {
                    throw new InvalidOperationException($"Method '{info.MethodName}' not found in type '{info.FullTypeName}'");
                }

                // Create instance if method is not static
                object instance = null;
                if (!targetMethod.IsStatic)
                {
                    instance = Activator.CreateInstance(targetType);
                }

                // Invoke the method directly to preserve pickfirst selection
                targetMethod.Invoke(instance, null);
            }
            catch (System.Exception ex)
            {
                if (ed != null)
                    ed.WriteMessage($"\n[InvokeLastAddin] Failed to invoke command '{info.CommandName}': {ex.Message}");
                throw;
            }
        }

        private static void SafeWriteMessage(Editor ed, string message)
        {
            try
            {
                // Try to use the original editor first
                ed.WriteMessage(message);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
                // If original editor fails (document closed), try application-level messaging
                try
                {
                    var activeDoc = AcApp.DocumentManager.MdiActiveDocument;
                    activeDoc?.Editor?.WriteMessage(message);
                }
                catch
                {
                    // If all fails, silently ignore - command executed but can't report status
                    System.Diagnostics.Debug.WriteLine($"[SafeWriteMessage] {message}");
                }
            }
            catch (System.Exception)
            {
                // Handle any other exception type
                try
                {
                    var activeDoc = AcApp.DocumentManager.MdiActiveDocument;
                    activeDoc?.Editor?.WriteMessage(message);
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine($"[SafeWriteMessage] {message}");
                }
            }
        }
    }
}