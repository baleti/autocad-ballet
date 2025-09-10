using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

// Alias to avoid Application naming ambiguity
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutocadBallet.InvokeAddinCommand))]

namespace AutocadBallet
{
    public class InvokeAddinCommand
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

        // Static dictionary to track assemblies loaded in this AppDomain
        private static Dictionary<string, WeakReference> loadedAssemblyCache = new Dictionary<string, WeakReference>();

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

        [CommandMethod("invoke-addin-command", CommandFlags.Session)]
        public void InvokeAddin()
        {
            var ed = AcAp.DocumentManager.MdiActiveDocument?.Editor;
            if (ed == null) return;

            // Resolve the path
            string dllPath = Environment.ExpandEnvironmentVariables(TargetDllPath);

            if (!File.Exists(dllPath))
            {
                ed.WriteMessage($"\nTarget DLL not found: {dllPath}");
                ed.WriteMessage($"\nPlease ensure autocad-ballet.dll exists in the expected location.");
                return;
            }

            try
            {
                // Get file info for cache management
                var fileInfo = new FileInfo(dllPath);
                string cacheKey = $"{dllPath}|{fileInfo.LastWriteTimeUtc.Ticks}|{fileInfo.Length}";

                Assembly assembly = null;

                // Check if we have a cached assembly with the same timestamp and size
                if (loadedAssemblyCache.ContainsKey(cacheKey))
                {
                    var weakRef = loadedAssemblyCache[cacheKey];
                    if (weakRef.IsAlive)
                    {
                        assembly = weakRef.Target as Assembly;
                        ed.WriteMessage("\nUsing cached assembly (no file changes detected).");
                    }
                }

                // If not cached or cache is invalid, load the assembly
                if (assembly == null)
                {
                    // Clear old cache entries for this file
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
                        assembly = LoadAssemblyWithoutRegistration(dllPath);

                        // Cache the assembly with a weak reference
                        loadedAssemblyCache[cacheKey] = new WeakReference(assembly);
                        ed.WriteMessage("\nLoaded fresh assembly from disk.");
                    }
                    finally
                    {
                        AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
                    }
                }

                // Extract all commands
                var commands = ExtractCommandsFromAssembly(assembly);

                if (commands.Count == 0)
                {
                    ed.WriteMessage("\nNo [CommandMethod] commands found in the assembly.");
                    return;
                }

                // Convert to format expected by DataGrid
                var entries = commands.Select(cmd => new Dictionary<string, object>
                {
                    ["Command"] = cmd.CommandName,
                    ["Type"] = cmd.TypeName,
                    ["Method"] = cmd.MethodName,
                    ["Full Name"] = cmd.FullTypeName
                }).ToList();

                var propertyNames = new List<string> { "Command" };

                // Show the DataGrid selector
                var selected = CustomGUIs.DataGrid(entries, propertyNames, false);

                if (selected == null || selected.Count == 0)
                {
                    ed.WriteMessage("\nCanceled.");
                    return;
                }

                // Get the selected command info
                var selectedEntry = selected.First();
                var selectedCmd = commands.First(c =>
                    c.CommandName == selectedEntry["Command"].ToString() &&
                    c.FullTypeName == selectedEntry["Full Name"].ToString());

                ed.WriteMessage($"\nInvoking command: {selectedCmd.CommandName}");

                // Invoke the chosen method
                InvokeCommandMethod(selectedCmd, assembly, ed);

                ed.WriteMessage($"\nCommand completed successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[InvokeAddin] Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    ed.WriteMessage($"\n  Inner: {ex.InnerException.Message}");
                }
                ed.WriteMessage($"\n  Stack: {ex.StackTrace}");
            }
            finally
            {
                // Clear session assemblies
                sessionAssemblies.Clear();
            }
        }

        private Assembly LoadAssemblyWithoutRegistration(string assemblyPath)
        {
            // Load the assembly in a way that doesn't trigger AutoCAD's automatic command registration
            // We do this by loading into a separate context first for inspection

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
                            // Use LoadFile or Load without triggering AutoCAD registration
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
            // The key is to NOT let AutoCAD's extension loader process it
            Assembly assembly = null;

            // Try to load in a way that bypasses AutoCAD's automatic processing
            try
            {
                // Method 1: Load from bytes (this usually works but sometimes AutoCAD still hooks it)
                assembly = Assembly.Load(assemblyBytes);
            }
            catch
            {
                // Method 2: If that fails, try reflection-only context first to validate
                try
                {
                    var refAssembly = Assembly.ReflectionOnlyLoadFrom(assemblyPath);
                    // If reflection-only succeeds, we know the assembly is valid
                    // Now load it for real
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
                var type = assembly.GetType(info.FullTypeName, throwOnError: true);
                if (type == null)
                    throw new TypeLoadException($"Could not find type: {info.FullTypeName}");

                var method = type.GetMethod(info.MethodName,
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);

                if (method == null)
                    throw new MissingMethodException(info.FullTypeName, info.MethodName);

                object instance = null;
                if (!method.IsStatic)
                {
                    instance = Activator.CreateInstance(type);
                }

                // Command methods in AutoCAD typically take no parameters
                method.Invoke(instance, parameters: null);
            }
            catch (TargetInvocationException tie)
            {
                if (ed != null)
                {
                    var message = tie.InnerException != null ? tie.InnerException.Message : tie.Message;
                    ed.WriteMessage($"\n[InvokeAddin] Command threw: {message}");
                }
                throw;
            }
            catch (System.Exception ex)
            {
                if (ed != null)
                    ed.WriteMessage($"\n[InvokeAddin] Failed to invoke: {ex.Message}");
                throw;
            }
        }
    }
}
