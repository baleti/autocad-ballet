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

        // Dictionary to store loaded assemblies (instance field, refreshes each command execution)
        private Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string, Assembly>();

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

        [CommandMethod("INVOKE-ADDIN-COMMAND", CommandFlags.Session)]
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
                // Register the assembly resolve event handler
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                // Load the main assembly
                Assembly assembly = LoadAssembly(dllPath);

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

                var propertyNames = new List<string> { "Command", "Type", "Method", "Full Name" };

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
            }
            finally
            {
                // Unregister the assembly resolve event handler
                AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
            }
        }

        private Assembly LoadAssembly(string assemblyPath)
        {
            string assemblyName = Path.GetFileNameWithoutExtension(assemblyPath);

            // Check if already loaded in this invocation
            if (loadedAssemblies.ContainsKey(assemblyName))
            {
                return loadedAssemblies[assemblyName];
            }

            // Load from bytes to avoid file lock
            byte[] assemblyBytes = File.ReadAllBytes(assemblyPath);
            Assembly assembly = Assembly.Load(assemblyBytes);
            loadedAssemblies[assemblyName] = assembly;

            // Load all DLLs in the same directory as potential dependencies
            string directory = Path.GetDirectoryName(assemblyPath);
            foreach (string dllFile in Directory.GetFiles(directory, "*.dll"))
            {
                if (dllFile != assemblyPath)
                {
                    string dllName = Path.GetFileNameWithoutExtension(dllFile);
                    if (!loadedAssemblies.ContainsKey(dllName))
                    {
                        try
                        {
                            byte[] dllBytes = File.ReadAllBytes(dllFile);
                            Assembly dllAssembly = Assembly.Load(dllBytes);
                            loadedAssemblies[dllName] = dllAssembly;
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

            return assembly;
        }

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // Parse the assembly name
            AssemblyName assemblyName = new AssemblyName(args.Name);
            string shortName = assemblyName.Name;

            // Check if we've already loaded this assembly
            if (loadedAssemblies.ContainsKey(shortName))
            {
                return loadedAssemblies[shortName];
            }

            // Look for the assembly in the same directory as the main DLL
            string dllPath = Environment.ExpandEnvironmentVariables(TargetDllPath);
            string directory = Path.GetDirectoryName(dllPath);
            string assemblyPath = Path.Combine(directory, shortName + ".dll");

            if (File.Exists(assemblyPath))
            {
                try
                {
                    Assembly assembly = LoadAssembly(assemblyPath);
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

            foreach (var type in assembly.GetTypes())
            {
                var methods = type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);
                foreach (var method in methods)
                {
                    foreach (var attr in method.GetCustomAttributes(false))
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
