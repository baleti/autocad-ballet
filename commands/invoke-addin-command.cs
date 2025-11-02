using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

using Mono.Cecil;

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

        private const string FolderName = "autocad-ballet";
        private const string RuntimeFolderName = "runtime";
        private const string ConfigFileName = "InvokeAddinCommand-last-dll-path";
        private const string LastCommandFileName = "InvokeAddinCommand-history";
        private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName, RuntimeFolderName);
        private static readonly string ConfigFilePath = Path.Combine(ConfigFolderPath, ConfigFileName);
        private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

        // Cached command info - tracks individual commands that have been loaded
        private class CachedCommandInfo
        {
            public Assembly Assembly { get; set; }
            public string OriginalName { get; set; }
            public string ModifiedName { get; set; }
            public string TypeFullName { get; set; }
            public string MethodName { get; set; }
        }

        // Static dictionary to track loaded commands per DLL version
        // Key format: "dllPath|timestamp|filesize|commandName"
        private static Dictionary<string, CachedCommandInfo> loadedCommandCache = new Dictionary<string, CachedCommandInfo>();

        // Instance dictionary for dependency resolution within a single invocation
        private Dictionary<string, Assembly> sessionAssemblies = new Dictionary<string, Assembly>();

        // Command info class
        private class CommandInfo
        {
            public string OriginalCommandName { get; }
            public string ModifiedCommandName { get; }
            public string TypeName { get; }
            public string FullTypeName { get; }
            public string MethodName { get; }

            public CommandInfo(string originalCommandName, string modifiedCommandName, string typeName, string fullTypeName, string methodName)
            {
                OriginalCommandName = originalCommandName;
                ModifiedCommandName = modifiedCommandName;
                TypeName = typeName;
                FullTypeName = fullTypeName;
                MethodName = methodName;
            }
        }

        [CommandMethod("invoke-addin-command", CommandFlags.Session | CommandFlags.UsePickSet)]
        public void InvokeAddin()
        {
            var ed = AcAp.DocumentManager.MdiActiveDocument?.Editor;
            if (ed == null) return;

            // Capture pickfirst selection at the start to preserve it
            // AutoCAD will clear it when the modal DataGrid shows, but we'll restore it via LISP
            Autodesk.AutoCAD.DatabaseServices.ObjectId[] pickfirstSelection = null;
            var selResult = ed.SelectImplied();
            if (selResult.Status == PromptStatus.OK)
            {
                pickfirstSelection = selResult.Value.GetObjectIds();
            }

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
                string baseCacheKey = $"{dllPath}|{fileInfo.LastWriteTimeUtc.Ticks}|{fileInfo.Length}";

                // Inspect DLL with Mono.Cecil to extract command names (don't load assembly yet)
                ed.WriteMessage("\nInspecting assembly...");
                var commandNames = InspectCommandNames(dllPath);

                if (commandNames.Count == 0)
                {
                    ed.WriteMessage("\nNo [CommandMethod] commands found in the assembly.");
                    return;
                }

                // Convert to format expected by DataGrid - show ORIGINAL names to user
                var entries = commandNames.Select(cmdName => new Dictionary<string, object>
                {
                    ["Command"] = cmdName
                }).ToList();

                var propertyNames = new List<string> { "Command" };

                // Show the DataGrid selector
                var selected = CustomGUIs.DataGrid(entries, propertyNames, false);

                if (selected == null || selected.Count == 0)
                {
                    ed.WriteMessage("\nCanceled.");
                    return;
                }

                // Get the selected command name
                var selectedEntry = selected.First();
                string selectedCommandName = selectedEntry["Command"].ToString();

                // Check if this specific command is already cached
                string commandCacheKey = $"{baseCacheKey}|{selectedCommandName}";
                CachedCommandInfo cachedCommand = null;

                if (loadedCommandCache.ContainsKey(commandCacheKey))
                {
                    cachedCommand = loadedCommandCache[commandCacheKey];
                    ed.WriteMessage($"\nUsing cached command: {selectedCommandName}");
                }
                else
                {
                    // Clear old cache entries for this DLL version (different timestamp/size)
                    var keysToRemove = loadedCommandCache.Keys
                        .Where(k => k.StartsWith(dllPath + "|") && !k.StartsWith(baseCacheKey))
                        .ToList();
                    foreach (var key in keysToRemove)
                    {
                        loadedCommandCache.Remove(key);
                    }

                    // Limit cache size to prevent memory leaks
                    if (loadedCommandCache.Count > 50)
                    {
                        var oldestKeys = loadedCommandCache.Keys.Take(loadedCommandCache.Count - 25).ToList();
                        foreach (var key in oldestKeys)
                        {
                            loadedCommandCache.Remove(key);
                        }
                    }

                    // Register the assembly resolve event handler for dependencies
                    AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                    try
                    {
                        // Rewrite ONLY the selected command with unique name and load
                        ed.WriteMessage($"\nRewriting and loading command: {selectedCommandName}...");
                        cachedCommand = RewriteAndLoadSingleCommand(dllPath, selectedCommandName, ed);

                        // Cache the command info
                        loadedCommandCache[commandCacheKey] = cachedCommand;
                        ed.WriteMessage($"\nLoaded command: {selectedCommandName}");
                    }
                    finally
                    {
                        AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
                    }
                }

                ed.WriteMessage($"\nInvoking command: {cachedCommand.OriginalName}");

                // Save the command and DLL path for invoke-last-addin-command
                SaveLastCommandInfo(cachedCommand, dllPath);

                // Invoke the chosen method (using MODIFIED name for AutoCAD)
                InvokeCommandMethod(cachedCommand, ed, pickfirstSelection);

                // Note: SendStringToExecute is asynchronous, so we can't report completion here
            }
            catch (System.Exception ex)
            {
                SafeWriteMessage(ed, $"\n[InvokeAddin] Error: {ex.Message}");
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

        private List<string> InspectCommandNames(string assemblyPath)
        {
            var commandNames = new List<string>();

            try
            {
                // Use Mono.Cecil to inspect the assembly WITHOUT loading it
                using (var ms = new MemoryStream(File.ReadAllBytes(assemblyPath)))
                {
                    var readerParams = new ReaderParameters { ReadSymbols = false };
                    var assemblyDef = AssemblyDefinition.ReadAssembly(ms, readerParams);

                    // Iterate through all types and methods to find CommandMethod attributes
                    foreach (var type in assemblyDef.MainModule.Types)
                    {
                        foreach (var method in type.Methods)
                        {
                            if (method.HasCustomAttributes)
                            {
                                var cmdAttr = method.CustomAttributes
                                    .FirstOrDefault(a => a.AttributeType.FullName == "Autodesk.AutoCAD.Runtime.CommandMethodAttribute");

                                if (cmdAttr != null && cmdAttr.ConstructorArguments.Count > 0)
                                {
                                    // Get the command name
                                    var nameArg = cmdAttr.ConstructorArguments[0];
                                    if (nameArg.Value is string commandName)
                                    {
                                        commandNames.Add(commandName);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new InvalidOperationException($"Failed to inspect assembly: {ex.Message}", ex);
            }

            return commandNames.OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private CachedCommandInfo RewriteAndLoadSingleCommand(string assemblyPath, string targetCommandName, Editor ed)
        {
            // Generate unique suffix for the target command to avoid registration conflicts
            string uniqueSuffix = "-" + DateTime.Now.Ticks.ToString();

            byte[] modifiedAssemblyBytes;
            string modifiedCommandName = null;
            string typeFullName = null;
            string methodName = null;

            // Use Mono.Cecil to REMOVE all other commands and rewrite only the target one
            using (var ms = new MemoryStream(File.ReadAllBytes(assemblyPath)))
            {
                var readerParams = new ReaderParameters { ReadSymbols = false };
                var assemblyDef = AssemblyDefinition.ReadAssembly(ms, readerParams);

                bool foundTargetCommand = false;

                // Iterate through all types and methods to find CommandMethod attributes
                foreach (var type in assemblyDef.MainModule.Types)
                {
                    foreach (var method in type.Methods)
                    {
                        if (method.HasCustomAttributes)
                        {
                            var cmdAttr = method.CustomAttributes
                                .FirstOrDefault(a => a.AttributeType.FullName == "Autodesk.AutoCAD.Runtime.CommandMethodAttribute");

                            if (cmdAttr != null && cmdAttr.ConstructorArguments.Count > 0)
                            {
                                // Get the original command name
                                var nameArg = cmdAttr.ConstructorArguments[0];
                                if (nameArg.Value is string originalName)
                                {
                                    if (originalName == targetCommandName)
                                    {
                                        // This is the target command - rewrite it with unique suffix
                                        modifiedCommandName = originalName + uniqueSuffix;

                                        cmdAttr.ConstructorArguments[0] = new CustomAttributeArgument(
                                            nameArg.Type,
                                            modifiedCommandName
                                        );

                                        // Store command info
                                        typeFullName = type.FullName;
                                        methodName = method.Name;
                                        foundTargetCommand = true;
                                    }
                                    else
                                    {
                                        // This is NOT the target command - REMOVE the attribute
                                        method.CustomAttributes.Remove(cmdAttr);
                                    }
                                }
                            }
                        }
                    }
                }

                if (!foundTargetCommand)
                {
                    throw new InvalidOperationException($"Command '{targetCommandName}' not found in assembly");
                }

                // Write modified assembly to memory
                using (var outputMs = new MemoryStream())
                {
                    assemblyDef.Write(outputMs);
                    modifiedAssemblyBytes = outputMs.ToArray();
                }
            }

            // Load dependencies first
            string directory = Path.GetDirectoryName(assemblyPath);
            string assemblyName = Path.GetFileNameWithoutExtension(assemblyPath);

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

            // Load the modified assembly
            Assembly assembly = Assembly.Load(modifiedAssemblyBytes);
            sessionAssemblies[assemblyName] = assembly;

            return new CachedCommandInfo
            {
                Assembly = assembly,
                OriginalName = targetCommandName,
                ModifiedName = modifiedCommandName,
                TypeFullName = typeFullName,
                MethodName = methodName
            };
        }

        private Assembly LoadAssemblyWithoutRegistration_OLD(string assemblyPath)
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
                // Method 2: If that fails, load directly from file
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

            // Check cached commands
            foreach (var kvp in loadedCommandCache)
            {
                var asm = kvp.Value.Assembly;
                if (asm != null && asm.GetName().Name == shortName)
                {
                    return asm;
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

        private List<CommandInfo> ExtractCommandsFromAssembly(Assembly assembly, Dictionary<string, string> commandMappings)
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
                                string modifiedCommandName = TryGetCommandName(attr) ?? method.Name;

                                // Find the original name from mappings (reverse lookup)
                                string originalCommandName = commandMappings
                                    .FirstOrDefault(kvp => kvp.Value == modifiedCommandName)
                                    .Key ?? modifiedCommandName;

                                list.Add(new CommandInfo(
                                    originalCommandName,
                                    modifiedCommandName,
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
                                    string modifiedCommandName = TryGetCommandName(attr) ?? method.Name;

                                    // Find the original name from mappings (reverse lookup)
                                    string originalCommandName = commandMappings
                                        .FirstOrDefault(kvp => kvp.Value == modifiedCommandName)
                                        .Key ?? modifiedCommandName;

                                    list.Add(new CommandInfo(
                                        originalCommandName,
                                        modifiedCommandName,
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

            return list.OrderBy(c => c.OriginalCommandName, StringComparer.OrdinalIgnoreCase).ToList();
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

        private void InvokeCommandMethod(CachedCommandInfo commandInfo, Editor ed, Autodesk.AutoCAD.DatabaseServices.ObjectId[] pickfirstSelection)
        {
            try
            {
                ed.WriteMessage($"\nExecuting command: {commandInfo.OriginalName}");

                var doc = AcAp.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    throw new InvalidOperationException("No active document");
                }

                string commandString;

                // Restore pickfirst selection using LISP before invoking the command
                // This ensures selection survives the async SendStringToExecute call
                // We use LISP because (sssetfirst) and (command) execute atomically in the same context
                if (pickfirstSelection != null && pickfirstSelection.Length > 0)
                {
                    var db = doc.Database;
                    var handles = new StringBuilder();
                    handles.Append("(progn (setq _tmp_ss (ssadd))");

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        foreach (var id in pickfirstSelection)
                        {
                            var ent = tr.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            var handle = ent.Handle.ToString();
                            handles.Append($" (ssadd (handent \"{handle}\") _tmp_ss)");
                        }
                        tr.Commit();
                    }

                    handles.Append($" (sssetfirst nil _tmp_ss) (command \"{commandInfo.ModifiedName}\") (princ)) ");
                    commandString = handles.ToString();
                }
                else
                {
                    commandString = commandInfo.ModifiedName + "\n";
                }

                // Use SendStringToExecute to invoke through AutoCAD's command pipeline
                doc.SendStringToExecute(commandString, true, false, false);
            }
            catch (System.Exception ex)
            {
                if (ed != null)
                    ed.WriteMessage($"\n[InvokeAddin] Failed to invoke command '{commandInfo.OriginalName}': {ex.Message}");
                throw;
            }
        }

        private void SaveLastCommandInfo(CachedCommandInfo commandInfo, string dllPath)
        {
            try
            {
                // Ensure the config directory exists
                if (!Directory.Exists(ConfigFolderPath))
                {
                    Directory.CreateDirectory(ConfigFolderPath);
                }

                // Save the DLL path
                File.WriteAllText(ConfigFilePath, dllPath);

                // Save the user-friendly ORIGINAL command name
                File.WriteAllText(LastCommandFilePath, commandInfo.OriginalName);
            }
            catch (System.Exception ex)
            {
                // Don't let save failures break the command execution
                var ed = AcAp.DocumentManager.MdiActiveDocument?.Editor;
                ed?.WriteMessage($"\nWarning: Failed to save command history: {ex.Message}");
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
                    var activeDoc = AcAp.DocumentManager.MdiActiveDocument;
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
                    var activeDoc = AcAp.DocumentManager.MdiActiveDocument;
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
