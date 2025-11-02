using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Mono.Cecil;
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

        [CommandMethod("invoke-last-addin-command", CommandFlags.Session | CommandFlags.UsePickSet)]
        public void Execute()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument?.Editor;
            if (ed == null) return;

            // Capture pickfirst selection at the start to preserve it
            // AutoCAD will clear it when the modal DataGrid shows, but we'll restore it via LISP
            Autodesk.AutoCAD.DatabaseServices.ObjectId[] pickfirstSelection = null;
            var selResult = ed.SelectImplied();
            if (selResult.Status == PromptStatus.OK)
            {
                pickfirstSelection = selResult.Value.GetObjectIds();
            }

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

                // Get file info for cache management
                var fileInfo = new FileInfo(dllPath);
                string baseCacheKey = $"{dllPath}|{fileInfo.LastWriteTimeUtc.Ticks}|{fileInfo.Length}";
                string commandCacheKey = $"{baseCacheKey}|{lastCommandName}";

                CachedCommandInfo cachedCommand = null;

                // Check if this specific command is already cached
                if (loadedCommandCache.ContainsKey(commandCacheKey))
                {
                    cachedCommand = loadedCommandCache[commandCacheKey];
                    ed.WriteMessage($"\nUsing cached command: {lastCommandName}");
                }
                else
                {
                    // Verify the command exists before loading (quick check using Mono.Cecil)
                    ed.WriteMessage("\nVerifying command exists...");
                    var commandNames = InspectCommandNames(dllPath);

                    if (!commandNames.Any(c => c.Equals(lastCommandName, StringComparison.OrdinalIgnoreCase)))
                    {
                        ed.WriteMessage($"\nCommand '{lastCommandName}' not found in the assembly.");
                        return;
                    }

                    // Clear old cache entries for this DLL version (different timestamp/size)
                    var keysToRemove = loadedCommandCache.Keys
                        .Where(k => k.StartsWith(dllPath + "|") && !k.StartsWith(baseCacheKey))
                        .ToList();
                    foreach (var key in keysToRemove)
                    {
                        loadedCommandCache.Remove(key);
                    }

                    // NOW load/rewrite ONLY this command
                    cachedCommand = LoadSingleCommand(dllPath, lastCommandName, ed);
                    if (cachedCommand == null) return;

                    // Cache it
                    loadedCommandCache[commandCacheKey] = cachedCommand;
                }

                // Invoke the command method through AutoCAD's command pipeline
                InvokeCommandMethod(cachedCommand, ed, pickfirstSelection);

                // Note: SendStringToExecute is asynchronous, so we can't report completion here
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

        private CachedCommandInfo LoadSingleCommand(string dllPath, string targetCommandName, Editor ed)
        {
            try
            {
                // Register the assembly resolve event handler for dependencies
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

                try
                {
                    // Rewrite ONLY the selected command with unique name and load
                    ed.WriteMessage($"\nRewriting and loading command: {targetCommandName}...");
                    var cachedCommand = RewriteAndLoadSingleCommand(dllPath, targetCommandName, ed);
                    ed.WriteMessage($"\nLoaded command: {targetCommandName}");
                    return cachedCommand;
                }
                finally
                {
                    AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[LoadCommand] Error: {ex.Message}");
                return null;
            }
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

                var doc = AcApp.DocumentManager.MdiActiveDocument;
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
                    ed.WriteMessage($"\n[InvokeLastAddin] Failed to invoke command '{commandInfo.OriginalName}': {ex.Message}");
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