using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

// Alias to avoid Application naming ambiguity
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutocadBallet.InvokeAddinCommand))]

namespace AutocadBallet
{
    /// <summary>
    /// Collectible load context that loads target DLL + its private deps from bytes
    /// </summary>
    internal sealed class AddinLoadContext : AssemblyLoadContext
    {
        private readonly string _baseDir;

        public AddinLoadContext(string baseDir)
            : base(isCollectible: true)
        {
            _baseDir = baseDir;
            this.Resolving += OnResolving;
        }

        public Assembly LoadFromPathNoLock(string path)
        {
            using var fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var ms = new System.IO.MemoryStream();
            fs.CopyTo(ms);
            ms.Position = 0;
            return LoadFromStream(ms);
        }

        private Assembly? OnResolving(AssemblyLoadContext alc, AssemblyName name)
        {
            // 1) If AutoCAD already has it in Default ALC, reuse that assembly
            var existing = AssemblyLoadContext.Default.Assemblies
                .FirstOrDefault(a => string.Equals(a.GetName().Name, name.Name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
                return existing;

            // 2) Otherwise, attempt to load from the same folder as the main add-in
            try
            {
                var candidate = Path.Combine(_baseDir, name.Name + ".dll");
                if (File.Exists(candidate))
                {
                    using var fs = File.Open(candidate, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var ms = new System.IO.MemoryStream();
                    fs.CopyTo(ms);
                    ms.Position = 0;
                    return LoadFromStream(ms);
                }
            }
            catch
            {
                // ignore and fall through
            }

            return null;
        }
    }

    public class InvokeAddinCommand
    {
        private const string TargetDllPath = @"%appdata%\autocad-ballet\commands.dll";

        [CommandMethod("INVOKEADDIN", CommandFlags.Session)]
        public void InvokeAddin()
        {
            var ed = AcAp.DocumentManager.MdiActiveDocument?.Editor;
            if (ed == null) return;

            // Resolve the path
            string dllPath = Environment.ExpandEnvironmentVariables(TargetDllPath);

            if (!File.Exists(dllPath))
            {
                ed.WriteMessage($"\nTarget DLL not found: {dllPath}");
                ed.WriteMessage("\nPlease ensure commands.dll exists in %appdata%\\autocad-ballet\\");
                return;
            }

            AddinLoadContext? alc = null;
            try
            {
                // Get all commands from the assembly
                var commands = GetCommandList(dllPath, out alc, out var asm);

                if (commands.Count == 0)
                {
                    ed.WriteMessage("\nNo [CommandMethod] commands found in the assembly.");
                    SafeUnload(alc);
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
                    SafeUnload(alc);
                    return;
                }

                // Get the selected command info
                var selectedEntry = selected.First();
                var selectedCmd = commands.First(c =>
                    c.CommandName == selectedEntry["Command"].ToString() &&
                    c.FullTypeName == selectedEntry["Full Name"].ToString());

                ed.WriteMessage($"\nInvoking command: {selectedCmd.CommandName}");

                // Invoke the chosen method
                InvokeCommandMethod(selectedCmd, asm, ed);

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
                // Always unload to free file lock
                SafeUnload(alc);
            }
        }

        private record CommandInfo(
            string CommandName,
            string TypeName,
            string FullTypeName,
            string MethodName);

        private List<CommandInfo> GetCommandList(string dllPath, out AddinLoadContext alc, out Assembly asm)
        {
            var baseDir = Path.GetDirectoryName(dllPath)!;
            alc = new AddinLoadContext(baseDir);
            asm = alc.LoadFromPathNoLock(dllPath);

            var list = new List<CommandInfo>();

            foreach (var t in asm.GetTypes())
            {
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);
                foreach (var m in methods)
                {
                    foreach (var attr in m.GetCustomAttributes(false))
                    {
                        // Check for CommandMethodAttribute
                        if (attr.GetType().FullName == "Autodesk.AutoCAD.Runtime.CommandMethodAttribute")
                        {
                            string commandName = TryGetCommandName(attr) ?? m.Name;
                            list.Add(new CommandInfo(
                                commandName,
                                t.Name,
                                t.FullName!,
                                m.Name));
                        }
                    }
                }
            }

            return list.OrderBy(c => c.CommandName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private string? TryGetCommandName(object cmdAttr)
        {
            var type = cmdAttr.GetType();

            // Try GlobalName property
            var prop = type.GetProperty("GlobalName", BindingFlags.Public | BindingFlags.Instance);
            if (prop != null)
            {
                var val = prop.GetValue(cmdAttr) as string;
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            // Try other string properties
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

        private void InvokeCommandMethod(CommandInfo info, Assembly asm, Editor? ed)
        {
            try
            {
                var type = asm.GetType(info.FullTypeName, throwOnError: true)!;
                var method = type.GetMethod(info.MethodName,
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);

                if (method == null)
                    throw new MissingMethodException(info.FullTypeName, info.MethodName);

                object? instance = null;
                if (!method.IsStatic)
                {
                    instance = Activator.CreateInstance(type);
                }

                // Command methods in AutoCAD typically take no parameters
                method.Invoke(instance, parameters: null);
            }
            catch (TargetInvocationException tie)
            {
                ed?.WriteMessage($"\n[InvokeAddin] Command threw: {tie.InnerException?.Message ?? tie.Message}");
                throw;
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\n[InvokeAddin] Failed to invoke: {ex.Message}");
                throw;
            }
        }

        private void SafeUnload(AddinLoadContext? alc)
        {
            if (alc == null) return;

            try
            {
                alc.Unload();
                // Encourage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch
            {
                // ignore
            }
        }
    }
}
