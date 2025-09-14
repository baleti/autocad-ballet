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

        [CommandMethod("invoke-last-addin-command", CommandFlags.UsePickSet)]
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

                ed.WriteMessage($"\nCommand completed successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[InvokeLastAddin] Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    ed.WriteMessage($"\n  Inner: {ex.InnerException.Message}");
                }
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
    }
}