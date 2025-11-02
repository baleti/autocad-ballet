using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.IO;
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
        private const string LastCommandFileName = "invoke-addin-command-last";
        private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName, RuntimeFolderName);
        private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

        [CommandMethod("invoke-last-addin-command", CommandFlags.Session | CommandFlags.UsePickSet)]
        public void Execute()
        {
            var ed = AcApp.DocumentManager.MdiActiveDocument?.Editor;
            if (ed == null) return;

            // Capture pickfirst selection at the start to preserve it
            Autodesk.AutoCAD.DatabaseServices.ObjectId[] pickfirstSelection = null;
            var selResult = ed.SelectImplied();
            if (selResult.Status == PromptStatus.OK)
            {
                pickfirstSelection = selResult.Value.GetObjectIds();
            }

            try
            {
                // Read last command name from file
                if (!File.Exists(LastCommandFilePath))
                {
                    ed.WriteMessage("\nNo command history found. Use invoke-addin-command first.\n");
                    return;
                }

                string commandName = File.ReadAllText(LastCommandFilePath).Trim();

                if (string.IsNullOrWhiteSpace(commandName))
                {
                    ed.WriteMessage("\nInvalid command history file.\n");
                    return;
                }

                // Use hardcoded DLL path
                string dllPath = Environment.ExpandEnvironmentVariables(TargetDllPath);

                if (!File.Exists(dllPath))
                {
                    ed.WriteMessage($"\nDLL not found: {dllPath}\n");
                    return;
                }

                ed.WriteMessage($"\nRe-invoking last command: {commandName}");

                // Use shared code from InvokeAddinCommand
                var sessionAssemblies = new System.Collections.Generic.Dictionary<string, System.Reflection.Assembly>();

                // Rewrite and load the command
                var cachedCommand = InvokeAddinCommand.RewriteAndLoadSingleCommand(dllPath, commandName, ed, sessionAssemblies);

                ed.WriteMessage($"\nLoaded command: {commandName}");

                // Invoke the command
                InvokeAddinCommand.InvokeCommandMethod(cachedCommand, ed, pickfirstSelection);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError invoking last command: {ex.Message}\n");
            }
        }
    }
}
