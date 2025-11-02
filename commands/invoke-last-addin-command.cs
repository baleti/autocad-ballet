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
        private const string FolderName = "autocad-ballet";
        private const string RuntimeFolderName = "runtime";
        private const string LastCommandFileName = "InvokeAddinCommand-history";
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
                // Read last command info from file
                if (!File.Exists(LastCommandFilePath))
                {
                    ed.WriteMessage("\nNo command history found. Use invoke-addin-command first.\n");
                    return;
                }

                var lines = File.ReadAllLines(LastCommandFilePath);
                if (lines.Length < 2)
                {
                    ed.WriteMessage("\nInvalid command history file.\n");
                    return;
                }

                string dllPath = Environment.ExpandEnvironmentVariables(lines[0]);
                string commandName = lines[1];

                if (!File.Exists(dllPath))
                {
                    ed.WriteMessage($"\nDLL not found: {dllPath}\n");
                    return;
                }

                ed.WriteMessage($"\nRe-invoking last command: {commandName}");
                ed.WriteMessage($"\nFrom: {dllPath}");

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
