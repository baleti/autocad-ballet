using System;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.StopRoslynServer))]

namespace AutoCADBallet
{
    public class StopRoslynServer
    {
        [CommandMethod("stop-roslyn-server", CommandFlags.Session)]
        public void StopRoslynServerCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            var server = StartRoslynServerInBackground.GetServerInstance();

            if (server == null || !server.IsRunning)
            {
                ed?.WriteMessage("\nNo Roslyn server is currently running.\n");
                return;
            }

            try
            {
                StartRoslynServerInBackground.StopServer();
                ed?.WriteMessage("\nRoslyn server stopped successfully.\n");
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\nError stopping Roslyn server: {ex.Message}\n");
            }
        }
    }
}
