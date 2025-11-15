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
            // Just call StopServer() directly - it will send HTTP shutdown request
            // This works even when assembly is hot-reloaded and static serverInstance is null
            StartRoslynServerInBackground.StopServer();
        }
    }
}
