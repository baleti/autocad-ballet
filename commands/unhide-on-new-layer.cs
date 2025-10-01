using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.UnhideOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class UnhideOnNewLayerCommand
    {
        [CommandMethod("unhide-on-new-layer", CommandFlags.Modal)]
        public void UnhideOnNewLayer()
        {
            AutoCADBallet.HideOnNewLayerCommand.HideLayerUtils.UnhideOnNewLayer(AcadApp.DocumentManager.MdiActiveDocument);
        }
    }
}
