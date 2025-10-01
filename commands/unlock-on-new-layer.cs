using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.UnlockOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class UnlockOnNewLayerCommand
    {
        [CommandMethod("unlock-on-new-layer", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void UnlockOnNewLayer()
        {
            AutoCADBallet.LockOnNewLayerCommand.LockLayerUtils.UnlockOnNewLayer(AcadApp.DocumentManager.MdiActiveDocument);
        }
    }
}
