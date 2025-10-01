using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.UnisolateLockOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class UnisolateLockOnNewLayerCommand
    {
        [CommandMethod("unisolate-lock-on-new-layer", CommandFlags.Modal)]
        public void UnisolateLockOnNewLayer()
        {
            AutoCADBallet.IsolateLockOnNewLayerCommand.IsolateLockUtils.UnisolateLock(AcadApp.DocumentManager.MdiActiveDocument);
        }
    }
}
