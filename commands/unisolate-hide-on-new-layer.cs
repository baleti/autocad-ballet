using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.UnisolateHideOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class UnisolateHideOnNewLayerCommand
    {
        [CommandMethod("unisolate-hide-on-new-layer", CommandFlags.Modal)]
        public void UnisolateHideOnNewLayer()
        {
            AutoCADBallet.IsolateHideOnNewLayerCommand.IsolateHideUtils.UnisolateHide(AcadApp.DocumentManager.MdiActiveDocument);
        }
    }
}
