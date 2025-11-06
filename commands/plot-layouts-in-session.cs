using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.PlotLayoutsInSession))]

namespace AutoCADBallet
{
    public class PlotLayoutsInSession
    {
        [CommandMethod("plot-layouts-in-session", CommandFlags.Session)]
        public void PlotLayoutsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            try
            {
                PlotLayouts.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in plot-layouts-in-session: {ex.Message}\n");
            }
        }
    }
}
