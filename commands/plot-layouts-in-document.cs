using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.PlotLayoutsInDocument))]

namespace AutoCADBallet
{
    public class PlotLayoutsInDocument
    {
        [CommandMethod("plot-layouts-in-document", CommandFlags.Modal)]
        public void PlotLayoutsInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                PlotLayouts.ExecuteDocumentScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in plot-layouts-in-document: {ex.Message}\n");
            }
        }
    }
}
