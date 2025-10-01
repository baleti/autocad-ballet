using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectSimilarByLayerInApplication))]

public class SelectSimilarByLayerInApplication
{
    [CommandMethod("select-similar-by-layer-in-application", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectSimilarByLayerInApplicationCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectSimilarByLayer.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-similar-by-layer-in-application: {ex.Message}\n");
        }
    }
}