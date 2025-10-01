using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectSimilarByLayerInView))]

public class SelectSimilarByLayerInView
{
    [CommandMethod("select-similar-by-layer-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectSimilarByLayerInViewCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectSimilarByLayer.ExecuteViewScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-similar-by-layer-in-view: {ex.Message}\n");
        }
    }
}