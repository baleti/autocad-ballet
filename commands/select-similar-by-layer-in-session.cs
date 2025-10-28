using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectSimilarByLayerInSession))]

public class SelectSimilarByLayerInSession
{
    [CommandMethod("select-similar-by-layer-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectSimilarByLayerInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectSimilarByLayer.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-similar-by-layer-in-session: {ex.Message}\n");
        }
    }
}