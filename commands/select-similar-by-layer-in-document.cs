using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectSimilarByLayerInDocument))]

public class SelectSimilarByLayerInDocument
{
    [CommandMethod("select-similar-by-layer-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectSimilarByLayerInDocumentCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            SelectSimilarByLayer.ExecuteDocumentScope(ed, db);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-similar-by-layer-in-document: {ex.Message}\n");
        }
    }
}