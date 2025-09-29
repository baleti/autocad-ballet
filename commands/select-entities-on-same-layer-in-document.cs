using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectEntitiesOnSameLayerInDocument))]

public class SelectEntitiesOnSameLayerInDocument
{
    [CommandMethod("select-entities-on-same-layer-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectEntitiesOnSameLayerInDocumentCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            SelectEntitiesOnSameLayer.ExecuteDocumentScope(ed, db);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-entities-on-same-layer-in-document: {ex.Message}\n");
        }
    }
}