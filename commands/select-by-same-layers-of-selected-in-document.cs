using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectBySameLayersOfSelectedInDocument))]

public class SelectBySameLayersOfSelectedInDocument
{
    [CommandMethod("select-by-same-layers-of-selected-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectBySameLayersOfSelectedInDocumentCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            SelectBySameLayersOfSelected.ExecuteDocumentScope(ed, db);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-same-layers-of-selected-in-document: {ex.Message}\n");
        }
    }
}