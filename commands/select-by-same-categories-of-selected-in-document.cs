using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectBySameCategoriesOfSelectedInDocument))]

public class SelectBySameCategoriesOfSelectedInDocument
{
    [CommandMethod("select-by-same-categories-of-selected-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectBySameCategoriesOfSelectedInDocumentCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            SelectBySameCategoriesOfSelected.ExecuteDocumentScope(ed, db);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-same-categories-of-selected-in-document: {ex.Message}\n");
        }
    }
}
