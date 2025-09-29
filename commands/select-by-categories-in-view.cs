using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectByCategoriesInView))]

public class SelectByCategoriesInView
{
    [CommandMethod("select-by-categories-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectByCategoriesInViewCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            SelectByCategories.ExecuteViewScope(ed, db);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-categories-in-view: {ex.Message}\n");
        }
    }
}