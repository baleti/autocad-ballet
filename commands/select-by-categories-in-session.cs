using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectByCategoriesInSession))]

public class SelectByCategoriesInSession
{
    [CommandMethod("select-by-categories-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectByCategoriesInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectByCategories.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-categories-in-session: {ex.Message}\n");
        }
    }
}