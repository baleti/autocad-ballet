using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectBySameCategoriesOfSelectedInSession))]

public class SelectBySameCategoriesOfSelectedInSession
{
    [CommandMethod("select-by-same-categories-of-selected-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectBySameCategoriesOfSelectedInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectBySameCategoriesOfSelected.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-same-categories-of-selected-in-session: {ex.Message}\n");
        }
    }
}
