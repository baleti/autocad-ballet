using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectBySameCategoriesOfSelectedInView))]

public class SelectBySameCategoriesOfSelectedInView
{
    [CommandMethod("select-by-same-categories-of-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectBySameCategoriesOfSelectedInViewCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectBySameCategoriesOfSelected.ExecuteViewScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-same-categories-of-selected-in-view: {ex.Message}\n");
        }
    }
}
