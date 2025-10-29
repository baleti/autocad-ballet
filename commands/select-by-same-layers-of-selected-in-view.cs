using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectBySameLayersOfSelectedInView))]

public class SelectBySameLayersOfSelectedInView
{
    [CommandMethod("select-by-same-layers-of-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectBySameLayersOfSelectedInViewCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectBySameLayersOfSelected.ExecuteViewScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-same-layers-of-selected-in-view: {ex.Message}\n");
        }
    }
}