using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectBySameLayersOfSelectedInSession))]

public class SelectBySameLayersOfSelectedInSession
{
    [CommandMethod("select-by-same-layers-of-selected-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectBySameLayersOfSelectedInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectBySameLayersOfSelected.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-same-layers-of-selected-in-session: {ex.Message}\n");
        }
    }
}