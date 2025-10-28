using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(DeleteSelectedInSession))]

public class DeleteSelectedInSession
{
    [CommandMethod("delete-selected-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void DeleteSelectedInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            DeleteSelected.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in delete-selected-in-session: {ex.Message}\n");
        }
    }
}