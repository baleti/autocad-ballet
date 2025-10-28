using Autodesk.AutoCAD.Runtime;
using AutoCADCommands;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(EditSelectedTextInSession))]

public class EditSelectedTextInSession
{
    [CommandMethod("edit-selected-text-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void EditSelectedTextInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            EditSelectedText.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in edit-selected-text-in-session: {ex.Message}\n");
        }
    }
}