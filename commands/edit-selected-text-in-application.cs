using Autodesk.AutoCAD.Runtime;
using AutoCADCommands;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(EditSelectedTextInApplication))]

public class EditSelectedTextInApplication
{
    [CommandMethod("edit-selected-text-in-application", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void EditSelectedTextInApplicationCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            EditSelectedText.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in edit-selected-text-in-application: {ex.Message}\n");
        }
    }
}