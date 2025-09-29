using Autodesk.AutoCAD.Runtime;
using AutoCADCommands;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(EditSelectedTextInView))]

public class EditSelectedTextInView
{
    [CommandMethod("edit-selected-text-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void EditSelectedTextInViewCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            EditSelectedText.ExecuteViewScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in edit-selected-text-in-view: {ex.Message}\n");
        }
    }
}