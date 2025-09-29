using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(DuplicateBlocksInApplication))]

public class DuplicateBlocksInApplication
{
    [CommandMethod("duplicate-blocks-in-application", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void DuplicateBlocksInApplicationCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            DuplicateBlocks.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in duplicate-blocks-in-application: {ex.Message}\n");
        }
    }
}