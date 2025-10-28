using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(DuplicateBlocksInSession))]

public class DuplicateBlocksInSession
{
    [CommandMethod("duplicate-blocks-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void DuplicateBlocksInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            DuplicateBlocks.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in duplicate-blocks-in-session: {ex.Message}\n");
        }
    }
}