using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectByLayersInSession))]

public class SelectByLayersInSession
{
    [CommandMethod("select-by-layers-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectByLayersInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectByLayers.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-by-layers-in-session: {ex.Message}\n");
        }
    }
}