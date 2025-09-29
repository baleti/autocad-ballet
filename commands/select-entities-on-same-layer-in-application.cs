using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectEntitiesOnSameLayerInApplication))]

public class SelectEntitiesOnSameLayerInApplication
{
    [CommandMethod("select-entities-on-same-layer-in-application", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectEntitiesOnSameLayerInApplicationCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectEntitiesOnSameLayer.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-entities-on-same-layer-in-application: {ex.Message}\n");
        }
    }
}