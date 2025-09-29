using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(SelectEntitiesOnSameLayerInView))]

public class SelectEntitiesOnSameLayerInView
{
    [CommandMethod("select-entities-on-same-layer-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectEntitiesOnSameLayerInViewCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            SelectEntitiesOnSameLayer.ExecuteViewScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in select-entities-on-same-layer-in-view: {ex.Message}\n");
        }
    }
}