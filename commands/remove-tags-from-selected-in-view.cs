using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.RemoveTagsFromSelectedInView))]

namespace AutoCADBallet
{
    public class RemoveTagsFromSelectedInView
    {
        [CommandMethod("remove-tags-from-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void RemoveTagsFromSelectedInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                RemoveTagsFromSelected.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in remove-tags-from-selected-in-view: {ex.Message}\n");
            }
        }
    }
}
