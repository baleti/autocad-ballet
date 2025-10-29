using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.RemoveTagsInView))]

namespace AutoCADBallet
{
    public class RemoveTagsInView
    {
        [CommandMethod("remove-tags-in-view", CommandFlags.Modal)]
        public void RemoveTagsInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                RemoveTags.ExecuteViewScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in remove-tags-in-view: {ex.Message}\n");
            }
        }
    }
}
