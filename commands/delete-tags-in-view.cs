using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DeleteTagsInView))]

namespace AutoCADBallet
{
    public class DeleteTagsInView
    {
        [CommandMethod("delete-tags-in-view", CommandFlags.Modal)]
        public void DeleteTagsInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                DeleteTags.ExecuteViewScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in delete-tags-in-view: {ex.Message}\n");
            }
        }
    }
}
