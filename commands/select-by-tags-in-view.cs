using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByTagsInView))]

namespace AutoCADBallet
{
    public class SelectByTagsInView
    {
        [CommandMethod("select-by-tags-in-view", CommandFlags.Modal)]
        public void SelectByTagsInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectByTags.ExecuteViewScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-tags-in-view: {ex.Message}\n");
            }
        }
    }
}
