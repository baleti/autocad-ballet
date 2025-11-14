using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByContentsInView))]

namespace AutoCADBallet
{
    public class SelectByContentsInView
    {
        [CommandMethod("select-by-contents-in-view", CommandFlags.Modal)]
        public void SelectByContentsInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectByContents.ExecuteViewScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-contents-in-view: {ex.Message}\n");
            }
        }
    }
}
