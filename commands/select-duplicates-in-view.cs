using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectDuplicatesInView))]

namespace AutoCADBallet
{
    public class SelectDuplicatesInView
    {
        [CommandMethod("select-duplicates-in-view", CommandFlags.Modal)]
        public void SelectDuplicatesInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectDuplicates.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-duplicates-in-view: {ex.Message}\n");
            }
        }
    }
}
