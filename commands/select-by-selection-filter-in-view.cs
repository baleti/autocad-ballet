using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySelectionFilterInView))]

namespace AutoCADBallet
{
    public class SelectBySelectionFilterInView
    {
        [CommandMethod("select-by-selection-filter-in-view", CommandFlags.Modal)]
        public void SelectBySelectionFilterInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectBySelectionFilter.ExecuteViewScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-selection-filter-in-view: {ex.Message}\n");
            }
        }
    }
}
