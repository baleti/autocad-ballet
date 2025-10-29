using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySelectionFiltersInView))]

namespace AutoCADBallet
{
    public class SelectBySelectionFiltersInView
    {
        [CommandMethod("select-by-selection-filters-in-view", CommandFlags.Modal)]
        public void SelectBySelectionFiltersInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectBySelectionFilters.ExecuteViewScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-selection-filters-in-view: {ex.Message}\n");
            }
        }
    }
}
