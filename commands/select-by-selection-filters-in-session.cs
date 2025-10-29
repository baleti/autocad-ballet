using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySelectionFiltersInSession))]

namespace AutoCADBallet
{
    public class SelectBySelectionFiltersInSession
    {
        [CommandMethod("select-by-selection-filters-in-session", CommandFlags.Modal)]
        public void SelectBySelectionFiltersInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectBySelectionFilters.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-selection-filters-in-session: {ex.Message}\n");
            }
        }
    }
}
