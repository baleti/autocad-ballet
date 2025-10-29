using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySelectionFiltersInDocument))]

namespace AutoCADBallet
{
    public class SelectBySelectionFiltersInDocument
    {
        [CommandMethod("select-by-selection-filters-in-document", CommandFlags.Modal)]
        public void SelectBySelectionFiltersInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectBySelectionFilters.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-selection-filters-in-document: {ex.Message}\n");
            }
        }
    }
}
