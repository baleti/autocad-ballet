using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySelectionFilterInDocument))]

namespace AutoCADBallet
{
    public class SelectBySelectionFilterInDocument
    {
        [CommandMethod("select-by-selection-filter-in-document", CommandFlags.Modal)]
        public void SelectBySelectionFilterInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectBySelectionFilter.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-selection-filter-in-document: {ex.Message}\n");
            }
        }
    }
}
