using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySiblingTagsOfSelectedInDocument))]

namespace AutoCADBallet
{
    public class SelectBySiblingTagsOfSelectedInDocument
    {
        [CommandMethod("select-by-sibling-tags-of-selected-in-document", CommandFlags.Modal)]
        public void SelectBySiblingTagsOfSelectedInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectBySiblingTagsOfSelected.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-sibling-tags-of-selected-in-document: {ex.Message}\n");
            }
        }
    }
}
