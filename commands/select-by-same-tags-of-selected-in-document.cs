using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySameTagsOfSelectedInDocument))]

namespace AutoCADBallet
{
    public class SelectBySameTagsOfSelectedInDocument
    {
        [CommandMethod("select-by-same-tags-of-selected-in-document", CommandFlags.Modal)]
        public void SelectBySameTagsOfSelectedInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectBySameTagsOfSelected.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-same-tags-of-selected-in-document: {ex.Message}\n");
            }
        }
    }
}
