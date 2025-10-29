using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByTagsOfSelectedInDocument))]

namespace AutoCADBallet
{
    public class SelectByTagsOfSelectedInDocument
    {
        [CommandMethod("select-by-tags-of-selected-in-document", CommandFlags.Modal)]
        public void SelectByTagsOfSelectedInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectByTagsOfSelected.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-tags-of-selected-in-document: {ex.Message}\n");
            }
        }
    }
}
