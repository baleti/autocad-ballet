using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByParentTagsOfSelectedInDocument))]

namespace AutoCADBallet
{
    public class SelectByParentTagsOfSelectedInDocument
    {
        [CommandMethod("select-by-parent-tags-of-selected-in-document", CommandFlags.Modal)]
        public void SelectByParentTagsOfSelectedInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectByParentTagsOfSelected.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-parent-tags-of-selected-in-document: {ex.Message}\n");
            }
        }
    }
}
