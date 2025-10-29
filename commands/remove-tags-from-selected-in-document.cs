using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.RemoveTagsFromSelectedInDocument))]

namespace AutoCADBallet
{
    public class RemoveTagsFromSelectedInDocument
    {
        [CommandMethod("remove-tags-from-selected-in-document", CommandFlags.Modal)]
        public void RemoveTagsFromSelectedInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                RemoveTagsFromSelected.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in remove-tags-from-selected-in-document: {ex.Message}\n");
            }
        }
    }
}
