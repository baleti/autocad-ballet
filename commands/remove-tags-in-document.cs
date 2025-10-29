using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.RemoveTagsInDocument))]

namespace AutoCADBallet
{
    public class RemoveTagsInDocument
    {
        [CommandMethod("remove-tags-in-document", CommandFlags.Modal)]
        public void RemoveTagsInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                RemoveTags.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in remove-tags-in-document: {ex.Message}\n");
            }
        }
    }
}
