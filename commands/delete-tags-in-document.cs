using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DeleteTagsInDocument))]

namespace AutoCADBallet
{
    public class DeleteTagsInDocument
    {
        [CommandMethod("delete-tags-in-document", CommandFlags.Modal)]
        public void DeleteTagsInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                DeleteTags.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in delete-tags-in-document: {ex.Message}\n");
            }
        }
    }
}
