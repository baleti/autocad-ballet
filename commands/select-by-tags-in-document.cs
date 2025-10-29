using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByTagsInDocument))]

namespace AutoCADBallet
{
    public class SelectByTagsInDocument
    {
        [CommandMethod("select-by-tags-in-document", CommandFlags.Modal)]
        public void SelectByTagsInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectByTags.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-tags-in-document: {ex.Message}\n");
            }
        }
    }
}
