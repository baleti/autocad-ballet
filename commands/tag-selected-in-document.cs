using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedInDocument))]

namespace AutoCADBallet
{
    public class TagSelectedInDocument
    {
        [CommandMethod("tag-selected-in-document", CommandFlags.Modal)]
        public void TagSelectedInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                TagSelected.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-in-document: {ex.Message}\n");
            }
        }
    }
}
