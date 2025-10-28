using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedRandomInDocument))]

namespace AutoCADBallet
{
    public class TagSelectedRandomInDocument
    {
        [CommandMethod("tag-selected-random-in-document", CommandFlags.Modal)]
        public void TagSelectedRandomInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                TagSelectedRandom.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-random-in-document: {ex.Message}\n");
            }
        }
    }
}
