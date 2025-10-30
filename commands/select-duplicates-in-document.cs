using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectDuplicatesInDocument))]

namespace AutoCADBallet
{
    public class SelectDuplicatesInDocument
    {
        [CommandMethod("select-duplicates-in-document", CommandFlags.Modal)]
        public void SelectDuplicatesInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectDuplicates.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-duplicates-in-document: {ex.Message}\n");
            }
        }
    }
}
