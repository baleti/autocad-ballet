using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectSimilarInDocument))]

namespace AutoCADBallet
{
    public class SelectSimilarInDocument
    {
        [CommandMethod("select-similar-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void SelectSimilarInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectSimilar.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-similar-in-document: {ex.Message}\n");
            }
        }
    }
}
