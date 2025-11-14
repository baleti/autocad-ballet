using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByContentsInDocument))]

namespace AutoCADBallet
{
    public class SelectByContentsInDocument
    {
        [CommandMethod("select-by-contents-in-document", CommandFlags.Modal)]
        public void SelectByContentsInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                SelectByContents.ExecuteDocumentScope(ed, db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-contents-in-document: {ex.Message}\n");
            }
        }
    }
}
