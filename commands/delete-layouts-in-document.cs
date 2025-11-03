using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DeleteLayoutsInDocument))]

namespace AutoCADBallet
{
    public class DeleteLayoutsInDocument
    {
        [CommandMethod("delete-layouts-in-document", CommandFlags.Modal)]
        public void DeleteLayoutsInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                DeleteLayouts.ExecuteDocumentScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in delete-layouts-in-document: {ex.Message}\n");
            }
        }
    }
}
