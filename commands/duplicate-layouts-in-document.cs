using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DuplicateLayoutsInDocument))]

namespace AutoCADBallet
{
    public class DuplicateLayoutsInDocument
    {
        [CommandMethod("duplicate-layouts-in-document", CommandFlags.Modal)]
        public void DuplicateLayoutsInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                DuplicateLayouts.ExecuteDocumentScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in duplicate-layouts-in-document: {ex.Message}\n");
            }
        }
    }
}
