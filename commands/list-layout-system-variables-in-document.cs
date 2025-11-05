using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.ListLayoutSystemVariablesInDocument))]

namespace AutoCADBallet
{
    public class ListLayoutSystemVariablesInDocument
    {
        [CommandMethod("list-layout-system-variables-in-document", CommandFlags.Modal)]
        public void ListLayoutSystemVariablesInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                ListLayoutSystemVariables.ExecuteDocumentScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in list-layout-system-variables-in-document: {ex.Message}\n");
            }
        }
    }
}
