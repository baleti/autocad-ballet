using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.ListSystemVariablesInDocument))]

namespace AutoCADBallet
{
    public class ListSystemVariablesInDocument
    {
        [CommandMethod("list-system-variables-in-document", CommandFlags.Modal)]
        public void ListSystemVariablesInDocumentCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                ListSystemVariables.ExecuteDocumentScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in list-system-variables-in-document: {ex.Message}\n");
            }
        }
    }
}
