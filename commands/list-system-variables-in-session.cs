using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.ListSystemVariablesInSession))]

namespace AutoCADBallet
{
    public class ListSystemVariablesInSession
    {
        [CommandMethod("list-system-variables-in-session", CommandFlags.Modal)]
        public void ListSystemVariablesInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                ListSystemVariables.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in list-system-variables-in-session: {ex.Message}\n");
            }
        }
    }
}
