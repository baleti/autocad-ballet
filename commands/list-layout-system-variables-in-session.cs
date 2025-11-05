using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.ListLayoutSystemVariablesInSession))]

namespace AutoCADBallet
{
    public class ListLayoutSystemVariablesInSession
    {
        [CommandMethod("list-layout-system-variables-in-session", CommandFlags.Modal)]
        public void ListLayoutSystemVariablesInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                ListLayoutSystemVariables.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in list-layout-system-variables-in-session: {ex.Message}\n");
            }
        }
    }
}
