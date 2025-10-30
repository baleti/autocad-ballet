using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectDuplicatesInSession))]

namespace AutoCADBallet
{
    public class SelectDuplicatesInSession
    {
        [CommandMethod("select-duplicates-in-session", CommandFlags.Modal)]
        public void SelectDuplicatesInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectDuplicates.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-duplicates-in-session: {ex.Message}\n");
            }
        }
    }
}
