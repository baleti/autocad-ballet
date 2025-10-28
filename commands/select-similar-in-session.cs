using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectSimilarInSession))]

namespace AutoCADBallet
{
    public class SelectSimilarInSession
    {
        [CommandMethod("select-similar-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void SelectSimilarInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectSimilar.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-similar-in-session: {ex.Message}\n");
            }
        }
    }
}
