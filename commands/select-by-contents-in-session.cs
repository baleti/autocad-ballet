using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByContentsInSession))]

namespace AutoCADBallet
{
    public class SelectByContentsInSession
    {
        [CommandMethod("select-by-contents-in-session", CommandFlags.Modal)]
        public void SelectByContentsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectByContents.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-contents-in-session: {ex.Message}\n");
            }
        }
    }
}
