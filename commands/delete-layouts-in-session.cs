using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DeleteLayoutsInSession))]

namespace AutoCADBallet
{
    public class DeleteLayoutsInSession
    {
        [CommandMethod("delete-layouts-in-session", CommandFlags.Session)]
        public void DeleteLayoutsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            try
            {
                DeleteLayouts.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in delete-layouts-in-session: {ex.Message}\n");
            }
        }
    }
}
