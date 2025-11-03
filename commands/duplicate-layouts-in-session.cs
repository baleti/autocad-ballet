using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DuplicateLayoutsInSession))]

namespace AutoCADBallet
{
    public class DuplicateLayoutsInSession
    {
        [CommandMethod("duplicate-layouts-in-session", CommandFlags.Session)]
        public void DuplicateLayoutsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            try
            {
                DuplicateLayouts.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in duplicate-layouts-in-session: {ex.Message}\n");
            }
        }
    }
}
