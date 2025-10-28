using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedInSession))]

namespace AutoCADBallet
{
    public class TagSelectedInSession
    {
        [CommandMethod("tag-selected-in-session", CommandFlags.Modal)]
        public void TagSelectedInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                TagSelected.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-in-session: {ex.Message}\n");
            }
        }
    }
}
