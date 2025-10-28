using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedRandomInSession))]

namespace AutoCADBallet
{
    public class TagSelectedRandomInSession
    {
        [CommandMethod("tag-selected-random-in-session", CommandFlags.Modal)]
        public void TagSelectedRandomInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                TagSelectedRandom.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-random-in-session: {ex.Message}\n");
            }
        }
    }
}
