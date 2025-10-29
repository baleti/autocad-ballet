using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.RemoveTagsInSession))]

namespace AutoCADBallet
{
    public class RemoveTagsInSession
    {
        [CommandMethod("remove-tags-in-session", CommandFlags.Modal)]
        public void RemoveTagsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                RemoveTags.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in remove-tags-in-session: {ex.Message}\n");
            }
        }
    }
}
