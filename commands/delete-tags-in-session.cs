using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DeleteTagsInSession))]

namespace AutoCADBallet
{
    public class DeleteTagsInSession
    {
        [CommandMethod("delete-tags-in-session", CommandFlags.Modal)]
        public void DeleteTagsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                DeleteTags.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in delete-tags-in-session: {ex.Message}\n");
            }
        }
    }
}
