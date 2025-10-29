using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByTagsInSession))]

namespace AutoCADBallet
{
    public class SelectByTagsInSession
    {
        [CommandMethod("select-by-tags-in-session", CommandFlags.Modal)]
        public void SelectByTagsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectByTags.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-tags-in-session: {ex.Message}\n");
            }
        }
    }
}
