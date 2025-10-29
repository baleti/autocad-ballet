using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.RemoveTagsFromSelectedInSession))]

namespace AutoCADBallet
{
    public class RemoveTagsFromSelectedInSession
    {
        [CommandMethod("remove-tags-from-selected-in-session", CommandFlags.Modal)]
        public void RemoveTagsFromSelectedInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                RemoveTagsFromSelected.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in remove-tags-from-selected-in-session: {ex.Message}\n");
            }
        }
    }
}
