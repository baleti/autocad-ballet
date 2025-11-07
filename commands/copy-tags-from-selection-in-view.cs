using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.CopyTagsFromSelectionInView))]

namespace AutoCADBallet
{
    public class CopyTagsFromSelectionInView
    {
        [CommandMethod("copy-tags-from-selection-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void CopyTagsFromSelectionInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                CopyTagsFromSelection.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in copy-tags-from-selection-in-view: {ex.Message}\n");
            }
        }
    }
}
