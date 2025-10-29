using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SaveSelectionFilterFromSelectedInView))]

namespace AutoCADBallet
{
    public class SaveSelectionFilterFromSelectedInView
    {
        [CommandMethod("save-selection-filter-from-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void SaveSelectionFilterFromSelectedInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SaveSelectionFilterFromSelected.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in save-selection-filter-from-selected-in-view: {ex.Message}\n");
            }
        }
    }
}
