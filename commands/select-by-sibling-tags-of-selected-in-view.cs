using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySiblingTagsOfSelectedInView))]

namespace AutoCADBallet
{
    public class SelectBySiblingTagsOfSelectedInView
    {
        [CommandMethod("select-by-sibling-tags-of-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void SelectBySiblingTagsOfSelectedInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectBySiblingTagsOfSelected.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-sibling-tags-of-selected-in-view: {ex.Message}\n");
            }
        }
    }
}
