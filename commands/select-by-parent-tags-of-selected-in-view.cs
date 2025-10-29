using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByParentTagsOfSelectedInView))]

namespace AutoCADBallet
{
    public class SelectByParentTagsOfSelectedInView
    {
        [CommandMethod("select-by-parent-tags-of-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void SelectByParentTagsOfSelectedInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectByParentTagsOfSelected.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-parent-tags-of-selected-in-view: {ex.Message}\n");
            }
        }
    }
}
