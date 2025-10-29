using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySameTagsOfSelectedInView))]

namespace AutoCADBallet
{
    public class SelectBySameTagsOfSelectedInView
    {
        [CommandMethod("select-by-same-tags-of-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void SelectBySameTagsOfSelectedInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectBySameTagsOfSelected.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-same-tags-of-selected-in-view: {ex.Message}\n");
            }
        }
    }
}
