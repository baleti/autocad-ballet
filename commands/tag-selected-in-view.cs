using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedInView))]

namespace AutoCADBallet
{
    public class TagSelectedInView
    {
        [CommandMethod("tag-selected-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void TagSelectedInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                TagSelected.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-in-view: {ex.Message}\n");
            }
        }
    }
}
