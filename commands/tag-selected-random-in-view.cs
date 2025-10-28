using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedRandomInView))]

namespace AutoCADBallet
{
    public class TagSelectedRandomInView
    {
        [CommandMethod("tag-selected-random-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void TagSelectedRandomInViewCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                TagSelectedRandom.ExecuteViewScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-random-in-view: {ex.Message}\n");
            }
        }
    }
}
