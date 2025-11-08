using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.AlignSelectedInViewLeft))]

namespace AutoCADBallet
{
    public class AlignSelectedInViewLeft
    {
        [CommandMethod("align-selected-in-view-left", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void AlignSelectedInViewLeftCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                AlignSelected.ExecuteViewScope(ed, AlignSelected.AlignmentType.Left);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in align-selected-in-view-left: {ex.Message}\n");
            }
        }
    }
}
