using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.AlignSelectedInViewBottom))]

namespace AutoCADBallet
{
    public class AlignSelectedInViewBottom
    {
        [CommandMethod("align-selected-in-view-bottom", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void AlignSelectedInViewBottomCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                AlignSelected.ExecuteViewScope(ed, AlignSelected.AlignmentType.Bottom);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in align-selected-in-view-bottom: {ex.Message}\n");
            }
        }
    }
}
