using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.AlignSelectedInViewCenterVertical))]

namespace AutoCADBallet
{
    public class AlignSelectedInViewCenterVertical
    {
        [CommandMethod("align-selected-in-view-center-vertical", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void AlignSelectedInViewCenterVerticalCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                AlignSelected.ExecuteViewScope(ed, AlignSelected.AlignmentType.CenterVertically);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in align-selected-in-view-center-vertical: {ex.Message}\n");
            }
        }
    }
}
