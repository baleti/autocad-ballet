using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.AlignSelectedInViewCenterHorizontal))]

namespace AutoCADBallet
{
    public class AlignSelectedInViewCenterHorizontal
    {
        [CommandMethod("align-selected-in-view-center-horizontal", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void AlignSelectedInViewCenterHorizontalCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                AlignSelected.ExecuteViewScope(ed, AlignSelected.AlignmentType.CenterHorizontally);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in align-selected-in-view-center-horizontal: {ex.Message}\n");
            }
        }
    }
}
