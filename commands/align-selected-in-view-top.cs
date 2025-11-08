using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.AlignSelectedInViewTop))]

namespace AutoCADBallet
{
    public class AlignSelectedInViewTop
    {
        [CommandMethod("align-selected-in-view-top", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void AlignSelectedInViewTopCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                AlignSelected.ExecuteViewScope(ed, AlignSelected.AlignmentType.Top);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in align-selected-in-view-top: {ex.Message}\n");
            }
        }
    }
}
