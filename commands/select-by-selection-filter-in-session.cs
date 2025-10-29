using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySelectionFilterInSession))]

namespace AutoCADBallet
{
    public class SelectBySelectionFilterInSession
    {
        [CommandMethod("select-by-selection-filter-in-session", CommandFlags.Modal)]
        public void SelectBySelectionFilterInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectBySelectionFilter.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-selection-filter-in-session: {ex.Message}\n");
            }
        }
    }
}
