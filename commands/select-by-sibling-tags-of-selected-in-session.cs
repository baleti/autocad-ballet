using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySiblingTagsOfSelectedInSession))]

namespace AutoCADBallet
{
    public class SelectBySiblingTagsOfSelectedInSession
    {
        [CommandMethod("select-by-sibling-tags-of-selected-in-session", CommandFlags.Modal)]
        public void SelectBySiblingTagsOfSelectedInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectBySiblingTagsOfSelected.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-sibling-tags-of-selected-in-session: {ex.Message}\n");
            }
        }
    }
}
