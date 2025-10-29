using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectBySameTagsOfSelectedInSession))]

namespace AutoCADBallet
{
    public class SelectBySameTagsOfSelectedInSession
    {
        [CommandMethod("select-by-same-tags-of-selected-in-session", CommandFlags.Modal)]
        public void SelectBySameTagsOfSelectedInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectBySameTagsOfSelected.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-same-tags-of-selected-in-session: {ex.Message}\n");
            }
        }
    }
}
