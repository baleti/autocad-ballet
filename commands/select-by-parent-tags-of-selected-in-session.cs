using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByParentTagsOfSelectedInSession))]

namespace AutoCADBallet
{
    public class SelectByParentTagsOfSelectedInSession
    {
        [CommandMethod("select-by-parent-tags-of-selected-in-session", CommandFlags.Modal)]
        public void SelectByParentTagsOfSelectedInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectByParentTagsOfSelected.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-parent-tags-of-selected-in-session: {ex.Message}\n");
            }
        }
    }
}
