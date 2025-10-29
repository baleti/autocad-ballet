using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SelectByTagsOfSelectedInSession))]

namespace AutoCADBallet
{
    public class SelectByTagsOfSelectedInSession
    {
        [CommandMethod("select-by-tags-of-selected-in-session", CommandFlags.Modal)]
        public void SelectByTagsOfSelectedInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                SelectByTagsOfSelected.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-tags-of-selected-in-session: {ex.Message}\n");
            }
        }
    }
}
