using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedInApplication))]

namespace AutoCADBallet
{
    public class TagSelectedInApplication
    {
        [CommandMethod("tag-selected-in-application", CommandFlags.Modal)]
        public void TagSelectedInApplicationCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                TagSelected.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-in-application: {ex.Message}\n");
            }
        }
    }
}
