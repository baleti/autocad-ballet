using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedRandomInApplication))]

namespace AutoCADBallet
{
    public class TagSelectedRandomInApplication
    {
        [CommandMethod("tag-selected-random-in-application", CommandFlags.Modal)]
        public void TagSelectedRandomInApplicationCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                TagSelectedRandom.ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in tag-selected-random-in-application: {ex.Message}\n");
            }
        }
    }
}
