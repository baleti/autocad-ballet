using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutocadBallet.TestHotReload))]

namespace AutocadBallet
{
    public class TestHotReload
    {
        [CommandMethod("test-hot-reload")]
        public void TestHotReloadCommand()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            // Version 1: Initial test message
            ed.WriteMessage("\n=== Hot Reload Test v1 ===");
            ed.WriteMessage("\nThis is the initial version of the test command.");
            ed.WriteMessage("\nIf you see this, hot reload is working!");
        }
    }
}
