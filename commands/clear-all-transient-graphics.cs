using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.ClearAllTransientGraphics))]

namespace AutoCADBallet
{
    public class ClearAllTransientGraphics
    {
        [CommandMethod("clear-all-transient-graphics", CommandFlags.Modal)]
        public void ClearAllTransientGraphicsCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                // Get document ID
                var db = doc.Database;
                var documentId = db.OriginalFileName;
                if (string.IsNullOrEmpty(documentId))
                {
                    documentId = doc.Name;
                }
                documentId = System.IO.Path.GetFileNameWithoutExtension(documentId);

                ed.WriteMessage($"\n[DEBUG] Document ID: {documentId}\n");

                // Load groups from persistent storage
                var groups = TransientGraphicsStorage.LoadGroups(documentId);
                ed.WriteMessage($"[DEBUG] Loaded {groups.Count} groups from storage\n");

                if (groups.Count == 0)
                {
                    ed.WriteMessage("\nNo transient graphics to clear (no groups in storage).\n");
                    return;
                }

                // Clear all markers from all groups
                int clearedMarkers = 0;
                int clearedGroups = 0;
                foreach (var group in groups)
                {
                    foreach (var marker in group.Markers)
                    {
                        ViewportTransientManager.TransientMgr.EraseTransients(
                            TransientDrawingMode.DirectTopmost,
                            marker,
                            ViewportTransientManager.IntCollection);
                        clearedMarkers++;
                    }
                    clearedGroups++;
                }

                // Force update
                doc.Editor.UpdateScreen();

                ed.WriteMessage($"\nCleared {clearedGroups} viewport outline group(s) ({clearedMarkers} transient markers).\n");

                // Clear the storage file
                TransientGraphicsStorage.ClearGroups(documentId);
                ed.WriteMessage("[DEBUG] Cleared groups storage file\n");

                // Reset marker counter
                ViewportTransientManager.Marker = 1000;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in clear-all-transient-graphics: {ex.Message}\n");
                ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
            }
        }
    }
}
