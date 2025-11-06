using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.ClearTransientGraphics))]

namespace AutoCADBallet
{
    public class ClearTransientGraphics
    {
        [CommandMethod("clear-transient-graphics", CommandFlags.Modal)]
        public void ClearTransientGraphicsCommand()
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

                // Convert to DataGrid format
                var gridData = new List<Dictionary<string, object>>();
                for (int i = 0; i < groups.Count; i++)
                {
                    var group = groups[i];
                    gridData.Add(new Dictionary<string, object>
                    {
                        ["Description"] = group.Description,
                        ["GroupIndex"] = i // Store index to identify which group to clear
                    });
                }

                // Define columns to display
                var columns = new List<string> { "Description" };

                // Show DataGrid and get user selection
                var selected = CustomGUIs.DataGrid(gridData, columns, spanAllScreens: false);

                if (selected == null || selected.Count == 0)
                {
                    ed.WriteMessage("\nNo groups selected for removal.\n");
                    return;
                }

                // Clear selected groups
                int clearedMarkers = 0;
                int clearedGroups = 0;
                var remainingGroups = new List<TransientGraphicsGroup>();

                for (int i = 0; i < groups.Count; i++)
                {
                    var group = groups[i];
                    bool isSelected = selected.Any(s => (int)s["GroupIndex"] == i);

                    if (isSelected)
                    {
                        // Clear this group's markers
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
                    else
                    {
                        // Keep this group
                        remainingGroups.Add(group);
                    }
                }

                // Force update
                doc.Editor.UpdateScreen();

                ed.WriteMessage($"\nCleared {clearedGroups} viewport outline group(s) ({clearedMarkers} transient markers).\n");

                // Save remaining groups back to storage
                if (remainingGroups.Count > 0)
                {
                    TransientGraphicsStorage.SaveGroups(documentId, remainingGroups);
                    ed.WriteMessage($"[DEBUG] {remainingGroups.Count} group(s) remain in storage\n");
                }
                else
                {
                    // All cleared, delete the file
                    TransientGraphicsStorage.ClearGroups(documentId);
                    ed.WriteMessage("[DEBUG] All groups cleared, removed storage file\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in clear-transient-graphics: {ex.Message}\n");
                ed.WriteMessage($"Stack trace: {ex.StackTrace}\n");
            }
        }
    }
}
