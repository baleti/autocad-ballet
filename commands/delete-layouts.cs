using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class DeleteLayouts
    {
        public static void ExecuteDocumentScope(Editor ed)
        {
            Document activeDoc = AcadApp.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;

            Database db = activeDoc.Database;
            string currentLayoutName = LayoutManager.Current.CurrentLayout;

            var allLayouts = new List<Dictionary<string, object>>();
            int currentViewIndex = -1;

            try
            {
                using (DocumentLock docLock = activeDoc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                        int viewIndex = 0;
                        foreach (DictionaryEntry entry in layoutDict)
                        {
                            Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                            if (layout != null && !layout.ModelType)
                            {
                                bool isCurrentView = layout.LayoutName == currentLayoutName;

                                if (isCurrentView)
                                {
                                    currentViewIndex = viewIndex;
                                }

                                allLayouts.Add(new Dictionary<string, object>
                                {
                                    ["layout"] = layout.LayoutName,
                                    ["tab order"] = layout.TabOrder,
                                    ["LayoutName"] = layout.LayoutName,
                                    ["TabOrder"] = layout.TabOrder,
                                    ["IsActive"] = isCurrentView,
                                    ["ObjectId"] = (ObjectId)entry.Value,
                                    ["Handle"] = layout.Handle.ToString()
                                });

                                viewIndex++;
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError reading layouts: {ex.Message}");
                return;
            }

            if (allLayouts.Count == 0)
            {
                ed.WriteMessage("\nNo paper space layouts found to delete.");
                return;
            }

            // Sort by tab order
            allLayouts = allLayouts.OrderBy(v =>
            {
                if (v["TabOrder"] == null) return int.MaxValue;
                if (int.TryParse(v["TabOrder"].ToString(), out int tabOrder))
                {
                    return tabOrder;
                }
                return int.MaxValue;
            }).ToList();

            // Update currentViewIndex after sorting
            currentViewIndex = allLayouts.FindIndex(v => (bool)v["IsActive"]);

            var propertyNames = new List<string> { "layout", "tab order" };
            var initialSelectionIndices = new List<int>();

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allLayouts, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    int totalDeleted = DeleteLayoutsInDocument(chosen, activeDoc, db, ed);

                    ed.WriteMessage($"\nTotal: {totalDeleted} layout(s) deleted successfully.");

                    if (totalDeleted > 0)
                    {
                        activeDoc.SendStringToExecute("_.REGENALL ", true, false, false);
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in delete-layouts-in-document command: {ex.Message}");
            }
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            string currentLayoutName = LayoutManager.Current.CurrentLayout;

            // Generate session identifier for this AutoCAD process
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            string currentSessionId = $"{processId}_{sessionId}";

            var allLayouts = new List<Dictionary<string, object>>();
            int currentViewIndex = -1;
            int viewIndex = 0;

            // Iterate through all open documents
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileName(doc.Name);
                string docFullPath = doc.Name;
                bool isActiveDoc = (doc == activeDoc);

                Database db = doc.Database;

                try
                {
                    using (DocumentLock docLock = doc.LockDocument())
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                            var layoutsInDoc = new List<Dictionary<string, object>>();

                            foreach (DictionaryEntry entry in layoutDict)
                            {
                                Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                                if (layout != null && !layout.ModelType)
                                {
                                    layoutsInDoc.Add(new Dictionary<string, object>
                                    {
                                        ["LayoutName"] = layout.LayoutName,
                                        ["TabOrder"] = layout.TabOrder,
                                        ["ObjectId"] = (ObjectId)entry.Value,
                                        ["Handle"] = layout.Handle.ToString()
                                    });
                                }
                            }

                            // Sort layouts by tab order
                            layoutsInDoc = layoutsInDoc.OrderBy(l =>
                            {
                                if (l["TabOrder"] == null) return int.MaxValue;
                                if (int.TryParse(l["TabOrder"].ToString(), out int tabOrder)) return tabOrder;
                                return int.MaxValue;
                            }).ToList();

                            // Add each layout to the combined list
                            foreach (var layoutInfo in layoutsInDoc)
                            {
                                string layoutName = layoutInfo["LayoutName"].ToString();
                                bool isCurrentView = isActiveDoc && layoutName == currentLayoutName;

                                if (isCurrentView)
                                {
                                    currentViewIndex = viewIndex;
                                }

                                allLayouts.Add(new Dictionary<string, object>
                                {
                                    ["layout"] = layoutName,
                                    ["document"] = Path.GetFileNameWithoutExtension(docName),
                                    ["autocad session"] = currentSessionId,
                                    ["LayoutName"] = layoutName,
                                    ["FullPath"] = docFullPath,
                                    ["TabOrder"] = layoutInfo["TabOrder"],
                                    ["IsActive"] = isCurrentView,
                                    ["DocumentObject"] = doc,
                                    ["ObjectId"] = layoutInfo["ObjectId"],
                                    ["DocumentPath"] = docFullPath,
                                    ["Handle"] = layoutInfo["Handle"]
                                });

                                viewIndex++;
                            }

                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception)
                {
                    // Silently skip documents that can't be read
                }
            }

            if (allLayouts.Count == 0)
            {
                ed.WriteMessage("\nNo paper space layouts found to delete.");
                return;
            }

            // Sort views by document name first, then by tab order
            allLayouts = allLayouts.OrderBy(v => v["document"].ToString())
                                  .ThenBy(v =>
                                  {
                                      if (v["TabOrder"] == null) return int.MaxValue;
                                      if (int.TryParse(v["TabOrder"].ToString(), out int tabOrder))
                                      {
                                          return tabOrder;
                                      }
                                      return int.MaxValue;
                                  })
                                  .ToList();

            // Update currentViewIndex after sorting
            currentViewIndex = allLayouts.FindIndex(v => (bool)v["IsActive"]);

            var propertyNames = new List<string> { "layout", "document", "autocad session" };
            var initialSelectionIndices = new List<int>();

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allLayouts, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    // Group selected layouts by document
                    var layoutsByDocument = chosen.GroupBy(l => l["DocumentObject"] as Document);

                    int totalDeleted = 0;

                    // Process all documents
                    foreach (var docGroup in layoutsByDocument)
                    {
                        Document targetDoc = docGroup.Key;
                        int deletedInDoc = DeleteLayoutsWithCloning(docGroup, targetDoc, ed);
                        totalDeleted += deletedInDoc;

                        string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                        if (deletedInDoc > 0)
                        {
                            ed.WriteMessage($"\nSuccessfully deleted {deletedInDoc} layout(s) in '{docName}'.");
                        }
                    }

                    ed.WriteMessage($"\n\nTotal: {totalDeleted} layout(s) deleted successfully.");

                    // Request regen for the active document only if layouts were deleted there
                    if (layoutsByDocument.Any(g => g.Key == activeDoc))
                    {
                        activeDoc.SendStringToExecute("_.REGENALL ", true, false, false);
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in delete-layouts-in-session command: {ex.Message}");
            }
        }

        private static int DeleteLayoutsInDocument(List<Dictionary<string, object>> selectedLayouts,
            Document targetDoc, Database targetDb, Editor ed)
        {
            int deletedCount = 0;

            try
            {
                using (DocumentLock docLock = targetDoc.LockDocument())
                {
                    using (Transaction tr = targetDb.TransactionManager.StartTransaction())
                    {
                        DBDictionary layoutDict = tr.GetObject(targetDb.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        LayoutManager layoutMgr = LayoutManager.Current;

                        foreach (var selectedLayout in selectedLayouts)
                        {
                            try
                            {
                                string layoutName = selectedLayout["LayoutName"].ToString();
                                ObjectId layoutId = (ObjectId)selectedLayout["ObjectId"];

                                // Check if this is the current layout - switch away if necessary
                                if (layoutMgr.CurrentLayout == layoutName)
                                {
                                    // Find another layout to switch to (prefer Model)
                                    string targetLayout = "Model";
                                    foreach (DictionaryEntry entry in layoutDict)
                                    {
                                        Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                                        if (layout != null && layout.LayoutName != layoutName)
                                        {
                                            targetLayout = layout.LayoutName;
                                            if (layout.ModelType) break; // Prefer Model space
                                        }
                                    }
                                    layoutMgr.CurrentLayout = targetLayout;
                                }

                                // Get the layout and its block table record
                                Layout layoutToDelete = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;
                                if (layoutToDelete != null)
                                {
                                    ObjectId btrId = layoutToDelete.BlockTableRecordId;

                                    // Delete the layout from the dictionary
                                    layoutDict.Remove(layoutName);

                                    // Erase the layout object
                                    layoutToDelete.Erase();

                                    // Delete the associated block table record
                                    if (btrId != ObjectId.Null)
                                    {
                                        BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForWrite) as BlockTableRecord;
                                        if (btr != null)
                                        {
                                            btr.Erase();
                                        }
                                    }

                                    deletedCount++;
                                    ed.WriteMessage($"\nDeleted layout '{layoutName}'");
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nError deleting layout '{selectedLayout["LayoutName"]}': {ex.Message}");
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                ed.WriteMessage($"\nError processing document '{docName}': {ex.Message}");
            }

            return deletedCount;
        }

        private static int DeleteLayoutsWithCloning(IGrouping<Document, Dictionary<string, object>> docGroup,
            Document targetDoc, Editor ed)
        {
            int deletedCount = 0;
            Database targetDb = targetDoc.Database;

            try
            {
                using (DocumentLock docLock = targetDoc.LockDocument())
                {
                    using (Transaction tr = targetDb.TransactionManager.StartTransaction())
                    {
                        DBDictionary layoutDict = tr.GetObject(targetDb.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        LayoutManager layoutMgr = LayoutManager.Current;

                        foreach (var selectedLayout in docGroup)
                        {
                            try
                            {
                                string layoutName = selectedLayout["LayoutName"].ToString();
                                ObjectId layoutId = (ObjectId)selectedLayout["ObjectId"];

                                // Check if this is the current layout - switch away if necessary
                                if (layoutMgr.CurrentLayout == layoutName)
                                {
                                    // Find another layout to switch to (prefer Model)
                                    string targetLayout = "Model";
                                    foreach (DictionaryEntry entry in layoutDict)
                                    {
                                        Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                                        if (layout != null && layout.LayoutName != layoutName)
                                        {
                                            targetLayout = layout.LayoutName;
                                            if (layout.ModelType) break; // Prefer Model space
                                        }
                                    }
                                    layoutMgr.CurrentLayout = targetLayout;
                                }

                                // Get the layout and its block table record
                                Layout layoutToDelete = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;
                                if (layoutToDelete != null)
                                {
                                    ObjectId btrId = layoutToDelete.BlockTableRecordId;

                                    // Delete the layout from the dictionary
                                    layoutDict.Remove(layoutName);

                                    // Erase the layout object
                                    layoutToDelete.Erase();

                                    // Delete the associated block table record
                                    if (btrId != ObjectId.Null)
                                    {
                                        BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForWrite) as BlockTableRecord;
                                        if (btr != null)
                                        {
                                            btr.Erase();
                                        }
                                    }

                                    deletedCount++;
                                    ed.WriteMessage($"\nDeleted layout '{layoutName}'");
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nError deleting layout '{selectedLayout["LayoutName"]}': {ex.Message}");
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                ed.WriteMessage($"\nError processing document '{docName}': {ex.Message}");
            }

            return deletedCount;
        }
    }
}
