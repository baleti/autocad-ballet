using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DuplicateLayoutsCommand))]

namespace AutoCADBallet
{
    public class DuplicateLayoutsCommand
    {
        [CommandMethod("duplicate-layouts")]
        public void DuplicateLayouts()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;
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

                // Get layouts from this document
                Database db = doc.Database;

                try
                {
                    // For non-active documents, we might need to lock the document
                    using (DocumentLock docLock = doc.LockDocument())
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                            var layoutsInDoc = new List<Dictionary<string, object>>();

                            foreach (DictionaryEntry entry in layoutDict)
                            {
                                Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                                if (layout != null)
                                {
                                    // Exclude model space layouts (Model tab)
                                    if (!layout.ModelType)
                                    {
                                        layoutsInDoc.Add(new Dictionary<string, object>
                                        {
                                            ["LayoutName"] = layout.LayoutName,
                                            ["TabOrder"] = layout.TabOrder,
                                            ["LayoutObject"] = layout,
                                            ["ObjectId"] = (ObjectId)entry.Value,
                                            ["Handle"] = layout.Handle.ToString()
                                        });
                                    }
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
                ed.WriteMessage("\nNo paper space layouts found to duplicate.");
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
            var initialSelectionIndices = currentViewIndex >= 0
                                            ? new List<int> { currentViewIndex }
                                            : new List<int>();

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allLayouts, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    // Group selected layouts by document
                    var layoutsByDocument = chosen.GroupBy(l => l["DocumentObject"] as Document);

                    int totalDuplicated = 0;

                    // Separate same-document and cross-document operations
                    var sameDocLayouts = layoutsByDocument.Where(g => g.Key == activeDoc).ToList();
                    var crossDocLayouts = layoutsByDocument.Where(g => g.Key != activeDoc).ToList();

                    // Process same-document layouts first (no document switching needed)
                    foreach (var docGroup in sameDocLayouts)
                    {
                        Document targetDoc = docGroup.Key;
                        totalDuplicated += ProcessLayoutsInDocument(docGroup, targetDoc, activeDoc, docs, ed);
                    }

                    // Process cross-document layouts (requires document switching with events)
                    foreach (var docGroup in crossDocLayouts)
                    {
                        Document targetDoc = docGroup.Key;
                        ProcessCrossDocumentLayouts(docGroup, targetDoc, activeDoc, docs, ed);
                    }

                    ed.WriteMessage($"\nSuccessfully duplicated {totalDuplicated} layout(s).");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in duplicate-layouts command: {ex.Message}");
            }
        }

        private int ProcessLayoutsInDocument(IGrouping<Document, Dictionary<string, object>> docGroup, Document targetDoc, Document activeDoc, DocumentCollection docs, Editor ed)
        {
            int duplicatedCount = 0;
            bool regenRequested = false;
            Editor logEditor = null;

            try
            {
                logEditor = targetDoc?.Editor ?? ed;

                using (DocumentLock docLock = targetDoc.LockDocument())
                {
                    using (Transaction tr = targetDoc.Database.TransactionManager.StartTransaction())
                    {
                        DBDictionary layoutDict = tr.GetObject(targetDoc.Database.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        LayoutManager layoutMgr = LayoutManager.Current;

                        // Process selected layouts in reverse tab order to maintain proper positioning
                        var selectedLayouts = docGroup.OrderByDescending(l =>
                        {
                            if (l["TabOrder"] == null) return int.MinValue;
                            if (int.TryParse(l["TabOrder"].ToString(), out int tabOrder)) return tabOrder;
                            return int.MinValue;
                        }).ToList();

                        foreach (var selectedLayout in selectedLayouts)
                        {
                            try
                            {
                                string originalName = selectedLayout["LayoutName"].ToString();
                                string newName = originalName + " - Copy";

                                // Ensure unique name
                                int counter = 1;
                                string baseName = newName;
                                while (layoutDict.Contains(newName))
                                {
                                    newName = baseName + $" ({counter})";
                                    counter++;
                                }

                                // Get the original layout's tab order
                                int originalTabOrder = int.Parse(selectedLayout["TabOrder"].ToString());

                                // Create new layout (only works on active document)
                                ObjectId newLayoutId = layoutMgr.CreateLayout(newName);
                                Layout newLayout = tr.GetObject(newLayoutId, OpenMode.ForWrite) as Layout;
                                Layout originalLayout = tr.GetObject((ObjectId)selectedLayout["ObjectId"], OpenMode.ForRead) as Layout;

                                if (newLayout != null && originalLayout != null)
                                {
                                    // Copy layout properties
                                    newLayout.CopyFrom(originalLayout);

                                    // Set the tab order to be directly after the original
                                    int targetTabOrder = originalTabOrder + 1;

                                    // Shift tab orders of all layouts that come after the original
                                    foreach (DictionaryEntry entry in layoutDict)
                                    {
                                        Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForWrite) as Layout;
                                        if (layout != null && layout.TabOrder >= targetTabOrder && layout.ObjectId != newLayoutId)
                                        {
                                            layout.TabOrder = layout.TabOrder + 1;
                                        }
                                    }

                                    // Set the new layout's tab order
                                    newLayout.TabOrder = targetTabOrder;

                                    // Copy all entities from original layout to new layout
                                    CopyLayoutEntities(selectedLayout, newLayout, tr, targetDoc);

                                    // Preserve layout-specific system variables
                                    PreserveLayoutSystemVariables(selectedLayout, newName, targetDoc, layoutMgr);

                                    regenRequested = true;
                                    string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                                    logEditor?.WriteMessage($"\nDuplicated layout '{originalName}' as '{newName}' in document '{docName}'");
                                    duplicatedCount++;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                                logEditor?.WriteMessage($"\nError duplicating layout '{selectedLayout["LayoutName"]}' in document '{docName}': {ex.Message}");
                            }
                        }

                        tr.Commit();
                    }
                }

                if (regenRequested)
                {
                    RequestLayoutRegen(targetDoc, targetDoc == activeDoc);
                }
            }
            catch (System.Exception ex)
            {
                string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                (logEditor ?? ed)?.WriteMessage($"\nError processing document '{docName}': {ex.Message}");
            }

            return duplicatedCount;
        }

        private void RequestLayoutRegen(Document targetDoc, bool activateDocument)
        {
            if (targetDoc == null)
            {
                return;
            }

            try
            {
                // Queue REGENALL so the graphics refresh once the command completes
                targetDoc.SendStringToExecute("_.REGENALL ", activateDocument, false, false);
            }
            catch
            {
                try
                {
                    targetDoc.Editor.Regen();
                }
                catch
                {
                    try
                    {
                        targetDoc.Editor.UpdateScreen();
                    }
                    catch
                    {
                        // If all regeneration attempts fail, leave the layouts as-is
                    }
                }
            }
        }

        private void ProcessCrossDocumentLayouts(IGrouping<Document, Dictionary<string, object>> docGroup, Document targetDoc, Document activeDoc, DocumentCollection docs, Editor ed)
        {
            // For cross-document operations, use the DocumentActivated event pattern from CLAUDE.md
            var layoutsToProcess = docGroup.ToList();

            if (layoutsToProcess.Count == 0) return;

            // Set up event handler for when document activation completes
            DocumentCollectionEventHandler handler = null;
            handler = (sender, e) =>
            {
                if (e.Document == targetDoc)
                {
                    // Unsubscribe from event to avoid memory leaks
                    docs.DocumentActivated -= handler;

                    try
                    {
                        // Now we can safely perform layout operations on the activated document
                        int duplicatedInDoc = ProcessLayoutsInDocument(
                            layoutsToProcess.GroupBy(l => targetDoc).First(),
                            targetDoc,
                            activeDoc,
                            docs,
                            ed
                        );

                        // Switch back to original document before writing status to its command line
                        docs.MdiActiveDocument = activeDoc;
                        activeDoc?.Editor?.WriteMessage($"\nCross-document operation completed: {duplicatedInDoc} layout(s) duplicated in '{Path.GetFileNameWithoutExtension(targetDoc.Name)}'");
                    }
                    catch (System.Exception ex)
                    {
                        string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                        try { docs.MdiActiveDocument = activeDoc; } catch { }
                        activeDoc?.Editor?.WriteMessage($"\nError in cross-document operation for '{docName}': {ex.Message}");
                    }
                }
            };

            docs.DocumentActivated += handler;
            docs.MdiActiveDocument = targetDoc;
        }

        private void CopyLayoutEntities(Dictionary<string, object> selectedLayout, Layout newLayout, Transaction targetTr, Document targetDoc)
        {
            try
            {
                Document sourceDoc = selectedLayout["DocumentObject"] as Document;
                ObjectId sourceLayoutId = (ObjectId)selectedLayout["ObjectId"];

                if (sourceDoc == null)
                    return;

                Database sourceDb = sourceDoc.Database;
                Database targetDb = targetDoc.Database;
                bool sameDocument = (sourceDoc == targetDoc);

                if (sameDocument)
                {
                    // Same document - use current transaction
                    Layout originalLayout = targetTr.GetObject(sourceLayoutId, OpenMode.ForRead) as Layout;
                    BlockTableRecord originalBtr = targetTr.GetObject(originalLayout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    BlockTableRecord newBtr = targetTr.GetObject(newLayout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                    if (originalBtr == null || newBtr == null)
                        return;

                    // Collect all entity IDs from the original layout
                    var entityIds = new ObjectIdCollection();
                    foreach (ObjectId entityId in originalBtr)
                    {
                        Entity entity = targetTr.GetObject(entityId, OpenMode.ForRead) as Entity;
                        if (entity != null)
                        {
                            entityIds.Add(entityId);
                        }
                    }

                    // If there are entities to copy
                    if (entityIds.Count > 0)
                    {
                        // Same database - use DeepClone
                        IdMapping idMap = new IdMapping();
                        targetDb.DeepCloneObjects(entityIds, newBtr.ObjectId, idMap, false);
                    }
                }
                else
                {
                    // Cross-document copying - need separate transactions
                    using (DocumentLock sourceLock = sourceDoc.LockDocument())
                    {
                        using (Transaction sourceTr = sourceDb.TransactionManager.StartTransaction())
                        {
                            Layout originalLayout = sourceTr.GetObject(sourceLayoutId, OpenMode.ForRead) as Layout;
                            BlockTableRecord originalBtr = sourceTr.GetObject(originalLayout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                            BlockTableRecord newBtr = targetTr.GetObject(newLayout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                            if (originalBtr == null || newBtr == null)
                                return;

                            // Collect all entity IDs from the original layout
                            var entityIds = new ObjectIdCollection();
                            foreach (ObjectId entityId in originalBtr)
                            {
                                Entity entity = sourceTr.GetObject(entityId, OpenMode.ForRead) as Entity;
                                if (entity != null)
                                {
                                    entityIds.Add(entityId);
                                }
                            }

                            // If there are entities to copy
                            if (entityIds.Count > 0)
                            {
                                // Cross-database - use WblockCloneObjects
                                IdMapping idMap = new IdMapping();
                                sourceDb.WblockCloneObjects(entityIds, newBtr.ObjectId, idMap, DuplicateRecordCloning.Replace, false);
                            }

                            sourceTr.Commit();
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                // If entity copying fails, continue without entities
                // The layout structure will still be duplicated
            }
        }

        private void PreserveLayoutSystemVariables(Dictionary<string, object> selectedLayout, string newLayoutName, Document targetDoc, LayoutManager layoutMgr)
        {
            try
            {
                Document sourceDoc = selectedLayout["DocumentObject"] as Document;
                string sourceLayoutName = selectedLayout["LayoutName"].ToString();

                if (sourceDoc == null)
                    return;

                // Get current system variable values from source layout
                Dictionary<string, object> sourceVariables = GetLayoutSystemVariables(sourceDoc, sourceLayoutName);

                // Apply these values to the new layout
                ApplyLayoutSystemVariables(targetDoc, newLayoutName, sourceVariables, layoutMgr);
            }
            catch (System.Exception)
            {
                // If system variable preservation fails, continue without it
                // The layout and entities will still be duplicated
            }
        }

        private Dictionary<string, object> GetLayoutSystemVariables(Document sourceDoc, string layoutName)
        {
            var variables = new Dictionary<string, object>();

            try
            {
                // For cross-document system variables, we'll use default safe values
                // Since PSLTSCALE and other variables can't be reliably read from non-active documents
                // without switching, we'll preserve the most common settings

                if (sourceDoc == AcadApp.DocumentManager.MdiActiveDocument)
                {
                    // Same document - can safely read current values
                    using (DocumentLock docLock = sourceDoc.LockDocument())
                    {
                        // Temporarily switch to the source layout to get its system variables
                        string currentLayout = LayoutManager.Current.CurrentLayout;
                        bool needToSwitchBack = (currentLayout != layoutName);

                        if (needToSwitchBack)
                        {
                            using (DocumentLock layoutLock = sourceDoc.LockDocument())
                            {
                                LayoutManager.Current.CurrentLayout = layoutName;
                            }
                        }

                        // Capture layout-specific system variables
                        variables["PSLTSCALE"] = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("PSLTSCALE");
                        variables["MSLTSCALE"] = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("MSLTSCALE");
                        variables["LTSCALE"] = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LTSCALE");
                        variables["LIMMIN"] = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LIMMIN");
                        variables["LIMMAX"] = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("LIMMAX");

                        // Switch back to original layout if we changed it
                        if (needToSwitchBack)
                        {
                            using (DocumentLock layoutLock = sourceDoc.LockDocument())
                            {
                                LayoutManager.Current.CurrentLayout = currentLayout;
                            }
                        }
                    }
                }
                else
                {
                    // Cross-document - use safe default values (most common settings)
                    // These will be overridden by the event-driven approach for cross-document
                    variables["PSLTSCALE"] = 0; // Most architectural drawings use 0
                    variables["MSLTSCALE"] = 1; // Standard setting
                    variables["LTSCALE"] = 1.0; // Standard scale
                    variables["LIMMIN"] = new double[] { 0.0, 0.0 }; // Standard limits
                    variables["LIMMAX"] = new double[] { 420.0, 297.0 }; // A3 metric default
                }
            }
            catch (System.Exception)
            {
                // If we can't get the variables, use safe defaults
                variables["PSLTSCALE"] = 0;
                variables["MSLTSCALE"] = 1;
                variables["LTSCALE"] = 1.0;
                variables["LIMMIN"] = new double[] { 0.0, 0.0 };
                variables["LIMMAX"] = new double[] { 420.0, 297.0 };
            }

            return variables;
        }

        private void ApplyLayoutSystemVariables(Document targetDoc, string layoutName, Dictionary<string, object> variables, LayoutManager layoutMgr)
        {
            try
            {
                if (variables.Count == 0)
                    return;

                using (DocumentLock docLock = targetDoc.LockDocument())
                {
                    // Temporarily switch to the new layout to set its system variables
                    string currentLayout = LayoutManager.Current.CurrentLayout;
                    bool needToSwitchBack = (currentLayout != layoutName);

                    if (needToSwitchBack)
                    {
                        LayoutManager.Current.CurrentLayout = layoutName;
                    }

                    // Apply the captured system variables
                    foreach (var variable in variables)
                    {
                        try
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(variable.Key, variable.Value);
                        }
                        catch (System.Exception)
                        {
                            // Skip individual variables that fail to set
                        }
                    }

                    // Switch back to original layout if we changed it
                    if (needToSwitchBack)
                    {
                        LayoutManager.Current.CurrentLayout = currentLayout;
                    }
                }
            }
            catch (System.Exception)
            {
                // If we can't apply the variables, continue without them
            }
        }
    }
}
