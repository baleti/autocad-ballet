using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(CopySelectedToViewsInProcess))]

public class CopySelectedToViewsInProcess
{
    [CommandMethod("copy-selected-to-views-in-process", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void CopyToViews()
    {
        DocumentCollection docs = AcadApp.DocumentManager;
        Document activeDoc = docs.MdiActiveDocument;
        if (activeDoc == null) return;

        Editor ed = activeDoc.Editor;
        Database db = activeDoc.Database;

        try
        {
            // Get current selection using pickfirst pattern (same as filter-selected)
            List<ObjectId> selectedObjects = new List<ObjectId>();

            // Get pickfirst set (pre-selected objects)
            var selResult = ed.SelectImplied();

            // If there is no pickfirst set, request user to select objects
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to copy: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                // Clear the pickfirst set since we're consuming it
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects selected.\n");
                return;
            }

            selectedObjects.AddRange(selResult.Value.GetObjectIds());

            // Get available layouts
            var allViews = GetAvailableLayouts(docs);
            if (allViews.Count == 0)
            {
                ed.WriteMessage("\nNo layouts found.\n");
                return;
            }

            // Show layout selection dialog
            var propertyNames = new List<string> { "layout", "document", "autocad session" };

            try
            {
                var chosen = CustomGUIs.DataGrid(allViews, propertyNames, false, null);

                if (chosen == null || chosen.Count == 0)
                {
                    ed.WriteMessage("\nNo layouts selected.\n");
                    return;
                }

                // Copy objects to selected layouts
                int copiedCount = 0;
                foreach (var selectedLayout in chosen)
                {
                    Document targetDoc = selectedLayout["DocumentObject"] as Document;
                    string targetLayout = selectedLayout["LayoutName"].ToString();

                    if (targetDoc != null)
                    {
                        if (CopyObjectsToLayout(selectedObjects, activeDoc, targetDoc, targetLayout))
                        {
                            copiedCount++;
                        }
                    }
                }

                ed.WriteMessage($"\nObjects copied to {copiedCount} layout(s).\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during layout selection: {ex.Message}\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nUnexpected error: {ex.Message}\n");
        }
    }


    private List<Dictionary<string, object>> GetAvailableLayouts(DocumentCollection docs)
    {
        var allViews = new List<Dictionary<string, object>>();
        // Generate session identifier for this AutoCAD process
        int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
        int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
        string currentSessionId = $"{processId}_{sessionId}";

        foreach (Document doc in docs)
        {
            string docName = Path.GetFileName(doc.Name);
            string docFullPath = doc.Name;
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
                            if (layout != null)
                            {
                                layoutsInDoc.Add(new Dictionary<string, object>
                                {
                                    ["LayoutName"] = layout.LayoutName,
                                    ["TabOrder"] = layout.TabOrder,
                                    ["LayoutObject"] = layout
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

                            allViews.Add(new Dictionary<string, object>
                            {
                                ["layout"] = layoutName,
                                ["document"] = Path.GetFileNameWithoutExtension(docName),
                                ["autocad session"] = currentSessionId,
                                ["LayoutName"] = layoutName,
                                ["FullPath"] = docFullPath,
                                ["TabOrder"] = layoutInfo["TabOrder"],
                                ["DocumentObject"] = doc
                            });
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception)
            {
                // Silently skip documents that can't be read
                continue;
            }
        }

        // Sort views by document name first, then by tab order
        return allViews.OrderBy(v => v["document"].ToString())
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
    }

    private bool CopyObjectsToLayout(List<ObjectId> sourceObjects, Document sourceDoc, Document targetDoc, string targetLayout)
    {
        try
        {
            if (sourceDoc == targetDoc)
            {
                // Same document - copy to different layout
                return CopyWithinDocument(sourceObjects, sourceDoc, targetLayout);
            }
            else
            {
                // Different document - copy across documents
                return CopyAcrossDocuments(sourceObjects, sourceDoc, targetDoc, targetLayout);
            }
        }
        catch (System.Exception)
        {
            return false;
        }
    }

    private bool CopyWithinDocument(List<ObjectId> sourceObjects, Document doc, string targetLayout)
    {
        Database db = doc.Database;

        using (DocumentLock docLock = doc.LockDocument())
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get target layout's block table record
                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (!layoutDict.Contains(targetLayout))
                        return false;

                    Layout layout = tr.GetObject(layoutDict.GetAt(targetLayout), OpenMode.ForRead) as Layout;
                    BlockTableRecord targetBtr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                    // Copy objects
                    IdMapping idMap = new IdMapping();
                    ObjectIdCollection sourceIds = new ObjectIdCollection(sourceObjects.ToArray());

                    db.DeepCloneObjects(sourceIds, targetBtr.ObjectId, idMap, false);

                    tr.Commit();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
    }

    private bool CopyAcrossDocuments(List<ObjectId> sourceObjects, Document sourceDoc, Document targetDoc, string targetLayout)
    {
        try
        {
            using (DocumentLock sourceLock = sourceDoc.LockDocument())
            using (DocumentLock targetLock = targetDoc.LockDocument())
            {
                Database sourceDb = sourceDoc.Database;
                Database targetDb = targetDoc.Database;

                using (Transaction sourceTr = sourceDb.TransactionManager.StartTransaction())
                using (Transaction targetTr = targetDb.TransactionManager.StartTransaction())
                {
                    // Get target layout's block table record
                    DBDictionary layoutDict = targetTr.GetObject(targetDb.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (!layoutDict.Contains(targetLayout))
                        return false;

                    Layout layout = targetTr.GetObject(layoutDict.GetAt(targetLayout), OpenMode.ForRead) as Layout;
                    BlockTableRecord targetBtr = targetTr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

                    // Copy objects across databases
                    IdMapping idMap = new IdMapping();
                    ObjectIdCollection sourceIds = new ObjectIdCollection(sourceObjects.ToArray());

                    sourceDb.WblockCloneObjects(sourceIds, targetBtr.ObjectId, idMap, DuplicateRecordCloning.Replace, false);

                    sourceTr.Commit();
                    targetTr.Commit();
                    return true;
                }
            }
        }
        catch
        {
            return false;
        }
    }
}