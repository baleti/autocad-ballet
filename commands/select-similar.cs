using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class SelectSimilar
    {
        // Helper class to store entity similarity information
        private class SimilarityKey
        {
            public string EntityType { get; set; }
            public string BlockName { get; set; }  // For BlockReferences
            public string Layer { get; set; }
            public short ColorIndex { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is SimilarityKey other)
                {
                    return EntityType == other.EntityType &&
                           BlockName == other.BlockName &&
                           Layer == other.Layer &&
                           ColorIndex == other.ColorIndex;
                }
                return false;
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + (EntityType?.GetHashCode() ?? 0);
                    hash = hash * 23 + (BlockName?.GetHashCode() ?? 0);
                    hash = hash * 23 + (Layer?.GetHashCode() ?? 0);
                    hash = hash * 23 + ColorIndex.GetHashCode();
                    return hash;
                }
            }
        }

        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;

            // Get pickfirst selection
            var selResult = ed.SelectImplied();

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects selected. Please select objects first, then run this command.\n");
                return;
            }

            var selectedIds = selResult.Value.GetObjectIds();
            var similarityKeys = new HashSet<SimilarityKey>();

            ed.WriteMessage($"\nDocument Mode: Finding entities similar to {selectedIds.Length} selected object(s)...\n");

            // Extract similarity keys from selected entities
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in selectedIds)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity != null)
                        {
                            var key = GetSimilarityKey(entity);
                            similarityKeys.Add(key);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                tr.Commit();
            }

            if (similarityKeys.Count == 0)
            {
                ed.WriteMessage("\nNo valid entities selected.\n");
                return;
            }

            // Find all similar entities across all layouts in the document
            var similarEntities = new List<SelectionItem>();
            var currentViewSimilarIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var isCurrentSpace = (btr.ObjectId == db.CurrentSpaceId);

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity != null)
                            {
                                var key = GetSimilarityKey(entity);
                                if (similarityKeys.Contains(key))
                                {
                                    similarEntities.Add(new SelectionItem
                                    {
                                        DocumentPath = doc.Name,
                                        Handle = entity.Handle.ToString(),
                                        SessionId = null
                                    });

                                    // Collect similar entities in current view for selection
                                    if (isCurrentSpace)
                                    {
                                        currentViewSimilarIds.Add(id);
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                tr.Commit();
            }

            // Save cross-layout selection (replacing previous document selection)
            try
            {
                var docName = Path.GetFileName(doc.Name);
                if (similarEntities.Count > 0)
                {
                    SelectionStorage.SaveSelection(similarEntities, docName);
                    ed.WriteMessage($"\nFound and stored {similarEntities.Count} similar entities across all layouts in document.\n");
                }
                else
                {
                    // Clear selection if no entities found
                    SelectionStorage.SaveSelection(new List<SelectionItem>(), docName);
                    ed.WriteMessage("\nNo similar entities found in document. Selection cleared.\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Update current view selection to include all similar entities
            if (currentViewSimilarIds.Count > 0)
            {
                // Validate ObjectIds before setting selection
                var validIds = new List<ObjectId>();
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var id in currentViewSimilarIds)
                    {
                        if (!id.IsNull && id.IsValid)
                        {
                            try
                            {
                                var testObj = tr.GetObject(id, OpenMode.ForRead, false);
                                if (testObj != null)
                                {
                                    validIds.Add(id);
                                }
                            }
                            catch
                            {
                                // Skip invalid ObjectIds
                            }
                        }
                    }
                    tr.Commit();
                }

                if (validIds.Count > 0)
                {
                    try
                    {
                        ed.SetImpliedSelection(validIds.ToArray());
                        ed.WriteMessage($"Selected {validIds.Count} similar entities in current view.\n");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError setting selection: {ex.Message}\n");
                    }
                }
                else
                {
                    ed.WriteMessage("No valid similar entities found in current view to select.\n");
                }
            }
            else
            {
                ed.WriteMessage("No similar entities found in current view to select.\n");
            }
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Get pickfirst selection from current document
            var selResult = ed.SelectImplied();

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects selected. Please select objects first, then run this command.\n");
                return;
            }

            var selectedIds = selResult.Value.GetObjectIds();
            var similarityKeys = new HashSet<SimilarityKey>();

            ed.WriteMessage($"\nSession Mode: Finding entities similar to {selectedIds.Length} selected object(s) across all open documents...\n");

            // Extract similarity keys from selected entities
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in selectedIds)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity != null)
                        {
                            var key = GetSimilarityKey(entity);
                            similarityKeys.Add(key);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                tr.Commit();
            }

            if (similarityKeys.Count == 0)
            {
                ed.WriteMessage("\nNo valid entities selected.\n");
                return;
            }

            // Find all similar entities across all open documents
            var similarEntities = new List<SelectionItem>();
            var currentViewSimilarIds = new List<ObjectId>();
            var docManager = AcadApp.DocumentManager;
            var currentDocPath = Path.GetFullPath(doc.Name);

            foreach (Autodesk.AutoCAD.ApplicationServices.Document openDoc in docManager)
            {
                string docPath = openDoc.Name;
                string docName = Path.GetFileName(docPath);
                bool isCurrentDoc = string.Equals(Path.GetFullPath(docPath), currentDocPath, StringComparison.OrdinalIgnoreCase);

                ed.WriteMessage($"\nScanning: {docName}...");

                try
                {
                    var refs = FindSimilarEntitiesInDocument(openDoc.Database, docPath, similarityKeys);
                    similarEntities.AddRange(refs);
                    ed.WriteMessage($" found {refs.Count} similar entities");

                    // For current document, also collect ObjectIds for current space/view
                    if (isCurrentDoc)
                    {
                        using (var tr = openDoc.Database.TransactionManager.StartTransaction())
                        {
                            foreach (var item in refs)
                            {
                                try
                                {
                                    var handle = Convert.ToInt64(item.Handle, 16);
                                    var objectId = openDoc.Database.GetObjectId(false, new Handle(handle), 0);

                                    if (!objectId.IsNull && objectId.IsValid)
                                    {
                                        var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                                        if (entity != null)
                                        {
                                            // Check if entity is in current space
                                            if (entity.BlockId == openDoc.Database.CurrentSpaceId)
                                            {
                                                currentViewSimilarIds.Add(objectId);
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
                }
            }

            // Clear all existing selections and save cross-document selection (session scope behavior)
            ClearAllStoredSelections();

            if (similarEntities.Count > 0)
            {
                try
                {
                    SelectionStorage.SaveSelection(similarEntities);
                    ed.WriteMessage($"\n\nFound and stored {similarEntities.Count} similar entities across all open documents.\n");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
                }
            }
            else
            {
                ed.WriteMessage("\n\nNo similar entities found in any open document.\n");
            }

            // Update current view selection to include all similar entities
            if (currentViewSimilarIds.Count > 0)
            {
                // Validate ObjectIds before setting selection
                var validIds = new List<ObjectId>();
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var id in currentViewSimilarIds)
                    {
                        if (!id.IsNull && id.IsValid)
                        {
                            try
                            {
                                var testObj = tr.GetObject(id, OpenMode.ForRead, false);
                                if (testObj != null)
                                {
                                    validIds.Add(id);
                                }
                            }
                            catch
                            {
                                // Skip invalid ObjectIds
                            }
                        }
                    }
                    tr.Commit();
                }

                if (validIds.Count > 0)
                {
                    try
                    {
                        ed.SetImpliedSelection(validIds.ToArray());
                        ed.WriteMessage($"Selected {validIds.Count} similar entities in current view.\n");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError setting selection: {ex.Message}\n");
                    }
                }
                else
                {
                    ed.WriteMessage("No valid similar entities found in current view to select.\n");
                }
            }
            else
            {
                ed.WriteMessage("No similar entities found in current view to select.\n");
            }
        }

        private static List<SelectionItem> FindSimilarEntitiesInDocument(Database db, string docPath, HashSet<SimilarityKey> similarityKeys)
        {
            var entities = new List<SelectionItem>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity != null)
                            {
                                var key = GetSimilarityKey(entity);
                                if (similarityKeys.Contains(key))
                                {
                                    entities.Add(new SelectionItem
                                    {
                                        DocumentPath = docPath,
                                        Handle = entity.Handle.ToString(),
                                        SessionId = null
                                    });
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                tr.Commit();
            }

            return entities;
        }

        private static SimilarityKey GetSimilarityKey(Entity entity)
        {
            var key = new SimilarityKey
            {
                EntityType = entity.GetType().Name,
                Layer = entity.Layer,
                ColorIndex = (short)entity.ColorIndex
            };

            // For BlockReferences, also match by block name
            if (entity is BlockReference blockRef)
            {
                using (var tr = blockRef.Database.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    key.BlockName = btr?.Name ?? "";
                }
            }

            return key;
        }

        private static void ClearAllStoredSelections()
        {
            try
            {
                // Clear all per-document selection files
                var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var selectionDir = Path.Combine(appDataPath, "autocad-ballet", "runtime", "selection");

                if (Directory.Exists(selectionDir))
                {
                    foreach (var file in Directory.GetFiles(selectionDir))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch
                        {
                            // Skip files that can't be deleted
                        }
                    }
                }

                // Also clear legacy global file for backward compatibility
                var legacyFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet", "selection");
                if (File.Exists(legacyFilePath))
                {
                    try
                    {
                        File.WriteAllLines(legacyFilePath, new string[0]);
                    }
                    catch
                    {
                        // Skip if can't clear legacy file
                    }
                }
            }
            catch
            {
                // If clearing fails, continue anyway - the save operation will overwrite
            }
        }
    }
}
