using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    /// <summary>
    /// Selects entities that have exactly the same tags as the currently selected entities (siblings)
    /// </summary>
    public static class SelectBySiblingTagsOfSelected
    {
        /// <summary>
        /// View scope: Select siblings by tags from pickfirst set
        /// </summary>
        public static void ExecuteViewScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Get pickfirst set or prompt for selection
            var selResult = ed.SelectImplied();

            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to find siblings of: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects selected.\n");
                return;
            }

            var selectedIds = selResult.Value.GetObjectIds();

            // Collect all unique tag sets from selected entities
            var tagSets = new List<HashSet<string>>();
            foreach (var objId in selectedIds)
            {
                try
                {
                    var tags = objId.GetTags(db);
                    if (tags.Count > 0)
                    {
                        tagSets.Add(new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase));
                    }
                }
                catch { }
            }

            if (tagSets.Count == 0)
            {
                ed.WriteMessage("\nSelected entities have no tags.\n");
                return;
            }

            // Find all entities in current space with matching tag sets
            var matchingEntities = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                    foreach (ObjectId entityId in currentSpace)
                    {
                        try
                        {
                            var tags = entityId.GetTags(db);
                            if (tags.Count == 0)
                                continue;

                            var entityTagSet = new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase);

                            // Check if entity has exactly the same tags as any selected entity
                            foreach (var targetTagSet in tagSets)
                            {
                                if (entityTagSet.SetEquals(targetTagSet))
                                {
                                    matchingEntities.Add(entityId);
                                    break;
                                }
                            }
                        }
                        catch { }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            if (matchingEntities.Count == 0)
            {
                ed.WriteMessage("\nNo sibling entities found with matching tags.\n");
                return;
            }

            // Set implied selection for view scope
            ed.SetImpliedSelection(matchingEntities.ToArray());

            ed.WriteMessage($"\nSelected {matchingEntities.Count} sibling entity(ies) with matching tags.\n");
        }

        /// <summary>
        /// Document scope: Select siblings by tags from stored selection
        /// </summary>
        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var docName = Path.GetFileName(doc.Name);

            var storedSelection = SelectionStorage.LoadSelection(docName);
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage($"\nNo stored selection found for document. Use 'select-by-categories-in-document' first.\n");
                return;
            }

            // Filter to current document only
            var currentDocPath = Path.GetFullPath(doc.Name);
            storedSelection = storedSelection.Where(item =>
            {
                try
                {
                    var itemPath = Path.GetFullPath(item.DocumentPath);
                    return string.Equals(itemPath, currentDocPath, StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return string.Equals(item.DocumentPath, doc.Name, StringComparison.OrdinalIgnoreCase);
                }
            }).ToList();

            if (storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection in current document.\n");
                return;
            }

            // Convert to ObjectIds and collect tag sets
            var tagSets = new List<HashSet<string>>();
            foreach (var item in storedSelection)
            {
                try
                {
                    var handle = Convert.ToInt64(item.Handle, 16);
                    var objectId = db.GetObjectId(false, new Handle(handle), 0);
                    if (objectId != ObjectId.Null)
                    {
                        var tags = objectId.GetTags(db);
                        if (tags.Count > 0)
                        {
                            tagSets.Add(new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase));
                        }
                    }
                }
                catch { }
            }

            if (tagSets.Count == 0)
            {
                ed.WriteMessage("\nStored selection has no entities with tags.\n");
                return;
            }

            // Find all entities in all layouts with matching tag sets
            var matchingSelectionItems = new List<SelectionItem>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                        foreach (ObjectId entityId in btr)
                        {
                            try
                            {
                                var tags = entityId.GetTags(db);
                                if (tags.Count == 0)
                                    continue;

                                var entityTagSet = new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase);

                                // Check if entity has exactly the same tags as any selected entity
                                foreach (var targetTagSet in tagSets)
                                {
                                    if (entityTagSet.SetEquals(targetTagSet))
                                    {
                                        var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                                        if (entity != null)
                                        {
                                            matchingSelectionItems.Add(new SelectionItem
                                            {
                                                DocumentPath = doc.Name,
                                                Handle = entity.Handle.ToString(),
                                                SessionId = doc.GetHashCode().ToString()
                                            });
                                        }
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            if (matchingSelectionItems.Count == 0)
            {
                ed.WriteMessage("\nNo sibling entities found with matching tags.\n");
                return;
            }

            // Save to stored selection
            SelectionStorage.SaveSelection(matchingSelectionItems, docName);

            ed.WriteMessage($"\nSelected {matchingSelectionItems.Count} sibling entity(ies) with matching tags in document.\n");
        }

        /// <summary>
        /// Session scope: Select siblings by tags from stored selection across all documents
        /// </summary>
        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            var storedSelection = SelectionStorage.LoadSelectionFromOpenDocuments();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use 'select-by-categories-in-session' first.\n");
                return;
            }

            // Collect tag sets from all selected entities across documents
            var tagSets = new List<HashSet<string>>();
            var docs = AcadApp.DocumentManager;

            foreach (var item in storedSelection)
            {
                // Find the document
                Document itemDoc = null;
                Database itemDb = null;

                foreach (Document openDoc in docs)
                {
                    try
                    {
                        if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(item.DocumentPath),
                            StringComparison.OrdinalIgnoreCase))
                        {
                            itemDoc = openDoc;
                            itemDb = openDoc.Database;
                            break;
                        }
                    }
                    catch { }
                }

                if (itemDb != null)
                {
                    try
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = itemDb.GetObjectId(false, new Handle(handle), 0);
                        if (objectId != ObjectId.Null)
                        {
                            var tags = objectId.GetTags(itemDb);
                            if (tags.Count > 0)
                            {
                                tagSets.Add(new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase));
                            }
                        }
                    }
                    catch { }
                }
            }

            if (tagSets.Count == 0)
            {
                ed.WriteMessage("\nStored selection has no entities with tags.\n");
                return;
            }

            // Find matching entities in all open documents
            var matchingSelectionItems = new List<SelectionItem>();

            foreach (Document openDoc in docs)
            {
                var openDb = openDoc.Database;

                using (var tr = openDb.TransactionManager.StartTransaction())
                {
                    try
                    {
                        var layoutDict = (DBDictionary)tr.GetObject(openDb.LayoutDictionaryId, OpenMode.ForRead);

                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                            foreach (ObjectId entityId in btr)
                            {
                                try
                                {
                                    var tags = entityId.GetTags(openDb);
                                    if (tags.Count == 0)
                                        continue;

                                    var entityTagSet = new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase);

                                    // Check if entity has exactly the same tags as any selected entity
                                    foreach (var targetTagSet in tagSets)
                                    {
                                        if (entityTagSet.SetEquals(targetTagSet))
                                        {
                                            var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                                            if (entity != null)
                                            {
                                                matchingSelectionItems.Add(new SelectionItem
                                                {
                                                    DocumentPath = openDoc.Name,
                                                    Handle = entity.Handle.ToString(),
                                                    SessionId = openDoc.GetHashCode().ToString()
                                                });
                                            }
                                            break;
                                        }
                                    }
                                }
                                catch { }
                            }
                        }

                        tr.Commit();
                    }
                    catch
                    {
                        tr.Abort();
                    }
                }
            }

            if (matchingSelectionItems.Count == 0)
            {
                ed.WriteMessage("\nNo sibling entities found with matching tags.\n");
                return;
            }

            // Group by document for reporting
            var groupedByDoc = matchingSelectionItems.GroupBy(item => Path.GetFileName(item.DocumentPath));

            // Save to stored selection for all documents
            foreach (var docGroup in groupedByDoc)
            {
                SelectionStorage.SaveSelection(docGroup.ToList(), docGroup.Key);
            }

            ed.WriteMessage($"\nSelected {matchingSelectionItems.Count} sibling entity(ies) with matching tags across {groupedByDoc.Count()} document(s).\n");
        }
    }
}
