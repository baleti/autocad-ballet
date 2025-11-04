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
    /// Selects entities that have at least one of the tags from the currently selected entities
    /// Shows a dialog to let user pick which tags to use for selection
    /// </summary>
    public static class SelectByTagsOfSelected
    {
        /// <summary>
        /// View scope: Select parents by tags from pickfirst set
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
                selectionOpts.MessageForAdding = "\nSelect objects to extract tags from: ";
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

            // Collect all unique tags from selected entities with counts
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var objId in selectedIds)
            {
                try
                {
                    var tags = objId.GetTags(db);
                    foreach (var tag in tags)
                    {
                        if (tagCounts.ContainsKey(tag))
                            tagCounts[tag]++;
                        else
                            tagCounts[tag] = 1;
                    }
                }
                catch { }
            }

            if (tagCounts.Count == 0)
            {
                ed.WriteMessage("\nSelected entities have no tags.\n");
                return;
            }

            // Show DataGrid with tags for user selection
            var entries = tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select(kvp => new Dictionary<string, object>
                {
                    { "TagName", kvp.Key },
                    { "UsageCount", kvp.Value }
                })
                .ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to find entities with...\n");

            var selectedTags = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: false, commandName: "select-by-tags-of-selected-in-view");

            if (selectedTags == null || selectedTags.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract selected tag names
            var tagNames = selectedTags
                .Where(tag => tag.ContainsKey("TagName"))
                .Select(tag => (string)tag["TagName"])
                .ToList();

            if (tagNames.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags selected.\n");
                return;
            }

            // Find all entities in current space with at least one of the selected tags
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

                            // Check if entity has at least one of the selected tags
                            if (tags.Any(tag => tagNames.Contains(tag, StringComparer.OrdinalIgnoreCase)))
                            {
                                matchingEntities.Add(entityId);
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
                ed.WriteMessage("\nNo entities found with selected tags.\n");
                return;
            }

            // Set implied selection for view scope
            ed.SetImpliedSelection(matchingEntities.ToArray());

            ed.WriteMessage($"\nSelected {matchingEntities.Count} entity(ies) with at least one of {tagNames.Count} tag(s).\n");
        }

        /// <summary>
        /// Document scope: Select parents by tags from stored selection
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

            // Convert to ObjectIds and collect tags with counts
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in storedSelection)
            {
                try
                {
                    var handle = Convert.ToInt64(item.Handle, 16);
                    var objectId = db.GetObjectId(false, new Handle(handle), 0);
                    if (objectId != ObjectId.Null)
                    {
                        var tags = objectId.GetTags(db);
                        foreach (var tag in tags)
                        {
                            if (tagCounts.ContainsKey(tag))
                                tagCounts[tag]++;
                            else
                                tagCounts[tag] = 1;
                        }
                    }
                }
                catch { }
            }

            if (tagCounts.Count == 0)
            {
                ed.WriteMessage("\nStored selection has no entities with tags.\n");
                return;
            }

            // Show DataGrid with tags for user selection
            var entries = tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select(kvp => new Dictionary<string, object>
                {
                    { "TagName", kvp.Key },
                    { "UsageCount", kvp.Value }
                })
                .ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to find entities with...\n");

            var selectedTags = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: false, commandName: "select-by-tags-of-selected-in-document");

            if (selectedTags == null || selectedTags.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract selected tag names
            var tagNames = selectedTags
                .Where(tag => tag.ContainsKey("TagName"))
                .Select(tag => (string)tag["TagName"])
                .ToList();

            if (tagNames.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags selected.\n");
                return;
            }

            // Find all entities in all layouts with at least one of the selected tags
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

                                // Check if entity has at least one of the selected tags
                                if (tags.Any(tag => tagNames.Contains(tag, StringComparer.OrdinalIgnoreCase)))
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
                ed.WriteMessage("\nNo entities found with selected tags.\n");
                return;
            }

            // Save to stored selection
            SelectionStorage.SaveSelection(matchingSelectionItems, docName);

            ed.WriteMessage($"\nSelected {matchingSelectionItems.Count} entity(ies) with at least one of {tagNames.Count} tag(s) in document.\n");
        }

        /// <summary>
        /// Session scope: Select parents by tags from stored selection across all documents
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

            // Collect tags from all selected entities across documents
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
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
                            foreach (var tag in tags)
                            {
                                if (tagCounts.ContainsKey(tag))
                                    tagCounts[tag]++;
                                else
                                    tagCounts[tag] = 1;
                            }
                        }
                    }
                    catch { }
                }
            }

            if (tagCounts.Count == 0)
            {
                ed.WriteMessage("\nStored selection has no entities with tags.\n");
                return;
            }

            // Show DataGrid with tags for user selection
            var entries = tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select(kvp => new Dictionary<string, object>
                {
                    { "TagName", kvp.Key },
                    { "UsageCount", kvp.Value }
                })
                .ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to find entities with...\n");

            var selectedTags = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: false, commandName: "select-by-tags-of-selected-in-session");

            if (selectedTags == null || selectedTags.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract selected tag names
            var tagNames = selectedTags
                .Where(tag => tag.ContainsKey("TagName"))
                .Select(tag => (string)tag["TagName"])
                .ToList();

            if (tagNames.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags selected.\n");
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

                                    // Check if entity has at least one of the selected tags
                                    if (tags.Any(tag => tagNames.Contains(tag, StringComparer.OrdinalIgnoreCase)))
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
                ed.WriteMessage("\nNo entities found with selected tags.\n");
                return;
            }

            // Group by document for reporting
            var groupedByDoc = matchingSelectionItems.GroupBy(item => Path.GetFileName(item.DocumentPath));

            // Save to stored selection for all documents
            foreach (var docGroup in groupedByDoc)
            {
                SelectionStorage.SaveSelection(docGroup.ToList(), docGroup.Key);
            }

            ed.WriteMessage($"\nSelected {matchingSelectionItems.Count} entity(ies) with at least one of {tagNames.Count} tag(s) across {groupedByDoc.Count()} document(s).\n");
        }
    }
}
