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
    public static class RemoveTagsFromSelected
    {
        /// <summary>
        /// Removes selected tags from entities in the current view
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
                selectionOpts.MessageForAdding = "\nSelect objects to remove tags from: ";
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

            // Get all tags from the selected entities
            var tagsDict = GetTagsFromEntities(selectedIds, db);

            if (tagsDict.Count == 0)
            {
                ed.WriteMessage("\nNo tags found on selected objects.\n");
                return;
            }

            // Show DataGrid with tags from selected entities
            var entries = tagsDict.Select(kvp => new Dictionary<string, object>
            {
                { "TagName", kvp.Key },
                { "UsageCount", kvp.Value }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to remove from selected objects...\n");

            var selectedTagsToRemove = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: false);

            if (selectedTagsToRemove == null || selectedTagsToRemove.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract tag names to remove
            var tagNamesToRemove = new List<string>();
            foreach (var tag in selectedTagsToRemove)
            {
                if (tag.ContainsKey("TagName"))
                {
                    tagNamesToRemove.Add((string)tag["TagName"]);
                }
            }

            if (tagNamesToRemove.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags selected.\n");
                return;
            }

            // Remove the selected tags from the selected entities
            RemoveTagsFromEntities(selectedIds, tagNamesToRemove, db, ed);

            ed.WriteMessage($"\nSuccessfully removed {tagNamesToRemove.Count} tag(s) from {selectedIds.Length} object(s).\n");
        }

        /// <summary>
        /// Removes selected tags from entities in stored selection (current document)
        /// </summary>
        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var docName = Path.GetFileName(doc.Name);

            var storedSelection = SelectionStorage.LoadSelection(docName);
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage($"\nNo stored selection found for document. Use 'select-by-tags-in-document' first.\n");
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

            // Convert to ObjectIds
            var objectIds = new List<ObjectId>();
            foreach (var item in storedSelection)
            {
                try
                {
                    var handle = Convert.ToInt64(item.Handle, 16);
                    var objectId = db.GetObjectId(false, new Handle(handle), 0);
                    if (objectId != ObjectId.Null)
                    {
                        objectIds.Add(objectId);
                    }
                }
                catch { }
            }

            if (objectIds.Count == 0)
            {
                ed.WriteMessage("\nNo valid objects found in stored selection.\n");
                return;
            }

            // Get all tags from the stored selection entities
            var tagsDict = GetTagsFromEntities(objectIds.ToArray(), db);

            if (tagsDict.Count == 0)
            {
                ed.WriteMessage("\nNo tags found on selected objects.\n");
                return;
            }

            // Show DataGrid with tags from selected entities
            var entries = tagsDict.Select(kvp => new Dictionary<string, object>
            {
                { "TagName", kvp.Key },
                { "UsageCount", kvp.Value }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to remove from selected objects...\n");

            var selectedTagsToRemove = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: false);

            if (selectedTagsToRemove == null || selectedTagsToRemove.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract tag names to remove
            var tagNamesToRemove = new List<string>();
            foreach (var tag in selectedTagsToRemove)
            {
                if (tag.ContainsKey("TagName"))
                {
                    tagNamesToRemove.Add((string)tag["TagName"]);
                }
            }

            if (tagNamesToRemove.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags selected.\n");
                return;
            }

            // Remove the selected tags from the entities
            RemoveTagsFromEntities(objectIds.ToArray(), tagNamesToRemove, db, ed);

            ed.WriteMessage($"\nSuccessfully removed {tagNamesToRemove.Count} tag(s) from {objectIds.Count} object(s).\n");
        }

        /// <summary>
        /// Removes selected tags from entities in stored selection (all open documents)
        /// </summary>
        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            var storedSelection = SelectionStorage.LoadSelectionFromOpenDocuments();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use 'select-by-tags-in-session' first.\n");
                return;
            }

            // Get all tags from all selected entities across open documents
            var docManager = AcadApp.DocumentManager;
            var allTagsDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Separate current document from external documents
            var currentDocEntities = new List<ObjectId>();
            var externalDocuments = new Dictionary<string, List<SelectionItem>>();

            foreach (var item in storedSelection)
            {
                try
                {
                    var itemPath = Path.GetFullPath(item.DocumentPath);
                    var currentPath = Path.GetFullPath(doc.Name);

                    if (string.Equals(itemPath, currentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);
                        if (objectId != ObjectId.Null)
                        {
                            currentDocEntities.Add(objectId);
                        }
                    }
                    else
                    {
                        if (!externalDocuments.ContainsKey(item.DocumentPath))
                        {
                            externalDocuments[item.DocumentPath] = new List<SelectionItem>();
                        }
                        externalDocuments[item.DocumentPath].Add(item);
                    }
                }
                catch { }
            }

            // Get tags from current document entities
            if (currentDocEntities.Count > 0)
            {
                var currentTags = GetTagsFromEntities(currentDocEntities.ToArray(), db);
                foreach (var kvp in currentTags)
                {
                    if (allTagsDict.ContainsKey(kvp.Key))
                        allTagsDict[kvp.Key] += kvp.Value;
                    else
                        allTagsDict[kvp.Key] = kvp.Value;
                }
            }

            // Get tags from external documents
            foreach (var kvp in externalDocuments)
            {
                var extDocPath = kvp.Key;
                var extItems = kvp.Value;

                // Find already open document
                Document extDoc = null;
                foreach (Document openDoc in docManager)
                {
                    try
                    {
                        if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(extDocPath),
                            StringComparison.OrdinalIgnoreCase))
                        {
                            extDoc = openDoc;
                            break;
                        }
                    }
                    catch { }
                }

                if (extDoc != null)
                {
                    var extDb = extDoc.Database;
                    var extObjectIds = new List<ObjectId>();

                    foreach (var item in extItems)
                    {
                        try
                        {
                            var handle = Convert.ToInt64(item.Handle, 16);
                            var objectId = extDb.GetObjectId(false, new Handle(handle), 0);
                            if (objectId != ObjectId.Null)
                            {
                                extObjectIds.Add(objectId);
                            }
                        }
                        catch { }
                    }

                    if (extObjectIds.Count > 0)
                    {
                        var extTags = GetTagsFromEntities(extObjectIds.ToArray(), extDb);
                        foreach (var tagKvp in extTags)
                        {
                            if (allTagsDict.ContainsKey(tagKvp.Key))
                                allTagsDict[tagKvp.Key] += tagKvp.Value;
                            else
                                allTagsDict[tagKvp.Key] = tagKvp.Value;
                        }
                    }
                }
            }

            if (allTagsDict.Count == 0)
            {
                ed.WriteMessage("\nNo tags found on selected objects.\n");
                return;
            }

            // Show DataGrid with tags from all selected entities
            var entries = allTagsDict.Select(kvp => new Dictionary<string, object>
            {
                { "TagName", kvp.Key },
                { "UsageCount", kvp.Value }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to remove from selected objects...\n");

            var selectedTagsToRemove = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: false);

            if (selectedTagsToRemove == null || selectedTagsToRemove.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract tag names to remove
            var tagNamesToRemove = new List<string>();
            foreach (var tag in selectedTagsToRemove)
            {
                if (tag.ContainsKey("TagName"))
                {
                    tagNamesToRemove.Add((string)tag["TagName"]);
                }
            }

            if (tagNamesToRemove.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags selected.\n");
                return;
            }

            // Remove tags from current document entities
            if (currentDocEntities.Count > 0)
            {
                RemoveTagsFromEntities(currentDocEntities.ToArray(), tagNamesToRemove, db, ed);
                ed.WriteMessage($"\nRemoved tags from {currentDocEntities.Count} object(s) in current document.\n");
            }

            // Remove tags from external document entities
            foreach (var kvp in externalDocuments)
            {
                var extDocPath = kvp.Key;
                var extItems = kvp.Value;

                // Find already open document
                Document extDoc = null;
                foreach (Document openDoc in docManager)
                {
                    try
                    {
                        if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(extDocPath),
                            StringComparison.OrdinalIgnoreCase))
                        {
                            extDoc = openDoc;
                            break;
                        }
                    }
                    catch { }
                }

                if (extDoc != null)
                {
                    var extDb = extDoc.Database;
                    var extObjectIds = new List<ObjectId>();

                    foreach (var item in extItems)
                    {
                        try
                        {
                            var handle = Convert.ToInt64(item.Handle, 16);
                            var objectId = extDb.GetObjectId(false, new Handle(handle), 0);
                            if (objectId != ObjectId.Null)
                            {
                                extObjectIds.Add(objectId);
                            }
                        }
                        catch { }
                    }

                    if (extObjectIds.Count > 0)
                    {
                        RemoveTagsFromEntities(extObjectIds.ToArray(), tagNamesToRemove, extDb, ed);
                        ed.WriteMessage($"\nRemoved tags from {extObjectIds.Count} object(s) in {Path.GetFileName(extDocPath)}.\n");
                    }
                }
                else
                {
                    ed.WriteMessage($"\nSkipping {Path.GetFileName(extDocPath)} - document not open.\n");
                }
            }

            ed.WriteMessage("\nTag removal operation completed.\n");
        }

        /// <summary>
        /// Helper method to get all tags from a collection of entities
        /// </summary>
        private static Dictionary<string, int> GetTagsFromEntities(ObjectId[] objectIds, Database db)
        {
            var tagsDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var objectId in objectIds)
            {
                try
                {
                    var tags = objectId.GetTags(db);
                    foreach (var tag in tags)
                    {
                        if (tagsDict.ContainsKey(tag))
                            tagsDict[tag]++;
                        else
                            tagsDict[tag] = 1;
                    }
                }
                catch { }
            }

            return tagsDict;
        }

        /// <summary>
        /// Helper method to remove tags from entities
        /// </summary>
        private static void RemoveTagsFromEntities(ObjectId[] objectIds, List<string> tagNamesToRemove,
            Database db, Editor ed)
        {
            foreach (var objectId in objectIds)
            {
                try
                {
                    objectId.RemoveTags(db, tagNamesToRemove.ToArray());
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError removing tags from entity: {ex.Message}\n");
                }
            }
        }
    }
}
