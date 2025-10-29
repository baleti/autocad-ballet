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
    public static class RemoveTags
    {
        /// <summary>
        /// Removes selected tags from all entities in the current view
        /// </summary>
        public static void ExecuteViewScope(Editor ed, Database db)
        {
            // Get all tags from the current view
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var tagToEntities = new Dictionary<string, List<ObjectId>>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in currentSpace)
                {
                    try
                    {
                        var tags = id.GetTags(db);

                        foreach (var tag in tags)
                        {
                            if (!tagCounts.ContainsKey(tag))
                            {
                                tagCounts[tag] = 0;
                                tagToEntities[tag] = new List<ObjectId>();
                            }

                            tagCounts[tag]++;
                            tagToEntities[tag].Add(id);
                        }
                    }
                    catch { }
                }

                tr.Commit();
            }

            if (tagCounts.Count == 0)
            {
                ed.WriteMessage("\nNo tags found in current view.\n");
                return;
            }

            // Show DataGrid with all tags from current view
            var entries = tagCounts.Select(kvp => new Dictionary<string, object>
            {
                { "TagName", kvp.Key },
                { "UsageCount", kvp.Value }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to remove from all entities in view...\n");

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

            // Remove the selected tags from all entities that have them
            var totalEntitiesModified = 0;
            foreach (var tagName in tagNamesToRemove)
            {
                if (tagToEntities.ContainsKey(tagName))
                {
                    var entities = tagToEntities[tagName];
                    RemoveTagsFromEntities(entities.ToArray(), new List<string> { tagName }, db, ed);
                    totalEntitiesModified += entities.Count;
                }
            }

            ed.WriteMessage($"\nSuccessfully removed {tagNamesToRemove.Count} tag(s) from {totalEntitiesModified} object(s) in current view.\n");
        }

        /// <summary>
        /// Removes selected tags from all entities in current document
        /// </summary>
        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;

            // Get all tags from the entire document
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var tagToEntities = new Dictionary<string, List<ObjectId>>(StringComparer.OrdinalIgnoreCase);

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
                            var tags = id.GetTags(db);

                            foreach (var tag in tags)
                            {
                                if (!tagCounts.ContainsKey(tag))
                                {
                                    tagCounts[tag] = 0;
                                    tagToEntities[tag] = new List<ObjectId>();
                                }

                                tagCounts[tag]++;
                                if (!tagToEntities[tag].Contains(id))
                                {
                                    tagToEntities[tag].Add(id);
                                }
                            }
                        }
                        catch { }
                    }
                }

                tr.Commit();
            }

            if (tagCounts.Count == 0)
            {
                ed.WriteMessage("\nNo tags found in document.\n");
                return;
            }

            // Show DataGrid with all tags from document
            var entries = tagCounts.Select(kvp => new Dictionary<string, object>
            {
                { "TagName", kvp.Key },
                { "UsageCount", kvp.Value }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to remove from all entities in document...\n");

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

            // Remove the selected tags from all entities that have them
            var totalEntitiesModified = 0;
            foreach (var tagName in tagNamesToRemove)
            {
                if (tagToEntities.ContainsKey(tagName))
                {
                    var entities = tagToEntities[tagName];
                    RemoveTagsFromEntities(entities.ToArray(), new List<string> { tagName }, db, ed);
                    totalEntitiesModified += entities.Count;
                }
            }

            ed.WriteMessage($"\nSuccessfully removed {tagNamesToRemove.Count} tag(s) from {totalEntitiesModified} object(s) in document.\n");
        }

        /// <summary>
        /// Removes selected tags from all entities across all open documents
        /// </summary>
        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var docManager = AcadApp.DocumentManager;

            // Get all tags from all open documents
            var allTagsDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var tagToDocuments = new Dictionary<string, Dictionary<string, List<ObjectId>>>(StringComparer.OrdinalIgnoreCase);

            foreach (Document openDoc in docManager)
            {
                var db = openDoc.Database;
                var docPath = openDoc.Name;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
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
                                    var tags = id.GetTags(db);

                                    foreach (var tag in tags)
                                    {
                                        if (!allTagsDict.ContainsKey(tag))
                                        {
                                            allTagsDict[tag] = 0;
                                            tagToDocuments[tag] = new Dictionary<string, List<ObjectId>>();
                                        }

                                        allTagsDict[tag]++;

                                        if (!tagToDocuments[tag].ContainsKey(docPath))
                                        {
                                            tagToDocuments[tag][docPath] = new List<ObjectId>();
                                        }

                                        if (!tagToDocuments[tag][docPath].Contains(id))
                                        {
                                            tagToDocuments[tag][docPath].Add(id);
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

            if (allTagsDict.Count == 0)
            {
                ed.WriteMessage("\nNo tags found across open documents.\n");
                return;
            }

            // Show DataGrid with all tags from all documents
            var entries = allTagsDict.Select(kvp => new Dictionary<string, object>
            {
                { "TagName", kvp.Key },
                { "UsageCount", kvp.Value }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect tags to remove from all entities across open documents...\n");

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

            // Remove the selected tags from all entities across all documents
            var totalEntitiesModified = 0;
            foreach (var tagName in tagNamesToRemove)
            {
                if (tagToDocuments.ContainsKey(tagName))
                {
                    foreach (var kvp in tagToDocuments[tagName])
                    {
                        var extDocPath = kvp.Key;
                        var entities = kvp.Value;

                        // Find the document
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
                            RemoveTagsFromEntities(entities.ToArray(), new List<string> { tagName }, extDoc.Database, ed);
                            totalEntitiesModified += entities.Count;
                            ed.WriteMessage($"\nRemoved tag '{tagName}' from {entities.Count} object(s) in {Path.GetFileName(extDocPath)}.\n");
                        }
                    }
                }
            }

            ed.WriteMessage($"\nSuccessfully removed {tagNamesToRemove.Count} tag(s) from {totalEntitiesModified} object(s) across open documents.\n");
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
