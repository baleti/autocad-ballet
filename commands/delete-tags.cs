using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class DeleteTags
    {
        public static void ExecuteViewScope(Editor ed, Database db)
        {
            ed.WriteMessage("\nView Mode: Gathering tags from current view/layout...\n");

            var tagGroups = new Dictionary<string, List<ObjectId>>();
            var spaceName = GetCurrentSpaceName(db);

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
                            if (!tagGroups.ContainsKey(tag))
                                tagGroups[tag] = new List<ObjectId>();

                            tagGroups[tag].Add(id);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                tr.Commit();
            }

            ShowDeleteDialogForView(ed, db, tagGroups, spaceName);
        }

        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            ed.WriteMessage("\nDocument Mode: Gathering tags from entire document...\n");

            var tagGroups = new Dictionary<string, List<TagEntityReference>>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Get all layouts
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead);
                            var tags = id.GetTags(db);

                            foreach (var tag in tags)
                            {
                                if (!tagGroups.ContainsKey(tag))
                                    tagGroups[tag] = new List<TagEntityReference>();

                                tagGroups[tag].Add(new TagEntityReference
                                {
                                    DocumentPath = doc.Name,
                                    DocumentName = Path.GetFileName(doc.Name),
                                    Handle = entity.Handle.ToString(),
                                    Tag = tag,
                                    SpaceName = spaceName
                                });
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

            ShowDeleteDialogForDocument(ed, db, tagGroups, doc.Name);
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            ed.WriteMessage("\nApplication Mode: Gathering tags from all open documents...\n");

            var docManager = AcadApp.DocumentManager;
            var allReferences = new List<TagEntityReference>();
            var tagGroups = new Dictionary<string, List<TagEntityReference>>();

            foreach (Autodesk.AutoCAD.ApplicationServices.Document doc in docManager)
            {
                string docPath = doc.Name;
                string docName = Path.GetFileName(docPath);

                ed.WriteMessage($"\nScanning: {docName}...");

                try
                {
                    var refs = GatherTagReferencesFromDocument(doc.Database, docPath, docName);
                    allReferences.AddRange(refs);

                    foreach (var tagRef in refs)
                    {
                        if (!tagGroups.ContainsKey(tagRef.Tag))
                            tagGroups[tagRef.Tag] = new List<TagEntityReference>();

                        tagGroups[tagRef.Tag].Add(tagRef);
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
                }
            }

            ShowDeleteDialogForApplication(ed, tagGroups);
        }

        private static void ShowDeleteDialogForView(Editor ed, Database db, Dictionary<string, List<ObjectId>> tagGroups, string spaceName)
        {
            if (tagGroups.Count == 0)
            {
                ed.WriteMessage($"\nNo tagged entities found in {spaceName}.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var tag in tagGroups.OrderBy(t => t.Key))
            {
                entries.Add(new Dictionary<string, object>
                {
                    { "Tag", tag.Key },
                    { "Count", tag.Value.Count },
                    { "Space", spaceName }
                });
            }

            var propertyNames = new List<string> { "Tag", "Count", "Space" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected for deletion.\n");
                return;
            }

            var selectedTags = new HashSet<string>(chosenRows.Select(row => row["Tag"].ToString()));
            var affectedEntities = new HashSet<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var tag in selectedTags)
                {
                    if (tagGroups.ContainsKey(tag))
                    {
                        foreach (var entityId in tagGroups[tag])
                        {
                            try
                            {
                                entityId.RemoveTags(db, tag);
                                affectedEntities.Add(entityId);
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nError removing tag '{tag}' from entity {entityId.Handle}: {ex.Message}\n");
                            }
                        }
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nRemoved {selectedTags.Count} tags from {affectedEntities.Count} entities in {spaceName}.\n");
        }

        private static void ShowDeleteDialogForDocument(Editor ed, Database db, Dictionary<string, List<TagEntityReference>> tagGroups, string documentName)
        {
            if (tagGroups.Count == 0)
            {
                ed.WriteMessage($"\nNo tagged entities found in document.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var tag in tagGroups.OrderBy(t => t.Key))
            {
                var docCounts = tag.Value.GroupBy(e => e.SpaceName)
                                         .Select(g => $"{g.Key}: {g.Count()}")
                                         .ToList();

                entries.Add(new Dictionary<string, object>
                {
                    { "Tag", tag.Key },
                    { "Total Count", tag.Value.Count },
                    { "Layouts", string.Join(", ", docCounts) }
                });
            }

            var propertyNames = new List<string> { "Tag", "Total Count", "Layouts" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected for deletion.\n");
                return;
            }

            var selectedTagNames = new HashSet<string>(chosenRows.Select(row => row["Tag"].ToString()));
            var selectedEntities = new List<TagEntityReference>();

            foreach (var tagName in selectedTagNames)
            {
                if (tagGroups.ContainsKey(tagName))
                {
                    selectedEntities.AddRange(tagGroups[tagName]);
                }
            }

            // Remove duplicates based on handle (entities can have multiple tags)
            selectedEntities = selectedEntities
                .GroupBy(e => e.Handle)
                .Select(g => g.First())
                .ToList();

            var affectedCount = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var entityRef in selectedEntities)
                {
                    try
                    {
                        var handle = Convert.ToInt64(entityRef.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (!objectId.IsNull && objectId.IsValid)
                        {
                            // Remove all selected tags from this entity
                            objectId.RemoveTags(db, selectedTagNames.ToArray());
                            affectedCount++;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError removing tags from entity {entityRef.Handle}: {ex.Message}\n");
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nRemoved {selectedTagNames.Count} tags from {affectedCount} entities in document.\n");
        }

        private static void ShowDeleteDialogForApplication(Editor ed, Dictionary<string, List<TagEntityReference>> tagGroups)
        {
            if (tagGroups.Count == 0)
            {
                ed.WriteMessage("\nNo tagged entities found across open documents.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var tag in tagGroups.OrderBy(t => t.Key))
            {
                var docCounts = tag.Value.GroupBy(e => e.DocumentName)
                                         .Select(g => $"{g.Key}: {g.Count()}")
                                         .ToList();

                entries.Add(new Dictionary<string, object>
                {
                    { "Tag", tag.Key },
                    { "Total Count", tag.Value.Count },
                    { "Documents", string.Join(", ", docCounts) }
                });
            }

            var propertyNames = new List<string> { "Tag", "Total Count", "Documents" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected for deletion.\n");
                return;
            }

            var selectedTagNames = new HashSet<string>(chosenRows.Select(row => row["Tag"].ToString()));
            var selectedEntities = new List<TagEntityReference>();

            foreach (var tagName in selectedTagNames)
            {
                if (tagGroups.ContainsKey(tagName))
                {
                    selectedEntities.AddRange(tagGroups[tagName]);
                }
            }

            // Group by document and remove tags from each document
            var entitiesByDocument = selectedEntities
                .GroupBy(e => e.DocumentPath)
                .ToDictionary(g => g.Key, g => g.ToList());

            var totalAffected = 0;

            foreach (var docGroup in entitiesByDocument)
            {
                var docPath = docGroup.Key;
                var docName = Path.GetFileName(docPath);

                // Find the document in the open documents
                Autodesk.AutoCAD.ApplicationServices.Document targetDoc = null;
                foreach (Autodesk.AutoCAD.ApplicationServices.Document doc in AcadApp.DocumentManager)
                {
                    if (string.Equals(Path.GetFullPath(doc.Name), Path.GetFullPath(docPath), StringComparison.OrdinalIgnoreCase))
                    {
                        targetDoc = doc;
                        break;
                    }
                }

                if (targetDoc == null)
                {
                    ed.WriteMessage($"\nWarning: Document {docName} not found in open documents.\n");
                    continue;
                }

                var db = targetDoc.Database;
                var affectedCount = 0;

                // Get unique entities in this document
                var uniqueEntities = docGroup.Value
                    .GroupBy(e => e.Handle)
                    .Select(g => g.First())
                    .ToList();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var entityRef in uniqueEntities)
                    {
                        try
                        {
                            var handle = Convert.ToInt64(entityRef.Handle, 16);
                            var objectId = db.GetObjectId(false, new Handle(handle), 0);

                            if (!objectId.IsNull && objectId.IsValid)
                            {
                                // Remove all selected tags from this entity
                                objectId.RemoveTags(db, selectedTagNames.ToArray());
                                affectedCount++;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError removing tags from entity {entityRef.Handle} in {docName}: {ex.Message}\n");
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\n  {docName}: Removed tags from {affectedCount} entities\n");
                totalAffected += affectedCount;
            }

            ed.WriteMessage($"\nTotal: Removed {selectedTagNames.Count} tags from {totalAffected} entities across {entitiesByDocument.Count} documents.\n");
        }

        private static List<TagEntityReference> GatherTagReferencesFromDocument(Database db, string docPath, string docName)
        {
            var references = new List<TagEntityReference>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead);
                            var tags = id.GetTags(db);

                            foreach (var tag in tags)
                            {
                                references.Add(new TagEntityReference
                                {
                                    DocumentPath = docPath,
                                    DocumentName = docName,
                                    Handle = entity.Handle.ToString(),
                                    Tag = tag,
                                    SpaceName = spaceName
                                });
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

            return references;
        }

        private static string GetCurrentSpaceName(Database db)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    if (db.TileMode)
                    {
                        return "Model";
                    }
                    else
                    {
                        var layoutMgr = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
                        return layoutMgr.CurrentLayout;
                    }
                }
            }
            catch
            {
                return "Unknown";
            }
        }
    }
}
