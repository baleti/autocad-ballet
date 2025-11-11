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
    // Simple class to store entity references with tags
    public class TagEntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string Tag { get; set; }
        public string SpaceName { get; set; }
        public string BlockPath { get; set; }  // Empty for top-level entities, "BlockA > BlockB" for nested
        public bool IsInBlock { get; set; }     // True if entity is within a block definition
    }

    public static class SelectByTags
    {
        public static void ExecuteViewScope(Editor ed, Database db)
        {
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();
            string modeMsg = searchInBlocks ? " (including blocks/xrefs)" : "";
            ed.WriteMessage($"\nView Mode: Gathering entities from current view/layout{modeMsg}...\n");

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

                        // If search-in-blocks mode is enabled and this is a block reference, gather entities from within
                        var entity = tr.GetObject(id, OpenMode.ForRead);
                        if (searchInBlocks && entity is BlockReference blockRef)
                        {
                            var blockEntities = BlockEntityUtilities.GetEntitiesInBlock(blockRef.BlockTableRecord, tr);
                            foreach (var blockEntityId in blockEntities)
                            {
                                try
                                {
                                    var blockTags = blockEntityId.GetTags(db);
                                    foreach (var tag in blockTags)
                                    {
                                        if (!tagGroups.ContainsKey(tag))
                                            tagGroups[tag] = new List<ObjectId>();

                                        tagGroups[tag].Add(blockEntityId);
                                    }
                                }
                                catch
                                {
                                    continue;
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

            ShowSelectionDialogForView(ed, tagGroups, spaceName);
        }

        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();
            string modeMsg = searchInBlocks ? " (including blocks/xrefs)" : "";
            ed.WriteMessage($"\nDocument Mode: Gathering entities from entire document{modeMsg}...\n");

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
                                    SpaceName = spaceName,
                                    BlockPath = "",
                                    IsInBlock = false
                                });
                            }

                            // If search-in-blocks mode is enabled and this is a block reference, gather entities from within
                            if (searchInBlocks && entity is BlockReference blockRef)
                            {
                                var blockName = BlockEntityUtilities.GetBlockName(blockRef, tr);
                                var blockEntities = BlockEntityUtilities.GetEntitiesInBlock(blockRef.BlockTableRecord, tr, blockName, blockRef.Handle.ToString());

                                foreach (var blockEntityId in blockEntities)
                                {
                                    try
                                    {
                                        var blockEntity = tr.GetObject(blockEntityId, OpenMode.ForRead);
                                        var blockTags = blockEntityId.GetTags(db);

                                        // Get block path from entity's block table record
                                        string blockPath = blockName;
                                        if (blockEntity is Entity ent)
                                        {
                                            var blockBtr = tr.GetObject(ent.BlockId, OpenMode.ForRead) as BlockTableRecord;
                                            blockPath = blockBtr?.Name ?? blockName;
                                        }

                                        foreach (var tag in blockTags)
                                        {
                                            if (!tagGroups.ContainsKey(tag))
                                                tagGroups[tag] = new List<TagEntityReference>();

                                            tagGroups[tag].Add(new TagEntityReference
                                            {
                                                DocumentPath = doc.Name,
                                                DocumentName = Path.GetFileName(doc.Name),
                                                Handle = blockEntity.Handle.ToString(),
                                                Tag = tag,
                                                SpaceName = spaceName,
                                                BlockPath = blockPath,
                                                IsInBlock = true
                                            });
                                        }
                                    }
                                    catch
                                    {
                                        continue;
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

            ShowSelectionDialogForDocument(ed, tagGroups, doc.Name);
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            ed.WriteMessage("\nApplication Mode: Gathering entities from all open documents...\n");

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

            ShowSelectionDialogForApplication(ed, tagGroups);
        }

        private static void ShowSelectionDialogForView(Editor ed, Dictionary<string, List<ObjectId>> tagGroups, string spaceName)
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
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            var selectedTags = new HashSet<string>(chosenRows.Select(row => row["Tag"].ToString()));
            var selectedIds = new List<ObjectId>();

            foreach (var tag in selectedTags)
            {
                if (tagGroups.ContainsKey(tag))
                {
                    selectedIds.AddRange(tagGroups[tag]);
                }
            }

            // Remove duplicates (entities can have multiple tags)
            selectedIds = selectedIds.Distinct().ToList();

            if (selectedIds.Count > 0)
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc.Database;

                // Build selection references for ALL entities (including those in blocks)
                var selectionItems = new List<SelectionItem>();
                var selectableIds = new List<ObjectId>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var id in selectedIds)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity != null)
                            {
                                // Add to selection storage (all entities, including block internals)
                                selectionItems.Add(new SelectionItem
                                {
                                    DocumentPath = doc.Name,
                                    Handle = entity.Handle.ToString(),
                                    SessionId = null
                                });

                                // Only add entities in current space to selectable list
                                if (entity.BlockId == db.CurrentSpaceId)
                                {
                                    selectableIds.Add(id);
                                }
                            }
                        }
                        catch
                        {
                            // Skip entities that can't be accessed
                            continue;
                        }
                    }
                    tr.Commit();
                }

                // Save selection references (including block entities) for filter commands
                try
                {
                    var docName = Path.GetFileName(doc.Name);
                    SelectionStorage.SaveSelection(selectionItems, docName);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Could not save selection references: {ex.Message}\n");
                }

                // Set implied selection for selectable entities only
                if (selectableIds.Count > 0)
                {
                    ed.SetImpliedSelection(selectableIds.ToArray());
                    ed.WriteMessage($"\nSelected {selectableIds.Count} entities from {selectedTags.Count} tags in {spaceName}.\n");

                    // Report if some entities were in blocks and stored as references
                    int blockedCount = selectionItems.Count - selectableIds.Count;
                    if (blockedCount > 0)
                    {
                        ed.WriteMessage($"  ({blockedCount} entities within blocks stored as references for filter commands)\n");
                    }
                }
                else
                {
                    ed.WriteMessage($"\nNo directly selectable entities. All {selectionItems.Count} entities within blocks stored as references.\n");
                }
            }
        }

        private static void ShowSelectionDialogForDocument(Editor ed, Dictionary<string, List<TagEntityReference>> tagGroups, string documentName)
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
                ed.WriteMessage("\nNo tags selected.\n");
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

            var selectionItems = selectedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            try
            {
                var docName = Path.GetFileName(documentName);
                SelectionStorage.SaveSelection(selectionItems, docName);
                ed.WriteMessage($"\nDocument selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
                ed.WriteMessage($"  Tags: {selectedTagNames.Count}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection to subset of selected entities in current space
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var currentViewIds = new List<ObjectId>();
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
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.BlockId == db.CurrentSpaceId)
                            {
                                currentViewIds.Add(objectId);
                            }
                        }
                    }
                    catch
                    {
                        // Skip invalid entities
                    }
                }
                tr.Commit();
            }

            if (currentViewIds.Count > 0)
            {
                try
                {
                    ed.SetImpliedSelection(currentViewIds.ToArray());
                    ed.WriteMessage($"  Selected {currentViewIds.Count} entities in current view.\n");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError setting current view selection: {ex.Message}\n");
                }
            }
        }

        private static void ShowSelectionDialogForApplication(Editor ed, Dictionary<string, List<TagEntityReference>> tagGroups)
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
                ed.WriteMessage("\nNo tags selected.\n");
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

            // Remove duplicates based on document path and handle
            selectedEntities = selectedEntities
                .GroupBy(e => new { e.DocumentPath, e.Handle })
                .Select(g => g.First())
                .ToList();

            var selectionItems = selectedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            try
            {
                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage($"\nApplication selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
                ed.WriteMessage($"  Tags: {selectedTagNames.Count}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection to subset of selected entities in current document's current space
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var currentDocPath = Path.GetFullPath(doc.Name);
            var currentViewIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var entityRef in selectedEntities)
                {
                    try
                    {
                        // Only process entities from current document
                        var entityDocPath = Path.GetFullPath(entityRef.DocumentPath);
                        if (!string.Equals(entityDocPath, currentDocPath, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var handle = Convert.ToInt64(entityRef.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (!objectId.IsNull && objectId.IsValid)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.BlockId == db.CurrentSpaceId)
                            {
                                currentViewIds.Add(objectId);
                            }
                        }
                    }
                    catch
                    {
                        // Skip invalid entities
                    }
                }
                tr.Commit();
            }

            if (currentViewIds.Count > 0)
            {
                try
                {
                    ed.SetImpliedSelection(currentViewIds.ToArray());
                    ed.WriteMessage($"  Selected {currentViewIds.Count} entities in current view.\n");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError setting current view selection: {ex.Message}\n");
                }
            }
        }

        private static List<TagEntityReference> GatherTagReferencesFromDocument(Database db, string docPath, string docName)
        {
            var references = new List<TagEntityReference>();
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();

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
                                    SpaceName = spaceName,
                                    BlockPath = "",
                                    IsInBlock = false
                                });
                            }

                            // If search-in-blocks mode is enabled and this is a block reference, gather entities from within
                            if (searchInBlocks && entity is BlockReference blockRef)
                            {
                                var blockName = BlockEntityUtilities.GetBlockName(blockRef, tr);
                                var blockEntities = BlockEntityUtilities.GetEntitiesInBlock(blockRef.BlockTableRecord, tr, blockName, blockRef.Handle.ToString());

                                foreach (var blockEntityId in blockEntities)
                                {
                                    try
                                    {
                                        var blockEntity = tr.GetObject(blockEntityId, OpenMode.ForRead);
                                        var blockTags = blockEntityId.GetTags(db);

                                        // Get block path from entity's block table record
                                        string blockPath = blockName;
                                        if (blockEntity is Entity ent)
                                        {
                                            var blockBtr = tr.GetObject(ent.BlockId, OpenMode.ForRead) as BlockTableRecord;
                                            blockPath = blockBtr?.Name ?? blockName;
                                        }

                                        foreach (var tag in blockTags)
                                        {
                                            references.Add(new TagEntityReference
                                            {
                                                DocumentPath = docPath,
                                                DocumentName = docName,
                                                Handle = blockEntity.Handle.ToString(),
                                                Tag = tag,
                                                SpaceName = spaceName,
                                                BlockPath = blockPath,
                                                IsInBlock = true
                                            });
                                        }
                                    }
                                    catch
                                    {
                                        continue;
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
