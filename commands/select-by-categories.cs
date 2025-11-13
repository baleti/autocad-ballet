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
    // Simple class to store entity references
    public class CategoryEntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string Category { get; set; }
        public string SpaceName { get; set; }
        public string BlockPath { get; set; }  // Empty for top-level entities, "BlockA > BlockB" for nested
        public bool IsInBlock { get; set; }     // True if entity is within a block definition
    }

    public static class SelectByCategories
    {
        public static void ExecuteViewScope(Editor ed, Database db)
        {
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();
            string modeMsg = searchInBlocks ? " (including blocks/xrefs)" : "";
            ed.WriteMessage($"\nView Mode: Gathering entities from current view/layout{modeMsg}...\n");

            var entityGroups = new Dictionary<string, List<ObjectId>>();
            var spaceName = GetCurrentSpaceName(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in currentSpace)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead);
                        var category = GetEntityCategory(entity);

                        if (!entityGroups.ContainsKey(category))
                            entityGroups[category] = new List<ObjectId>();

                        entityGroups[category].Add(id);

                        // If search-in-blocks mode is enabled and this is a block reference, gather entities from within
                        if (searchInBlocks && entity is BlockReference blockRef)
                        {
                            var blockEntities = BlockEntityUtilities.GetEntitiesInBlock(blockRef.BlockTableRecord, tr);
                            foreach (var blockEntityId in blockEntities)
                            {
                                try
                                {
                                    var blockEntity = tr.GetObject(blockEntityId, OpenMode.ForRead);
                                    var blockCategory = GetEntityCategory(blockEntity);

                                    if (!entityGroups.ContainsKey(blockCategory))
                                        entityGroups[blockCategory] = new List<ObjectId>();

                                    entityGroups[blockCategory].Add(blockEntityId);
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

                // Also gather layers (layers are document-wide, not space-specific)
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                var layerIds = new List<ObjectId>();

                foreach (ObjectId layerId in layerTable)
                {
                    try
                    {
                        var layer = tr.GetObject(layerId, OpenMode.ForRead);
                        layerIds.Add(layerId);
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (layerIds.Count > 0)
                {
                    entityGroups["Layer"] = layerIds;
                }

                tr.Commit();
            }

            ShowSelectionDialogForView(ed, entityGroups, spaceName);
        }

        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();
            string modeMsg = searchInBlocks ? " (including blocks/xrefs)" : "";
            ed.WriteMessage($"\nDocument Mode: Gathering entities from entire document{modeMsg}...\n");

            var entityGroups = new Dictionary<string, List<CategoryEntityReference>>();

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
                            var category = GetEntityCategory(entity);

                            if (!entityGroups.ContainsKey(category))
                                entityGroups[category] = new List<CategoryEntityReference>();

                            entityGroups[category].Add(new CategoryEntityReference
                            {
                                DocumentPath = doc.Name,
                                DocumentName = Path.GetFileName(doc.Name),
                                Handle = entity.Handle.ToString(),
                                Category = category,
                                SpaceName = spaceName,
                                BlockPath = "",
                                IsInBlock = false
                            });

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
                                        var blockCategory = GetEntityCategory(blockEntity);

                                        // Get block path from entity's block table record
                                        string blockPath = blockName;
                                        if (blockEntity is Entity ent)
                                        {
                                            var blockBtr = tr.GetObject(ent.BlockId, OpenMode.ForRead) as BlockTableRecord;
                                            blockPath = blockBtr?.Name ?? blockName;
                                        }

                                        if (!entityGroups.ContainsKey(blockCategory))
                                            entityGroups[blockCategory] = new List<CategoryEntityReference>();

                                        entityGroups[blockCategory].Add(new CategoryEntityReference
                                        {
                                            DocumentPath = doc.Name,
                                            DocumentName = Path.GetFileName(doc.Name),
                                            Handle = blockEntity.Handle.ToString(),
                                            Category = blockCategory,
                                            SpaceName = spaceName,
                                            BlockPath = blockPath,
                                            IsInBlock = true
                                        });
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

                // Also gather layers (layers are document-wide)
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                foreach (ObjectId layerId in layerTable)
                {
                    try
                    {
                        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);

                        if (!entityGroups.ContainsKey("Layer"))
                            entityGroups["Layer"] = new List<CategoryEntityReference>();

                        entityGroups["Layer"].Add(new CategoryEntityReference
                        {
                            DocumentPath = doc.Name,
                            DocumentName = Path.GetFileName(doc.Name),
                            Handle = layer.Handle.ToString(),
                            Category = "Layer",
                            SpaceName = "Document-wide",
                            BlockPath = "",
                            IsInBlock = false
                        });
                    }
                    catch
                    {
                        continue;
                    }
                }

                tr.Commit();
            }

            ShowSelectionDialogForDocument(ed, db, entityGroups, doc.Name);
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            ed.WriteMessage("\nApplication Mode: Gathering entities from all open documents...\n");

            var docManager = AcadApp.DocumentManager;
            var allReferences = new List<CategoryEntityReference>();
            var categoryGroups = new Dictionary<string, List<CategoryEntityReference>>();

            foreach (Autodesk.AutoCAD.ApplicationServices.Document openDoc in docManager)
            {
                string docPath = openDoc.Name;
                string docName = Path.GetFileName(docPath);

                ed.WriteMessage($"\nScanning: {docName}...");

                try
                {
                    var refs = GatherEntityReferencesFromDocument(openDoc.Database, docPath, docName);
                    allReferences.AddRange(refs);

                    var layoutRefs = GatherLayoutReferencesFromDocument(openDoc.Database, docPath, docName);
                    allReferences.AddRange(layoutRefs);

                    var layerRefs = GatherLayerReferencesFromDocument(openDoc.Database, docPath, docName);
                    allReferences.AddRange(layerRefs);

                    foreach (var entityRef in refs.Concat(layoutRefs).Concat(layerRefs))
                    {
                        if (!categoryGroups.ContainsKey(entityRef.Category))
                            categoryGroups[entityRef.Category] = new List<CategoryEntityReference>();

                        categoryGroups[entityRef.Category].Add(entityRef);
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
                }
            }

            ShowSelectionDialogForApplication(ed, db, categoryGroups);
        }

        private static void ShowSelectionDialogForView(Editor ed, Dictionary<string, List<ObjectId>> entityGroups, string spaceName)
        {
            if (entityGroups.Count == 0)
            {
                ed.WriteMessage($"\nNo entities found in {spaceName}.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var cat in entityGroups.OrderBy(c => c.Key))
            {
                entries.Add(new Dictionary<string, object>
                {
                    { "Category", cat.Key },
                    { "Count", cat.Value.Count },
                    { "Space", spaceName }
                });
            }

            var propertyNames = new List<string> { "Category", "Count", "Space" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo categories selected.\n");
                return;
            }

            var selectedCategories = new HashSet<string>(chosenRows.Select(row => row["Category"].ToString()));
            var selectedIds = new List<ObjectId>();

            foreach (var cat in selectedCategories)
            {
                if (entityGroups.ContainsKey(cat))
                {
                    selectedIds.AddRange(entityGroups[cat]);
                }
            }

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
                    ed.WriteMessage($"\nSelected {selectableIds.Count} entities from {selectedCategories.Count} categories in {spaceName}.\n");

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

        private static void ShowSelectionDialogForDocument(Editor ed, Database db, Dictionary<string, List<CategoryEntityReference>> entityGroups, string documentName)
        {
            if (entityGroups.Count == 0)
            {
                ed.WriteMessage($"\nNo entities found in document.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var cat in entityGroups.OrderBy(c => c.Key))
            {
                var docCounts = cat.Value.GroupBy(e => e.SpaceName)
                                         .Select(g => $"{g.Key}: {g.Count()}")
                                         .ToList();

                entries.Add(new Dictionary<string, object>
                {
                    { "Category", cat.Key },
                    { "Total Count", cat.Value.Count },
                    { "Layouts", string.Join(", ", docCounts) }
                });
            }

            var propertyNames = new List<string> { "Category", "Total Count", "Layouts" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo categories selected.\n");
                return;
            }

            var selectedCategoryNames = new HashSet<string>(chosenRows.Select(row => row["Category"].ToString()));
            var selectedEntities = new List<CategoryEntityReference>();

            foreach (var categoryName in selectedCategoryNames)
            {
                if (entityGroups.ContainsKey(categoryName))
                {
                    selectedEntities.AddRange(entityGroups[categoryName]);
                }
            }

            var selectionItems = selectedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            // Save selection (replacing previous document selection)
            try
            {
                var docName = Path.GetFileName(documentName);
                SelectionStorage.SaveSelection(selectionItems, docName);
                ed.WriteMessage($"\nDocument selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
                ed.WriteMessage($"  Categories: {selectedCategoryNames.Count}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection to subset of selected entities in current space
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

        private static void ShowSelectionDialogForApplication(Editor ed, Database db, Dictionary<string, List<CategoryEntityReference>> categoryGroups)
        {
            if (categoryGroups.Count == 0)
            {
                ed.WriteMessage("\nNo entities found across open documents.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var cat in categoryGroups.OrderBy(c => c.Key))
            {
                var docCounts = cat.Value.GroupBy(e => e.DocumentName)
                                         .Select(g => $"{g.Key}: {g.Count()}")
                                         .ToList();

                entries.Add(new Dictionary<string, object>
                {
                    { "Category", cat.Key },
                    { "Total Count", cat.Value.Count },
                    { "Documents", string.Join(", ", docCounts) }
                });
            }

            var propertyNames = new List<string> { "Category", "Total Count", "Documents" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo categories selected.\n");
                return;
            }

            var selectedCategoryNames = new HashSet<string>(chosenRows.Select(row => row["Category"].ToString()));
            var selectedEntities = new List<CategoryEntityReference>();

            foreach (var categoryName in selectedCategoryNames)
            {
                if (categoryGroups.ContainsKey(categoryName))
                {
                    selectedEntities.AddRange(categoryGroups[categoryName]);
                }
            }

            var selectionItems = selectedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            // Clear all existing selections and save (session scope behavior)
            ClearAllStoredSelections();

            try
            {
                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage($"\nApplication selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
                ed.WriteMessage($"  Categories: {selectedCategoryNames.Count}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection to subset of selected entities in current document's current space
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
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

        private static List<CategoryEntityReference> GatherEntityReferencesFromDocument(Database db, string docPath, string docName)
        {
            var references = new List<CategoryEntityReference>();
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
                            var category = GetEntityCategory(entity);

                            references.Add(new CategoryEntityReference
                            {
                                DocumentPath = docPath,
                                DocumentName = docName,
                                Handle = entity.Handle.ToString(),
                                Category = category,
                                SpaceName = spaceName,
                                BlockPath = "",
                                IsInBlock = false
                            });

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
                                        var blockCategory = GetEntityCategory(blockEntity);

                                        // Get block path from entity's block table record
                                        string blockPath = blockName;
                                        if (blockEntity is Entity ent)
                                        {
                                            var blockBtr = tr.GetObject(ent.BlockId, OpenMode.ForRead) as BlockTableRecord;
                                            blockPath = blockBtr?.Name ?? blockName;
                                        }

                                        references.Add(new CategoryEntityReference
                                        {
                                            DocumentPath = docPath,
                                            DocumentName = docName,
                                            Handle = blockEntity.Handle.ToString(),
                                            Category = blockCategory,
                                            SpaceName = spaceName,
                                            BlockPath = blockPath,
                                            IsInBlock = true
                                        });
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

        private static List<CategoryEntityReference> GatherLayoutReferencesFromDocument(Database db, string docPath, string docName)
        {
            var references = new List<CategoryEntityReference>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                    references.Add(new CategoryEntityReference
                    {
                        DocumentPath = docPath,
                        DocumentName = docName,
                        Handle = layout.Handle.ToString(),
                        Category = "Layout",
                        SpaceName = layout.LayoutName,
                        BlockPath = "",
                        IsInBlock = false
                    });
                }

                tr.Commit();
            }

            return references;
        }

        private static List<CategoryEntityReference> GatherLayerReferencesFromDocument(Database db, string docPath, string docName)
        {
            var references = new List<CategoryEntityReference>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                foreach (ObjectId layerId in layerTable)
                {
                    try
                    {
                        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);

                        references.Add(new CategoryEntityReference
                        {
                            DocumentPath = docPath,
                            DocumentName = docName,
                            Handle = layer.Handle.ToString(),
                            Category = "Layer",
                            SpaceName = "Document-wide",
                            BlockPath = "",
                            IsInBlock = false
                        });
                    }
                    catch
                    {
                        continue;
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

        private static bool IsRevisionCloud(Entity entity)
        {
            try
            {
                var xData = entity.XData;
                if (xData != null)
                {
                    foreach (TypedValue typedValue in xData)
                    {
                        if (typedValue.TypeCode == (int)DxfCode.ExtendedDataRegAppName)
                        {
                            string appName = typedValue.Value?.ToString();
                            // Check for both "RevcloudProps" and "RevCloudProps" (version differences)
                            if (appName != null &&
                                (appName.Equals("RevcloudProps", StringComparison.OrdinalIgnoreCase) ||
                                 appName.Equals("RevCloudProps", StringComparison.OrdinalIgnoreCase)))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            catch
            {
                // If we can't read XData, assume it's not a revision cloud
            }
            return false;
        }

        public static string GetEntityCategory(DBObject entity)
        {
            string typeName = entity.GetType().Name;

            // Check for revision clouds before checking for regular polylines
            // Revision clouds are polylines with XData registered application name "RevcloudProps"
            if (entity is Polyline polyline && IsRevisionCloud(polyline))
                return "Revision Cloud";

            if (entity is LayerTableRecord)
                return "Layer";
            else if (entity is BlockReference)
            {
                var br = entity as BlockReference;
                using (var tr = br.Database.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr.IsFromExternalReference)
                        return "XRef";
                    else if (btr.IsAnonymous)
                        return "Dynamic Block";
                    else
                        return "Block Reference";
                }
            }
            else if (entity is Dimension)
            {
                if (entity is AlignedDimension)
                    return "Aligned Dimension";
                else if (entity is RotatedDimension)
                    return "Linear Dimension";
                else if (entity is RadialDimension)
                    return "Radial Dimension";
                else if (entity is DiametricDimension)
                    return "Diametric Dimension";
                else if (entity is OrdinateDimension)
                    return "Ordinate Dimension";
                else if (entity is ArcDimension)
                    return "Arc Dimension";
                else if (entity is RadialDimensionLarge)
                    return "Jogged Dimension";
                else
                    return "Dimension";
            }
            else if (entity is MText)
                return "MText";
            else if (entity is DBText)
                return "Text";
            else if (entity is Polyline)
                return "Polyline";
            else if (entity is Polyline2d)
                return "Polyline2D";
            else if (entity is Polyline3d)
                return "Polyline3D";
            else if (entity is Line)
                return "Line";
            else if (entity is Circle)
                return "Circle";
            else if (entity is Arc)
                return "Arc";
            else if (entity is Ellipse)
                return "Ellipse";
            else if (entity is Spline)
                return "Spline";
            else if (entity is Hatch)
                return "Hatch";
            else if (entity is Solid)
                return "2D Solid";
            else if (entity is Leader)
                return "Leader";
            else if (entity is MLeader)
                return "Multileader";
            else if (entity is Table)
                return "Table";
            else if (entity is Viewport)
                return "Viewport";
            else if (entity is RasterImage)
                return "Raster Image";
            else if (entity is Wipeout)
                return "Wipeout";
            else if (entity is DBPoint)
                return "Point";
            else if (entity is Ray)
                return "Ray";
            else if (entity is Xline)
                return "Construction Line";
            else if (entity is Layout)
                return "Layout";
            else
            {
                return typeName.Replace("Autodesk.AutoCAD.", "");
            }
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
