using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Register the command class
// Command registration removed - using scope-specific commands

/// <summary>
/// Helper class to extract entity data for filtering and display
/// </summary>
public static class FilterEntityDataHelper
{
    private static string GetCurrentSessionId()
    {
        // Generate unique session identifier combining process ID and session ID (same as SelectionStorage)
        int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
        int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
        return $"{processId}_{sessionId}";
    }

    public static List<Dictionary<string, object>> GetEntityData(Editor ed, SelectionScope scope, out ObjectId[] originalSelection, bool selectedOnly = false, bool includeProperties = false)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        var entityData = new List<Dictionary<string, object>>();
        originalSelection = null;

        // Handle selection based on current scope
        if (scope == SelectionScope.view)
        {
            // Get pickfirst set (pre-selected objects)
            var selResult = ed.SelectImplied();

            // Store original selection for potential restoration on cancel
            if (selResult.Status == PromptStatus.OK && selResult.Value.Count > 0)
            {
                originalSelection = selResult.Value.GetObjectIds();
            }

            // If there is no pickfirst set, request user to select objects
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to filter: ";
                selResult = ed.GetSelection(selectionOpts);

                // Store the new selection for potential restoration
                if (selResult.Status == PromptStatus.OK && selResult.Value.Count > 0)
                {
                    originalSelection = selResult.Value.GetObjectIds();
                }
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                // Clear the pickfirst set since we're consuming it
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            // Collect entities from pickfirst selection
            if (selResult.Status == PromptStatus.OK && selResult.Value.Count > 0)
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // Build block hierarchy cache ONCE at the start (huge performance optimization)
                    var blockHierarchyCache = BuildBlockHierarchyCache(db, tr);

                    foreach (var objectId in selResult.Value.GetObjectIds())
                    {
                        try
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead);
                            if (entity != null)
                            {
                                var data = GetEntityDataDictionary(entity, doc.Name, null, includeProperties, tr, blockHierarchyCache);
                                data["ObjectId"] = objectId; // Store for selection
                                entityData.Add(data);
                            }
                        }
                        catch
                        {
                            // Skip problematic entities
                            continue;
                        }
                    }
                    tr.Commit();
                }
            }

            // ALSO load stored selection references (may include entities in blocks)
            // Only do this if select-in-blocks mode is enabled
            if (SelectInBlocksMode.IsEnabled())
            {
                var docName = Path.GetFileName(doc.Name);
                var storedSelection = SelectionStorage.LoadSelection(docName);

                if (storedSelection != null && storedSelection.Count > 0)
                {
                    // Filter to current session
                    var currentSessionId = GetCurrentSessionId();
                    storedSelection = storedSelection.Where(item =>
                        string.IsNullOrEmpty(item.SessionId) || item.SessionId == currentSessionId).ToList();

                    // Process stored selection items from current document
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        // Build block hierarchy cache ONCE at the start (huge performance optimization)
                        var blockHierarchyCache = BuildBlockHierarchyCache(db, tr);

                        foreach (var item in storedSelection)
                        {
                            try
                            {
                                // Only process items from current document
                                if (Path.GetFullPath(item.DocumentPath) == Path.GetFullPath(doc.Name))
                                {
                                    var handle = Convert.ToInt64(item.Handle, 16);
                                    var objectId = db.GetObjectId(false, new Handle(handle), 0);

                                    if (objectId != ObjectId.Null)
                                    {
                                        var entity = tr.GetObject(objectId, OpenMode.ForRead);
                                        if (entity != null)
                                        {
                                            // Check if this entity is already in entityData (from pickfirst)
                                            bool alreadyAdded = entityData.Any(d => d.ContainsKey("ObjectId") && ((ObjectId)d["ObjectId"]) == objectId);
                                            if (!alreadyAdded)
                                            {
                                                var data = GetEntityDataDictionary(entity, item.DocumentPath, null, includeProperties, tr, blockHierarchyCache);
                                                data["ObjectId"] = objectId; // Store for selection
                                                entityData.Add(data);
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Skip problematic entities
                                continue;
                            }
                        }
                        tr.Commit();
                    }
                }
            }

            if (entityData.Count == 0)
            {
                throw new InvalidOperationException("No selection found. Please select entities first or use 'select-by-categories-in-view' to create a stored selection.");
            }

            return entityData;
        }

        if (selectedOnly)
        {
            // Get entities from stored selection based on scope
            List<SelectionItem> storedSelection;

            var loadTimer = System.Diagnostics.Stopwatch.StartNew();
            if (scope == SelectionScope.document)
            {
                // Document scope - load selection for current document only
                var docName = Path.GetFileNameWithoutExtension(doc.Name);
                storedSelection = SelectionStorage.LoadSelection(docName);
            }
            else if (scope == SelectionScope.application)
            {
                // Application scope - load from all open documents (session scope)
                storedSelection = SelectionStorage.LoadSelectionFromOpenDocuments();
            }
            else
            {
                // View scope - load from global file
                storedSelection = SelectionStorage.LoadSelection();
            }
            loadTimer.Stop();

            if (storedSelection == null || storedSelection.Count == 0)
            {
                if (scope == SelectionScope.document)
                {
                    throw new InvalidOperationException($"No stored selection found for current document '{Path.GetFileNameWithoutExtension(doc.Name)}'. Use commands like 'select-by-categories-in-document' first.");
                }
                else
                {
                    throw new InvalidOperationException("No stored selection found. Use commands like 'select-by-categories' first.");
                }
            }

            var filterTimer = System.Diagnostics.Stopwatch.StartNew();
            // Filter to current session to avoid confusion with selections from different AutoCAD processes
            var currentSessionId = GetCurrentSessionId();
            var beforeFilterCount = storedSelection.Count;
            storedSelection = storedSelection.Where(item =>
                string.IsNullOrEmpty(item.SessionId) || item.SessionId == currentSessionId).ToList();
            filterTimer.Stop();

            // Check if filtering resulted in empty selection
            if (storedSelection.Count == 0)
            {
                if (scope == SelectionScope.document)
                {
                    throw new InvalidOperationException($"No stored selection found for current document '{Path.GetFileName(doc.Name)}'. The stored selection may be from other documents in process scope.");
                }
                else
                {
                    throw new InvalidOperationException("No stored selection found after filtering. Use commands like 'select-by-categories' first.");
                }
            }

            var processTimer = System.Diagnostics.Stopwatch.StartNew();
            int currentDocCount = 0;
            int externalDocCount = 0;
            var externalDocTimer = new System.Diagnostics.Stopwatch();

            // Separate current document items from external document items
            var currentDocItems = new List<SelectionItem>();
            var externalDocItems = new List<SelectionItem>();

            foreach (var item in storedSelection)
            {
                try
                {
                    if (Path.GetFullPath(item.DocumentPath) == Path.GetFullPath(doc.Name))
                    {
                        currentDocItems.Add(item);
                    }
                    else
                    {
                        externalDocItems.Add(item);
                    }
                }
                catch
                {
                    // If path comparison fails, treat as external
                    externalDocItems.Add(item);
                }
            }

            // Process current document items in a single transaction with cache
            if (currentDocItems.Count > 0)
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // Build block hierarchy cache ONCE for all entities (huge performance optimization)
                    var blockHierarchyCache = BuildBlockHierarchyCache(db, tr);

                    foreach (var item in currentDocItems)
                    {
                        try
                        {
                            currentDocCount++;
                            var handle = Convert.ToInt64(item.Handle, 16);
                            var objectId = db.GetObjectId(false, new Handle(handle), 0);

                            if (objectId != ObjectId.Null)
                            {
                                var entity = tr.GetObject(objectId, OpenMode.ForRead);
                                if (entity != null)
                                {
                                    var data = GetEntityDataDictionary(entity, item.DocumentPath, null, includeProperties, tr, blockHierarchyCache);
                                    data["ObjectId"] = objectId; // Store for selection
                                    entityData.Add(data);
                                }
                            }
                        }
                        catch
                        {
                            // Skip problematic entities
                            continue;
                        }
                    }
                    tr.Commit();
                }
            }

            // Process external document items
            foreach (var item in externalDocItems)
            {
                try
                {
                    externalDocCount++;
                    externalDocTimer.Start();
                    var data = GetExternalEntityData(item.DocumentPath, item.Handle, includeProperties);
                    externalDocTimer.Stop();
                    entityData.Add(data);
                }
                catch
                {
                    // Skip problematic entities
                    continue;
                }
            }
            processTimer.Stop();
            if (externalDocCount > 0)
            {
                // Reset timers for next run
                _diagCheckOpenTimer.Reset();
                _diagOpenDocTimer.Reset();
                _diagReadEntityTimer.Reset();
                _diagCloseDocTimer.Reset();
            }
        }
        else
        {
            // Get all entities from current scope (fallback - should not be used by filter-selection)
            var entities = GatherEntitiesFromScope(db, scope);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Build block hierarchy cache ONCE at the start (huge performance optimization)
                var blockHierarchyCache = BuildBlockHierarchyCache(db, tr);

                foreach (var objectId in entities)
                {
                    try
                    {
                        var entity = tr.GetObject(objectId, OpenMode.ForRead);
                        if (entity != null)
                        {
                            var data = GetEntityDataDictionary(entity, doc.Name, null, includeProperties, tr, blockHierarchyCache);
                            entityData.Add(data);
                        }
                    }
                    catch
                    {
                        // Skip problematic entities
                        continue;
                    }
                }
                tr.Commit();
            }
        }

        return entityData;
    }

    // Diagnostic counters for GetExternalEntityData
    private static System.Diagnostics.Stopwatch _diagCheckOpenTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagOpenDocTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagReadEntityTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagCloseDocTimer = new System.Diagnostics.Stopwatch();

    public static Dictionary<string, object> GetExternalEntityData(string documentPath, string handle, bool includeProperties = false)
    {
        var data = new Dictionary<string, object>
        {
            ["Name"] = "External Reference",
            ["Category"] = "External Entity",
            ["Tags"] = "",
            ["Layer"] = "N/A",
            ["Color"] = "N/A",
            ["LineType"] = "N/A",
            ["Layout"] = "N/A",
            ["Contents"] = "N/A",
            ["DocumentPath"] = documentPath,
            ["DocumentName"] = Path.GetFileName(documentPath),
            ["Handle"] = handle,
            ["Id"] = handle,
            ["IsExternal"] = true,
            ["DisplayName"] = $"External: {Path.GetFileName(documentPath)}",
            ["DynamicBlockName"] = "N/A"
        };

        try
        {
            // Try to open the external document and get real entity properties
            var docs = AcadApp.DocumentManager;
            Document externalDoc = null;
            bool docWasAlreadyOpen = false;

            _diagCheckOpenTimer.Start();
            // Check if the document is already open in the current session
            foreach (Document openDoc in docs)
            {
                if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(documentPath), StringComparison.OrdinalIgnoreCase))
                {
                    externalDoc = openDoc;
                    docWasAlreadyOpen = true;
                    break;
                }
            }
            _diagCheckOpenTimer.Stop();

            // If not already open, try to open it temporarily
            if (externalDoc == null && File.Exists(documentPath))
            {
                try
                {
                    _diagOpenDocTimer.Start();
                    externalDoc = docs.Open(documentPath, false); // Open read-only
                    _diagOpenDocTimer.Stop();
                    docWasAlreadyOpen = false;
                }
                catch
                {
                    _diagOpenDocTimer.Stop();
                    // If we can't open the document, return the N/A data
                    return data;
                }
            }

            // If we have the external document, get the entity properties
            if (externalDoc != null)
            {
                try
                {
                    _diagReadEntityTimer.Start();
                    var handleValue = Convert.ToInt64(handle, 16);
                    var objectId = externalDoc.Database.GetObjectId(false, new Handle(handleValue), 0);

                    if (objectId != ObjectId.Null)
                    {
                        using (var tr = externalDoc.Database.TransactionManager.StartTransaction())
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead);
                            if (entity != null)
                            {
                                // Get the real entity data
                                data = GetEntityDataDictionary(entity, documentPath, null, includeProperties, tr);
                                data["IsExternal"] = true;
                                data["DisplayName"] = $"External: {data["Name"]}";
                            }
                            tr.Commit();
                        }
                    }
                    _diagReadEntityTimer.Stop();
                }
                finally
                {
                    // Close the document if we opened it temporarily
                    if (!docWasAlreadyOpen && externalDoc != null)
                    {
                        try
                        {
                            _diagCloseDocTimer.Start();
                            externalDoc.CloseAndDiscard();
                            _diagCloseDocTimer.Stop();
                        }
                        catch
                        {
                            _diagCloseDocTimer.Stop();
                            // Ignore close errors
                        }
                    }
                }
            }
        }
        catch
        {
            // If anything goes wrong, return the default N/A data
            // The data dictionary is already initialized with N/A values above
        }

        return data;
    }

    /// <summary>
    /// Information about a parent block in the hierarchy
    /// </summary>
    private class ParentBlockInfo
    {
        public string Name { get; set; }
        public string Type { get; set; } // "Block Reference", "Dynamic Block", "XRef"
    }

    /// <summary>
    /// Builds a cache mapping block definition ObjectIds to their parent block ObjectIds.
    /// This eliminates the need to scan all blocks repeatedly in FindParentBlockReference.
    /// Returns a dictionary where Key = referenced block definition, Value = parent block containing the reference.
    /// </summary>
    public static Dictionary<ObjectId, ObjectId> BuildBlockHierarchyCache(Database db, Transaction tr)
    {
        var cache = new Dictionary<ObjectId, ObjectId>();
        var cacheTimer = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            var blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (blockTable == null)
                return cache;

            int blocksScanned = 0;
            int entitiesScanned = 0;
            int referencesFound = 0;

            // Scan all block definitions once
            foreach (ObjectId blockId in blockTable)
            {
                var btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null || btr.IsLayout)
                    continue;

                blocksScanned++;

                // Look through entities in this block for BlockReferences
                foreach (ObjectId entityId in btr)
                {
                    entitiesScanned++;

                    var entity = tr.GetObject(entityId, OpenMode.ForRead);
                    if (entity is BlockReference blockRef)
                    {
                        // Map: referenced block definition → parent block that contains it
                        var referencedBlockId = blockRef.BlockTableRecord;

                        // Only store if not already in cache (first parent wins)
                        if (!cache.ContainsKey(referencedBlockId))
                        {
                            cache[referencedBlockId] = btr.ObjectId;
                            referencesFound++;
                        }
                    }
                }
            }

            cacheTimer.Stop();
            _diagCacheBuildTime = cacheTimer.ElapsedMilliseconds;
            _diagCacheBlocksScanned = blocksScanned;
            _diagCacheEntitiesScanned = entitiesScanned;
            _diagCacheReferencesFound = referencesFound;
        }
        catch
        {
            // Return whatever cache we've built
        }

        return cache;
    }

    /// <summary>
    /// Fast lookup using pre-built cache instead of scanning entire database.
    /// </summary>
    private static ObjectId FindParentBlockReferenceWithCache(ObjectId blockId, Dictionary<ObjectId, ObjectId> cache)
    {
        if (cache != null && cache.TryGetValue(blockId, out var parentBlockId))
        {
            return parentBlockId;
        }
        return ObjectId.Null;
    }

    /// <summary>
    /// Gets the parent block hierarchy for an entity.
    /// Returns list of block info from outermost to innermost (reversed for display).
    /// Returns empty list if entity is not in a block.
    /// </summary>
    private static List<ParentBlockInfo> GetParentBlockHierarchy(Entity entity, Transaction tr, Dictionary<ObjectId, ObjectId> blockHierarchyCache = null)
    {
        var hierarchy = new List<ParentBlockInfo>();
        _diagParentBlockCallCount++;

        try
        {
            var db = entity.Database;
            var currentBlockId = entity.BlockId;

            // Walk up the block hierarchy
            while (currentBlockId != ObjectId.Null)
            {
                _diagParentBlockIterations++;

                _diagParentBlockGetObjectTimer.Start();
                var btr = tr.GetObject(currentBlockId, OpenMode.ForRead) as BlockTableRecord;
                _diagParentBlockGetObjectTimer.Stop();

                if (btr == null)
                    break;

                // Stop if we reached a layout block (Model or Paper space)
                _diagParentBlockCheckLayoutTimer.Start();
                bool isLayout = btr.IsLayout;
                _diagParentBlockCheckLayoutTimer.Stop();

                if (isLayout)
                    break;

                // Determine block type and name
                _diagParentBlockDetermineTypeTimer.Start();
                string blockName = btr.Name;
                string blockType = "Block Reference";

                if (btr.IsFromExternalReference)
                {
                    blockType = "XRef";
                }
                else if (btr.IsDynamicBlock)
                {
                    blockType = "Dynamic Block";
                }
                else if (btr.IsAnonymous)
                {
                    // Anonymous blocks could be from dynamic blocks
                    // Try to get the original dynamic block name
                    blockType = "Dynamic Block";
                }
                _diagParentBlockDetermineTypeTimer.Stop();

                // Add this block to the hierarchy
                hierarchy.Add(new ParentBlockInfo
                {
                    Name = blockName,
                    Type = blockType
                });

                // Find if this block is itself nested in another block
                // Use cache for fast lookup if available, otherwise fall back to database scan
                _diagParentBlockFindParentTimer.Start();
                if (blockHierarchyCache != null)
                {
                    currentBlockId = FindParentBlockReferenceWithCache(btr.ObjectId, blockHierarchyCache);
                }
                else
                {
                    currentBlockId = FindParentBlockReference(btr, db, tr);
                }
                _diagParentBlockFindParentTimer.Stop();
            }
        }
        catch
        {
            // Return empty list on error
        }

        // Reverse the hierarchy so outermost block is first (ParentBlock1 = main parent)
        hierarchy.Reverse();
        return hierarchy;
    }

    /// <summary>
    /// Finds if a block definition is referenced by a BlockReference that is itself inside another block.
    /// Returns the ObjectId of the parent block, or ObjectId.Null if not nested.
    /// </summary>
    private static ObjectId FindParentBlockReference(BlockTableRecord btr, Database db, Transaction tr)
    {
        try
        {
            _diagFindParentGetBlockTableTimer.Start();
            // Search through all block definitions for references to this block
            var blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            _diagFindParentGetBlockTableTimer.Stop();

            if (blockTable == null)
                return ObjectId.Null;

            _diagFindParentIterateBlocksTimer.Start();
            foreach (ObjectId blockId in blockTable)
            {
                _diagFindParentBlockCount++;

                _diagFindParentGetBlockDefTimer.Start();
                var searchBtr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                _diagFindParentGetBlockDefTimer.Stop();

                if (searchBtr == null || searchBtr.IsLayout)
                    continue; // Skip layout blocks (Model, Paper spaces)

                _diagFindParentIterateEntitiesTimer.Start();
                // Look through entities in this block for a reference to our target block
                foreach (ObjectId entityId in searchBtr)
                {
                    _diagFindParentEntityCount++;

                    _diagFindParentGetEntityTimer.Start();
                    var entity = tr.GetObject(entityId, OpenMode.ForRead);
                    _diagFindParentGetEntityTimer.Stop();

                    if (entity is BlockReference blockRef)
                    {
                        _diagFindParentCheckRefTimer.Start();
                        // Check if this block reference points to our target block
                        if (blockRef.BlockTableRecord == btr.ObjectId)
                        {
                            _diagFindParentCheckRefTimer.Stop();
                            _diagFindParentIterateEntitiesTimer.Stop();
                            _diagFindParentIterateBlocksTimer.Stop();
                            // Found a reference! Return the block this reference is in
                            return searchBtr.ObjectId;
                        }
                        _diagFindParentCheckRefTimer.Stop();
                    }
                }
                _diagFindParentIterateEntitiesTimer.Stop();
            }
            _diagFindParentIterateBlocksTimer.Stop();
        }
        catch
        {
            // Return null on error
        }

        return ObjectId.Null;
    }

    /// <summary>
    /// Processes parent block hierarchy into separate columns.
    /// Only adds columns if there are entities inside blocks.
    /// Uses "Parent Block" for single level, "Parent Block 1", "Parent Block 2", etc. for nested blocks.
    /// Also adds "Parent Block Type" column showing block type (Block Reference, Dynamic Block, XRef).
    /// </summary>
    public static void ProcessParentBlockColumns(List<Dictionary<string, object>> entityData)
    {
        if (entityData == null || entityData.Count == 0)
            return;

        // Determine maximum nesting depth across all entities
        int maxDepth = 0;
        bool hasAnyBlockEntities = false;

        foreach (var entity in entityData)
        {
            if (entity.ContainsKey("_ParentBlocks") && entity["_ParentBlocks"] is List<ParentBlockInfo> parentBlocks)
            {
                if (parentBlocks.Count > 0)
                {
                    hasAnyBlockEntities = true;
                    maxDepth = Math.Max(maxDepth, parentBlocks.Count);
                }
            }
        }

        // Only add columns if there are entities in blocks
        if (!hasAnyBlockEntities)
        {
            // Remove _ParentBlocks from all entities
            foreach (var entity in entityData)
            {
                entity.Remove("_ParentBlocks");
            }
            return;
        }

        // Determine column names based on nesting depth
        var columnNames = new List<string>();
        if (maxDepth == 1)
        {
            columnNames.Add("ParentBlock");
        }
        else
        {
            for (int i = 1; i <= maxDepth; i++)
            {
                columnNames.Add($"ParentBlock{i}");
            }
        }

        // Add parent block columns to each entity
        foreach (var entity in entityData)
        {
            if (entity.ContainsKey("_ParentBlocks") && entity["_ParentBlocks"] is List<ParentBlockInfo> parentBlocks)
            {
                // Populate columns with block names
                for (int i = 0; i < columnNames.Count; i++)
                {
                    entity[columnNames[i]] = i < parentBlocks.Count ? parentBlocks[i].Name : "";
                }

                // Add parent block type column (using the last/innermost block's type - the immediate parent)
                entity["ParentBlockType"] = parentBlocks.Count > 0 ? parentBlocks[parentBlocks.Count - 1].Type : "";

                // Remove internal property
                entity.Remove("_ParentBlocks");
            }
            else
            {
                // Entity not in a block - add empty values
                foreach (var columnName in columnNames)
                {
                    entity[columnName] = "";
                }
                entity["ParentBlockType"] = "";

                // Remove internal property even if type cast failed
                entity.Remove("_ParentBlocks");
            }
        }
    }

    // Diagnostics for performance profiling
    private static System.Diagnostics.Stopwatch _diagGetLayoutNameTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagGetParentBlocksTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagGetEntityAttributesTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagAddBlockAttributesTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagAddExtensionDataTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagGetCategoryTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagGetEntityAreaTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagGetEntityLengthTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagGetEntityElevationTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagGetGeometryPropsTimer = new System.Diagnostics.Stopwatch();

    // Detailed GetParentBlockHierarchy diagnostics
    private static System.Diagnostics.Stopwatch _diagParentBlockGetObjectTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagParentBlockCheckLayoutTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagParentBlockDetermineTypeTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagParentBlockFindParentTimer = new System.Diagnostics.Stopwatch();
    private static long _diagParentBlockCallCount = 0;
    private static long _diagParentBlockIterations = 0;

    // Detailed FindParentBlockReference diagnostics
    private static System.Diagnostics.Stopwatch _diagFindParentGetBlockTableTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagFindParentIterateBlocksTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagFindParentGetBlockDefTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagFindParentIterateEntitiesTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagFindParentGetEntityTimer = new System.Diagnostics.Stopwatch();
    private static System.Diagnostics.Stopwatch _diagFindParentCheckRefTimer = new System.Diagnostics.Stopwatch();
    private static long _diagFindParentBlockCount = 0;
    private static long _diagFindParentEntityCount = 0;

    // Cache build diagnostics
    private static long _diagCacheBuildTime = 0;
    private static int _diagCacheBlocksScanned = 0;
    private static int _diagCacheEntitiesScanned = 0;
    private static int _diagCacheReferencesFound = 0;

    private static long _diagEntityCount = 0;

    public static void ResetDiagnostics()
    {
        _diagGetLayoutNameTimer.Reset();
        _diagGetParentBlocksTimer.Reset();
        _diagGetEntityAttributesTimer.Reset();
        _diagAddBlockAttributesTimer.Reset();
        _diagAddExtensionDataTimer.Reset();
        _diagGetCategoryTimer.Reset();
        _diagGetEntityAreaTimer.Reset();
        _diagGetEntityLengthTimer.Reset();
        _diagGetEntityElevationTimer.Reset();
        _diagGetGeometryPropsTimer.Reset();

        _diagParentBlockGetObjectTimer.Reset();
        _diagParentBlockCheckLayoutTimer.Reset();
        _diagParentBlockDetermineTypeTimer.Reset();
        _diagParentBlockFindParentTimer.Reset();
        _diagParentBlockCallCount = 0;
        _diagParentBlockIterations = 0;

        _diagFindParentGetBlockTableTimer.Reset();
        _diagFindParentIterateBlocksTimer.Reset();
        _diagFindParentGetBlockDefTimer.Reset();
        _diagFindParentIterateEntitiesTimer.Reset();
        _diagFindParentGetEntityTimer.Reset();
        _diagFindParentCheckRefTimer.Reset();
        _diagFindParentBlockCount = 0;
        _diagFindParentEntityCount = 0;

        _diagCacheBuildTime = 0;
        _diagCacheBlocksScanned = 0;
        _diagCacheEntitiesScanned = 0;
        _diagCacheReferencesFound = 0;

        _diagEntityCount = 0;
    }

    public static string GetDiagnosticsSummary()
    {
        if (_diagEntityCount == 0) return "No diagnostics collected";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"GetEntityDataDictionary Performance Breakdown ({_diagEntityCount} entities):");
        sb.AppendLine($"  - GetEntityLayoutName: {_diagGetLayoutNameTimer.ElapsedMilliseconds}ms (avg {_diagGetLayoutNameTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
        sb.AppendLine($"  - GetParentBlockHierarchy: {_diagGetParentBlocksTimer.ElapsedMilliseconds}ms (avg {_diagGetParentBlocksTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
        sb.AppendLine($"  - GetEntityAttributes: {_diagGetEntityAttributesTimer.ElapsedMilliseconds}ms (avg {_diagGetEntityAttributesTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
        sb.AppendLine($"  - AddBlockAttributes: {_diagAddBlockAttributesTimer.ElapsedMilliseconds}ms (avg {_diagAddBlockAttributesTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
        sb.AppendLine($"  - AddExtensionData: {_diagAddExtensionDataTimer.ElapsedMilliseconds}ms (avg {_diagAddExtensionDataTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
        sb.AppendLine($"  - GetEntityCategory: {_diagGetCategoryTimer.ElapsedMilliseconds}ms (avg {_diagGetCategoryTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");

        var geomTotal = _diagGetEntityAreaTimer.ElapsedMilliseconds + _diagGetEntityLengthTimer.ElapsedMilliseconds +
                        _diagGetEntityElevationTimer.ElapsedMilliseconds + _diagGetGeometryPropsTimer.ElapsedMilliseconds;
        if (geomTotal > 0)
        {
            sb.AppendLine($"  GEOMETRY PROPERTIES:");
            sb.AppendLine($"    - GetEntityArea: {_diagGetEntityAreaTimer.ElapsedMilliseconds}ms (avg {_diagGetEntityAreaTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
            sb.AppendLine($"    - GetEntityLength: {_diagGetEntityLengthTimer.ElapsedMilliseconds}ms (avg {_diagGetEntityLengthTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
            sb.AppendLine($"    - GetEntityElevation: {_diagGetEntityElevationTimer.ElapsedMilliseconds}ms (avg {_diagGetEntityElevationTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
            sb.AppendLine($"    - GetEntityGeometryProperties: {_diagGetGeometryPropsTimer.ElapsedMilliseconds}ms (avg {_diagGetGeometryPropsTimer.ElapsedMilliseconds / _diagEntityCount}ms per entity)");
            sb.AppendLine($"    - Geometry total: {geomTotal}ms");
        }

        // Cache build diagnostics
        if (_diagCacheBuildTime > 0)
        {
            sb.AppendLine($"  BLOCK HIERARCHY CACHE:");
            sb.AppendLine($"    - Build time: {_diagCacheBuildTime}ms (ONE-TIME cost)");
            sb.AppendLine($"    - Blocks scanned: {_diagCacheBlocksScanned}");
            sb.AppendLine($"    - Entities scanned: {_diagCacheEntitiesScanned}");
            sb.AppendLine($"    - References found: {_diagCacheReferencesFound}");
        }

        // Parent block hierarchy detailed breakdown
        if (_diagParentBlockCallCount > 0)
        {
            sb.AppendLine($"  PARENT BLOCK HIERARCHY DETAILS:");
            sb.AppendLine($"    - Called {_diagParentBlockCallCount} times with {_diagParentBlockIterations} total iterations");
            sb.AppendLine($"    - GetObject (BlockTableRecord): {_diagParentBlockGetObjectTimer.ElapsedMilliseconds}ms (avg {(_diagParentBlockIterations > 0 ? _diagParentBlockGetObjectTimer.ElapsedMilliseconds / _diagParentBlockIterations : 0)}ms per iteration)");
            sb.AppendLine($"    - Check IsLayout: {_diagParentBlockCheckLayoutTimer.ElapsedMilliseconds}ms");
            sb.AppendLine($"    - Determine block type: {_diagParentBlockDetermineTypeTimer.ElapsedMilliseconds}ms");
            sb.AppendLine($"    - FindParentBlockReference: {_diagParentBlockFindParentTimer.ElapsedMilliseconds}ms (avg {(_diagParentBlockIterations > 0 ? _diagParentBlockFindParentTimer.ElapsedMilliseconds / _diagParentBlockIterations : 0)}ms per iteration)");
            var parentBlockDetailTotal = _diagParentBlockGetObjectTimer.ElapsedMilliseconds + _diagParentBlockCheckLayoutTimer.ElapsedMilliseconds +
                                        _diagParentBlockDetermineTypeTimer.ElapsedMilliseconds + _diagParentBlockFindParentTimer.ElapsedMilliseconds;
            sb.AppendLine($"    - Detailed total: {parentBlockDetailTotal}ms");

            // FindParentBlockReference detailed breakdown
            if (_diagFindParentBlockCount > 0)
            {
                sb.AppendLine($"  FINDPARENTBLOCKREFERENCE DETAILS:");
                sb.AppendLine($"    - Called {_diagParentBlockIterations} times (once per hierarchy iteration)");
                sb.AppendLine($"    - Total block definitions checked: {_diagFindParentBlockCount}");
                sb.AppendLine($"    - Total entities checked: {_diagFindParentEntityCount}");
                sb.AppendLine($"    - Avg blocks per call: {(_diagParentBlockIterations > 0 ? _diagFindParentBlockCount / _diagParentBlockIterations : 0)}");
                sb.AppendLine($"    - Avg entities per call: {(_diagParentBlockIterations > 0 ? _diagFindParentEntityCount / _diagParentBlockIterations : 0)}");
                sb.AppendLine($"    BREAKDOWN:");
                sb.AppendLine($"      - Get BlockTable: {_diagFindParentGetBlockTableTimer.ElapsedMilliseconds}ms");
                sb.AppendLine($"      - Iterate blocks (outer loop): {_diagFindParentIterateBlocksTimer.ElapsedMilliseconds}ms");
                sb.AppendLine($"      - Get block definitions: {_diagFindParentGetBlockDefTimer.ElapsedMilliseconds}ms (avg {(_diagFindParentBlockCount > 0 ? _diagFindParentGetBlockDefTimer.ElapsedMilliseconds / _diagFindParentBlockCount : 0)}ms per block)");
                sb.AppendLine($"      - Iterate entities (inner loop): {_diagFindParentIterateEntitiesTimer.ElapsedMilliseconds}ms");
                sb.AppendLine($"      - Get entity from block: {_diagFindParentGetEntityTimer.ElapsedMilliseconds}ms (avg {(_diagFindParentEntityCount > 0 ? _diagFindParentGetEntityTimer.ElapsedMilliseconds / _diagFindParentEntityCount : 0)}ms per entity)");
                sb.AppendLine($"      - Check if BlockReference matches: {_diagFindParentCheckRefTimer.ElapsedMilliseconds}ms");
                var findParentTotal = _diagFindParentGetBlockTableTimer.ElapsedMilliseconds + _diagFindParentGetBlockDefTimer.ElapsedMilliseconds +
                                     _diagFindParentGetEntityTimer.ElapsedMilliseconds + _diagFindParentCheckRefTimer.ElapsedMilliseconds;
                sb.AppendLine($"      - Accounted time: {findParentTotal}ms");
                sb.AppendLine($"      - Unaccounted (loop overhead): {_diagParentBlockFindParentTimer.ElapsedMilliseconds - findParentTotal}ms");
            }
        }

        var totalTracked = _diagGetLayoutNameTimer.ElapsedMilliseconds + _diagGetParentBlocksTimer.ElapsedMilliseconds +
                          _diagGetEntityAttributesTimer.ElapsedMilliseconds + _diagAddBlockAttributesTimer.ElapsedMilliseconds +
                          _diagAddExtensionDataTimer.ElapsedMilliseconds + _diagGetCategoryTimer.ElapsedMilliseconds + geomTotal;
        sb.AppendLine($"  - Total tracked: {totalTracked}ms");
        return sb.ToString();
    }

    public static Dictionary<string, object> GetEntityDataDictionary(DBObject entity, string documentPath, string spaceName, bool includeProperties, Transaction tr = null, Dictionary<ObjectId, ObjectId> blockHierarchyCache = null)
    {
        _diagEntityCount++;

        // Include parent blocks ONLY when SelectInBlocksMode is enabled
        // (SelectInBlocksMode controls whether we care about block nesting)
        bool includeParentBlocks = SelectInBlocksMode.IsEnabled();

        string entityName = "";
        string layer = "";
        string color = "";
        string lineType = "";
        string layoutName = "";

        if (entity is Entity ent)
        {
            layer = ent.Layer;
            color = ent.Color.ToString();
            lineType = ent.Linetype;

            _diagGetLayoutNameTimer.Start();
            layoutName = GetEntityLayoutName(ent);
            _diagGetLayoutNameTimer.Stop();

            // Get entity-specific name
            if (entity is BlockReference br)
            {
                if (tr != null)
                {
                    var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    entityName = btr?.Name ?? "Block";
                }
                else
                {
                    // Fallback if no transaction provided
                    using (var tempTr = br.Database.TransactionManager.StartTransaction())
                    {
                        var btr = tempTr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        entityName = btr?.Name ?? "Block";
                        tempTr.Commit();
                    }
                }
            }
            else if (entity is MText mtext)
            {
                entityName = mtext.Contents.Length > 30 ?
                    mtext.Contents.Substring(0, 30) + "..." :
                    mtext.Contents;
            }
            else if (entity is DBText text)
            {
                entityName = text.TextString.Length > 30 ?
                    text.TextString.Substring(0, 30) + "..." :
                    text.TextString;
            }
            else if (entity is Dimension dim)
            {
                entityName = dim.DimensionText;
            }
        }
        else if (entity is Layout layout)
        {
            // For Layout entities, show the layout name (what appears in AutoCAD tabs)
            entityName = layout.LayoutName;
        }
        else if (entity is LayerTableRecord layerRecord)
        {
            // For Layer entities, show the layer name
            entityName = layerRecord.Name;
            // Override the default layer/color/linetype properties with layer-specific values
            layer = layerRecord.Name;
            color = layerRecord.Color.ToString();
            lineType = GetLayerLinetype(layerRecord);
        }

        _diagGetCategoryTimer.Start();
        var category = GetEntityCategory(entity);
        _diagGetCategoryTimer.Stop();

        var data = new Dictionary<string, object>
        {
            ["Name"] = entityName,
            ["Category"] = category,
            ["Layer"] = layer,
            ["Color"] = color,
            ["LineType"] = lineType,
            ["Layout"] = layoutName,
            ["Contents"] = GetEntityContents(entity),
            ["DocumentPath"] = documentPath,
            ["DocumentName"] = Path.GetFileName(documentPath),
            ["Handle"] = entity.Handle.ToString(),
            ["Id"] = entity.ObjectId.Handle.Value,
            ["IsExternal"] = false,
            ["ObjectId"] = entity.ObjectId, // Store for selection
            ["DynamicBlockName"] = "", // Will be populated for dynamic blocks
            ["XrefPath"] = "" // Will be populated for xrefs
        };

        data["DisplayName"] = !string.IsNullOrEmpty(entityName) ? entityName : data["Category"].ToString();

        // Add parent block hierarchy for entities inside blocks (only when SelectInBlocksMode is enabled)
        if (includeParentBlocks && tr != null && entity is Entity entityObj)
        {
            _diagGetParentBlocksTimer.Start();

            // OPTIMIZATION: Only check parent blocks if entity is actually in a block (not in a layout)
            var blockId = entityObj.BlockId;
            var parentBlocks = new List<ParentBlockInfo>();

            if (blockId != ObjectId.Null)
            {
                try
                {
                    var btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                    // Only compute hierarchy if entity is NOT in a layout block
                    if (btr != null && !btr.IsLayout)
                    {
                        parentBlocks = GetParentBlockHierarchy(entityObj, tr, blockHierarchyCache);
                    }
                }
                catch
                {
                    // If we can't determine, assume no parent blocks
                }
            }

            _diagGetParentBlocksTimer.Stop();
            data["_ParentBlocks"] = parentBlocks; // Internal property for processing
        }
        else
        {
            data["_ParentBlocks"] = new List<ParentBlockInfo>();
        }

        // For BlockReferences, add dynamic block parent name if applicable
        if (entity is BlockReference blockRef)
        {
            if (tr != null)
            {
                var dynamicBlockTableRecordId = blockRef.DynamicBlockTableRecord;
                if (dynamicBlockTableRecordId != ObjectId.Null && dynamicBlockTableRecordId != blockRef.BlockTableRecord)
                {
                    // This is a dynamic block - get the parent block name
                    var dynamicBtr = tr.GetObject(dynamicBlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    data["DynamicBlockName"] = dynamicBtr?.Name ?? "";
                }

                // For xrefs, also store the xref path
                var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (btr != null && btr.IsFromExternalReference)
                {
                    data["XrefPath"] = btr.PathName ?? "";
                }
            }
            else
            {
                // Fallback if no transaction provided
                using (var tempTr = blockRef.Database.TransactionManager.StartTransaction())
                {
                    var dynamicBlockTableRecordId = blockRef.DynamicBlockTableRecord;
                    if (dynamicBlockTableRecordId != ObjectId.Null && dynamicBlockTableRecordId != blockRef.BlockTableRecord)
                    {
                        // This is a dynamic block - get the parent block name
                        var dynamicBtr = tempTr.GetObject(dynamicBlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                        data["DynamicBlockName"] = dynamicBtr?.Name ?? "";
                    }

                    // For xrefs, also store the xref path
                    var btr = tempTr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null && btr.IsFromExternalReference)
                    {
                        data["XrefPath"] = btr.PathName ?? "";
                    }

                    tempTr.Commit();
                }
            }
        }

        // Add all attributes from unified attribute system (including Tags)
        try
        {
            _diagGetEntityAttributesTimer.Start();
            var attributes = EntityAttributes.GetEntityAttributes(entity.ObjectId, entity.Database);
            _diagGetEntityAttributesTimer.Stop();

            foreach (var attr in attributes)
            {
                // Tags get special treatment (no prefix) for backward compatibility
                // All other attributes get "attr_" prefix
                var columnName = attr.Key.Equals("Tags", StringComparison.OrdinalIgnoreCase)
                    ? "Tags"
                    : $"attr_{attr.Key}";
                data[columnName] = attr.Value.AsString();
            }
        }
        catch
        {
            _diagGetEntityAttributesTimer.Stop();
            // Skip if attributes can't be read
        }

        // Tags column will be split into tag_1, tag_2, tag_3, etc. during post-processing
        // Ensure Tags exists as empty string for entities without tags (for consistent processing)
        if (!data.ContainsKey("Tags"))
        {
            data["Tags"] = "";
        }

        // Add space information if available
        if (!string.IsNullOrEmpty(spaceName))
        {
            data["Space"] = spaceName;
        }

        // Add block attributes if entity is a block reference
        if (entity is BlockReference)
        {
            _diagAddBlockAttributesTimer.Start();
            AddBlockAttributes((BlockReference)entity, data);
            _diagAddBlockAttributesTimer.Stop();
        }

        // Add plot settings for Layout entities
        if (entity is Layout layoutEntity)
        {
            AddLayoutPlotSettings(layoutEntity, data);
        }

        // Add layer-specific properties for Layer entities
        if (entity is LayerTableRecord layerTableRecord)
        {
            AddLayerProperties(layerTableRecord, data);
        }

        // Add XData and extension dictionary data
        _diagAddExtensionDataTimer.Start();
        AddExtensionData(entity, data);
        _diagAddExtensionDataTimer.Stop();

        // Include properties if requested
        if (includeProperties && entity is Entity entityWithProps)
        {
            try
            {
                // Add common AutoCAD properties
                _diagGetEntityAreaTimer.Start();
                data["Area"] = GetEntityArea(entityWithProps);
                _diagGetEntityAreaTimer.Stop();

                _diagGetEntityLengthTimer.Start();
                data["Length"] = GetEntityLength(entityWithProps);
                _diagGetEntityLengthTimer.Stop();

                _diagGetEntityElevationTimer.Start();
                data["Elevation"] = GetEntityElevation(entityWithProps);
                _diagGetEntityElevationTimer.Stop();

                // Add geometry properties
                _diagGetGeometryPropsTimer.Start();
                var geometryProps = GetEntityGeometryProperties(entityWithProps);
                foreach (var prop in geometryProps)
                {
                    data[prop.Key] = prop.Value;
                }
                _diagGetGeometryPropsTimer.Stop();
            }
            catch
            {
                _diagGetEntityAreaTimer.Stop();
                _diagGetEntityLengthTimer.Stop();
                _diagGetEntityElevationTimer.Stop();
                _diagGetGeometryPropsTimer.Stop();
                /* Skip if properties can't be read */
            }
        }

        return data;
    }

    private static List<ObjectId> GatherEntitiesFromScope(Database db, SelectionScope scope)
    {
        var entities = new List<ObjectId>();

        using (var tr = db.TransactionManager.StartTransaction())
        {
            switch (scope)
            {
                case SelectionScope.view:
                    var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id in currentSpace)
                    {
                        entities.Add(id);
                    }
                    break;

                case SelectionScope.document:
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                        foreach (ObjectId id in btr)
                        {
                            entities.Add(id);
                        }
                    }
                    break;

                default:
                    // For session scope - fall back to Document scope for now
                    goto case SelectionScope.document;
            }

            tr.Commit();
        }

        return entities;
    }

    private static string GetEntityContents(DBObject entity)
    {
        if (entity is MText mtext)
        {
            return mtext.Contents;
        }
        else if (entity is DBText text)
        {
            return text.TextString;
        }
        else if (entity is Dimension dim)
        {
            return dim.DimensionText;
        }
        else if (entity is Layout layout)
        {
            return layout.LayoutName;
        }

        return "";
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

    private static string GetEntityCategory(DBObject entity)
    {
        // Use the same categorization logic as select-by-categories.cs
        string typeName = entity.GetType().Name;

        // Check for revision clouds before checking for regular polylines
        // Revision clouds are polylines with XData registered application name "RevcloudProps"
        if (entity is Polyline polyline && IsRevisionCloud(polyline))
            return "Revision Cloud";

        if (entity is LayerTableRecord)
            return "Layer";
        else if (entity is Layout)
            return "Layout";
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
        else
        {
            return typeName.Replace("Autodesk.AutoCAD.", "");
        }
    }

    private static string GetEntityArea(Entity entity)
    {
        try
        {
            if (entity is Autodesk.AutoCAD.DatabaseServices.Region region)
            {
                return region.Area.ToString("F3");
            }
            else if (entity is Hatch hatch)
            {
                return hatch.Area.ToString("F3");
            }
            else if (entity is Circle circle)
            {
                return (Math.PI * circle.Radius * circle.Radius).ToString("F3");
            }
            else if (entity is Polyline pline && pline.Closed)
            {
                return pline.Area.ToString("F3");
            }
        }
        catch { }
        return "";
    }

    private static string GetEntityLength(Entity entity)
    {
        try
        {
            if (entity is Curve curve)
            {
                return curve.GetDistanceAtParameter(curve.EndParam).ToString("F3");
            }
        }
        catch { }
        return "";
    }

    private static string GetEntityElevation(Entity entity)
    {
        try
        {
            if (entity is Entity ent)
            {
                // return ent.Elevation.ToString("F3"); // Not available on base Entity
            }
        }
        catch { }
        return "";
    }

    private static Dictionary<string, object> GetEntityGeometryProperties(Entity entity)
    {
        var props = new Dictionary<string, object>();

        try
        {
            // Get center coordinates and dimensions based on entity type
            if (entity is Circle circle)
            {
                props["CenterX"] = circle.Center.X.ToString("F3");
                props["CenterY"] = circle.Center.Y.ToString("F3");
                props["CenterZ"] = circle.Center.Z.ToString("F3");
                props["Radius"] = circle.Radius.ToString("F3");
                props["Diameter"] = (circle.Radius * 2).ToString("F3");
                props["Circumference"] = (2 * Math.PI * circle.Radius).ToString("F3");
            }
            else if (entity is Arc arc)
            {
                props["CenterX"] = arc.Center.X.ToString("F3");
                props["CenterY"] = arc.Center.Y.ToString("F3");
                props["CenterZ"] = arc.Center.Z.ToString("F3");
                props["Radius"] = arc.Radius.ToString("F3");
                props["StartAngle"] = (arc.StartAngle * 180 / Math.PI).ToString("F1"); // Convert to degrees
                props["EndAngle"] = (arc.EndAngle * 180 / Math.PI).ToString("F1");
                props["ArcLength"] = arc.Length.ToString("F3");
            }
            else if (entity is Line line)
            {
                var midPoint = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                props["CenterX"] = midPoint.X.ToString("F3");
                props["CenterY"] = midPoint.Y.ToString("F3");
                props["CenterZ"] = midPoint.Z.ToString("F3");
                props["StartX"] = line.StartPoint.X.ToString("F3");
                props["StartY"] = line.StartPoint.Y.ToString("F3");
                props["StartZ"] = line.StartPoint.Z.ToString("F3");
                props["EndX"] = line.EndPoint.X.ToString("F3");
                props["EndY"] = line.EndPoint.Y.ToString("F3");
                props["EndZ"] = line.EndPoint.Z.ToString("F3");
                props["LineLength"] = line.Length.ToString("F3");
                props["Angle"] = (line.Angle * 180 / Math.PI).ToString("F1");
            }
            else if (entity is Polyline pline)
            {
                var bounds = pline.GeometricExtents;
                var center = bounds.MinPoint + (bounds.MaxPoint - bounds.MinPoint) * 0.5;
                props["CenterX"] = center.X.ToString("F3");
                props["CenterY"] = center.Y.ToString("F3");
                props["CenterZ"] = center.Z.ToString("F3");
                props["Width"] = (bounds.MaxPoint.X - bounds.MinPoint.X).ToString("F3");
                props["Height"] = (bounds.MaxPoint.Y - bounds.MinPoint.Y).ToString("F3");
                props["Depth"] = (bounds.MaxPoint.Z - bounds.MinPoint.Z).ToString("F3");
                props["IsClosed"] = pline.Closed.ToString();
                props["VertexCount"] = pline.NumberOfVertices.ToString();
                if (pline.HasWidth)
                {
                    props["ConstantWidth"] = pline.ConstantWidth.ToString("F3");
                }
            }
            else if (entity is Ellipse ellipse)
            {
                props["CenterX"] = ellipse.Center.X.ToString("F3");
                props["CenterY"] = ellipse.Center.Y.ToString("F3");
                props["CenterZ"] = ellipse.Center.Z.ToString("F3");
                props["MajorRadius"] = ellipse.MajorRadius.ToString("F3");
                props["MinorRadius"] = ellipse.MinorRadius.ToString("F3");
                props["RadiusRatio"] = ellipse.RadiusRatio.ToString("F3");
            }
            else if (entity is Spline spline)
            {
                var bounds = spline.GeometricExtents;
                var center = bounds.MinPoint + (bounds.MaxPoint - bounds.MinPoint) * 0.5;
                props["CenterX"] = center.X.ToString("F3");
                props["CenterY"] = center.Y.ToString("F3");
                props["CenterZ"] = center.Z.ToString("F3");
                props["Width"] = (bounds.MaxPoint.X - bounds.MinPoint.X).ToString("F3");
                props["Height"] = (bounds.MaxPoint.Y - bounds.MinPoint.Y).ToString("F3");
                props["Depth"] = (bounds.MaxPoint.Z - bounds.MinPoint.Z).ToString("F3");
                props["Degree"] = spline.Degree.ToString();
                props["IsPeriodic"] = spline.IsPeriodic.ToString();
            }
            else if (entity is BlockReference blockRef)
            {
                props["CenterX"] = blockRef.Position.X.ToString("F3");
                props["CenterY"] = blockRef.Position.Y.ToString("F3");
                props["CenterZ"] = blockRef.Position.Z.ToString("F3");
                props["ScaleX"] = blockRef.ScaleFactors.X.ToString("F3");
                props["ScaleY"] = blockRef.ScaleFactors.Y.ToString("F3");
                props["ScaleZ"] = blockRef.ScaleFactors.Z.ToString("F3");
                props["Rotation"] = (blockRef.Rotation * 180 / Math.PI).ToString("F1");

                // Get bounds if possible
                try
                {
                    var bounds = blockRef.GeometricExtents;
                    props["Width"] = (bounds.MaxPoint.X - bounds.MinPoint.X).ToString("F3");
                    props["Height"] = (bounds.MaxPoint.Y - bounds.MinPoint.Y).ToString("F3");
                    props["Depth"] = (bounds.MaxPoint.Z - bounds.MinPoint.Z).ToString("F3");
                }
                catch { /* Bounds may not be available for some blocks */ }
            }
            else if (entity is DBText text)
            {
                props["CenterX"] = text.Position.X.ToString("F3");
                props["CenterY"] = text.Position.Y.ToString("F3");
                props["CenterZ"] = text.Position.Z.ToString("F3");
                props["TextHeight"] = text.Height.ToString("F3");
                props["Rotation"] = (text.Rotation * 180 / Math.PI).ToString("F1");
                props["WidthFactor"] = text.WidthFactor.ToString("F3");
            }
            else if (entity is MText mtext)
            {
                props["CenterX"] = mtext.Location.X.ToString("F3");
                props["CenterY"] = mtext.Location.Y.ToString("F3");
                props["CenterZ"] = mtext.Location.Z.ToString("F3");
                props["TextHeight"] = mtext.TextHeight.ToString("F3");
                props["Width"] = mtext.Width.ToString("F3");
                props["Rotation"] = (mtext.Rotation * 180 / Math.PI).ToString("F1");

                // Get actual height if possible
                try
                {
                    var bounds = mtext.GeometricExtents;
                    props["ActualHeight"] = (bounds.MaxPoint.Y - bounds.MinPoint.Y).ToString("F3");
                    props["ActualWidth"] = (bounds.MaxPoint.X - bounds.MinPoint.X).ToString("F3");
                }
                catch { /* Bounds may not be available */ }
            }
            else if (entity is Hatch hatch)
            {
                try
                {
                    var bounds = hatch.GeometricExtents;
                    var center = bounds.MinPoint + (bounds.MaxPoint - bounds.MinPoint) * 0.5;
                    props["CenterX"] = center.X.ToString("F3");
                    props["CenterY"] = center.Y.ToString("F3");
                    props["CenterZ"] = center.Z.ToString("F3");
                    props["Width"] = (bounds.MaxPoint.X - bounds.MinPoint.X).ToString("F3");
                    props["Height"] = (bounds.MaxPoint.Y - bounds.MinPoint.Y).ToString("F3");
                    props["PatternName"] = hatch.PatternName ?? "";
                    props["PatternScale"] = hatch.PatternScale.ToString("F3");
                    props["NumLoops"] = hatch.NumberOfLoops.ToString();
                }
                catch { /* Some hatches may not have valid bounds */ }
            }
            else if (entity is Dimension dim)
            {
                try
                {
                    var bounds = dim.GeometricExtents;
                    var center = bounds.MinPoint + (bounds.MaxPoint - bounds.MinPoint) * 0.5;
                    props["CenterX"] = center.X.ToString("F3");
                    props["CenterY"] = center.Y.ToString("F3");
                    props["CenterZ"] = center.Z.ToString("F3");
                    props["Width"] = (bounds.MaxPoint.X - bounds.MinPoint.X).ToString("F3");
                    props["Height"] = (bounds.MaxPoint.Y - bounds.MinPoint.Y).ToString("F3");
                    props["TextHeight"] = ""; // Text height not directly accessible from base Dimension
                    props["Measurement"] = dim.Measurement.ToString("F3");
                }
                catch { /* Some dimensions may not have valid bounds */ }
            }
            else if (entity is DBPoint point)
            {
                props["CenterX"] = point.Position.X.ToString("F3");
                props["CenterY"] = point.Position.Y.ToString("F3");
                props["CenterZ"] = point.Position.Z.ToString("F3");
            }
            else
            {
                // For other entity types, try to get geometric extents
                try
                {
                    var bounds = entity.GeometricExtents;
                    var center = bounds.MinPoint + (bounds.MaxPoint - bounds.MinPoint) * 0.5;
                    props["CenterX"] = center.X.ToString("F3");
                    props["CenterY"] = center.Y.ToString("F3");
                    props["CenterZ"] = center.Z.ToString("F3");
                    props["Width"] = (bounds.MaxPoint.X - bounds.MinPoint.X).ToString("F3");
                    props["Height"] = (bounds.MaxPoint.Y - bounds.MinPoint.Y).ToString("F3");
                    props["Depth"] = (bounds.MaxPoint.Z - bounds.MinPoint.Z).ToString("F3");
                }
                catch
                {
                    // Entity doesn't have geometric extents or they can't be calculated
                    props["CenterX"] = "";
                    props["CenterY"] = "";
                    props["CenterZ"] = "";
                    props["Width"] = "";
                    props["Height"] = "";
                    props["Depth"] = "";
                }
            }
        }
        catch { /* Skip if geometry properties can't be read */ }

        return props;
    }

    private static void AddBlockAttributes(BlockReference blockRef, Dictionary<string, object> data)
    {
        try
        {
            using (var tr = blockRef.Database.TransactionManager.StartTransaction())
            {
                var hasAttributes = false;
                foreach (ObjectId attId in blockRef.AttributeCollection)
                {
                    var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (attRef != null)
                    {
                        hasAttributes = true;
                        var columnName = $"attr_{attRef.Tag.ToLower()}";
                        data[columnName] = attRef.TextString;
                    }
                }

                // If no attributes found but block definition has attribute definitions, show empty values
                if (!hasAttributes)
                {
                    var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null)
                    {
                        foreach (ObjectId id in btr)
                        {
                            var attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                            if (attDef != null)
                            {
                                var columnName = $"attr_{attDef.Tag.ToLower()}";
                                data[columnName] = "";
                            }
                        }
                    }
                }

                tr.Commit();
            }
        }
        catch
        {
            // Skip if attributes can't be read
        }
    }

    private static string GetEntityLayoutName(Entity entity)
    {
        try
        {
            using (var tr = entity.Database.TransactionManager.StartTransaction())
            {
                // Get the layout dictionary
                var layoutDict = (DBDictionary)tr.GetObject(entity.Database.LayoutDictionaryId, OpenMode.ForRead);

                // Check each layout to see if it contains this entity
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                    // Check if this entity belongs to this layout's block table record
                    if (entity.BlockId == layout.BlockTableRecordId)
                    {
                        tr.Commit();
                        return layout.LayoutName;
                    }
                }

                tr.Commit();
            }
        }
        catch
        {
            // If we can't determine the layout, return empty string
        }

        return "";
    }

    private static void AddLayoutPlotSettings(Layout layout, Dictionary<string, object> data)
    {
        try
        {
            using (var tr = layout.Database.TransactionManager.StartTransaction())
            {
                var plotSettings = tr.GetObject(layout.ObjectId, OpenMode.ForRead) as PlotSettings;
                if (plotSettings != null)
                {
                    // Paper size
                    data["PaperSize"] = plotSettings.CanonicalMediaName ?? "";

                    // Plot style table
                    data["PlotStyleTable"] = plotSettings.CurrentStyleSheet ?? "";

                    // Drawing orientation
                    data["PlotRotation"] = plotSettings.PlotRotation.ToString();

                    // Plot device/printer
                    data["PlotConfigurationName"] = plotSettings.PlotConfigurationName ?? "";

                    // Plot scale
                    var scale = plotSettings.CustomPrintScale;
                    data["PlotScale"] = $"{scale.Numerator}:{scale.Denominator}";

                    // Plot type
                    data["PlotType"] = plotSettings.PlotType.ToString();

                    // Plot centered
                    data["PlotCentered"] = plotSettings.PlotCentered.ToString();

                    // Scale lineweights
                    data["ScaleLineweights"] = plotSettings.ScaleLineweights.ToString();

                    // Print lineweights
                    data["PrintLineweights"] = plotSettings.PrintLineweights.ToString();

                    // Plot hidden
                    data["PlotHidden"] = plotSettings.PlotHidden.ToString();

                    // Plot transparency
                    data["PlotTransparency"] = plotSettings.PlotTransparency.ToString();

                    // Standard scale
                    data["UseStandardScale"] = plotSettings.UseStandardScale.ToString();
                    if (plotSettings.UseStandardScale)
                    {
                        data["StandardScale"] = plotSettings.StdScale.ToString();
                    }
                    else
                    {
                        data["StandardScale"] = "Custom";
                    }

                    // Plot area
                    if (plotSettings.PlotViewportBorders)
                    {
                        data["PlotViewportBorders"] = "True";
                    }
                    else
                    {
                        data["PlotViewportBorders"] = "False";
                    }

                    // Plot margins (read-only property)
                    data["PlotMargins"] = "Read-only";

                    // Plot paper units
                    data["PlotPaperUnits"] = plotSettings.PlotPaperUnits.ToString();
                }
                tr.Commit();
            }
        }
        catch
        {
            // If we can't read plot settings, add empty values
            data["PaperSize"] = "";
            data["PlotStyleTable"] = "";
            data["PlotRotation"] = "";
            data["PlotConfigurationName"] = "";
            data["PlotScale"] = "";
            data["PlotType"] = "";
            data["PlotCentered"] = "";
            data["ScaleLineweights"] = "";
            data["PrintLineweights"] = "";
            data["PlotHidden"] = "";
            data["PlotTransparency"] = "";
            data["UseStandardScale"] = "";
            data["StandardScale"] = "";
            data["PlotViewportBorders"] = "";
            data["PlotMargins"] = "";
            data["PlotPaperUnits"] = "";
        }
    }

    private static string GetLayerLinetype(LayerTableRecord layer)
    {
        try
        {
            using (var tr = layer.Database.TransactionManager.StartTransaction())
            {
                var linetypeId = layer.LinetypeObjectId;
                if (linetypeId != ObjectId.Null)
                {
                    var linetype = tr.GetObject(linetypeId, OpenMode.ForRead) as LinetypeTableRecord;
                    if (linetype != null)
                    {
                        return linetype.Name;
                    }
                }
                tr.Commit();
            }
        }
        catch
        {
            // If we can't read the linetype, return empty string
        }

        return "";
    }

    private static void AddLayerProperties(LayerTableRecord layer, Dictionary<string, object> data)
    {
        try
        {
            // Layer state properties
            data["IsFrozen"] = layer.IsFrozen.ToString();
            data["IsLocked"] = layer.IsLocked.ToString();
            data["IsOff"] = layer.IsOff.ToString();
            data["IsPlottable"] = layer.IsPlottable.ToString();

            // Lineweight
            data["LineWeight"] = layer.LineWeight.ToString();

            // Transparency
            data["Transparency"] = layer.Transparency.ToString();

            // Plot style name (if available)
            try
            {
                data["PlotStyleName"] = layer.PlotStyleName ?? "";
            }
            catch
            {
                data["PlotStyleName"] = "";
            }

            // Material (if available)
            try
            {
                if (layer.MaterialId != ObjectId.Null)
                {
                    using (var tr = layer.Database.TransactionManager.StartTransaction())
                    {
                        var material = tr.GetObject(layer.MaterialId, OpenMode.ForRead) as Material;
                        if (material != null)
                        {
                            data["Material"] = material.Name;
                        }
                        else
                        {
                            data["Material"] = "";
                        }
                        tr.Commit();
                    }
                }
                else
                {
                    data["Material"] = "";
                }
            }
            catch
            {
                data["Material"] = "";
            }

            // Description
            data["Description"] = layer.Description ?? "";

            // ViewportVisibilityDefault
            data["ViewportVisibilityDefault"] = layer.ViewportVisibilityDefault.ToString();

            // IsInUse (read-only check if layer is used by any entities)
            data["IsInUse"] = layer.IsUsed.ToString();

            // IsReferenced (for xref layers)
            data["IsReferenced"] = layer.IsDependent.ToString();
        }
        catch
        {
            // If we can't read layer properties, add empty values
            data["IsFrozen"] = "";
            data["IsLocked"] = "";
            data["IsOff"] = "";
            data["IsPlottable"] = "";
            data["LineWeight"] = "";
            data["Transparency"] = "";
            data["PlotStyleName"] = "";
            data["Material"] = "";
            data["Description"] = "";
            data["ViewportVisibilityDefault"] = "";
            data["IsInUse"] = "";
            data["IsReferenced"] = "";
        }
    }

    private static void AddExtensionData(DBObject entity, Dictionary<string, object> data)
    {
        try
        {
            // Add XData
            var xData = entity.XData;
            if (xData != null)
            {
                var valuesByApp = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                string currentApp = null;

                foreach (TypedValue typedValue in xData)
                {
                    if (typedValue.TypeCode == (int)DxfCode.ExtendedDataRegAppName)
                    {
                        currentApp = typedValue.Value?.ToString();

                        if (string.IsNullOrWhiteSpace(currentApp))
                        {
                            currentApp = null;
                            continue;
                        }

                        if (!valuesByApp.ContainsKey(currentApp))
                        {
                            valuesByApp[currentApp] = new List<string>();
                        }
                    }
                    else if (currentApp != null)
                    {
                        var formatted = FormatXDataValue(typedValue);
                        if (!string.IsNullOrEmpty(formatted))
                        {
                            valuesByApp[currentApp].Add(formatted);
                        }
                    }
                }

                var usedColumnNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var kvp in valuesByApp)
                {
                    if (!kvp.Value.Any())
                    {
                        continue;
                    }

                    var baseColumnName = $"xdata_{SanitizeXDataAppName(kvp.Key)}";
                    var columnName = baseColumnName;
                    var counter = 2;

                    while (usedColumnNames.Contains(columnName) || data.ContainsKey(columnName))
                    {
                        columnName = $"{baseColumnName}_{counter}";
                        counter++;
                    }

                    usedColumnNames.Add(columnName);
                    data[columnName] = string.Join("; ", kvp.Value);
                }
            }

            // Add Extension Dictionary data
            if (entity.ExtensionDictionary != ObjectId.Null)
            {
                using (var tr = entity.Database.TransactionManager.StartTransaction())
                {
                    var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    if (extDict != null && extDict.Count > 0)
                    {
                        var dictKeys = new List<string>();
                        var dictValues = new List<string>();

                        foreach (DBDictionaryEntry entry in extDict)
                        {
                            dictKeys.Add(entry.Key);
                            try
                            {
                                var dictObj = tr.GetObject(entry.Value, OpenMode.ForRead);
                                if (dictObj != null)
                                {
                                    // Try to get meaningful data from common dictionary objects
                                    if (dictObj is Xrecord xrec)
                                    {
                                        var xrecData = new List<string>();
                                        foreach (TypedValue tv in xrec.Data)
                                        {
                                            if (tv.Value != null)
                                            {
                                                xrecData.Add(tv.Value.ToString());
                                            }
                                        }
                                        if (xrecData.Any())
                                        {
                                            dictValues.Add(string.Join("|", xrecData.Take(2))); // Limit to first 2 values
                                        }
                                    }
                                    else
                                    {
                                        dictValues.Add(dictObj.GetType().Name);
                                    }
                                }
                            }
                            catch
                            {
                                dictValues.Add("(Error reading)");
                            }
                        }

                        if (dictKeys.Any())
                        {
                            data["ext_dict_keys"] = string.Join(", ", dictKeys);
                        }
                        if (dictValues.Any())
                        {
                            data["ext_dict_values"] = string.Join(", ", dictValues);
                        }
                    }
                    tr.Commit();
                }
            }
        }
        catch
        {
            // Skip if extension data can't be read
        }
    }

    private static string FormatXDataValue(TypedValue typedValue)
    {
        if (typedValue.Value == null)
        {
            return string.Empty;
        }

        switch (typedValue.Value)
        {
            case string text:
                return text;
            case double real:
                return real.ToString("G", CultureInfo.InvariantCulture);
            case float singleValue:
                return singleValue.ToString("G", CultureInfo.InvariantCulture);
            case short int16:
                return int16.ToString(CultureInfo.InvariantCulture);
            case int int32:
                return int32.ToString(CultureInfo.InvariantCulture);
            case long int64:
                return int64.ToString(CultureInfo.InvariantCulture);
            case Point3d point3d:
                return string.Format(CultureInfo.InvariantCulture, "{0:G}, {1:G}, {2:G}", point3d.X, point3d.Y, point3d.Z);
            case Point2d point2d:
                return string.Format(CultureInfo.InvariantCulture, "{0:G}, {1:G}", point2d.X, point2d.Y);
            case bool boolean:
                return boolean ? "True" : "False";
            case Autodesk.AutoCAD.DatabaseServices.Handle handle:
                return handle.ToString();
            case ObjectId objectId:
                return objectId.Handle.ToString();
            case byte[] bytes:
                return BitConverter.ToString(bytes);
            default:
                return typedValue.Value.ToString();
        }
    }

    private static string SanitizeXDataAppName(string regAppName)
    {
        if (string.IsNullOrWhiteSpace(regAppName))
        {
            return "unnamed";
        }

        var sanitized = new string(regAppName.Select(ch =>
            char.IsLetterOrDigit(ch) ? char.ToLowerInvariant(ch) : '_').ToArray()).Trim('_');

        return string.IsNullOrEmpty(sanitized) ? "unnamed" : sanitized;
    }

    public static bool IsGeometryProperty(string propertyName)
    {
        var geometryProps = new HashSet<string>
        {
            "CenterX", "CenterY", "CenterZ", "Width", "Height", "Depth",
            "Radius", "Diameter", "Circumference", "MajorRadius", "MinorRadius", "RadiusRatio",
            "StartX", "StartY", "StartZ", "EndX", "EndY", "EndZ", "LineLength", "ArcLength",
            "StartAngle", "EndAngle", "Angle", "Rotation", "ScaleX", "ScaleY", "ScaleZ",
            "TextHeight", "ActualHeight", "ActualWidth", "WidthFactor", "PatternScale",
            "Measurement", "IsClosed", "VertexCount", "ConstantWidth", "Degree", "IsPeriodic",
            "PatternName", "NumLoops", "Area", "Length", "Elevation"
        };

        return geometryProps.Contains(propertyName);
    }

    public static int GetGeometryPropertyOrder(string propertyName)
    {
        var orderMap = new Dictionary<string, int>
        {
            // Position properties first
            ["CenterX"] = 1, ["CenterY"] = 2, ["CenterZ"] = 3,
            ["StartX"] = 4, ["StartY"] = 5, ["StartZ"] = 6,
            ["EndX"] = 7, ["EndY"] = 8, ["EndZ"] = 9,

            // Dimensions
            ["Width"] = 10, ["Height"] = 11, ["Depth"] = 12,
            ["Radius"] = 13, ["Diameter"] = 14, ["MajorRadius"] = 15, ["MinorRadius"] = 16,
            ["TextHeight"] = 17, ["ActualHeight"] = 18, ["ActualWidth"] = 19,

            // Measurements
            ["Area"] = 20, ["Length"] = 21, ["LineLength"] = 22, ["ArcLength"] = 23,
            ["Circumference"] = 24, ["Measurement"] = 25,

            // Angles and rotation
            ["Angle"] = 30, ["StartAngle"] = 31, ["EndAngle"] = 32, ["Rotation"] = 33,

            // Scale factors
            ["ScaleX"] = 40, ["ScaleY"] = 41, ["ScaleZ"] = 42, ["WidthFactor"] = 43,
            ["RadiusRatio"] = 44, ["PatternScale"] = 45,

            // Boolean and count properties
            ["IsClosed"] = 50, ["IsPeriodic"] = 51, ["VertexCount"] = 52, ["NumLoops"] = 53,

            // Misc properties
            ["ConstantWidth"] = 60, ["Degree"] = 61, ["PatternName"] = 62, ["Elevation"] = 63
        };

        return orderMap.TryGetValue(propertyName, out int order) ? order : 999;
    }

    /// <summary>
    /// Post-processes entity data to:
    /// 1. Split Tags column into tag_1, tag_2, tag_3, etc. (one tag per column)
    /// 2. Remove attribute columns (Tags, attr_*) that have no data across all entities
    /// 3. Remove DynamicBlockName column if no dynamic blocks are present
    /// </summary>
    public static void ProcessTagsAndAttributes(List<Dictionary<string, object>> entityData)
    {
        if (entityData == null || entityData.Count == 0)
            return;

        // Step 0: Check if any dynamic blocks exist and remove column if not
        bool hasDynamicBlocks = false;
        foreach (var entity in entityData)
        {
            if (entity.TryGetValue("DynamicBlockName", out object dynBlockName) &&
                dynBlockName != null &&
                !string.IsNullOrWhiteSpace(dynBlockName.ToString()) &&
                dynBlockName.ToString() != "N/A")
            {
                hasDynamicBlocks = true;
                break;
            }
        }

        if (!hasDynamicBlocks)
        {
            // Remove DynamicBlockName column from all entities
            foreach (var entity in entityData)
            {
                entity.Remove("DynamicBlockName");
            }
        }

        // Step 1: Determine maximum number of tags across all entities
        int maxTagCount = 0;
        foreach (var entity in entityData)
        {
            if (entity.TryGetValue("Tags", out object tagsObj) && tagsObj != null)
            {
                string tagsStr = tagsObj.ToString();
                if (!string.IsNullOrWhiteSpace(tagsStr))
                {
                    var tags = tagsStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(t => t.Trim())
                                     .Where(t => !string.IsNullOrEmpty(t))
                                     .ToList();
                    maxTagCount = Math.Max(maxTagCount, tags.Count);
                }
            }
        }

        // Step 2: Split Tags into separate columns if any tags exist
        if (maxTagCount > 0)
        {
            foreach (var entity in entityData)
            {
                // Get tags for this entity
                List<string> tags = new List<string>();
                if (entity.TryGetValue("Tags", out object tagsObj) && tagsObj != null)
                {
                    string tagsStr = tagsObj.ToString();
                    if (!string.IsNullOrWhiteSpace(tagsStr))
                    {
                        tags = tagsStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(t => t.Trim())
                                     .Where(t => !string.IsNullOrEmpty(t))
                                     .ToList();
                    }
                }

                // Remove the original Tags column
                entity.Remove("Tags");

                // Add individual tag columns (tag_1, tag_2, tag_3, etc.)
                for (int i = 0; i < maxTagCount; i++)
                {
                    string columnName = $"tag_{i + 1}";
                    entity[columnName] = i < tags.Count ? tags[i] : "";
                }
            }
        }
        else
        {
            // No tags exist - remove Tags column from all entities
            foreach (var entity in entityData)
            {
                entity.Remove("Tags");
            }
        }

        // Step 3: Check if any block references exist in the selection
        bool hasBlockReferences = false;
        foreach (var entity in entityData)
        {
            if (entity.TryGetValue("Category", out object category) && category != null)
            {
                string categoryStr = category.ToString();
                if (categoryStr == "Block Reference" || categoryStr == "Dynamic Block" || categoryStr == "XRef")
                {
                    hasBlockReferences = true;
                    break;
                }
            }
        }

        // Step 4: Find all attribute columns (attr_*) across ALL entities
        var allKeys = entityData
            .SelectMany(e => e.Keys)
            .Where(k => k.StartsWith("attr_"))
            .Distinct()
            .ToList();
        var emptyAttributeColumns = new HashSet<string>();

        // Only remove empty attribute columns if NO block references are present
        // If block references exist, keep all attribute columns even if empty
        if (!hasBlockReferences)
        {
            foreach (var attrKey in allKeys)
            {
                bool hasData = false;
                foreach (var entity in entityData)
                {
                    if (entity.TryGetValue(attrKey, out object value) && value != null)
                    {
                        string strValue = value.ToString();
                        if (!string.IsNullOrWhiteSpace(strValue))
                        {
                            hasData = true;
                            break;
                        }
                    }
                }

                if (!hasData)
                {
                    emptyAttributeColumns.Add(attrKey);
                }
            }

            // Step 5: Remove empty attribute columns from all entities
            if (emptyAttributeColumns.Count > 0)
            {
                foreach (var entity in entityData)
                {
                    foreach (var emptyCol in emptyAttributeColumns)
                    {
                        entity.Remove(emptyCol);
                    }
                }
            }
        }
    }
}

/// <summary>
/// Selection scope enum for filter commands
/// </summary>
public enum SelectionScope
{
    view,           // Current active space/layout only
    document,       // Entire current document (all layouts)
    application     // All opened documents in current AutoCAD process (was "process")
}

/// <summary>
/// Base class for commands that display AutoCAD entities in a custom data grid for filtering and reselection.
/// Works with the stored selection system used by commands like select-by-categories.
/// </summary>
public abstract class FilterElementsBase
{
    public abstract bool SpanAllScreens { get; }
    public abstract bool UseSelectedElements { get; }
    public abstract bool IncludeProperties { get; }
    public abstract SelectionScope Scope { get; }

    // Virtual method to get command name for search history
    public virtual string GetCommandName()
    {
        return $"filter-selection-in-{Scope.ToString().ToLowerInvariant()}";
    }

    public virtual void Execute()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ObjectId[] originalSelection = null;

        var totalTimer = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            var timer = System.Diagnostics.Stopwatch.StartNew();
            var entityData = FilterEntityDataHelper.GetEntityData(ed, Scope, out originalSelection, UseSelectedElements, IncludeProperties);
            timer.Stop();

            if (!entityData.Any())
            {
                ed.WriteMessage("\nNo entities found.\n");
                // Restore original selection if we had one in view scope
                if (originalSelection != null && originalSelection.Length > 0 && Scope == SelectionScope.view)
                {
                    ed.SetImpliedSelection(originalSelection);
                }
                return;
            }

            timer.Restart();
            // Add index to each entity for mapping back after user selection
            for (int i = 0; i < entityData.Count; i++)
            {
                entityData[i]["OriginalIndex"] = i;
            }
            timer.Stop();

            timer.Restart();
            // Post-process: Split Tags into separate columns (tag_1, tag_2, tag_3, etc.)
            // and remove columns that have no data across all entities
            FilterEntityDataHelper.ProcessTagsAndAttributes(entityData);
            timer.Stop();

            timer.Restart();
            // DIAGNOSTIC: Check for _ParentBlocks BEFORE ProcessParentBlockColumns
            var entitiesWithParentBlocksBefore = entityData.Count(e => e.ContainsKey("_ParentBlocks"));
            if (entitiesWithParentBlocksBefore > 0)
            {
                var sampleEntity = entityData.First(e => e.ContainsKey("_ParentBlocks"));
            }

            // Process parent block hierarchy into separate columns
            FilterEntityDataHelper.ProcessParentBlockColumns(entityData);
            timer.Stop();

            timer.Restart();
            // DIAGNOSTIC: Check for _ParentBlocks AFTER ProcessParentBlockColumns
            var entitiesWithParentBlocksAfter = entityData.Count(e => e.ContainsKey("_ParentBlocks"));

            // DEFENSIVE: Ensure _ParentBlocks is removed from all entities (should already be done by ProcessParentBlockColumns)
            var removedCount = 0;
            foreach (var entity in entityData)
            {
                if (entity.Remove("_ParentBlocks"))
                {
                    removedCount++;
                }
            }
            timer.Stop();

            timer.Restart();
            // DIAGNOSTIC: Final check before collecting property names
            var entitiesWithParentBlocksFinal = entityData.Count(e => e.ContainsKey("_ParentBlocks"));

            // Get ALL unique property names from ALL entities (union, not just first entity)
            var propertyNames = entityData
                .SelectMany(e => e.Keys)
                .Distinct()
                .Where(k => !k.EndsWith("ObjectId") && k != "OriginalIndex" && k != "_ParentBlocks")
                .ToList();

            // DIAGNOSTIC: Check if _ParentBlocks is in propertyNames
            var allKeysBeforeFilter = entityData.SelectMany(e => e.Keys).Distinct().ToList();
            var hasParentBlocksKey = allKeysBeforeFilter.Contains("_ParentBlocks");

            // DIAGNOSTIC: Show all keys that contain "parent" or "block" (case insensitive)
            var parentBlockRelatedKeys = allKeysBeforeFilter.Where(k =>
                k.ToLower().Contains("parent") || k.ToLower().Contains("block")).ToList();
            if (parentBlockRelatedKeys.Any())
            {
            }

            // Normalize all entities to have all property names (fill missing keys with empty values)
            foreach (var entity in entityData)
            {
                foreach (var propName in propertyNames)
                {
                    if (!entity.ContainsKey(propName))
                    {
                        entity[propName] = "";
                    }
                }
            }
            timer.Stop();

            timer.Restart();
            // Reorder to put most useful columns first
            var orderedProps = new List<string> { "Name", "DynamicBlockName", "Contents", "Category", "Layer", "Layout", "DocumentName", "Color", "LineType", "Handle" };

            // Extract tag columns (tag_1, tag_2, tag_3, etc.) to place after Category
            var tagColumns = propertyNames.Where(p => p.StartsWith("tag_")).OrderBy(p => p).ToList();

            // Extract parent block columns (ParentBlock or ParentBlock1, ParentBlock2, etc.) to place after IsExternal
            var parentBlockColumns = propertyNames
                .Where(p => p == "ParentBlock" || p.StartsWith("ParentBlock") && char.IsDigit(p.Last()))
                .OrderBy(p => p == "ParentBlock" ? 0 : int.Parse(p.Substring("ParentBlock".Length)))
                .ToList();

            // Extract ParentBlockType column to place after parent block columns
            var parentBlockTypeColumn = propertyNames.Contains("ParentBlockType") ? new[] { "ParentBlockType" } : new string[0];

            var remainingProps = propertyNames
                .Except(orderedProps)
                .Except(tagColumns)
                .Except(parentBlockColumns)
                .Except(parentBlockTypeColumn);

            // Separate geometry properties, attributes and extension data for better organization
            var geometryProps = remainingProps.Where(p => FilterEntityDataHelper.IsGeometryProperty(p)).OrderBy(p => FilterEntityDataHelper.GetGeometryPropertyOrder(p));
            var attributeProps = remainingProps.Where(p => p.StartsWith("attr_")).OrderBy(p => p);
            var extensionProps = remainingProps.Where(p => p.StartsWith("xdata_") || p.StartsWith("ext_dict_")).OrderBy(p => p);

            // Get IsExternal column separately to place parent block columns after it
            var isExternalColumn = propertyNames.Contains("IsExternal") ? new[] { "IsExternal" } : new string[0];

            var otherProps = remainingProps
                .Where(p => !p.StartsWith("attr_") && !p.StartsWith("xdata_") && !p.StartsWith("ext_dict_")
                    && p != "DocumentPath" && p != "DisplayName" && p != "IsExternal"
                    && !FilterEntityDataHelper.IsGeometryProperty(p))
                .OrderBy(p => p);
            var documentPathProp = propertyNames.Contains("DocumentPath") ? new[] { "DocumentPath" } : new string[0];

            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(tagColumns)  // Add tag columns right after the ordered props
                .Concat(geometryProps)
                .Concat(attributeProps)
                .Concat(extensionProps)
                .Concat(isExternalColumn)  // IsExternal column
                .Concat(parentBlockColumns)  // Parent block columns after IsExternal
                .Concat(parentBlockTypeColumn)  // Parent block type after parent block columns
                .Concat(otherProps)
                .Concat(documentPathProp)
                .ToList();
            timer.Stop();

            // Reset the edits flag at the start of each DataGrid session
            CustomGUIs.ResetEditsAppliedFlag();

            totalTimer.Stop();

            timer.Restart();
            var chosenRows = CustomGUIs.DataGrid(entityData, propertyNames, SpanAllScreens, null, null, false, GetCommandName());
            timer.Stop();

            // Check if any edits were applied during DataGrid interaction
            bool editsWereApplied = CustomGUIs.HasPendingEdits() || CustomGUIs.WereEditsApplied();

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo entities selected.\n");
                // Only restore original selection if user canceled and we're in view scope AND no edits were applied
                if (!editsWereApplied && originalSelection != null && originalSelection.Length > 0 && Scope == SelectionScope.view)
                {
                    ed.SetImpliedSelection(originalSelection);
                    ed.WriteMessage($"Restored original selection of {originalSelection.Length} entities.\n");
                }
                else if (editsWereApplied)
                {
                    ed.WriteMessage("Entity modifications were applied. Selection not changed.\n");
                }
                return;
            }

            // Collect ObjectIds for selection
            var selectedIds = new List<ObjectId>();
            var externalEntities = new List<Dictionary<string, object>>();

            foreach (var row in chosenRows)
            {
                // Check if this is an external entity
                if (row.TryGetValue("IsExternal", out var isExternal) && (bool)isExternal)
                {
                    externalEntities.Add(row);
                }
                else if (row.TryGetValue("ObjectId", out var objIdObj) && objIdObj is ObjectId objectId)
                {
                    // Validate ObjectId before adding to avoid eInvalidInput errors
                    if (!objectId.IsNull && objectId.IsValid)
                    {
                        try
                        {
                            // Quick validation by checking if object exists in current database
                            using (var tr = doc.Database.TransactionManager.StartTransaction())
                            {
                                var testObj = tr.GetObject(objectId, OpenMode.ForRead, false);
                                if (testObj != null)
                                {
                                    selectedIds.Add(objectId);
                                }
                                tr.Commit();
                            }
                        }
                        catch
                        {
                            // Skip invalid ObjectIds to prevent eInvalidInput errors
                            ed.WriteMessage($"\nSkipping invalid ObjectId: {objectId}\n");
                        }
                    }
                }
            }

            // Set selection for current document entities (only if no edits were applied)
            if (selectedIds.Count > 0 && !editsWereApplied)
            {
                try
                {
                    // For session scope, we need to actually set the AutoCAD selection (not just save to storage)
                    // to properly narrow down the selection as expected
                    if (Scope == SelectionScope.application)
                    {
                        ed.SetImpliedSelection(selectedIds.ToArray());
                    }
                    else
                    {
                        ed.SetImpliedSelectionEx(selectedIds.ToArray());
                    }
                    ed.WriteMessage($"\n{selectedIds.Count} entities selected in current document.\n");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                {
                    ed.WriteMessage($"\nError setting selection: {acEx.ErrorStatus} - {acEx.Message}\n");
                    if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.InvalidInput)
                    {
                        ed.WriteMessage("\nSome ObjectIds are invalid. This can happen when entities are deleted or database changes occur.\n");
                    }
                }
            }
            else if (selectedIds.Count > 0 && editsWereApplied)
            {
                ed.WriteMessage($"\nEntity modifications were applied. Selection not changed (would have selected {selectedIds.Count} entities).\n");
            }

            // Report external entities
            if (externalEntities.Count > 0)
            {
                ed.WriteMessage($"\nNote: {externalEntities.Count} external entities were found but cannot be selected:\n");
                foreach (var ext in externalEntities.Take(5)) // Show first 5
                {
                    ed.WriteMessage($"  {ext["DocumentName"]} - Handle: {ext["Handle"]}\n");
                }
                if (externalEntities.Count > 5)
                {
                    ed.WriteMessage($"  ... and {externalEntities.Count - 5} more\n");
                }
            }

            // Save the selected entities to selection storage for other commands (except in view mode)
            if ((selectedIds.Count > 0 || externalEntities.Count > 0) && Scope != SelectionScope.view)
            {
                var selectionItems = new List<SelectionItem>();

                // Add current document entities
                foreach (var id in selectedIds)
                {
                    selectionItems.Add(new SelectionItem
                    {
                        DocumentPath = doc.Name,
                        Handle = id.Handle.ToString(),
                        SessionId = null // Will be auto-generated by SelectionStorage
                    });
                }

                // Add external entities (from other documents)
                foreach (var ext in externalEntities)
                {
                    selectionItems.Add(new SelectionItem
                    {
                        DocumentPath = ext["DocumentPath"].ToString(),
                        Handle = ext["Handle"].ToString(),
                        SessionId = null // Will be auto-generated by SelectionStorage
                    });
                }

                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage("Selection saved for use with other commands.\n");
            }
        }
        catch (InvalidOperationException ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
            // Restore original selection if we had one in view scope
            if (originalSelection != null && originalSelection.Length > 0 && Scope == SelectionScope.view)
            {
                ed.SetImpliedSelection(originalSelection);
                ed.WriteMessage($"Restored original selection of {originalSelection.Length} entities.\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nUnexpected error: {ex.Message}\n");
        }
    }
}


// FilterSelectionElements and FilterSelectionImpl classes removed
// Use scope-specific commands instead:
// - filter-selection-in-view
// - filter-selection-in-document
// - filter-selection-in-session
