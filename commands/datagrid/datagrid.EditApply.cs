using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using AutoCADBallet;

public partial class CustomGUIs
{
    /// <summary>Apply pending cell edits to actual AutoCAD entities</summary>
    public static void ApplyCellEditsToEntities()
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        ed.WriteMessage($"\n=== ApplyCellEditsToEntities START ===");
        ed.WriteMessage($"\nPending edits count: {_pendingCellEdits.Count}");

        if (_pendingCellEdits.Count == 0)
        {
            ed.WriteMessage("\nNo pending edits - returning early");
            return;
        }

        try
        {
            var editsByDocument = GroupEditsByDocument();
            ed.WriteMessage($"\nEdits grouped by {editsByDocument.Count} document(s)");

            int totalApplied = 0;
            int totalProcessed = 0;

            string currentDocPath = System.IO.Path.GetFullPath(doc.Name);
            if (editsByDocument.ContainsKey(currentDocPath))
            {
                var currentDocEdits = editsByDocument[currentDocPath];
                ed.WriteMessage($"\nProcessing {currentDocEdits.Count} edits for current document");

                int applied = ApplyEditsToDocument(doc, currentDocEdits);
                totalApplied += applied;
                totalProcessed += currentDocEdits.Count;

                editsByDocument.Remove(currentDocPath);
            }

            foreach (var docEdits in editsByDocument)
            {
                string externalDocPath = docEdits.Key;
                var edits = docEdits.Value;
                ed.WriteMessage($"\nProcessing {edits.Count} edits for external document: {System.IO.Path.GetFileName(externalDocPath)}");
                int applied = ApplyEditsToExternalDocument(externalDocPath, edits);
                totalApplied += applied;
                totalProcessed += edits.Count;
            }

            ed.WriteMessage($"\nFinal result: Applied {totalApplied} entity modifications out of {totalProcessed} processed.");
            if (totalApplied > 0) _editsWereApplied = true;
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in ApplyCellEditsToEntities: {ex.Message}");
            ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
        }
        finally
        {
            ed.WriteMessage($"\nClearing {_pendingCellEdits.Count} pending edits...");
            _pendingCellEdits.Clear();
            ed.WriteMessage($"\n=== ApplyCellEditsToEntities END ===");
        }
    }

    private class PendingEdit
    {
        public int RowIndex { get; set; }
        public string ColumnName { get; set; }
        public string NewValue { get; set; }
        public string OriginalValue { get; set; }  // Store original value before edit
        public Dictionary<string, object> Entry { get; set; }
    }

    private static Dictionary<string, List<PendingEdit>> GroupEditsByDocument()
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var editsByDocument = new Dictionary<string, List<PendingEdit>>(StringComparer.OrdinalIgnoreCase);

        foreach (var kvp in _pendingCellEdits)
        {
            string editKey = kvp.Key;
            string newValue = kvp.Value?.ToString() ?? "";

            // Parse edit key: "InternalID|ColumnName"
            int pipePos = editKey.LastIndexOf('|');
            if (pipePos <= 0) continue;

            string columnName = editKey.Substring(pipePos + 1);
            string internalIdStr = editKey.Substring(0, pipePos);

            if (!long.TryParse(internalIdStr, out long internalId))
            {
                ed.WriteMessage($"\nWARNING: Invalid internal ID in edit key '{editKey}' - edit will be skipped");
                continue;
            }

            // Find the entry in _modifiedEntries by matching internal ID
            Dictionary<string, object> entry = _modifiedEntries.FirstOrDefault(e =>
            {
                long entryId = CustomGUIs.GetInternalId(e);
                return entryId == internalId;
            });

            if (entry == null)
            {
                ed.WriteMessage($"\nWARNING: Could not find entry with internal ID '{internalId}' - edit will be skipped");
                continue;
            }

            // Get document path from entry (if it exists)
            string documentPath = doc.Name; // Default to current document
            if (entry.TryGetValue("DocumentPath", out var docPathObj))
            {
                documentPath = docPathObj.ToString();
            }

            string fullDocPath;
            try
            {
                fullDocPath = System.IO.Path.GetFullPath(documentPath);
            }
            catch
            {
                fullDocPath = documentPath;
            }

            var pendingEdit = new PendingEdit
            {
                RowIndex = 0, // Not used anymore
                ColumnName = columnName,
                NewValue = newValue,
                Entry = entry
            };

            if (!editsByDocument.ContainsKey(fullDocPath))
                editsByDocument[fullDocPath] = new List<PendingEdit>();
            editsByDocument[fullDocPath].Add(pendingEdit);

            ed.WriteMessage($"\nGrouped edit for {System.IO.Path.GetFileName(fullDocPath)}: InternalID={internalId}, Column={columnName}, Value='{newValue}'");
        }
        return editsByDocument;
    }

    private static int ApplyEditsToDocument(Document doc, List<PendingEdit> edits)
    {
        var ed = doc.Editor;
        var db = doc.Database;
        int appliedCount = 0;
        ed.WriteMessage("\nStarting transaction for current document...");
        using (DocumentLock docLock = doc.LockDocument())
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var edit in edits)
                {
                    try
                    {
                        ObjectId objectId = ObjectId.Null;
                        if (edit.Entry.TryGetValue("ObjectId", out var objIdValue) && objIdValue is ObjectId validObjectId)
                        {
                            objectId = validObjectId;
                        }
                        else if (edit.Entry.TryGetValue("Handle", out var handleValue))
                        {
                            try
                            {
                                var handle = Convert.ToInt64(handleValue.ToString(), 16);
                                objectId = doc.Database.GetObjectId(false, new Autodesk.AutoCAD.DatabaseServices.Handle(handle), 0);
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to reconstruct ObjectId from Handle {handleValue}: {ex.Message}");
                            }
                        }

                        if (objectId != ObjectId.Null)
                        {
                            ed.WriteMessage($"\nApplying edit: {edit.ColumnName} = '{edit.NewValue}' to ObjectId {objectId}");
                            var dbObject = tr.GetObject(objectId, OpenMode.ForWrite);
                            if (dbObject != null)
                            {
                                ApplyEditToDBObject(dbObject, edit.ColumnName, edit.NewValue, edit.Entry, tr);
                                appliedCount++;
                                ed.WriteMessage($"\nEdit applied successfully");
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"\nSkipping edit - no valid ObjectId or Handle found");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError applying edit to {edit.ColumnName}: {ex.Message}");
                        continue;
                    }
                }

                ed.WriteMessage($"\nCommitting transaction with {appliedCount} modifications...");
                tr.Commit();
                ed.WriteMessage($"\nTransaction committed successfully!");
            }
        }
        return appliedCount;
    }

    private static int ApplyEditsToExternalDocument(string documentPath, List<PendingEdit> edits)
    {
        var docs = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
        var currentDoc = docs.MdiActiveDocument;
        var ed = currentDoc.Editor;
        int appliedCount = 0;

        try
        {
            Document targetDoc = null;
            foreach (Document d in docs)
            {
                try
                {
                    if (System.IO.Path.GetFullPath(d.Name).Equals(System.IO.Path.GetFullPath(documentPath), StringComparison.OrdinalIgnoreCase))
                    {
                        targetDoc = d;
                        break;
                    }
                }
                catch { }
            }

            if (targetDoc == null)
            {
                ed.WriteMessage($"\nDocument not found or not open: {documentPath}");
                return 0;
            }

            ed.WriteMessage($"\nProcessing edits in external document: {System.IO.Path.GetFileName(documentPath)}");
            using (var docLock = targetDoc.LockDocument())
            {
                var targetDb = targetDoc.Database;
                ed.WriteMessage("\nStarting transaction for external document...");
                using (var tr = targetDb.TransactionManager.StartTransaction())
                {
                    foreach (var edit in edits)
                    {
                        try
                        {
                            ObjectId objectId = ObjectId.Null;
                            if (edit.Entry.TryGetValue("Handle", out var handleValue))
                            {
                                try
                                {
                                    var handle = Convert.ToInt64(handleValue.ToString(), 16);
                                    objectId = targetDb.GetObjectId(false, new Autodesk.AutoCAD.DatabaseServices.Handle(handle), 0);
                                }
                                catch (System.Exception ex)
                                {
                                    ed.WriteMessage($"\nFailed to reconstruct ObjectId from Handle {handleValue} in external document: {ex.Message}");
                                    continue;
                                }
                            }

                            if (objectId != ObjectId.Null)
                            {
                                ed.WriteMessage($"\nApplying edit to external document: {edit.ColumnName} = '{edit.NewValue}' to ObjectId {objectId}");
                                var dbObject = tr.GetObject(objectId, OpenMode.ForWrite);
                                if (dbObject != null)
                                {
                                    ApplyEditToDBObjectInExternalDocument(dbObject, edit.ColumnName, edit.NewValue, edit.Entry, tr);
                                    appliedCount++;
                                    ed.WriteMessage($"\nExternal edit applied successfully");
                                }
                            }
                            else
                            {
                                ed.WriteMessage($"\nSkipping external edit - no valid Handle found");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError applying external edit to {edit.ColumnName}: {ex.Message}");
                            continue;
                        }
                    }

                    ed.WriteMessage($"\nCommitting external transaction with {appliedCount} modifications...");
                    tr.Commit();
                    ed.WriteMessage($"\nExternal transaction committed successfully!");
                }
            }

            ed.WriteMessage($"\nCompleted processing {appliedCount} edits in external document");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError processing external document {System.IO.Path.GetFileName(documentPath)}: {ex.Message}");
            ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            return 0;
        }
        return appliedCount;
    }

    private static void ApplyEditToDBObjectInExternalDocument(DBObject dbObject, string columnName, string newValue, Dictionary<string, object> entry, Transaction tr)
    {
        var currentDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = currentDoc.Editor;
        try
        {
            string lowerColumnName = columnName.ToLowerInvariant();

            // Handle tag_* columns (tag_1, tag_2, tag_3, etc.)
            if (lowerColumnName.StartsWith("tag_") && dbObject is Entity)
            {
                // Collect all tag_* values from the entry dictionary
                var allTags = new List<string>();
                foreach (var kvp in entry)
                {
                    if (kvp.Key.ToLowerInvariant().StartsWith("tag_"))
                    {
                        string tagValue = kvp.Value?.ToString() ?? "";
                        // Update the current column with the new value
                        if (string.Equals(kvp.Key, columnName, StringComparison.OrdinalIgnoreCase))
                        {
                            tagValue = newValue;
                        }
                        // Add non-empty tags to the list
                        if (!string.IsNullOrWhiteSpace(tagValue))
                        {
                            allTags.Add(tagValue.Trim());
                        }
                    }
                }

                // Apply all tags to the entity (or clear if empty)
                if (allTags.Count > 0)
                {
                    dbObject.ObjectId.SetTags(dbObject.Database, allTags.ToArray());
                }
                else
                {
                    dbObject.ObjectId.ClearTags(dbObject.Database);
                }
                return;
            }

            switch (lowerColumnName)
            {
                case "tags":
                    if (dbObject is Entity)
                    {
                        // Empty string or single space means clear all tags
                        if (string.IsNullOrEmpty(newValue) || newValue.Trim() == "")
                        {
                            dbObject.ObjectId.ClearTags(dbObject.Database);
                        }
                        else
                        {
                            // Parse comma-separated tags and set them
                            var tags = newValue.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Select(t => t.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .ToArray();
                            if (tags.Length > 0)
                            {
                                dbObject.ObjectId.SetTags(dbObject.Database, tags);
                            }
                            else
                            {
                                // If all tags are empty after parsing, clear tags
                                dbObject.ObjectId.ClearTags(dbObject.Database);
                            }
                        }
                    }
                    break;
                case "contents":
                    if (dbObject is MText mtextExt) mtextExt.Contents = newValue;
                    else if (dbObject is DBText textExt) textExt.TextString = newValue;
                    else if (dbObject is Dimension dimExt) dimExt.DimensionText = newValue;
                    break;
                case "layout":
                    // For Layout entities: rename the layout tab
                    // For other entities: find and rename the layout that the entity belongs to
                    if (dbObject is Layout layout)
                    {
                        try
                        {
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                            {
                                ed.WriteMessage("\n  >> Cannot rename Model layout");
                                return;
                            }
                            var layoutDict = (DBDictionary)tr.GetObject(dbObject.Database.LayoutDictionaryId, OpenMode.ForRead);
                            if (layoutDict.Contains(newValue))
                            {
                                int counter = 1; string uniqueName = newValue;
                                while (layoutDict.Contains(uniqueName)) uniqueName = $"{newValue}_{counter++}";
                                newValue = uniqueName;
                            }
                            layout.LayoutName = newValue;
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotApplicable) return;
                            throw;
                        }
                    }
                    else if (dbObject is Entity entityForLayoutRename)
                    {
                        // Rename the layout that this entity belongs to
                        RenameEntityLayout(entityForLayoutRename, newValue, entry, tr);
                    }
                    break;
                case "name":
                case "dynamicblockname":
                    if (dbObject is Layout layout2 && lowerColumnName == "name")
                    {
                        try
                        {
                            if (string.Equals(layout2.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) return;
                            var layoutDict = (DBDictionary)tr.GetObject(dbObject.Database.LayoutDictionaryId, OpenMode.ForRead);
                            if (layoutDict.Contains(newValue))
                            {
                                int counter = 1; string uniqueName = newValue;
                                while (layoutDict.Contains(uniqueName)) uniqueName = $"{newValue}_{counter++}";
                                newValue = uniqueName;
                            }
                            layout2.LayoutName = newValue;
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotApplicable) return;
                            throw;
                        }
                    }
                    else if (dbObject is MText mtextExt2 && lowerColumnName == "name") mtextExt2.Contents = newValue;
                    else if (dbObject is DBText textExt2 && lowerColumnName == "name") textExt2.TextString = newValue;
                    else if (dbObject is BlockReference blockRef)
                    {
                        var blockTable = (BlockTable)tr.GetObject(dbObject.Database.BlockTableId, OpenMode.ForRead);
                        var currentBtr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                        // For dynamic blocks, check the parent block name; for regular blocks, use the current block name
                        string currentBlockName = currentBtr?.Name ?? "";
                        if (lowerColumnName == "dynamicblockname" && blockRef.DynamicBlockTableRecord != ObjectId.Null)
                        {
                            var dynamicBtr = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            currentBlockName = dynamicBtr?.Name ?? "";
                        }

                        // Check if already using the target block definition (avoid unnecessary swap)
                        if (!string.IsNullOrEmpty(currentBlockName) && string.Equals(currentBlockName, newValue, StringComparison.OrdinalIgnoreCase))
                        {
                            // Already using the target block, no action needed
                            return;
                        }

                        // Check if the new name matches an existing block definition
                        if (blockTable.Has(newValue))
                        {
                            // Swap to existing block definition and copy matching attributes
                            SwapBlockReference(blockRef, newValue, tr, ed);
                        }
                        else if (lowerColumnName == "name")
                        {
                            // For "name" column: rename the current block definition
                            var blockTableRecord = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            if (blockTableRecord != null)
                            {
                                blockTableRecord.Name = newValue;
                            }
                        }
                        else if (lowerColumnName == "dynamicblockname")
                        {
                            // For "dynamicblockname" column: this case should not be reached if validation works correctly
                            // The validation happens earlier in EditMode before pending edits are created
                            ed.WriteMessage($"\n  >> Unexpected: DynamicBlockName edit reached apply phase for non-existent block '{newValue}'");
                        }
                    }
                    else if (dbObject is LayerTableRecord layerRec && lowerColumnName == "name")
                    {
                        // Rename layer with validation
                        if (string.Equals(layerRec.Name, "0", StringComparison.OrdinalIgnoreCase))
                        {
                            ed.WriteMessage("\n  >> Cannot rename layer 0");
                            return;
                        }
                        var layerTable = (LayerTable)tr.GetObject(dbObject.Database.LayerTableId, OpenMode.ForRead);
                        if (layerTable.Has(newValue))
                        {
                            ed.WriteMessage($"\n  >> Layer '{newValue}' already exists");
                            return;
                        }
                        layerRec.Name = newValue;
                    }
                    else if (dbObject is TextStyleTableRecord tStyle && lowerColumnName == "name") tStyle.Name = newValue;
                    else if (dbObject is LinetypeTableRecord ltype && lowerColumnName == "name") ltype.Name = newValue;
                    else if (dbObject is DimStyleTableRecord dStyle && lowerColumnName == "name") dStyle.Name = newValue;
                    else if (dbObject is UcsTableRecord ucs && lowerColumnName == "name") ucs.Name = newValue;
                    break;
                case "papersize":
                case "plotstyletable":
                case "plotrotation":
                case "plotscale":
                case "plottype":
                case "plotcentered":
                    if (dbObject is Layout layoutForPlotSettings)
                        ApplyPlotSettingEdit(layoutForPlotSettings, columnName, newValue, tr);
                    break;
                case "centerx": case "centery": case "centerz": case "scalex": case "scaley": case "scalez": case "rotation": case "width": case "height": case "radius": case "textheight": case "widthfactor":
                    if (dbObject is Entity externalGeometryEntity)
                        ApplyGeometryPropertyEdit(externalGeometryEntity, columnName, newValue, tr);
                    break;
                case "xrefpath":
                    if (dbObject is BlockReference xrefBlockRefExt)
                    {
                        ApplyXrefPathEdit(xrefBlockRefExt, newValue, tr);
                    }
                    break;
                case "layer":
                    if (dbObject is LayerTableRecord layerForRenameExt)
                    {
                        // Renaming the layer itself
                        if (string.Equals(layerForRenameExt.Name, "0", StringComparison.OrdinalIgnoreCase))
                        {
                            ed.WriteMessage("\n  >> Cannot rename layer 0");
                            return;
                        }
                        var layerTableExt = (LayerTable)tr.GetObject(dbObject.Database.LayerTableId, OpenMode.ForRead);
                        if (layerTableExt.Has(newValue))
                        {
                            ed.WriteMessage($"\n  >> Layer '{newValue}' already exists");
                            return;
                        }
                        layerForRenameExt.Name = newValue;
                    }
                    else if (dbObject is Entity entityExt)
                    {
                        entityExt.Layer = newValue;
                    }
                    else
                    {
                        ed.WriteMessage("\n  >> Not an Entity or Layer");
                    }
                    break;
                case "color":
                    if (dbObject is LayerTableRecord layerForColorExt)
                    {
                        // Parse color for layer
                        if (TryParseColor(newValue, out Autodesk.AutoCAD.Colors.Color colorExt))
                        {
                            layerForColorExt.Color = colorExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid color value: {newValue}");
                        }
                    }
                    else if (dbObject is Entity entity2Ext && int.TryParse(newValue, out int colorIndexExt))
                    {
                        entity2Ext.ColorIndex = colorIndexExt;
                    }
                    break;
                case "linetype":
                    if (dbObject is LayerTableRecord layerForLinetypeExt)
                    {
                        // Set linetype for layer
                        var linetypeTableExt = (LinetypeTable)tr.GetObject(dbObject.Database.LinetypeTableId, OpenMode.ForRead);
                        if (linetypeTableExt.Has(newValue))
                        {
                            var linetypeIdExt = linetypeTableExt[newValue];
                            layerForLinetypeExt.LinetypeObjectId = linetypeIdExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Linetype '{newValue}' not found");
                        }
                    }
                    else if (dbObject is Entity entity3Ext)
                    {
                        entity3Ext.Linetype = newValue;
                    }
                    break;
                case "isfrozen":
                    if (dbObject is LayerTableRecord layerForFrozenExt)
                    {
                        if (bool.TryParse(newValue, out bool isFrozenExt))
                        {
                            layerForFrozenExt.IsFrozen = isFrozenExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "islocked":
                    if (dbObject is LayerTableRecord layerForLockedExt)
                    {
                        if (bool.TryParse(newValue, out bool isLockedExt))
                        {
                            layerForLockedExt.IsLocked = isLockedExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "isoff":
                    if (dbObject is LayerTableRecord layerForOffExt)
                    {
                        if (bool.TryParse(newValue, out bool isOffExt))
                        {
                            layerForOffExt.IsOff = isOffExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "isplottable":
                    if (dbObject is LayerTableRecord layerForPlottableExt)
                    {
                        if (bool.TryParse(newValue, out bool isPlottableExt))
                        {
                            layerForPlottableExt.IsPlottable = isPlottableExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "lineweight":
                    if (dbObject is LayerTableRecord layerForLineweightExt)
                    {
                        if (TryParseLineWeight(newValue, out LineWeight lineWeightExt))
                        {
                            layerForLineweightExt.LineWeight = lineWeightExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid lineweight value: {newValue}");
                        }
                    }
                    break;
                case "transparency":
                    if (dbObject is LayerTableRecord layerForTransparencyExt)
                    {
                        if (TryParseTransparency(newValue, out Autodesk.AutoCAD.Colors.Transparency transparencyExt))
                        {
                            layerForTransparencyExt.Transparency = transparencyExt;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid transparency value: {newValue}");
                        }
                    }
                    break;
                case "description":
                    if (dbObject is LayerTableRecord layerForDescriptionExt)
                    {
                        layerForDescriptionExt.Description = newValue;
                    }
                    break;
                default:
                    if (columnName.StartsWith("attr_"))
                    {
                        if (dbObject is Entity entity4) ApplyBlockAttributeEdit(entity4, columnName, newValue, tr);
                    }
                    else if (columnName.StartsWith("xdata_"))
                    {
                        if (dbObject is Entity entity5) ApplyXDataEdit(entity5, columnName, newValue);
                    }
                    else if (columnName.StartsWith("ext_dict_"))
                    {
                        if (dbObject is Entity entity6) ApplyExtensionDictEdit(entity6, columnName, newValue, tr);
                    }
                    break;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR: {ex.Message}");
            throw;
        }
    }

    private static void ApplyEditToDBObject(DBObject dbObject, string columnName, string newValue, Dictionary<string, object> entry, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        try
        {
            string lowerColumnName = columnName.ToLowerInvariant();

            // Handle tag_* columns (tag_1, tag_2, tag_3, etc.)
            if (lowerColumnName.StartsWith("tag_") && dbObject is Entity)
            {
                // Collect all tag_* values from the entry dictionary
                var allTags = new List<string>();
                foreach (var kvp in entry)
                {
                    if (kvp.Key.ToLowerInvariant().StartsWith("tag_"))
                    {
                        string tagValue = kvp.Value?.ToString() ?? "";
                        // Update the current column with the new value
                        if (string.Equals(kvp.Key, columnName, StringComparison.OrdinalIgnoreCase))
                        {
                            tagValue = newValue;
                        }
                        // Add non-empty tags to the list
                        if (!string.IsNullOrWhiteSpace(tagValue))
                        {
                            allTags.Add(tagValue.Trim());
                        }
                    }
                }

                // Apply all tags to the entity (or clear if empty)
                if (allTags.Count > 0)
                {
                    dbObject.ObjectId.SetTags(dbObject.Database, allTags.ToArray());
                }
                else
                {
                    dbObject.ObjectId.ClearTags(dbObject.Database);
                }
                return;
            }

            switch (lowerColumnName)
            {
                case "tags":
                    if (dbObject is Entity)
                    {
                        // Empty string or single space means clear all tags
                        if (string.IsNullOrEmpty(newValue) || newValue.Trim() == "")
                        {
                            dbObject.ObjectId.ClearTags(dbObject.Database);
                        }
                        else
                        {
                            // Parse comma-separated tags and set them
                            var tags = newValue.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Select(t => t.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .ToArray();
                            if (tags.Length > 0)
                            {
                                dbObject.ObjectId.SetTags(dbObject.Database, tags);
                            }
                            else
                            {
                                // If all tags are empty after parsing, clear tags
                                dbObject.ObjectId.ClearTags(dbObject.Database);
                            }
                        }
                    }
                    break;
                case "layer":
                    if (dbObject is LayerTableRecord layerForRename)
                    {
                        // Renaming the layer itself
                        if (string.Equals(layerForRename.Name, "0", StringComparison.OrdinalIgnoreCase))
                        {
                            ed.WriteMessage("\n  >> Cannot rename layer 0");
                            return;
                        }
                        var layerTable = (LayerTable)tr.GetObject(dbObject.Database.LayerTableId, OpenMode.ForRead);
                        if (layerTable.Has(newValue))
                        {
                            ed.WriteMessage($"\n  >> Layer '{newValue}' already exists");
                            return;
                        }
                        layerForRename.Name = newValue;
                    }
                    else if (dbObject is Entity entity)
                    {
                        entity.Layer = newValue;
                    }
                    else
                    {
                        ed.WriteMessage("\n  >> Not an Entity or Layer");
                    }
                    break;
                case "color":
                    if (dbObject is LayerTableRecord layerForColor)
                    {
                        // Parse color for layer
                        if (TryParseColor(newValue, out Autodesk.AutoCAD.Colors.Color color))
                        {
                            layerForColor.Color = color;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid color value: {newValue}");
                        }
                    }
                    else if (dbObject is Entity entity2 && int.TryParse(newValue, out int colorIndex))
                    {
                        entity2.ColorIndex = colorIndex;
                    }
                    break;
                case "linetype":
                    if (dbObject is LayerTableRecord layerForLinetype)
                    {
                        // Set linetype for layer
                        var linetypeTable = (LinetypeTable)tr.GetObject(dbObject.Database.LinetypeTableId, OpenMode.ForRead);
                        if (linetypeTable.Has(newValue))
                        {
                            var linetypeId = linetypeTable[newValue];
                            layerForLinetype.LinetypeObjectId = linetypeId;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Linetype '{newValue}' not found");
                        }
                    }
                    else if (dbObject is Entity entity3)
                    {
                        entity3.Linetype = newValue;
                    }
                    break;
                case "isfrozen":
                    if (dbObject is LayerTableRecord layerForFrozen)
                    {
                        if (bool.TryParse(newValue, out bool isFrozen))
                        {
                            layerForFrozen.IsFrozen = isFrozen;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "islocked":
                    if (dbObject is LayerTableRecord layerForLocked)
                    {
                        if (bool.TryParse(newValue, out bool isLocked))
                        {
                            layerForLocked.IsLocked = isLocked;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "isoff":
                    if (dbObject is LayerTableRecord layerForOff)
                    {
                        if (bool.TryParse(newValue, out bool isOff))
                        {
                            layerForOff.IsOff = isOff;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "isplottable":
                    if (dbObject is LayerTableRecord layerForPlottable)
                    {
                        if (bool.TryParse(newValue, out bool isPlottable))
                        {
                            layerForPlottable.IsPlottable = isPlottable;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid boolean value: {newValue}");
                        }
                    }
                    break;
                case "lineweight":
                    if (dbObject is LayerTableRecord layerForLineweight)
                    {
                        if (TryParseLineWeight(newValue, out LineWeight lineWeight))
                        {
                            layerForLineweight.LineWeight = lineWeight;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid lineweight value: {newValue}");
                        }
                    }
                    break;
                case "transparency":
                    if (dbObject is LayerTableRecord layerForTransparency)
                    {
                        if (TryParseTransparency(newValue, out Autodesk.AutoCAD.Colors.Transparency transparency))
                        {
                            layerForTransparency.Transparency = transparency;
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Invalid transparency value: {newValue}");
                        }
                    }
                    break;
                case "description":
                    if (dbObject is LayerTableRecord layerForDescription)
                    {
                        layerForDescription.Description = newValue;
                    }
                    break;
                case "contents":
                    if (dbObject is MText mtextContents) mtextContents.Contents = newValue;
                    else if (dbObject is DBText textContents) textContents.TextString = newValue;
                    else if (dbObject is Dimension dimContents) dimContents.DimensionText = newValue;
                    break;
                case "layout":
                    // For Layout entities: rename the layout tab
                    // For other entities: find and rename the layout that the entity belongs to
                    if (dbObject is Layout layout)
                    {
                        try
                        {
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                            {
                                ed.WriteMessage("\n  >> Cannot rename Model layout");
                                return;
                            }
                            var layoutDict = (DBDictionary)tr.GetObject(dbObject.Database.LayoutDictionaryId, OpenMode.ForRead);
                            if (layoutDict.Contains(newValue))
                            {
                                int counter = 1; string uniqueName = newValue;
                                while (layoutDict.Contains(uniqueName)) uniqueName = $"{newValue}_{counter++}";
                                newValue = uniqueName;
                            }
                            layout.LayoutName = newValue;
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotApplicable) return;
                            throw;
                        }
                    }
                    else if (dbObject is Entity entityForLayoutRename)
                    {
                        // Rename the layout that this entity belongs to
                        RenameEntityLayout(entityForLayoutRename, newValue, entry, tr);
                    }
                    break;
                case "name":
                case "dynamicblockname":
                    if (dbObject is Layout layout2 && lowerColumnName == "name")
                    {
                        try
                        {
                            if (string.Equals(layout2.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) return;
                            var layoutDict = (DBDictionary)tr.GetObject(dbObject.Database.LayoutDictionaryId, OpenMode.ForRead);
                            if (layoutDict.Contains(newValue))
                            {
                                int counter = 1; string uniqueName = newValue;
                                while (layoutDict.Contains(uniqueName)) uniqueName = $"{newValue}_{counter++}";
                                newValue = uniqueName;
                            }
                            layout2.LayoutName = newValue;
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotApplicable) return;
                            throw;
                        }
                    }
                    else if (dbObject is MText mtext2 && lowerColumnName == "name") mtext2.Contents = newValue;
                    else if (dbObject is DBText text2 && lowerColumnName == "name") text2.TextString = newValue;
                    else if (dbObject is BlockReference br)
                    {
                        var blockTable = (BlockTable)tr.GetObject(dbObject.Database.BlockTableId, OpenMode.ForRead);
                        var currentBtr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                        // For dynamic blocks, check the parent block name; for regular blocks, use the current block name
                        string currentBlockName = currentBtr?.Name ?? "";
                        if (lowerColumnName == "dynamicblockname" && br.DynamicBlockTableRecord != ObjectId.Null)
                        {
                            var dynamicBtr = tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            currentBlockName = dynamicBtr?.Name ?? "";
                        }

                        // Check if already using the target block definition (avoid unnecessary swap)
                        if (!string.IsNullOrEmpty(currentBlockName) && string.Equals(currentBlockName, newValue, StringComparison.OrdinalIgnoreCase))
                        {
                            // Already using the target block, no action needed
                            return;
                        }

                        // Check if the new name matches an existing block definition
                        if (blockTable.Has(newValue))
                        {
                            // Swap to existing block definition and copy matching attributes
                            SwapBlockReference(br, newValue, tr, ed);
                        }
                        else if (lowerColumnName == "name")
                        {
                            // For "name" column: rename the current block definition
                            var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            if (btr != null)
                            {
                                btr.Name = newValue;
                            }
                        }
                        else if (lowerColumnName == "dynamicblockname")
                        {
                            // For "dynamicblockname" column: this case should not be reached if validation works correctly
                            // The validation happens earlier in EditMode before pending edits are created
                            ed.WriteMessage($"\n  >> Unexpected: DynamicBlockName edit reached apply phase for non-existent block '{newValue}'");
                        }
                    }
                    else if (dbObject is LayerTableRecord layerRec && lowerColumnName == "name")
                    {
                        // Rename layer with validation
                        if (string.Equals(layerRec.Name, "0", StringComparison.OrdinalIgnoreCase))
                        {
                            ed.WriteMessage("\n  >> Cannot rename layer 0");
                            return;
                        }
                        var layerTable = (LayerTable)tr.GetObject(dbObject.Database.LayerTableId, OpenMode.ForRead);
                        if (layerTable.Has(newValue))
                        {
                            ed.WriteMessage($"\n  >> Layer '{newValue}' already exists");
                            return;
                        }
                        layerRec.Name = newValue;
                    }
                    else if (dbObject is TextStyleTableRecord tStyle && lowerColumnName == "name") tStyle.Name = newValue;
                    else if (dbObject is LinetypeTableRecord ltype && lowerColumnName == "name") ltype.Name = newValue;
                    else if (dbObject is DimStyleTableRecord dStyle && lowerColumnName == "name") dStyle.Name = newValue;
                    else if (dbObject is UcsTableRecord ucs && lowerColumnName == "name") ucs.Name = newValue;
                    break;
                case "papersize":
                case "plotstyletable":
                case "plotrotation":
                case "plotscale":
                case "plottype":
                case "plotcentered":
                    if (dbObject is Layout layoutForPlotSettings) ApplyPlotSettingEdit(layoutForPlotSettings, columnName, newValue, tr);
                    break;
                case "centerx": case "centery": case "centerz": case "scalex": case "scaley": case "scalez": case "rotation": case "width": case "height": case "radius": case "textheight": case "widthfactor":
                    if (dbObject is Entity geometryEntity) ApplyGeometryPropertyEdit(geometryEntity, columnName, newValue, tr);
                    break;
                case "xrefpath":
                    if (dbObject is BlockReference xrefBlockRef)
                    {
                        ApplyXrefPathEdit(xrefBlockRef, newValue, tr);
                    }
                    break;
                default:
                    if (columnName.StartsWith("attr_")) { if (dbObject is Entity e4) ApplyBlockAttributeEdit(e4, columnName, newValue, tr); }
                    else if (columnName.StartsWith("xdata_")) { if (dbObject is Entity e5) ApplyXDataEdit(e5, columnName, newValue); }
                    else if (columnName.StartsWith("ext_dict_")) { if (dbObject is Entity e6) ApplyExtensionDictEdit(e6, columnName, newValue, tr); }
                    break;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR: {ex.Message}");
            throw;
        }
    }

    private static void ApplyBlockAttributeEdit(Entity entity, string columnName, string newValue, Transaction tr)
    {
        if (entity is BlockReference blockRef)
        {
            string attributeTag = columnName.Substring(5);
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            bool found = false;
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef != null)
                {
                    if (attRef.Tag.ToLowerInvariant() == attributeTag.ToLowerInvariant())
                    {
                        attRef.TextString = newValue;
                        found = true;
                        break;
                    }
                }
            }
            if (!found) ed.WriteMessage($"\nAttribute '{attributeTag}' not found in block!");
        }
    }

    private static void ApplyXDataEdit(Entity entity, string columnName, string newValue)
    {
        string appName = columnName.Substring(6);
        var rb = new ResultBuffer(
            new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, newValue)
        );
        entity.XData = rb;
        rb.Dispose();
    }

    private static void ApplyExtensionDictEdit(Entity entity, string columnName, string newValue, Transaction tr)
    {
        if (entity.ExtensionDictionary == ObjectId.Null) entity.CreateExtensionDictionary();
        string key = columnName.Substring(9);
        var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
        var xrec = new Xrecord { Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, newValue)) };
        if (extDict.Contains(key)) extDict.SetAt(key, xrec); else extDict.SetAt(key, xrec);
        tr.AddNewlyCreatedDBObject(xrec, true);
    }

    private static void ApplyPlotSettingEdit(Layout layout, string columnName, string newValue, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        try
        {
            var plotSettings = tr.GetObject(layout.ObjectId, OpenMode.ForWrite) as PlotSettings;
            if (plotSettings == null) { ed.WriteMessage($"\n  >> Failed to get PlotSettings from layout"); return; }
            switch (columnName.ToLowerInvariant())
            {
                case "papersize":
                    try
                    {
                        var psv = PlotSettingsValidator.Current;
                        string mediaName = newValue;

                        // If empty, set to default
                        if (string.IsNullOrEmpty(mediaName))
                        {
                            ed.WriteMessage($"\n  >> Paper size cannot be empty.");
                            break;
                        }

                        // Get current plotter configuration to validate media
                        string deviceName = plotSettings.PlotConfigurationName ?? "";
                        if (string.IsNullOrEmpty(deviceName))
                        {
                            ed.WriteMessage($"\n  >> No plotter device configured for paper size change.");
                            break;
                        }

                        // Try to set the canonical media name
                        psv.SetCanonicalMediaName(plotSettings, mediaName);
                        ed.WriteMessage($"\n  >> Paper size set to: {mediaName}");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n  >> Failed to set paper size to '{newValue}': {ex.Message}");
                        ed.WriteMessage($"\n  >> Ensure the paper size name is valid for the current plotter device.");
                    }
                    break;
                case "plotstyletable":
                    try
                    {
                        string styleTable = newValue;
                        if (string.IsNullOrEmpty(styleTable) || string.Equals(styleTable, "None", StringComparison.OrdinalIgnoreCase)) styleTable = "None";
                        else if (!styleTable.ToLowerInvariant().EndsWith(".ctb") && !styleTable.ToLowerInvariant().EndsWith(".stb")) styleTable = newValue + ".ctb";
                        var psv = PlotSettingsValidator.Current;
                        psv.SetCurrentStyleSheet(plotSettings, styleTable);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n  >> Failed to set style table: {ex.Message}");
                    }
                    break;
                case "plotrotation":
                case "plotscale":
                case "plottype":
                case "plotcentered":
                    // Placeholder: implement as needed with PlotSettingsValidator if required.
                    break;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in ApplyPlotSettingEdit: {ex.Message}");
        }
    }

    private static void ApplyGeometryPropertyEdit(Entity entity, string columnName, string newValue, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ed.WriteMessage($"\n  >> ApplyGeometryPropertyEdit: Entity={entity.GetType().Name}, Column='{columnName}', NewValue='{newValue}'");

        if (!double.TryParse(newValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double value))
        {
            ed.WriteMessage($"\n  >> ERROR: Failed to parse '{newValue}' as double. Length={newValue?.Length}, IsNull={newValue == null}");
            if (newValue != null)
            {
                ed.WriteMessage($"\n  >> String bytes: {string.Join(",", System.Text.Encoding.UTF8.GetBytes(newValue))}");
            }
            return;
        }

        ed.WriteMessage($"\n  >> Successfully parsed '{newValue}' as {value}");

        switch (entity)
        {
            case Circle circle: ApplyCircleGeometryEdit(circle, columnName, value); break;
            case Arc arc: ApplyArcGeometryEdit(arc, columnName, value); break;
            case Line line: ApplyLineGeometryEdit(line, columnName, value); break;
            case Polyline pl: ApplyPolylineGeometryEdit(pl, columnName, value); break;
            case Ellipse ellipse: ApplyEllipseGeometryEdit(ellipse, columnName, value); break;
            case BlockReference br: ApplyBlockReferenceGeometryEdit(br, columnName, value); break;
            case DBText t: ApplyDBTextGeometryEdit(t, columnName, value); break;
            case MText mt: ApplyMTextGeometryEdit(mt, columnName, value); break;
            case Viewport vp: ApplyViewportGeometryEdit(vp, columnName, value); break;
        }
    }

    private static void ApplyCircleGeometryEdit(Circle circle, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx": circle.Center = new Point3d(value, circle.Center.Y, circle.Center.Z); break;
            case "centery": circle.Center = new Point3d(circle.Center.X, value, circle.Center.Z); break;
            case "centerz": circle.Center = new Point3d(circle.Center.X, circle.Center.Y, value); break;
            case "radius": if (value > 0) circle.Radius = value; break;
        }
    }

    private static void ApplyArcGeometryEdit(Arc arc, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx": arc.Center = new Point3d(value, arc.Center.Y, arc.Center.Z); break;
            case "centery": arc.Center = new Point3d(arc.Center.X, value, arc.Center.Z); break;
            case "centerz": arc.Center = new Point3d(arc.Center.X, arc.Center.Y, value); break;
            case "radius": if (value > 0) arc.Radius = value; break;
        }
    }

    private static void ApplyLineGeometryEdit(Line line, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx":
                var currentCenter = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var offset = new Vector3d(value - currentCenter.X, 0, 0);
                line.StartPoint += offset; line.EndPoint += offset; break;
            case "centery":
                var currentCenterY = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var offsetY = new Vector3d(0, value - currentCenterY.Y, 0);
                line.StartPoint += offsetY; line.EndPoint += offsetY; break;
            case "centerz":
                var currentCenterZ = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var offsetZ = new Vector3d(0, 0, value - currentCenterZ.Z);
                line.StartPoint += offsetZ; line.EndPoint += offsetZ; break;
            case "rotation":
                var angleRadians = value * Math.PI / 180.0;
                var center = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var transform = Matrix3d.Rotation(angleRadians, Vector3d.ZAxis, center);
                line.TransformBy(transform); break;
        }
    }

    private static void ApplyPolylineGeometryEdit(Polyline polyline, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx":
            case "centery":
            case "centerz":
                var bounds = polyline.GeometricExtents;
                var currentCenter = bounds.MinPoint + (bounds.MaxPoint - bounds.MinPoint) * 0.5;
                Vector3d offset = new Vector3d(0, 0, 0);
                if (columnName == "centerx") offset = new Vector3d(value - currentCenter.X, 0, 0);
                else if (columnName == "centery") offset = new Vector3d(0, value - currentCenter.Y, 0);
                else if (columnName == "centerz") offset = new Vector3d(0, 0, value - currentCenter.Z);
                polyline.TransformBy(Matrix3d.Displacement(offset));
                break;
        }
    }

    private static void ApplyEllipseGeometryEdit(Ellipse ellipse, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx": ellipse.Center = new Point3d(value, ellipse.Center.Y, ellipse.Center.Z); break;
            case "centery": ellipse.Center = new Point3d(ellipse.Center.X, value, ellipse.Center.Z); break;
            case "centerz": ellipse.Center = new Point3d(ellipse.Center.X, ellipse.Center.Y, value); break;
        }
    }

    private static void ApplyBlockReferenceGeometryEdit(BlockReference blockRef, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx":
                var offsetX = new Vector3d(value - blockRef.Position.X, 0, 0);
                blockRef.TransformBy(Matrix3d.Displacement(offsetX));
                break;
            case "centery":
                var offsetY = new Vector3d(0, value - blockRef.Position.Y, 0);
                blockRef.TransformBy(Matrix3d.Displacement(offsetY));
                break;
            case "centerz":
                var offsetZ = new Vector3d(0, 0, value - blockRef.Position.Z);
                blockRef.TransformBy(Matrix3d.Displacement(offsetZ));
                break;
            case "scalex":
                if (value > 0) blockRef.ScaleFactors = new Scale3d(value, blockRef.ScaleFactors.Y, blockRef.ScaleFactors.Z);
                break;
            case "scaley":
                if (value > 0) blockRef.ScaleFactors = new Scale3d(blockRef.ScaleFactors.X, value, blockRef.ScaleFactors.Z);
                break;
            case "scalez":
                if (value > 0) blockRef.ScaleFactors = new Scale3d(blockRef.ScaleFactors.X, blockRef.ScaleFactors.Y, value);
                break;
            case "rotation":
                blockRef.Rotation = value * Math.PI / 180.0;
                break;
        }
    }

    private static void ApplyDBTextGeometryEdit(DBText dbText, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx": dbText.Position = new Point3d(value, dbText.Position.Y, dbText.Position.Z); break;
            case "centery": dbText.Position = new Point3d(dbText.Position.X, value, dbText.Position.Z); break;
            case "centerz": dbText.Position = new Point3d(dbText.Position.X, dbText.Position.Y, value); break;
            case "textheight": if (value > 0) dbText.Height = value; break;
            case "widthfactor": if (value > 0) dbText.WidthFactor = value; break;
            case "rotation": dbText.Rotation = value * Math.PI / 180.0; break;
        }
    }

    private static void ApplyMTextGeometryEdit(MText mText, string columnName, double value)
    {
        switch (columnName.ToLowerInvariant())
        {
            case "centerx": mText.Location = new Point3d(value, mText.Location.Y, mText.Location.Z); break;
            case "centery": mText.Location = new Point3d(mText.Location.X, value, mText.Location.Z); break;
            case "centerz": mText.Location = new Point3d(mText.Location.X, mText.Location.Y, value); break;
            case "textheight": if (value > 0) mText.TextHeight = value; break;
            case "width": if (value > 0) mText.Width = value; break;
            case "rotation": mText.Rotation = value * Math.PI / 180.0; break;
        }
    }

    private static void ApplyViewportGeometryEdit(Viewport viewport, string columnName, double value)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ed.WriteMessage($"\n  >> ApplyViewportGeometryEdit: Column='{columnName}', Value={value}");
        ed.WriteMessage($"\n  >> Current CenterPoint: {viewport.CenterPoint}");

        try
        {
            switch (columnName.ToLowerInvariant())
            {
                case "centerx":
                    var newPoint = new Point3d(value, viewport.CenterPoint.Y, viewport.CenterPoint.Z);
                    ed.WriteMessage($"\n  >> Setting CenterPoint from {viewport.CenterPoint} to {newPoint}");
                    viewport.CenterPoint = newPoint;
                    ed.WriteMessage($"\n  >> New CenterPoint: {viewport.CenterPoint}");
                    break;
                case "centery":
                    viewport.CenterPoint = new Point3d(viewport.CenterPoint.X, value, viewport.CenterPoint.Z);
                    break;
                case "centerz":
                    viewport.CenterPoint = new Point3d(viewport.CenterPoint.X, viewport.CenterPoint.Y, value);
                    break;
                case "width":
                    if (value > 0) viewport.Width = value;
                    break;
                case "height":
                    if (value > 0) viewport.Height = value;
                    break;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in ApplyViewportGeometryEdit: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
        }
    }

    private static void ApplyXrefPathEdit(BlockReference blockRef, string newPath, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = blockRef.Database;

        ed.WriteMessage($"\n  >> ApplyXrefPathEdit: BlockReference={blockRef.Name}, NewPath='{newPath}'");

        try
        {
            // Get the BlockTableRecord for this block reference
            var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

            if (btr == null)
            {
                ed.WriteMessage("\n  >> ERROR: Could not get BlockTableRecord");
                return;
            }

            // Check if this is actually an xref
            if (!btr.IsFromExternalReference)
            {
                ed.WriteMessage("\n  >> WARNING: Block is not an external reference (xref). Skipping path change.");
                return;
            }

            ed.WriteMessage($"\n  >> Current xref path: '{btr.PathName}'");

            // Validate the new path
            if (string.IsNullOrWhiteSpace(newPath))
            {
                ed.WriteMessage("\n  >> ERROR: New path cannot be empty");
                return;
            }

            // Only process DWG files for xrefs
            string extension = System.IO.Path.GetExtension(newPath).ToLowerInvariant();
            if (extension != ".dwg")
            {
                ed.WriteMessage($"\n  >> ERROR: XRef path must point to a .dwg file (got '{extension}')");
                return;
            }

            // Enable xref editing on the database
            bool wasXrefEditEnabled = db.XrefEditEnabled;
            db.XrefEditEnabled = true;

            try
            {
                // Set the new path
                btr.PathName = newPath;
                ed.WriteMessage($"\n  >> Successfully changed xref path to: '{newPath}'");
                ed.WriteMessage("\n  >> NOTE: You may need to reload the xref for changes to take effect.");
            }
            finally
            {
                // Restore original XrefEditEnabled state
                db.XrefEditEnabled = wasXrefEditEnabled;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in ApplyXrefPathEdit: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
        }
    }

    private static void RenameEntityLayout(Entity entity, string newLayoutName, Dictionary<string, object> entry, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = entity.Database;

        ed.WriteMessage($"\n  >> RenameEntityLayout: Entity={entity.GetType().Name}, NewLayoutName='{newLayoutName}'");

        try
        {
            // Get the ACTUAL current layout name by querying the entity's BlockId
            // This is more reliable than reading from the entry dictionary which may have been modified
            string currentLayoutName = "";

            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            // Find which layout this entity belongs to by checking its BlockId
            foreach (DBDictionaryEntry layoutEntry in layoutDict)
            {
                var tempLayout = (Layout)tr.GetObject(layoutEntry.Value, OpenMode.ForRead);
                if (entity.BlockId == tempLayout.BlockTableRecordId)
                {
                    currentLayoutName = tempLayout.LayoutName;
                    break;
                }
            }

            if (string.IsNullOrEmpty(currentLayoutName))
            {
                ed.WriteMessage("\n  >> ERROR: Could not determine current layout name from entity's BlockId");
                return;
            }

            ed.WriteMessage($"\n  >> Current layout (from entity BlockId): '{currentLayoutName}'");
            ed.WriteMessage($"\n  >> New layout name: '{newLayoutName}'");

            // Check if the user is trying to rename to the same name (no-op)
            if (string.Equals(currentLayoutName, newLayoutName, StringComparison.OrdinalIgnoreCase))
            {
                ed.WriteMessage("\n  >> Layout name unchanged, skipping");
                return;
            }

            // Cannot rename Model layout
            if (string.Equals(currentLayoutName, "Model", StringComparison.OrdinalIgnoreCase))
            {
                ed.WriteMessage("\n  >> Cannot rename Model layout");
                return;
            }

            // Find the layout by name
            if (!layoutDict.Contains(currentLayoutName))
            {
                ed.WriteMessage($"\n  >> ERROR: Layout '{currentLayoutName}' not found in layout dictionary");
                return;
            }

            // Get the Layout object
            var layoutId = layoutDict.GetAt(currentLayoutName);
            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForWrite);

            // Validate layout name for invalid characters
            char[] invalidChars = new char[] { '<', '>', '/', '\\', '"', ':', ';', '?', ',', '*', '|', '=', '`' };
            var foundInvalidChars = newLayoutName.Where(c => invalidChars.Contains(c)).Distinct().ToArray();
            if (foundInvalidChars.Length > 0)
            {
                string invalidCharsStr = string.Join(", ", foundInvalidChars.Select(c => $"'{c}'"));
                ed.WriteMessage($"\n  >> ERROR: Layout name contains invalid characters: {invalidCharsStr}");
                ed.WriteMessage($"\n  >> Invalid characters in AutoCAD layout names: < > / \\ \" : ; ? , * | = `");
                return;
            }

            // Check if the new name already exists
            string finalLayoutName = newLayoutName;
            if (layoutDict.Contains(newLayoutName))
            {
                int counter = 1;
                string uniqueName = newLayoutName;
                while (layoutDict.Contains(uniqueName))
                {
                    uniqueName = $"{newLayoutName}_{counter++}";
                }
                finalLayoutName = uniqueName;
                ed.WriteMessage($"\n  >> Layout name '{newLayoutName}' already exists, using '{finalLayoutName}' instead");
            }

            // Rename the layout
            layout.LayoutName = finalLayoutName;
            ed.WriteMessage($"\n  >> Successfully renamed layout from '{currentLayoutName}' to '{finalLayoutName}'");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in RenameEntityLayout: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
        }
    }

    private static void MoveEntityToLayout(Entity entity, string layoutName, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = entity.Database;

        ed.WriteMessage($"\n  >> MoveEntityToLayout: Entity={entity.GetType().Name}, TargetLayout='{layoutName}'");

        try
        {
            // Get the layout dictionary
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            // Check if the target layout exists
            if (!layoutDict.Contains(layoutName))
            {
                ed.WriteMessage($"\n  >> ERROR: Layout '{layoutName}' does not exist");
                return;
            }

            // Get the target layout
            var layoutId = layoutDict.GetAt(layoutName);
            var targetLayout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);

            // Get the target layout's block table record
            var targetBtr = (BlockTableRecord)tr.GetObject(targetLayout.BlockTableRecordId, OpenMode.ForWrite);

            // Get the current block table record that owns this entity
            var currentBtr = (BlockTableRecord)tr.GetObject(entity.OwnerId, OpenMode.ForWrite);

            // Check if entity is already in the target layout
            if (entity.OwnerId == targetLayout.BlockTableRecordId)
            {
                ed.WriteMessage($"\n  >> Entity is already in layout '{layoutName}'");
                return;
            }

            // Clone the entity to the target layout
            var idCollection = new ObjectIdCollection();
            idCollection.Add(entity.ObjectId);

            // Use DeepClone to copy the entity to the new layout
            var idMapping = new IdMapping();
            currentBtr.Database.DeepCloneObjects(idCollection, targetBtr.ObjectId, idMapping, false);

            // Get the cloned entity's ObjectId
            IdPair idPair = idMapping[entity.ObjectId];
            if (idPair.IsCloned)
            {
                ed.WriteMessage($"\n  >> Successfully cloned entity to layout '{layoutName}'");

                // Erase the original entity
                entity.UpgradeOpen();
                entity.Erase();
                ed.WriteMessage($"\n  >> Original entity erased");
            }
            else
            {
                ed.WriteMessage($"\n  >> ERROR: Failed to clone entity to target layout");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in MoveEntityToLayout: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
        }
    }

    private static void SwapBlockReference(BlockReference blockRef, string newBlockName, Transaction tr, Autodesk.AutoCAD.EditorInput.Editor ed)
    {
        var db = blockRef.Database;

        try
        {
            // Get the block table
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            // Get the new block definition
            if (!blockTable.Has(newBlockName))
            {
                ed.WriteMessage($"\n  >> ERROR: Block '{newBlockName}' not found in block table");
                return;
            }

            var newBlockId = blockTable[newBlockName];
            var newBtr = (BlockTableRecord)tr.GetObject(newBlockId, OpenMode.ForRead);

            // Get old block attributes before swapping
            var oldAttributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef != null)
                {
                    oldAttributes[attRef.Tag] = attRef.TextString;
                }
            }

            // Get the old block definition name for logging
            var oldBtr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            string oldBlockName = oldBtr?.Name ?? "Unknown";

            // Upgrade block reference to write mode
            blockRef.UpgradeOpen();

            // Collect old attribute IDs before changing the block
            var oldAttributeIds = new List<ObjectId>();
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                oldAttributeIds.Add(attId);
            }

            // Change the block reference to point to the new block definition
            blockRef.BlockTableRecord = newBlockId;

            // Remove old attributes after changing the block reference
            foreach (ObjectId attId in oldAttributeIds)
            {
                if (!attId.IsErased)
                {
                    var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                    if (attRef != null)
                    {
                        attRef.Erase();
                    }
                }
            }

            // Add new attributes from the new block definition
            int copiedCount = 0;
            int newCount = 0;
            foreach (ObjectId defId in newBtr)
            {
                var attDef = tr.GetObject(defId, OpenMode.ForRead) as AttributeDefinition;
                if (attDef != null && !attDef.Constant)
                {
                    var attRef = new AttributeReference();

                    // CRITICAL: Set database defaults to associate AttributeReference with correct database
                    // This prevents eWrongDatabase errors when appending to the block reference
                    attRef.SetDatabaseDefaults(db);

                    attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);

                    // Try to copy value from old attributes if tag matches
                    if (oldAttributes.TryGetValue(attDef.Tag, out string oldValue))
                    {
                        attRef.TextString = oldValue;
                        copiedCount++;
                    }
                    else
                    {
                        attRef.TextString = attDef.TextString; // Use default value
                        newCount++;
                    }

                    // Add the attribute to the block reference
                    blockRef.AttributeCollection.AppendAttribute(attRef);
                    tr.AddNewlyCreatedDBObject(attRef, true);
                }
            }

            // CRITICAL: Refresh the block reference to ensure AttributeCollection reflects the new attributes
            // Without this, subsequent reads of blockRef.AttributeCollection will return empty/stale data
            blockRef.RecordGraphicsModified(true);

            ed.WriteMessage($"\n  >> Swapped '{oldBlockName}' → '{newBlockName}' ({copiedCount} attrs copied, {newCount} new)");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in SwapBlockReference: {ex.Message}");
            throw;
        }
    }

    /// <summary>Try to parse a color string (supports color names and RGB format)</summary>
    private static bool TryParseColor(string colorString, out Autodesk.AutoCAD.Colors.Color color)
    {
        color = null;
        try
        {
            colorString = colorString?.Trim();
            if (string.IsNullOrEmpty(colorString))
                return false;

            // Try to parse as color index (e.g., "1", "256")
            if (int.TryParse(colorString, out int colorIndex))
            {
                color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, (short)colorIndex);
                return true;
            }

            // Try to parse RGB format like "RGB:255,0,0" or just "255,0,0"
            string rgbPattern = colorString.Replace("RGB:", "").Replace("rgb:", "");
            string[] parts = rgbPattern.Split(',');
            if (parts.Length == 3 &&
                byte.TryParse(parts[0].Trim(), out byte r) &&
                byte.TryParse(parts[1].Trim(), out byte g) &&
                byte.TryParse(parts[2].Trim(), out byte b))
            {
                color = Autodesk.AutoCAD.Colors.Color.FromRgb(r, g, b);
                return true;
            }

            // Could add color name parsing here if needed (e.g., "Red", "Blue")
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Try to parse a lineweight string</summary>
    private static bool TryParseLineWeight(string lineWeightString, out LineWeight lineWeight)
    {
        lineWeight = LineWeight.ByLayer;
        try
        {
            lineWeightString = lineWeightString?.Trim();
            if (string.IsNullOrEmpty(lineWeightString))
                return false;

            // Try parsing as enum name (e.g., "LineWeight000", "LineWeight025")
            if (Enum.TryParse<LineWeight>(lineWeightString, true, out lineWeight))
                return true;

            // Try parsing numeric values (e.g., "0", "25", "50")
            if (int.TryParse(lineWeightString, out int lwValue))
            {
                // Map common lineweight values
                switch (lwValue)
                {
                    case 0: lineWeight = LineWeight.LineWeight000; return true;
                    case 5: lineWeight = LineWeight.LineWeight005; return true;
                    case 9: lineWeight = LineWeight.LineWeight009; return true;
                    case 13: lineWeight = LineWeight.LineWeight013; return true;
                    case 15: lineWeight = LineWeight.LineWeight015; return true;
                    case 18: lineWeight = LineWeight.LineWeight018; return true;
                    case 20: lineWeight = LineWeight.LineWeight020; return true;
                    case 25: lineWeight = LineWeight.LineWeight025; return true;
                    case 30: lineWeight = LineWeight.LineWeight030; return true;
                    case 35: lineWeight = LineWeight.LineWeight035; return true;
                    case 40: lineWeight = LineWeight.LineWeight040; return true;
                    case 50: lineWeight = LineWeight.LineWeight050; return true;
                    case 53: lineWeight = LineWeight.LineWeight053; return true;
                    case 60: lineWeight = LineWeight.LineWeight060; return true;
                    case 70: lineWeight = LineWeight.LineWeight070; return true;
                    case 80: lineWeight = LineWeight.LineWeight080; return true;
                    case 90: lineWeight = LineWeight.LineWeight090; return true;
                    case 100: lineWeight = LineWeight.LineWeight100; return true;
                    case 106: lineWeight = LineWeight.LineWeight106; return true;
                    case 120: lineWeight = LineWeight.LineWeight120; return true;
                    case 140: lineWeight = LineWeight.LineWeight140; return true;
                    case 158: lineWeight = LineWeight.LineWeight158; return true;
                    case 200: lineWeight = LineWeight.LineWeight200; return true;
                    case 211: lineWeight = LineWeight.LineWeight211; return true;
                    case -1: lineWeight = LineWeight.ByLayer; return true;
                    case -2: lineWeight = LineWeight.ByBlock; return true;
                    case -3: lineWeight = LineWeight.ByLineWeightDefault; return true;
                    default: return false;
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Try to parse a transparency string</summary>
    private static bool TryParseTransparency(string transparencyString, out Autodesk.AutoCAD.Colors.Transparency transparency)
    {
        transparency = new Autodesk.AutoCAD.Colors.Transparency();
        try
        {
            transparencyString = transparencyString?.Trim();
            if (string.IsNullOrEmpty(transparencyString))
                return false;

            // Try parsing as percentage (e.g., "50%")
            if (transparencyString.EndsWith("%"))
            {
                string percentStr = transparencyString.TrimEnd('%');
                if (int.TryParse(percentStr, out int percent))
                {
                    // Convert percentage to alpha value (0-255)
                    // 0% = opaque (alpha 255), 100% = fully transparent (alpha 0)
                    byte alpha = (byte)((100 - percent) * 255 / 100);
                    transparency = new Autodesk.AutoCAD.Colors.Transparency(alpha);
                    return true;
                }
            }

            // Try parsing as alpha value (0-255)
            if (byte.TryParse(transparencyString, out byte alpha2))
            {
                transparency = new Autodesk.AutoCAD.Colors.Transparency(alpha2);
                return true;
            }

            return false;
        }
        catch
        {
            return false;
        }
    }
}
