using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Register the command class
[assembly: CommandClass(typeof(FilterSelectionElements))]

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

    public static List<Dictionary<string, object>> GetEntityData(Editor ed, bool selectedOnly = false, bool includeProperties = false)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        var entityData = new List<Dictionary<string, object>>();
        var currentScope = SelectionScopeManager.CurrentScope;

        // Handle selection based on current scope
        if (currentScope == SelectionScopeManager.SelectionScope.view)
        {
            // Get pickfirst set (pre-selected objects)
            var selResult = ed.SelectImplied();

            // If there is no pickfirst set, request user to select objects
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to filter: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                // Clear the pickfirst set since we're consuming it
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status == PromptStatus.OK && selResult.Value.Count > 0)
            {
                // Get current selection
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var objectId in selResult.Value.GetObjectIds())
                    {
                        try
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead);
                            if (entity != null)
                            {
                                var data = GetEntityDataDictionary(entity, doc.Name, null, includeProperties);
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
                return entityData;
            }
            else
            {
                throw new InvalidOperationException("No selection found. Please select entities first when in 'view' scope.");
            }
        }

        if (selectedOnly)
        {
            // Get entities from stored selection
            var storedSelection = SelectionStorage.LoadSelection();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                throw new InvalidOperationException("No stored selection found. Use commands like 'select-by-category' first.");
            }

            // Filter to current session to avoid confusion with selections from different AutoCAD processes
            var currentSessionId = GetCurrentSessionId();
            storedSelection = storedSelection.Where(item =>
                string.IsNullOrEmpty(item.SessionId) || item.SessionId == currentSessionId).ToList();

            // If in document scope, filter to current document only
            if (currentScope == SelectionScopeManager.SelectionScope.document)
            {
                var currentDocPath = Path.GetFullPath(doc.Name);
                var originalCount = storedSelection.Count;

                storedSelection = storedSelection.Where(item =>
                {
                    try
                    {
                        var itemPath = Path.GetFullPath(item.DocumentPath);
                        return string.Equals(itemPath, currentDocPath, StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        // If path resolution fails, fall back to direct comparison
                        return string.Equals(item.DocumentPath, doc.Name, StringComparison.OrdinalIgnoreCase);
                    }
                }).ToList();

            }

            // Check if filtering resulted in empty selection
            if (storedSelection.Count == 0)
            {
                if (currentScope == SelectionScopeManager.SelectionScope.document)
                {
                    throw new InvalidOperationException($"No stored selection found for current document '{Path.GetFileName(doc.Name)}'. The stored selection may be from other documents in process scope.");
                }
                else
                {
                    throw new InvalidOperationException("No stored selection found after filtering. Use commands like 'select-by-category' first.");
                }
            }

            // Process stored selection items
            foreach (var item in storedSelection)
            {
                try
                {
                    // Check if this is from the current document
                    if (Path.GetFullPath(item.DocumentPath) == Path.GetFullPath(doc.Name))
                    {
                        // Current document - get entity directly
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (objectId != ObjectId.Null)
                        {
                            using (var tr = db.TransactionManager.StartTransaction())
                            {
                                var entity = tr.GetObject(objectId, OpenMode.ForRead);
                                if (entity != null)
                                {
                                    var data = GetEntityDataDictionary(entity, item.DocumentPath, null, includeProperties);
                                    data["ObjectId"] = objectId; // Store for selection
                                    entityData.Add(data);
                                }
                                tr.Commit();
                            }
                        }
                    }
                    else
                    {
                        // Different document - retrieve properties from external document
                        var data = GetExternalEntityData(item.DocumentPath, item.Handle);
                        entityData.Add(data);
                    }
                }
                catch
                {
                    // Skip problematic entities
                    continue;
                }
            }
        }
        else
        {
            // Get all entities from current scope (fallback - should not be used by filter-selection)
            var entities = GatherEntitiesFromScope(db, currentScope);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var objectId in entities)
                {
                    try
                    {
                        var entity = tr.GetObject(objectId, OpenMode.ForRead);
                        if (entity != null)
                        {
                            var data = GetEntityDataDictionary(entity, doc.Name, null, includeProperties);
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

    private static Dictionary<string, object> GetExternalEntityData(string documentPath, string handle)
    {
        var data = new Dictionary<string, object>
        {
            ["Name"] = "External Reference",
            ["Category"] = "External Entity",
            ["Layer"] = "N/A",
            ["Color"] = "N/A",
            ["LineType"] = "N/A",
            ["Layout"] = "N/A",
            ["DocumentPath"] = documentPath,
            ["DocumentName"] = Path.GetFileName(documentPath),
            ["Handle"] = handle,
            ["Id"] = handle,
            ["IsExternal"] = true,
            ["DisplayName"] = $"External: {Path.GetFileName(documentPath)}"
        };

        try
        {
            // Try to open the external document and get real entity properties
            var docs = AcadApp.DocumentManager;
            Document externalDoc = null;
            bool docWasAlreadyOpen = false;

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

            // If not already open, try to open it temporarily
            if (externalDoc == null && File.Exists(documentPath))
            {
                try
                {
                    externalDoc = docs.Open(documentPath, false); // Open read-only
                    docWasAlreadyOpen = false;
                }
                catch
                {
                    // If we can't open the document, return the N/A data
                    return data;
                }
            }

            // If we have the external document, get the entity properties
            if (externalDoc != null)
            {
                try
                {
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
                                data = GetEntityDataDictionary(entity, documentPath, null, false);
                                data["IsExternal"] = true;
                                data["DisplayName"] = $"External: {data["Name"]}";
                            }
                            tr.Commit();
                        }
                    }
                }
                finally
                {
                    // Close the document if we opened it temporarily
                    if (!docWasAlreadyOpen && externalDoc != null)
                    {
                        try
                        {
                            externalDoc.CloseAndDiscard();
                        }
                        catch
                        {
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

    private static Dictionary<string, object> GetEntityDataDictionary(DBObject entity, string documentPath, string spaceName, bool includeProperties)
    {
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
            layoutName = GetEntityLayoutName(ent);

            // Get entity-specific name
            if (entity is BlockReference br)
            {
                using (var tr = br.Database.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    entityName = btr?.Name ?? "Block";
                    tr.Commit();
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

        var data = new Dictionary<string, object>
        {
            ["Name"] = entityName,
            ["Category"] = GetEntityCategory(entity),
            ["Layer"] = layer,
            ["Color"] = color,
            ["LineType"] = lineType,
            ["Layout"] = layoutName,
            ["DocumentPath"] = documentPath,
            ["DocumentName"] = Path.GetFileName(documentPath),
            ["Handle"] = entity.Handle.ToString(),
            ["Id"] = entity.ObjectId.Handle.Value,
            ["IsExternal"] = false,
            ["ObjectId"] = entity.ObjectId // Store for selection
        };

        data["DisplayName"] = !string.IsNullOrEmpty(entityName) ? entityName : data["Category"].ToString();

        // Add space information if available
        if (!string.IsNullOrEmpty(spaceName))
        {
            data["Space"] = spaceName;
        }

        // Add block attributes if entity is a block reference
        if (entity is BlockReference blockRef)
        {
            AddBlockAttributes(blockRef, data);
        }

        // Add XData and extension dictionary data
        AddExtensionData(entity, data);

        // Include properties if requested
        if (includeProperties && entity is Entity entityWithProps)
        {
            try
            {
                // Add common AutoCAD properties
                data["Area"] = GetEntityArea(entityWithProps);
                data["Length"] = GetEntityLength(entityWithProps);
                // data["Thickness"] = entityWithProps.Thickness; // Not available on base Entity
                data["Elevation"] = GetEntityElevation(entityWithProps);
            }
            catch { /* Skip if properties can't be read */ }
        }

        return data;
    }

    private static List<ObjectId> GatherEntitiesFromScope(Database db, SelectionScopeManager.SelectionScope scope)
    {
        var entities = new List<ObjectId>();

        using (var tr = db.TransactionManager.StartTransaction())
        {
            switch (scope)
            {
                case SelectionScopeManager.SelectionScope.view:
                    var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id in currentSpace)
                    {
                        entities.Add(id);
                    }
                    break;

                case SelectionScopeManager.SelectionScope.document:
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
                    // For Process, Desktop, Network - fall back to Document scope for now
                    goto case SelectionScopeManager.SelectionScope.document;
            }

            tr.Commit();
        }

        return entities;
    }

    private static string GetEntityCategory(DBObject entity)
    {
        // Use the same categorization logic as select-by-category.cs
        string typeName = entity.GetType().Name;

        if (entity is BlockReference)
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

    private static void AddExtensionData(DBObject entity, Dictionary<string, object> data)
    {
        try
        {
            // Add XData
            var xData = entity.XData;
            if (xData != null)
            {
                var xDataApps = new List<string>();
                var xDataValues = new List<string>();

                foreach (TypedValue typedValue in xData)
                {
                    if (typedValue.TypeCode == (int)DxfCode.ExtendedDataRegAppName)
                    {
                        xDataApps.Add(typedValue.Value.ToString());
                    }
                    else if (typedValue.TypeCode == (int)DxfCode.ExtendedDataAsciiString ||
                             typedValue.TypeCode == (int)DxfCode.ExtendedDataReal ||
                             typedValue.TypeCode == (int)DxfCode.ExtendedDataInteger16 ||
                             typedValue.TypeCode == (int)DxfCode.ExtendedDataInteger32)
                    {
                        if (typedValue.Value != null)
                        {
                            xDataValues.Add(typedValue.Value.ToString());
                        }
                    }
                }

                if (xDataApps.Any())
                {
                    data["xdata_apps"] = string.Join(", ", xDataApps);
                }
                if (xDataValues.Any())
                {
                    data["xdata_values"] = string.Join(", ", xDataValues.Take(3)); // Limit to first 3 values
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
}

/// <summary>
/// Base class for commands that display AutoCAD entities in a custom data grid for filtering and reselection.
/// Works with the stored selection system used by commands like select-by-category.
/// </summary>
public abstract class FilterElementsBase
{
    public abstract bool SpanAllScreens { get; }
    public abstract bool UseSelectedElements { get; }
    public abstract bool IncludeProperties { get; }

    public void Execute()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            var entityData = FilterEntityDataHelper.GetEntityData(ed, UseSelectedElements, IncludeProperties);

            if (!entityData.Any())
            {
                ed.WriteMessage("\nNo entities found.\n");
                return;
            }

            // Add index to each entity for mapping back after user selection
            for (int i = 0; i < entityData.Count; i++)
            {
                entityData[i]["OriginalIndex"] = i;
            }

            // Get property names, excluding internal fields
            var propertyNames = entityData.First().Keys
                .Where(k => !k.EndsWith("ObjectId") && k != "OriginalIndex")
                .ToList();

            // Reorder to put most useful columns first
            var orderedProps = new List<string> { "Name", "Category", "Layer", "Layout", "DocumentName", "Color", "LineType", "Handle" };
            var remainingProps = propertyNames.Except(orderedProps);

            // Separate attributes and extension data for better organization
            var attributeProps = remainingProps.Where(p => p.StartsWith("attr_")).OrderBy(p => p);
            var extensionProps = remainingProps.Where(p => p.StartsWith("xdata_") || p.StartsWith("ext_dict_")).OrderBy(p => p);
            var otherProps = remainingProps.Where(p => !p.StartsWith("attr_") && !p.StartsWith("xdata_") && !p.StartsWith("ext_dict_") && p != "DocumentPath" && p != "DisplayName").OrderBy(p => p);
            var documentPathProp = propertyNames.Contains("DocumentPath") ? new[] { "DocumentPath" } : new string[0];

            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(attributeProps)
                .Concat(extensionProps)
                .Concat(otherProps)
                .Concat(documentPathProp)
                .ToList();

            var chosenRows = CustomGUIs.DataGrid(entityData, propertyNames, SpanAllScreens, null);
            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo entities selected.\n");
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
                    selectedIds.Add(objectId);
                }
            }

            // Set selection for current document entities
            if (selectedIds.Count > 0)
            {
                // For process scope, we need to actually set the AutoCAD selection (not just save to storage)
                // to properly narrow down the selection as expected
                if (SelectionScopeManager.CurrentScope == SelectionScopeManager.SelectionScope.process)
                {
                    ed.SetImpliedSelection(selectedIds.ToArray());
                }
                else
                {
                    ed.SetImpliedSelectionEx(selectedIds.ToArray());
                }
                ed.WriteMessage($"\n{selectedIds.Count} entities selected in current document.\n");
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
            if ((selectedIds.Count > 0 || externalEntities.Count > 0) && SelectionScopeManager.CurrentScope != SelectionScopeManager.SelectionScope.view)
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
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nUnexpected error: {ex.Message}\n");
        }
    }
}

/// <summary>
/// Concrete command class
/// </summary>
public class FilterSelectionElements
{
    [CommandMethod("filter-selection", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void FilterSelectionCommand()
    {
        var command = new FilterSelectionImpl();
        command.Execute();
    }
}

public class FilterSelectionImpl : FilterElementsBase
{
    public override bool SpanAllScreens => false;
    public override bool UseSelectedElements => SelectionScopeManager.CurrentScope != SelectionScopeManager.SelectionScope.view;
    public override bool IncludeProperties => true;
}