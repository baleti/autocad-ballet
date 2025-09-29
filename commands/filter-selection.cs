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
            // Get entities from stored selection based on scope
            List<SelectionItem> storedSelection;

            if (scope == SelectionScope.document)
            {
                // Document scope - load selection for current document only
                var docName = Path.GetFileNameWithoutExtension(doc.Name);
                storedSelection = SelectionStorage.LoadSelection(docName);
            }
            else if (scope == SelectionScope.application)
            {
                // Application scope - load from all open documents
                storedSelection = SelectionStorage.LoadSelectionFromAllDocuments();
            }
            else
            {
                // View scope - load from global file
                storedSelection = SelectionStorage.LoadSelection();
            }

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

            // Filter to current session to avoid confusion with selections from different AutoCAD processes
            var currentSessionId = GetCurrentSessionId();
            storedSelection = storedSelection.Where(item =>
                string.IsNullOrEmpty(item.SessionId) || item.SessionId == currentSessionId).ToList();

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
                        var data = GetExternalEntityData(item.DocumentPath, item.Handle, includeProperties);
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
            var entities = GatherEntitiesFromScope(db, scope);

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

    public static Dictionary<string, object> GetExternalEntityData(string documentPath, string handle, bool includeProperties = false)
    {
        var data = new Dictionary<string, object>
        {
            ["Name"] = "External Reference",
            ["Category"] = "External Entity",
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
                                data = GetEntityDataDictionary(entity, documentPath, null, includeProperties);
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

    public static Dictionary<string, object> GetEntityDataDictionary(DBObject entity, string documentPath, string spaceName, bool includeProperties)
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
            ["Contents"] = GetEntityContents(entity),
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

        // Add plot settings for Layout entities
        if (entity is Layout layoutEntity)
        {
            AddLayoutPlotSettings(layoutEntity, data);
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
                data["Elevation"] = GetEntityElevation(entityWithProps);

                // Add geometry properties
                var geometryProps = GetEntityGeometryProperties(entityWithProps);
                foreach (var prop in geometryProps)
                {
                    data[prop.Key] = prop.Value;
                }
            }
            catch { /* Skip if properties can't be read */ }
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
                    // For application scope - fall back to Document scope for now
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

    private static string GetEntityCategory(DBObject entity)
    {
        // Use the same categorization logic as select-by-categories.cs
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

    public virtual void Execute()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ObjectId[] originalSelection = null;

        try
        {
            var entityData = FilterEntityDataHelper.GetEntityData(ed, Scope, out originalSelection, UseSelectedElements, IncludeProperties);

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
            var orderedProps = new List<string> { "Name", "Contents", "Category", "Layer", "Layout", "DocumentName", "Color", "LineType", "Handle" };
            var remainingProps = propertyNames.Except(orderedProps);

            // Separate geometry properties, attributes and extension data for better organization
            var geometryProps = remainingProps.Where(p => FilterEntityDataHelper.IsGeometryProperty(p)).OrderBy(p => FilterEntityDataHelper.GetGeometryPropertyOrder(p));
            var attributeProps = remainingProps.Where(p => p.StartsWith("attr_")).OrderBy(p => p);
            var extensionProps = remainingProps.Where(p => p.StartsWith("xdata_") || p.StartsWith("ext_dict_")).OrderBy(p => p);
            var otherProps = remainingProps.Where(p => !p.StartsWith("attr_") && !p.StartsWith("xdata_") && !p.StartsWith("ext_dict_") && p != "DocumentPath" && p != "DisplayName" && !FilterEntityDataHelper.IsGeometryProperty(p)).OrderBy(p => p);
            var documentPathProp = propertyNames.Contains("DocumentPath") ? new[] { "DocumentPath" } : new string[0];

            propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
                .Concat(geometryProps)
                .Concat(attributeProps)
                .Concat(extensionProps)
                .Concat(otherProps)
                .Concat(documentPathProp)
                .ToList();

            // Reset the edits flag at the start of each DataGrid session
            CustomGUIs.ResetEditsAppliedFlag();

            var chosenRows = CustomGUIs.DataGrid(entityData, propertyNames, SpanAllScreens, null);

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
                    // For application scope, we need to actually set the AutoCAD selection (not just save to storage)
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
// - filter-selection-in-application
