using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;

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
        public Dictionary<string, object> Entry { get; set; }
    }

    private static Dictionary<string, List<PendingEdit>> GroupEditsByDocument()
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var editsByDocument = new Dictionary<string, List<PendingEdit>>(StringComparer.OrdinalIgnoreCase);

        foreach (var kvp in _pendingCellEdits)
        {
            string[] parts = kvp.Key.Split('|');
            if (parts.Length == 2)
            {
                int rowIndex = int.Parse(parts[0]);
                string columnName = parts[1];
                string newValue = kvp.Value?.ToString() ?? "";

                if (rowIndex >= 0 && rowIndex < _cachedFilteredData.Count)
                {
                    var entry = _cachedFilteredData[rowIndex];
                    string documentPath;
                    if (entry.TryGetValue("DocumentPath", out var docPathObj))
                        documentPath = System.IO.Path.GetFullPath(docPathObj.ToString());
                    else
                        documentPath = System.IO.Path.GetFullPath(doc.Name);

                    var pendingEdit = new PendingEdit
                    {
                        RowIndex = rowIndex,
                        ColumnName = columnName,
                        NewValue = newValue,
                        Entry = entry
                    };

                    if (!editsByDocument.ContainsKey(documentPath))
                        editsByDocument[documentPath] = new List<PendingEdit>();
                    editsByDocument[documentPath].Add(pendingEdit);

                    ed.WriteMessage($"\nGrouped edit for {System.IO.Path.GetFileName(documentPath)}: {columnName} = '{newValue}'");
                }
            }
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
                                ApplyEditToDBObject(dbObject, edit.ColumnName, edit.NewValue, tr);
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
                                    ApplyEditToDBObjectInExternalDocument(dbObject, edit.ColumnName, edit.NewValue, tr);
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

    private static void ApplyEditToDBObjectInExternalDocument(DBObject dbObject, string columnName, string newValue, Transaction tr)
    {
        var currentDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = currentDoc.Editor;
        ed.WriteMessage($"\n  >> ApplyEditToDBObjectInExternalDocument: DBObject={dbObject.GetType().Name}, Column='{columnName}', Value='{newValue}'");
        try
        {
            switch (columnName.ToLowerInvariant())
            {
                case "contents":
                    if (dbObject is MText mtextExt) mtextExt.Contents = newValue;
                    else if (dbObject is DBText textExt) textExt.TextString = newValue;
                    else if (dbObject is Dimension dimExt) dimExt.DimensionText = newValue;
                    break;
                case "layout":
                case "name":
                    if (dbObject is Layout layout)
                    {
                        try
                        {
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) return;
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
                    else if (dbObject is MText mtextExt2) mtextExt2.Contents = newValue;
                    else if (dbObject is DBText textExt2) textExt2.TextString = newValue;
                    else if (dbObject is BlockReference blockRef)
                    {
                        var blockTableRecord = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                        if (blockTableRecord != null)
                        {
                            var blockTable = (BlockTable)tr.GetObject(dbObject.Database.BlockTableId, OpenMode.ForRead);
                            if (blockTable.Has(newValue))
                            {
                                int counter = 1; string uniqueName = newValue;
                                while (blockTable.Has(uniqueName)) uniqueName = $"{newValue}_{counter++}";
                                newValue = uniqueName;
                            }
                            blockTableRecord.Name = newValue;
                        }
                    }
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
            ed.WriteMessage($"\n  >> ERROR in ApplyEditToDBObjectInExternalDocument: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
            throw;
        }
    }

    private static void ApplyEditToDBObject(DBObject dbObject, string columnName, string newValue, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ed.WriteMessage($"\n  >> ApplyEditToDBObject: DBObject={dbObject.GetType().Name}, Column='{columnName}', Value='{newValue}'");
        try
        {
            switch (columnName.ToLowerInvariant())
            {
                case "layer":
                    if (dbObject is Entity entity) entity.Layer = newValue; else ed.WriteMessage("\n  >> Not an Entity");
                    break;
                case "color":
                    if (dbObject is Entity entity2 && int.TryParse(newValue, out int colorIndex)) entity2.ColorIndex = colorIndex;
                    break;
                case "linetype":
                    if (dbObject is Entity entity3) entity3.Linetype = newValue;
                    break;
                case "contents":
                    if (dbObject is MText mtextContents) mtextContents.Contents = newValue;
                    else if (dbObject is DBText textContents) textContents.TextString = newValue;
                    else if (dbObject is Dimension dimContents) dimContents.DimensionText = newValue;
                    break;
                case "layout":
                case "name":
                    if (dbObject is Layout layout)
                    {
                        try
                        {
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase)) return;
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
                    else if (dbObject is MText mtext2) mtext2.Contents = newValue;
                    else if (dbObject is DBText text2) text2.TextString = newValue;
                    else if (dbObject is BlockReference br)
                    {
                        var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                        if (btr != null)
                        {
                            var blockTable = (BlockTable)tr.GetObject(dbObject.Database.BlockTableId, OpenMode.ForRead);
                            if (blockTable.Has(newValue))
                            {
                                int counter = 1; string uniqueName = newValue;
                                while (blockTable.Has(uniqueName)) uniqueName = $"{newValue}_{counter++}";
                                newValue = uniqueName;
                            }
                            btr.Name = newValue;
                        }
                    }
                    else if (dbObject is LayerTableRecord layerRec) layerRec.Name = newValue;
                    else if (dbObject is TextStyleTableRecord tStyle) tStyle.Name = newValue;
                    else if (dbObject is LinetypeTableRecord ltype) ltype.Name = newValue;
                    else if (dbObject is DimStyleTableRecord dStyle) dStyle.Name = newValue;
                    else if (dbObject is UcsTableRecord ucs) ucs.Name = newValue;
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
                default:
                    if (columnName.StartsWith("attr_")) { if (dbObject is Entity e4) ApplyBlockAttributeEdit(e4, columnName, newValue, tr); }
                    else if (columnName.StartsWith("xdata_")) { if (dbObject is Entity e5) ApplyXDataEdit(e5, columnName, newValue); }
                    else if (columnName.StartsWith("ext_dict_")) { if (dbObject is Entity e6) ApplyExtensionDictEdit(e6, columnName, newValue, tr); }
                    break;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in ApplyEditToDBObject: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
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
        if (!double.TryParse(newValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double value)) return;
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
        }
    }

    private static void ApplyCircleGeometryEdit(Circle circle, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx": circle.Center = new Point3d(value, circle.Center.Y, circle.Center.Z); break;
            case "centery": circle.Center = new Point3d(circle.Center.X, value, circle.Center.Z); break;
            case "centerz": circle.Center = new Point3d(circle.Center.X, circle.Center.Y, value); break;
            case "radius": if (value > 0) circle.Radius = value; break;
        }
    }

    private static void ApplyArcGeometryEdit(Arc arc, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx": arc.Center = new Point3d(value, arc.Center.Y, arc.Center.Z); break;
            case "centery": arc.Center = new Point3d(arc.Center.X, value, arc.Center.Z); break;
            case "centerz": arc.Center = new Point3d(arc.Center.X, arc.Center.Y, value); break;
            case "radius": if (value > 0) arc.Radius = value; break;
        }
    }

    private static void ApplyLineGeometryEdit(Line line, string columnName, double value)
    {
        switch (columnName)
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
        switch (columnName)
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
        switch (columnName)
        {
            case "centerx": ellipse.Center = new Point3d(value, ellipse.Center.Y, ellipse.Center.Z); break;
            case "centery": ellipse.Center = new Point3d(ellipse.Center.X, value, ellipse.Center.Z); break;
            case "centerz": ellipse.Center = new Point3d(ellipse.Center.X, ellipse.Center.Y, value); break;
        }
    }

    private static void ApplyBlockReferenceGeometryEdit(BlockReference blockRef, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx": blockRef.Position = new Point3d(value, blockRef.Position.Y, blockRef.Position.Z); break;
            case "centery": blockRef.Position = new Point3d(blockRef.Position.X, value, blockRef.Position.Z); break;
            case "centerz": blockRef.Position = new Point3d(blockRef.Position.X, blockRef.Position.Y, value); break;
            case "scalex": if (value > 0) blockRef.ScaleFactors = new Scale3d(value, blockRef.ScaleFactors.Y, blockRef.ScaleFactors.Z); break;
            case "scaley": if (value > 0) blockRef.ScaleFactors = new Scale3d(blockRef.ScaleFactors.X, value, blockRef.ScaleFactors.Z); break;
            case "scalez": if (value > 0) blockRef.ScaleFactors = new Scale3d(blockRef.ScaleFactors.X, blockRef.ScaleFactors.Y, value); break;
            case "rotation": blockRef.Rotation = value * Math.PI / 180.0; break;
        }
    }

    private static void ApplyDBTextGeometryEdit(DBText dbText, string columnName, double value)
    {
        switch (columnName)
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
        switch (columnName)
        {
            case "centerx": mText.Location = new Point3d(value, mText.Location.Y, mText.Location.Z); break;
            case "centery": mText.Location = new Point3d(mText.Location.X, value, mText.Location.Z); break;
            case "centerz": mText.Location = new Point3d(mText.Location.X, mText.Location.Y, value); break;
            case "textheight": if (value > 0) mText.TextHeight = value; break;
            case "width": if (value > 0) mText.Width = value; break;
            case "rotation": mText.Rotation = value * Math.PI / 180.0; break;
        }
    }
}
