using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class SelectDuplicates
    {
        // View scope: Find duplicates in current view using pickfirst or selection prompt
        public static void ExecuteViewScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Get entities from current space
            var entities = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    if (id.ObjectClass.DxfName != "VIEWPORT")
                    {
                        entities.Add(id);
                    }
                }

                tr.Commit();
            }

            if (entities.Count == 0)
            {
                ed.WriteMessage("\nNo entities found in current view.\n");
                return;
            }

            // Find duplicates
            var duplicateIds = FindDuplicates(db, entities);

            if (duplicateIds.Count == 0)
            {
                ed.WriteMessage("\nNo duplicate entities found.\n");
                return;
            }

            // Set selection
            ed.SetImpliedSelection(duplicateIds.ToArray());
            ed.WriteMessage($"\nSelected {duplicateIds.Count} duplicate entities.\n");
        }

        // Document scope: Find duplicates in entire document
        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            // Get all entities from all layouts
            var entities = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Get model space
                var modelBtr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId id in modelBtr)
                {
                    if (id.ObjectClass.DxfName != "VIEWPORT")
                    {
                        entities.Add(id);
                    }
                }

                // Get all paper spaces
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (!layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase))
                    {
                        var paperBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                        foreach (ObjectId id in paperBtr)
                        {
                            if (id.ObjectClass.DxfName != "VIEWPORT")
                            {
                                entities.Add(id);
                            }
                        }
                    }
                }

                tr.Commit();
            }

            if (entities.Count == 0)
            {
                ed.WriteMessage("\nNo entities found in document.\n");
                return;
            }

            // Find duplicates
            var duplicateIds = FindDuplicates(db, entities);

            if (duplicateIds.Count == 0)
            {
                ed.WriteMessage("\nNo duplicate entities found.\n");
                return;
            }

            // Set selection
            ed.SetImpliedSelection(duplicateIds.ToArray());
            ed.WriteMessage($"\nSelected {duplicateIds.Count} duplicate entities across entire document.\n");
        }

        // Session scope: Find duplicates across all open documents
        public static void ExecuteApplicationScope(Editor ed)
        {
            var docs = AcadApp.DocumentManager;
            var currentDoc = docs.MdiActiveDocument;
            var allDuplicates = new List<ObjectId>();

            // Process each open document
            foreach (Document doc in docs)
            {
                var db = doc.Database;
                var entities = new List<ObjectId>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // Get model space
                    var modelBtr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    foreach (ObjectId id in modelBtr)
                    {
                        if (id.ObjectClass.DxfName != "VIEWPORT")
                        {
                            entities.Add(id);
                        }
                    }

                    // Get all paper spaces
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        if (!layout.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase))
                        {
                            var paperBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                            foreach (ObjectId id in paperBtr)
                            {
                                if (id.ObjectClass.DxfName != "VIEWPORT")
                                {
                                    entities.Add(id);
                                }
                            }
                        }
                    }

                    tr.Commit();
                }

                // Find duplicates in this document
                var duplicateIds = FindDuplicates(db, entities);

                // Only add duplicates from current document to selection
                if (doc == currentDoc)
                {
                    allDuplicates.AddRange(duplicateIds);
                }
            }

            if (allDuplicates.Count == 0)
            {
                ed.WriteMessage("\nNo duplicate entities found in current document.\n");
                return;
            }

            // Set selection in current document
            ed.SetImpliedSelection(allDuplicates.ToArray());
            ed.WriteMessage($"\nSelected {allDuplicates.Count} duplicate entities in current document.\n");
        }

        // Core duplicate detection logic
        private static List<ObjectId> FindDuplicates(Database db, List<ObjectId> entities)
        {
            var duplicateIds = new List<ObjectId>();
            var signatureGroups = new Dictionary<string, List<ObjectId>>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in entities)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    try
                    {
                        var signature = GetEntitySignature(entity);
                        if (!signatureGroups.ContainsKey(signature))
                        {
                            signatureGroups[signature] = new List<ObjectId>();
                        }
                        signatureGroups[signature].Add(id);
                    }
                    catch
                    {
                        // Skip entities that fail signature generation
                        continue;
                    }
                }

                // Print diagnostics for duplicate hatches
                var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                foreach (var kvp in signatureGroups)
                {
                    if (kvp.Value.Count > 1)
                    {
                        var firstEntity = tr.GetObject(kvp.Value[0], OpenMode.ForRead) as Entity;
                        if (firstEntity is Hatch)
                        {
                            ed.WriteMessage($"\n=== DUPLICATE HATCH GROUP ({kvp.Value.Count} hatches) ===\n");
                            ed.WriteMessage($"Signature: {kvp.Key}\n");

                            for (int i = 0; i < kvp.Value.Count; i++)
                            {
                                var hatch = tr.GetObject(kvp.Value[i], OpenMode.ForRead) as Hatch;
                                if (hatch != null)
                                {
                                    ed.WriteMessage($"\nHatch #{i + 1} (Handle: {hatch.Handle}):\n");
                                    ed.WriteMessage($"  Layer: {hatch.Layer}\n");
                                    ed.WriteMessage($"  PatternName: {hatch.PatternName}\n");
                                    ed.WriteMessage($"  PatternScale: {hatch.PatternScale}\n");
                                    ed.WriteMessage($"  PatternAngle: {hatch.PatternAngle}\n");
                                    ed.WriteMessage($"  PatternType: {hatch.PatternType}\n");
                                    ed.WriteMessage($"  HatchObjectType: {hatch.HatchObjectType}\n");
                                    ed.WriteMessage($"  NumberOfLoops: {hatch.NumberOfLoops}\n");
                                    ed.WriteMessage($"  Elevation: {hatch.Elevation}\n");
                                    ed.WriteMessage($"  Normal: {hatch.Normal}\n");
                                    ed.WriteMessage($"  Area: {hatch.Area}\n");
                                    ed.WriteMessage($"  Associative: {hatch.Associative}\n");

                                    // Print loop details
                                    for (int loopIndex = 0; loopIndex < hatch.NumberOfLoops; loopIndex++)
                                    {
                                        var loop = hatch.GetLoopAt(loopIndex);
                                        ed.WriteMessage($"  Loop {loopIndex}: Type={loop.LoopType}, IsPolyline={loop.IsPolyline}, Curves={loop.Curves?.Count ?? 0}\n");

                                        if (loop.IsPolyline && loop.Polyline != null && loop.Polyline.Count > 0)
                                        {
                                            ed.WriteMessage($"    Polyline vertices: {loop.Polyline.Count}\n");
                                            for (int v = 0; v < Math.Min(5, loop.Polyline.Count); v++)
                                            {
                                                var bulge = loop.Polyline[v];
                                                ed.WriteMessage($"      [{v}] {bulge.Vertex.X:F12}, {bulge.Vertex.Y:F12}\n");
                                            }
                                            if (loop.Polyline.Count > 5)
                                            {
                                                ed.WriteMessage($"      ... ({loop.Polyline.Count - 5} more vertices)\n");
                                            }
                                        }
                                    }

                                    // Print geometric extents
                                    try
                                    {
                                        var extents = hatch.GeometricExtents;
                                        ed.WriteMessage($"  Extents: Min({extents.MinPoint.X:F6},{extents.MinPoint.Y:F6}) Max({extents.MaxPoint.X:F6},{extents.MaxPoint.Y:F6})\n");
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }

                // Collect all entities that have duplicates, but skip the first one in each group
                foreach (var group in signatureGroups.Values)
                {
                    if (group.Count > 1)
                    {
                        // Skip the first entity, select all others as duplicates
                        duplicateIds.AddRange(group.Skip(1));
                    }
                }

                tr.Commit();
            }

            return duplicateIds;
        }

        // Generate a signature string for an entity based on its properties and geometry
        private static string GetEntitySignature(Entity entity)
        {
            var parts = new List<string>();

            // Entity type
            parts.Add(entity.GetType().Name);

            // Common properties
            parts.Add(entity.Layer ?? "0");
            parts.Add(entity.Color.ColorIndex.ToString());
            parts.Add(entity.Linetype ?? "");
            parts.Add(entity.LineWeight.ToString());
            parts.Add(entity.LinetypeScale.ToString("F12"));
            parts.Add(entity.Transparency.ToString());

            // Geometric properties based on entity type
            switch (entity)
            {
                case Line line:
                    parts.Add(PointToString(line.StartPoint));
                    parts.Add(PointToString(line.EndPoint));
                    break;

                case Circle circle:
                    parts.Add(PointToString(circle.Center));
                    parts.Add(circle.Radius.ToString("F12"));
                    parts.Add(VectorToString(circle.Normal));
                    break;

                case Arc arc:
                    parts.Add(PointToString(arc.Center));
                    parts.Add(arc.Radius.ToString("F12"));
                    parts.Add(arc.StartAngle.ToString("F12"));
                    parts.Add(arc.EndAngle.ToString("F12"));
                    parts.Add(VectorToString(arc.Normal));
                    break;

                case Polyline pline:
                    parts.Add(pline.NumberOfVertices.ToString());
                    parts.Add(pline.Closed.ToString());
                    parts.Add(pline.Elevation.ToString("F12"));
                    parts.Add(VectorToString(pline.Normal));
                    for (int i = 0; i < pline.NumberOfVertices; i++)
                    {
                        parts.Add(Point2dToString(pline.GetPoint2dAt(i)));
                        parts.Add(pline.GetBulgeAt(i).ToString("F12"));
                    }
                    break;

                case Polyline2d pline2d:
                    parts.Add(pline2d.PolyType.ToString());
                    parts.Add(pline2d.Closed.ToString());
                    parts.Add(pline2d.Elevation.ToString("F12"));
                    parts.Add(VectorToString(pline2d.Normal));
                    // Note: Vertex iteration would require opening vertex objects
                    break;

                case Polyline3d pline3d:
                    parts.Add(pline3d.PolyType.ToString());
                    parts.Add(pline3d.Closed.ToString());
                    // Note: Vertex iteration would require opening vertex objects
                    break;

                case DBText text:
                    parts.Add(text.TextString ?? "");
                    parts.Add(PointToString(text.Position));
                    parts.Add(text.Height.ToString("F12"));
                    parts.Add(text.Rotation.ToString("F12"));
                    parts.Add(text.WidthFactor.ToString("F12"));
                    parts.Add(text.Oblique.ToString("F12"));
                    parts.Add(text.TextStyleName ?? "");
                    parts.Add(text.HorizontalMode.ToString());
                    parts.Add(text.VerticalMode.ToString());
                    break;

                case MText mtext:
                    parts.Add(mtext.Contents ?? "");
                    parts.Add(PointToString(mtext.Location));
                    parts.Add(mtext.TextHeight.ToString("F12"));
                    parts.Add(mtext.Rotation.ToString("F12"));
                    parts.Add(mtext.Width.ToString("F12"));
                    parts.Add(mtext.TextStyleName ?? "");
                    parts.Add(mtext.Attachment.ToString());
                    break;

                case BlockReference blockRef:
                    parts.Add(blockRef.Name ?? "");
                    parts.Add(PointToString(blockRef.Position));
                    parts.Add(blockRef.Rotation.ToString("F12"));
                    parts.Add(VectorToString(blockRef.ScaleFactors));
                    parts.Add(VectorToString(blockRef.Normal));
                    break;

                case Ellipse ellipse:
                    parts.Add(PointToString(ellipse.Center));
                    parts.Add(VectorToString(ellipse.MajorAxis));
                    parts.Add(ellipse.MinorRadius.ToString("F12"));
                    parts.Add(ellipse.StartAngle.ToString("F12"));
                    parts.Add(ellipse.EndAngle.ToString("F12"));
                    parts.Add(VectorToString(ellipse.Normal));
                    break;

                case Spline spline:
                    parts.Add(spline.Degree.ToString());
                    parts.Add(spline.Closed.ToString());
                    parts.Add(spline.IsPeriodic.ToString());
                    parts.Add(spline.NumControlPoints.ToString());
                    for (int i = 0; i < spline.NumControlPoints; i++)
                    {
                        parts.Add(PointToString(spline.GetControlPointAt(i)));
                    }
                    break;

                case Hatch hatch:
                    parts.Add(hatch.PatternName ?? "");
                    parts.Add(hatch.PatternScale.ToString("F12"));
                    parts.Add(hatch.PatternAngle.ToString("F12"));
                    parts.Add(hatch.NumberOfLoops.ToString());
                    parts.Add(hatch.Elevation.ToString("F12"));
                    parts.Add(VectorToString(hatch.Normal));
                    parts.Add(hatch.Area.ToString("F12"));

                    // Include all loop boundary geometry
                    for (int loopIndex = 0; loopIndex < hatch.NumberOfLoops; loopIndex++)
                    {
                        var loop = hatch.GetLoopAt(loopIndex);
                        parts.Add($"Loop{loopIndex}:Type={loop.LoopType}");

                        if (loop.IsPolyline && loop.Polyline != null && loop.Polyline.Count > 0)
                        {
                            parts.Add($"Polyline:{loop.Polyline.Count}");
                            foreach (BulgeVertex bulge in loop.Polyline)
                            {
                                parts.Add($"{bulge.Vertex.X:F12},{bulge.Vertex.Y:F12},{bulge.Bulge:F12}");
                            }
                        }
                        else if (loop.Curves != null && loop.Curves.Count > 0)
                        {
                            parts.Add($"Curves:{loop.Curves.Count}");
                            foreach (Curve2d curve in loop.Curves)
                            {
                                // Include curve type and basic properties
                                parts.Add(curve.GetType().Name);

                                if (curve is LineSegment2d line2d)
                                {
                                    parts.Add(Point2dToString(line2d.StartPoint));
                                    parts.Add(Point2dToString(line2d.EndPoint));
                                }
                                else if (curve is CircularArc2d arc2d)
                                {
                                    parts.Add(Point2dToString(arc2d.Center));
                                    parts.Add(arc2d.Radius.ToString("F12"));
                                    parts.Add(arc2d.StartAngle.ToString("F12"));
                                    parts.Add(arc2d.EndAngle.ToString("F12"));
                                }
                                else if (curve is EllipticalArc2d ellipse2d)
                                {
                                    parts.Add(Point2dToString(ellipse2d.Center));
                                    parts.Add(ellipse2d.MajorRadius.ToString("F12"));
                                    parts.Add(ellipse2d.MinorRadius.ToString("F12"));
                                    parts.Add(ellipse2d.StartAngle.ToString("F12"));
                                    parts.Add(ellipse2d.EndAngle.ToString("F12"));
                                }
                            }
                        }
                    }
                    break;

                case Dimension dim:
                    parts.Add(dim.DimensionText ?? "");
                    parts.Add(PointToString(dim.TextPosition));
                    parts.Add(dim.DimensionStyleName ?? "");
                    break;

                case Leader leader:
                    parts.Add(leader.NumVertices.ToString());
                    for (int i = 0; i < leader.NumVertices; i++)
                    {
                        parts.Add(PointToString(leader.VertexAt(i)));
                    }
                    parts.Add(leader.DimensionStyleName ?? "");
                    break;

                default:
                    // For unknown entity types, try to use geometric extents
                    try
                    {
                        var extents = entity.GeometricExtents;
                        parts.Add(PointToString(extents.MinPoint));
                        parts.Add(PointToString(extents.MaxPoint));
                    }
                    catch
                    {
                        // If extents are not available, use ObjectId as fallback
                        parts.Add(entity.ObjectId.ToString());
                    }
                    break;
            }

            return string.Join("|", parts);
        }

        private static string PointToString(Point3d point)
        {
            return $"{point.X:F12},{point.Y:F12},{point.Z:F12}";
        }

        private static string Point2dToString(Point2d point)
        {
            return $"{point.X:F12},{point.Y:F12}";
        }

        private static string VectorToString(Vector3d vector)
        {
            return $"{vector.X:F12},{vector.Y:F12},{vector.Z:F12}";
        }

        private static string VectorToString(Scale3d scale)
        {
            return $"{scale.X:F12},{scale.Y:F12},{scale.Z:F12}";
        }
    }
}
