using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.OutlineAllViewports))]

namespace AutoCADBallet
{
    // Shared static transient manager - EXACT pattern from test-transient-clear.cs
    public static class ViewportTransientManager
    {
        public static TransientManager TransientMgr = TransientManager.CurrentTransientManager;
        public static IntegerCollection IntCollection = new IntegerCollection();
        public static int Marker = 1000;
    }

    // Helper class for text positioning with collision detection
    public class TextBounds
    {
        public double MinX { get; set; }
        public double MaxX { get; set; }
        public double MinY { get; set; }
        public double MaxY { get; set; }

        public bool Overlaps(TextBounds other)
        {
            return !(MaxX < other.MinX || MinX > other.MaxX || MaxY < other.MinY || MinY > other.MaxY);
        }
    }

    public static class TextPositionHelper
    {
        private static List<TextBounds> _placedTextBounds = new List<TextBounds>();

        public static void Reset()
        {
            _placedTextBounds.Clear();
        }

        public static Autodesk.AutoCAD.Geometry.Point3d FindNonOverlappingPosition(
            double startX, double startY, double startZ,
            double textHeight, List<string> textLines,
            out double finalWidth)
        {
            // Calculate text bounds (approximate width based on character count)
            double maxLineLength = textLines.Max(line => line.Length);
            double textWidth = maxLineLength * textHeight * 0.6; // Approximate character width
            double totalHeight = textLines.Count * textHeight * 1.2; // Line spacing: 1.2x

            finalWidth = textWidth;

            // Try positions: right, below, above, left
            var positions = new[]
            {
                new { X = startX + textHeight * 0.5, Y = startY, Label = "right" },
                new { X = startX, Y = startY - totalHeight - textHeight * 0.5, Label = "below" },
                new { X = startX, Y = startY + textHeight * 0.5, Label = "above" },
                new { X = startX - textWidth - textHeight * 0.5, Y = startY, Label = "left" }
            };

            foreach (var pos in positions)
            {
                var bounds = new TextBounds
                {
                    MinX = pos.X,
                    MaxX = pos.X + textWidth,
                    MinY = pos.Y - totalHeight,
                    MaxY = pos.Y
                };

                // Check if this position overlaps with any existing text
                bool overlaps = _placedTextBounds.Any(existing => existing.Overlaps(bounds));
                if (!overlaps)
                {
                    // Found a non-overlapping position
                    _placedTextBounds.Add(bounds);
                    return new Autodesk.AutoCAD.Geometry.Point3d(pos.X, pos.Y, startZ);
                }
            }

            // If all positions overlap, use wrapped text at original position
            _placedTextBounds.Add(new TextBounds
            {
                MinX = startX,
                MaxX = startX + textWidth,
                MinY = startY - totalHeight,
                MaxY = startY
            });

            return new Autodesk.AutoCAD.Geometry.Point3d(startX, startY, startZ);
        }
    }

    public class OutlineAllViewports
    {

        [CommandMethod("outline-all-viewports", CommandFlags.Modal)]
        public void OutlineAllViewportsCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            // Check if we're NOT in Model Space
            if (!db.TileMode)
            {
                ed.WriteMessage("\nThis command must be run from Model Space (not Paper Space layouts).\n");
                ed.WriteMessage("Switch to Model tab and run the command again.\n");
                return;
            }

            try
            {
                // Reset text position tracking for collision detection
                TextPositionHelper.Reset();

                var outlineViewports = new OutlineViewports();
                var viewports = CollectAllViewports(db);

                if (viewports.Count == 0)
                {
                    ed.WriteMessage("\nNo viewports found in any layouts.\n");
                    return;
                }

                // Get document ID for persistent storage
                var documentId = db.OriginalFileName;
                if (string.IsNullOrEmpty(documentId))
                {
                    documentId = doc.Name; // Use current name if not saved
                }
                // Sanitize for filename
                documentId = Path.GetFileNameWithoutExtension(documentId);

                ed.WriteMessage($"\n[DEBUG] Document ID: {documentId}\n");

                // Load existing groups
                var existingGroups = TransientGraphicsStorage.LoadGroups(documentId);
                ed.WriteMessage($"[DEBUG] Loaded {existingGroups.Count} existing groups from storage\n");

                // Track new groups
                var newGroups = new List<TransientGraphicsGroup>();

                // Outline all viewports
                int outlined = 0;
                foreach (var vp in viewports)
                {
                    int markerBefore = ViewportTransientManager.Marker;
                    if (OutlineViewportDirect(db, vp))
                    {
                        outlined++;

                        // Get the viewport to calculate scale
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var viewport = tr.GetObject(vp.ViewportId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Viewport;
                            if (viewport != null)
                            {
                                var scale = viewport.ViewHeight / viewport.Height;
                                var scaleText = $"1:{Math.Round(scale, 2)}";

                                // Track which markers were used for this viewport
                                var markers = new List<int>();
                                for (int i = markerBefore; i < ViewportTransientManager.Marker; i++)
                                {
                                    markers.Add(i);
                                }

                                // Create description string
                                var description = $"viewport: {vp.LayoutName}, scale: {scaleText}, marker count: {markers.Count}";

                                newGroups.Add(new TransientGraphicsGroup
                                {
                                    Description = description,
                                    Markers = markers
                                });
                            }
                            tr.Commit();
                        }
                    }
                }

                ed.WriteMessage($"[DEBUG] Created {newGroups.Count} new groups\n");

                // Combine and save all groups
                var allGroups = existingGroups.Concat(newGroups).ToList();
                TransientGraphicsStorage.SaveGroups(documentId, allGroups);
                ed.WriteMessage($"[DEBUG] Saved {allGroups.Count} total groups to storage\n");

                ed.WriteMessage($"\n{outlined} viewport(s) outlined with transient graphics.\n");
                ed.WriteMessage("Use 'clear-all-transient-graphics' to remove all outlines.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in outline-all-viewports: {ex.Message}\n");
            }
        }

        private List<OutlineViewports.ViewportInfo> CollectAllViewports(Database db)
        {
            var viewports = new List<OutlineViewports.ViewportInfo>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                    // Skip Model Space
                    if (layout.LayoutName == "Model")
                        continue;

                    var blockTableRecord = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                    foreach (ObjectId objId in blockTableRecord)
                    {
                        var entity = tr.GetObject(objId, OpenMode.ForRead);

                        if (entity is Autodesk.AutoCAD.DatabaseServices.Viewport vp && vp.Number != 1) // Skip the overall viewport (number 1)
                        {
                            viewports.Add(new OutlineViewports.ViewportInfo
                            {
                                LayoutName = layout.LayoutName,
                                ViewportHandle = vp.Handle.ToString(),
                                ViewportId = objId,
                                Center = vp.CenterPoint,
                                Width = vp.Width,
                                Height = vp.Height
                            });
                        }
                    }
                }

                tr.Commit();
            }

            return viewports;
        }

        private bool OutlineViewportDirect(Database db, OutlineViewports.ViewportInfo vpInfo)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var viewport = tr.GetObject(vpInfo.ViewportId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Viewport;
                    if (viewport == null)
                        return false;

                    // Get viewport boundary in paper space
                    var boundary = GetViewportBoundary(viewport);

                    // Transform to model space coordinates
                    var modelSpacePoints = new List<Autodesk.AutoCAD.Geometry.Point3d>();
                    foreach (var pt in boundary)
                    {
                        var transformed = TransformPaperToModel(pt, viewport);
                        modelSpacePoints.Add(transformed);
                    }

                    // Create transient graphics entities
                    var entities = new List<Entity>();

                    // Create polyline outline
                    using (var polyline = new Autodesk.AutoCAD.DatabaseServices.Polyline())
                    {
                        for (int i = 0; i < modelSpacePoints.Count; i++)
                        {
                            polyline.AddVertexAt(i, new Autodesk.AutoCAD.Geometry.Point2d(modelSpacePoints[i].X, modelSpacePoints[i].Y), 0, 0, 0);
                        }
                        polyline.Closed = true;
                        polyline.ColorIndex = 1; // Red

                        // Clone the polyline for transient graphics
                        entities.Add(polyline.Clone() as Entity);
                    }

                    // Create text labels at top-right corner of viewport bounding box
                    var minX = modelSpacePoints.Min(p => p.X);
                    var maxX = modelSpacePoints.Max(p => p.X);
                    var minY = modelSpacePoints.Min(p => p.Y);
                    var maxY = modelSpacePoints.Max(p => p.Y);
                    var boundingWidth = maxX - minX;
                    var boundingHeight = maxY - minY;

                    // Calculate text height as 5% of the smaller dimension of bounding box
                    var textHeight = Math.Min(boundingWidth, boundingHeight) * 0.05;
                    if (textHeight < 1.0) textHeight = 1.0; // Minimum text height

                    // Calculate viewport scale
                    var scale = viewport.ViewHeight / viewport.Height;
                    var scaleText = $"1:{Math.Round(scale, 2)}";

                    // Build multi-line text content
                    var textLines = new List<string>
                    {
                        vpInfo.LayoutName,
                        $"Scale: {scaleText}"
                    };

                    // Find non-overlapping position for text (tries right, below, above, left)
                    var preferredX = maxX + textHeight * 0.5; // Prefer right of top-right corner
                    var preferredY = maxY - textHeight * 0.6; // Aligned to top
                    var preferredZ = modelSpacePoints[0].Z;

                    double textWidth;
                    var textPosition = TextPositionHelper.FindNonOverlappingPosition(
                        preferredX, preferredY, preferredZ,
                        textHeight, textLines, out textWidth);

                    // Create each line of text at the determined position
                    for (int i = 0; i < textLines.Count; i++)
                    {
                        var linePosition = new Autodesk.AutoCAD.Geometry.Point3d(
                            textPosition.X,
                            textPosition.Y - (i * textHeight * 1.2), // Line spacing: 1.2x text height
                            textPosition.Z
                        );

                        using (var text = new DBText())
                        {
                            text.Position = linePosition;
                            text.TextString = textLines[i];
                            text.Height = textHeight;
                            text.ColorIndex = 3; // Green
                            text.HorizontalMode = TextHorizontalMode.TextLeft;
                            text.VerticalMode = TextVerticalMode.TextTop;
                            text.AlignmentPoint = linePosition;

                            entities.Add(text.Clone() as Entity);
                        }
                    }

                    // Add transients directly - EXACT pattern from test-transient-clear.cs
                    foreach (var entity in entities)
                    {
                        ViewportTransientManager.TransientMgr.AddTransient(entity, TransientDrawingMode.DirectTopmost, ViewportTransientManager.Marker++, ViewportTransientManager.IntCollection);
                    }

                    tr.Commit();
                    return true;
                }
            }
            catch (System.Exception ex)
            {
                var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"\nError outlining viewport {vpInfo.LayoutName}: {ex.Message}\n");
                return false;
            }
        }

        private List<Autodesk.AutoCAD.Geometry.Point3d> GetViewportBoundary(Autodesk.AutoCAD.DatabaseServices.Viewport viewport)
        {
            var boundary = new List<Autodesk.AutoCAD.Geometry.Point3d>();

            // Check if viewport has a non-rectangular clip boundary
            // This matches LISP: (cdr (assoc 340 vpt)) - DXF code 340 is the clip entity
            if (viewport.NonRectClipEntityId != ObjectId.Null)
            {
                // Get the clipping boundary polyline
                using (var tr = viewport.Database.TransactionManager.StartTransaction())
                {
                    var clipEntity = tr.GetObject(viewport.NonRectClipEntityId, OpenMode.ForRead);

                    if (clipEntity is Autodesk.AutoCAD.DatabaseServices.Polyline pline)
                    {
                        // Extract vertices from the polyline
                        for (int i = 0; i < pline.NumberOfVertices; i++)
                        {
                            var point2d = pline.GetPoint2dAt(i);
                            boundary.Add(new Autodesk.AutoCAD.Geometry.Point3d(point2d.X, point2d.Y, 0));
                        }
                    }
                    else if (clipEntity is Polyline2d pline2d)
                    {
                        // Handle 2D polyline (old-style)
                        foreach (ObjectId vertexId in pline2d)
                        {
                            var vertex = tr.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                            if (vertex != null)
                            {
                                var pos = vertex.Position;
                                boundary.Add(new Autodesk.AutoCAD.Geometry.Point3d(pos.X, pos.Y, 0));
                            }
                        }
                    }

                    tr.Commit();
                }
            }

            // If no clip boundary or failed to get vertices, use rectangular boundary
            if (boundary.Count == 0)
            {
                var center = viewport.CenterPoint;
                var halfWidth = viewport.Width / 2.0;
                var halfHeight = viewport.Height / 2.0;

                // Create rectangle corners in paper space
                boundary.Add(new Autodesk.AutoCAD.Geometry.Point3d(center.X - halfWidth, center.Y - halfHeight, 0));
                boundary.Add(new Autodesk.AutoCAD.Geometry.Point3d(center.X + halfWidth, center.Y - halfHeight, 0));
                boundary.Add(new Autodesk.AutoCAD.Geometry.Point3d(center.X + halfWidth, center.Y + halfHeight, 0));
                boundary.Add(new Autodesk.AutoCAD.Geometry.Point3d(center.X - halfWidth, center.Y + halfHeight, 0));
            }

            return boundary;
        }

        private Autodesk.AutoCAD.Geometry.Point3d TransformPaperToModel(Autodesk.AutoCAD.Geometry.Point3d paperPoint, Autodesk.AutoCAD.DatabaseServices.Viewport viewport)
        {
            // This is a direct port of the LISP PCS2WCS function
            // Translates PCS (Paper Coordinate System) point to WCS (World Coordinate System)

            // Get DXF properties from viewport
            var vpCenter = viewport.CenterPoint;           // DXF 10 - viewport center in paper space
            var viewCenter = viewport.ViewCenter;          // DXF 12 - view center (2D in DCS)
            var viewTarget = viewport.ViewTarget;          // DXF 17 - view target (3D in WCS)
            var viewNormal = viewport.ViewDirection;       // DXF 16 - view direction normal
            var viewHeight = viewport.ViewHeight;          // DXF 45 - view height
            var vpHeight = viewport.Height;                // DXF 41 - viewport height
            var twistAngle = viewport.TwistAngle;          // DXF 51 - twist angle

            // Calculate scale (from LISP: scl (/ (cdr (assoc 45 enx)) (cdr (assoc 41 enx))))
            double scale = viewHeight / vpHeight;

            // LISP step 1: (vxs pnt scl) - multiply point by scale
            double scaledX = paperPoint.X * scale;
            double scaledY = paperPoint.Y * scale;
            double scaledZ = paperPoint.Z * scale;

            // LISP step 2: (vxs (cdr (assoc 10 enx)) (- scl)) - subtract scaled viewport center
            scaledX -= vpCenter.X * scale;
            scaledY -= vpCenter.Y * scale;
            scaledZ -= vpCenter.Z * scale;

            // LISP step 3: Add view center (cdr (assoc 12 enx))
            scaledX += viewCenter.X;
            scaledY += viewCenter.Y;
            // Note: viewCenter is 2D, no Z component to add

            // LISP step 4: Apply rotation matrix for twist angle
            // ang (- (cdr (assoc 51 enx))) - negative twist angle
            double angle = -twistAngle;
            double cosA = Math.Cos(angle);
            double sinA = Math.Sin(angle);

            // Rotation matrix: [[cos -sin 0], [sin cos 0], [0 0 1]]
            double rotatedX = scaledX * cosA - scaledY * sinA;
            double rotatedY = scaledX * sinA + scaledY * cosA;
            double rotatedZ = scaledZ;

            // LISP step 5: Transform by view normal using matrix multiplication
            // mat (mxm (trans-identity-to-normal nor) (rotation-matrix ang))
            // In LISP: (mapcar (function (lambda ( v ) (trans v 0 nor t))) identity-matrix)

            // Create transformation matrix from WCS to DCS based on view normal
            var normal = viewNormal.GetNormal();

            // Use AutoCAD's trans function equivalent: transform from WCS (0) to view normal coordinate system
            // This creates the transformation matrix that aligns world axes to the view
            var xAxisWCS = new Autodesk.AutoCAD.Geometry.Vector3d(1, 0, 0);
            var yAxisWCS = new Autodesk.AutoCAD.Geometry.Vector3d(0, 1, 0);
            var zAxisWCS = new Autodesk.AutoCAD.Geometry.Vector3d(0, 0, 1);

            // Transform world axes by view normal (equivalent to LISP trans v 0 nor t)
            var matrix = Autodesk.AutoCAD.Geometry.Matrix3d.PlaneToWorld(normal);
            var xAxisDCS = xAxisWCS.TransformBy(matrix);
            var yAxisDCS = yAxisWCS.TransformBy(matrix);
            var zAxisDCS = zAxisWCS.TransformBy(matrix);

            // Apply the transformed matrix to rotated point (mxv mat point)
            var transformedX = rotatedX * xAxisDCS.X + rotatedY * xAxisDCS.Y + rotatedZ * xAxisDCS.Z;
            var transformedY = rotatedX * yAxisDCS.X + rotatedY * yAxisDCS.Y + rotatedZ * yAxisDCS.Z;
            var transformedZ = rotatedX * zAxisDCS.X + rotatedY * zAxisDCS.Y + rotatedZ * zAxisDCS.Z;

            // LISP final step: (mapcar '+ transformed-point (cdr (assoc 17 enx)))
            // Add view target
            var finalX = transformedX + viewTarget.X;
            var finalY = transformedY + viewTarget.Y;
            var finalZ = transformedZ + viewTarget.Z;

            return new Autodesk.AutoCAD.Geometry.Point3d(finalX, finalY, finalZ);
        }

        private Autodesk.AutoCAD.Geometry.Point3d CalculateCenter(List<Autodesk.AutoCAD.Geometry.Point3d> points)
        {
            if (points.Count == 0)
                return Autodesk.AutoCAD.Geometry.Point3d.Origin;

            double sumX = 0, sumY = 0, sumZ = 0;
            foreach (var pt in points)
            {
                sumX += pt.X;
                sumY += pt.Y;
                sumZ += pt.Z;
            }

            return new Autodesk.AutoCAD.Geometry.Point3d(sumX / points.Count, sumY / points.Count, sumZ / points.Count);
        }

    }
}
