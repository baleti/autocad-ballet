using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.InsertVertex))]

namespace AutoCADBallet
{
    public class InsertVertex
    {
        [CommandMethod("insert-vertex", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void InsertVertexCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // Prompt user to select a polyline
                var selectionOpts = new PromptEntityOptions("\nSelect polyline: ");
                selectionOpts.SetRejectMessage("\nMust be a polyline.");
                selectionOpts.AddAllowedClass(typeof(Polyline), true);

                var selectionResult = ed.GetEntity(selectionOpts);

                if (selectionResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nCommand cancelled.\n");
                    return;
                }

                var polylineId = selectionResult.ObjectId;

                // Prompt user to pick a point on the polyline
                var pointOpts = new PromptPointOptions("\nSpecify point on polyline to insert vertex: ");
                var pointResult = ed.GetPoint(pointOpts);

                if (pointResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nCommand cancelled.\n");
                    return;
                }

                var insertPoint = pointResult.Value;

                // Insert the vertex
                using (var docLock = doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var polyline = (Polyline)tr.GetObject(polylineId, OpenMode.ForWrite);

                        // Find the closest point on the polyline and which segment it's on
                        var closestPoint = polyline.GetClosestPointTo(insertPoint, false);

                        // Get the parameter at the closest point
                        var param = polyline.GetParameterAtPoint(closestPoint);

                        // The parameter value represents the vertex index + fractional distance
                        // For example, param = 1.5 means halfway between vertex 1 and vertex 2
                        int segmentIndex = (int)Math.Floor(param);

                        // Insert the vertex after the segment index
                        // We need to insert at segmentIndex + 1
                        int insertIndex = segmentIndex + 1;

                        // Make sure we don't exceed the number of vertices
                        if (insertIndex > polyline.NumberOfVertices)
                        {
                            insertIndex = polyline.NumberOfVertices;
                        }

                        // Get width and bulge properties from the segment being split
                        double startWidth = 0;
                        double endWidth = 0;
                        double bulge = 0;

                        // Check if polyline has a constant width
                        if (polyline.ConstantWidth > 0)
                        {
                            startWidth = polyline.ConstantWidth;
                            endWidth = polyline.ConstantWidth;
                        }
                        else
                        {
                            // Get the widths from the segment we're splitting
                            // The segment from segmentIndex to segmentIndex+1 has these properties
                            startWidth = polyline.GetStartWidthAt(segmentIndex);
                            endWidth = polyline.GetEndWidthAt(segmentIndex);

                            // Interpolate width based on position along segment
                            double segmentFraction = param - segmentIndex;
                            double interpolatedWidth = startWidth + (endWidth - startWidth) * segmentFraction;

                            // Use interpolated width for both start and end of new vertex
                            startWidth = interpolatedWidth;
                            endWidth = interpolatedWidth;
                        }

                        // Get the bulge from the segment (for arc segments)
                        bulge = polyline.GetBulgeAt(segmentIndex);

                        // If there's a bulge, we need to split it proportionally
                        if (Math.Abs(bulge) > 1e-10)
                        {
                            double segmentFraction = param - segmentIndex;

                            // Split the arc: adjust the bulge at the original vertex
                            // and calculate the bulge for the new segment
                            // This is complex geometry, so for now we'll zero out the bulge
                            // to maintain straight segments (can be enhanced later)
                            double newBulge1 = bulge * segmentFraction;
                            double newBulge2 = bulge * (1 - segmentFraction);

                            // Set the first segment's bulge
                            polyline.SetBulgeAt(segmentIndex, newBulge1);
                            bulge = newBulge2;
                        }

                        // Add the vertex at the closest point on the polyline with inherited properties
                        polyline.AddVertexAt(insertIndex, new Point2d(closestPoint.X, closestPoint.Y), bulge, startWidth, endWidth);

                        ed.WriteMessage($"\nVertex inserted at segment {segmentIndex} with width {startWidth:F3}.\n");

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in insert-vertex: {ex.Message}\n");
            }
        }
    }
}
