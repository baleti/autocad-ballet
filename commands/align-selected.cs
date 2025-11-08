using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace AutoCADBallet
{
    public static class AlignSelected
    {
        public enum AlignmentType
        {
            Top,
            Bottom,
            Left,
            Right,
            CenterVertically,
            CenterHorizontally
        }

        public static void ExecuteViewScope(Editor ed, AlignmentType alignmentType)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Get pickfirst set (pre-selected objects)
            var selResult = ed.SelectImplied();

            // If there is no pickfirst set, request user to select objects
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to align: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                // Clear the pickfirst set since we're consuming it
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects selected.\n");
                return;
            }

            var selectedObjects = selResult.Value.GetObjectIds();

            if (selectedObjects.Length < 2)
            {
                ed.WriteMessage("\nNeed at least 2 objects to align.\n");
                return;
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Collect entity extents
                var entityData = new List<EntityExtentData>();

                foreach (ObjectId objId in selectedObjects)
                {
                    var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    try
                    {
                        var extents = entity.GeometricExtents;
                        entityData.Add(new EntityExtentData
                        {
                            ObjectId = objId,
                            Extents = extents
                        });
                    }
                    catch
                    {
                        // Skip entities without valid extents
                        continue;
                    }
                }

                if (entityData.Count < 2)
                {
                    ed.WriteMessage("\nNeed at least 2 objects with valid extents to align.\n");
                    tr.Abort();
                    return;
                }

                // Calculate alignment reference
                double alignmentValue = CalculateAlignmentReference(entityData, alignmentType);

                // Align entities
                int alignedCount = 0;
                foreach (var data in entityData)
                {
                    var entity = tr.GetObject(data.ObjectId, OpenMode.ForWrite) as Entity;
                    if (entity == null) continue;

                    Vector3d displacement = CalculateDisplacement(data.Extents, alignmentValue, alignmentType);

                    if (!displacement.IsZeroLength())
                    {
                        Matrix3d transform = Matrix3d.Displacement(displacement);
                        entity.TransformBy(transform);
                        alignedCount++;
                    }
                }

                tr.Commit();

                string alignmentName = GetAlignmentName(alignmentType);
                ed.WriteMessage($"\nAligned {alignedCount} objects ({alignmentName}).\n");
            }
        }

        private static double CalculateAlignmentReference(List<EntityExtentData> entityData, AlignmentType alignmentType)
        {
            switch (alignmentType)
            {
                case AlignmentType.Top:
                    // Align to the highest top edge
                    return entityData.Max(e => e.Extents.MaxPoint.Y);

                case AlignmentType.Bottom:
                    // Align to the lowest bottom edge
                    return entityData.Min(e => e.Extents.MinPoint.Y);

                case AlignmentType.Left:
                    // Align to the leftmost left edge
                    return entityData.Min(e => e.Extents.MinPoint.X);

                case AlignmentType.Right:
                    // Align to the rightmost right edge
                    return entityData.Max(e => e.Extents.MaxPoint.X);

                case AlignmentType.CenterVertically:
                    // Align to the overall vertical center (center on vertical axis)
                    double minX = entityData.Min(e => e.Extents.MinPoint.X);
                    double maxX = entityData.Max(e => e.Extents.MaxPoint.X);
                    return (minX + maxX) / 2.0;

                case AlignmentType.CenterHorizontally:
                    // Align to the overall horizontal center (center on horizontal axis)
                    double minY = entityData.Min(e => e.Extents.MinPoint.Y);
                    double maxY = entityData.Max(e => e.Extents.MaxPoint.Y);
                    return (minY + maxY) / 2.0;

                default:
                    return 0.0;
            }
        }

        private static Vector3d CalculateDisplacement(Extents3d extents, double alignmentValue, AlignmentType alignmentType)
        {
            switch (alignmentType)
            {
                case AlignmentType.Top:
                    // Move entity so its top edge is at alignmentValue
                    double currentTop = extents.MaxPoint.Y;
                    return new Vector3d(0, alignmentValue - currentTop, 0);

                case AlignmentType.Bottom:
                    // Move entity so its bottom edge is at alignmentValue
                    double currentBottom = extents.MinPoint.Y;
                    return new Vector3d(0, alignmentValue - currentBottom, 0);

                case AlignmentType.Left:
                    // Move entity so its left edge is at alignmentValue
                    double currentLeft = extents.MinPoint.X;
                    return new Vector3d(alignmentValue - currentLeft, 0, 0);

                case AlignmentType.Right:
                    // Move entity so its right edge is at alignmentValue
                    double currentRight = extents.MaxPoint.X;
                    return new Vector3d(alignmentValue - currentRight, 0, 0);

                case AlignmentType.CenterVertically:
                    // Move entity so it aligns to vertical centerline (horizontal movement)
                    double currentCenterX = (extents.MinPoint.X + extents.MaxPoint.X) / 2.0;
                    return new Vector3d(alignmentValue - currentCenterX, 0, 0);

                case AlignmentType.CenterHorizontally:
                    // Move entity so it aligns to horizontal centerline (vertical movement)
                    double currentCenterY = (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0;
                    return new Vector3d(0, alignmentValue - currentCenterY, 0);

                default:
                    return new Vector3d(0, 0, 0);
            }
        }

        private static string GetAlignmentName(AlignmentType alignmentType)
        {
            switch (alignmentType)
            {
                case AlignmentType.Top: return "top";
                case AlignmentType.Bottom: return "bottom";
                case AlignmentType.Left: return "left";
                case AlignmentType.Right: return "right";
                case AlignmentType.CenterVertically: return "center vertically";
                case AlignmentType.CenterHorizontally: return "center horizontally";
                default: return "unknown";
            }
        }

        private class EntityExtentData
        {
            public ObjectId ObjectId { get; set; }
            public Extents3d Extents { get; set; }
        }
    }
}
