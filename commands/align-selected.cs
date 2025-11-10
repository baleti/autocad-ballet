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
                // Build group membership map
                var entityToGroup = BuildEntityToGroupMap(db, tr, selectedObjects);

                // Group entities by their group membership or individual entities
                var alignmentUnits = GroupEntitiesForAlignment(selectedObjects, entityToGroup);

                // Collect extents for each alignment unit
                var unitData = new List<AlignmentUnitData>();

                foreach (var unit in alignmentUnits)
                {
                    Extents3d? combinedExtents = null;

                    foreach (ObjectId objId in unit)
                    {
                        var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        if (entity == null) continue;

                        try
                        {
                            var extents = entity.GeometricExtents;
                            if (combinedExtents == null)
                            {
                                combinedExtents = extents;
                            }
                            else
                            {
                                // Extents3d is a struct, so we need to get the value, modify it, and store it back
                                var temp = combinedExtents.Value;
                                temp.AddExtents(extents);
                                combinedExtents = temp;
                            }
                        }
                        catch
                        {
                            // Skip entities without valid extents
                            continue;
                        }
                    }

                    if (combinedExtents != null)
                    {
                        unitData.Add(new AlignmentUnitData
                        {
                            EntityIds = unit,
                            Extents = combinedExtents.Value
                        });
                    }
                }

                if (unitData.Count < 2)
                {
                    ed.WriteMessage("\nNeed at least 2 objects/groups with valid extents to align.\n");
                    tr.Abort();
                    return;
                }

                // Calculate alignment reference
                double alignmentValue = CalculateAlignmentReference(unitData, alignmentType);

                // Align units (entities or groups)
                int alignedCount = 0;
                foreach (var data in unitData)
                {
                    Vector3d displacement = CalculateDisplacement(data.Extents, alignmentValue, alignmentType);

                    if (!displacement.IsZeroLength())
                    {
                        Matrix3d transform = Matrix3d.Displacement(displacement);

                        // Apply transformation to all entities in the unit
                        foreach (ObjectId objId in data.EntityIds)
                        {
                            var entity = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                            if (entity != null)
                            {
                                entity.TransformBy(transform);
                            }
                        }
                        alignedCount++;
                    }
                }

                tr.Commit();

                string alignmentName = GetAlignmentName(alignmentType);
                ed.WriteMessage($"\nAligned {alignedCount} objects/groups ({alignmentName}).\n");
            }
        }

        private static Dictionary<ObjectId, ObjectId> BuildEntityToGroupMap(Database db, Transaction tr, ObjectId[] selectedObjects)
        {
            var entityToGroup = new Dictionary<ObjectId, ObjectId>();

            try
            {
                // Access the group dictionary
                var groupDict = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in groupDict)
                {
                    var group = tr.GetObject(entry.Value, OpenMode.ForRead) as Group;
                    if (group == null || !group.Selectable) continue;

                    // Get all entities in this group
                    var groupEntities = group.GetAllEntityIds();

                    // Map each selected entity to its group
                    foreach (ObjectId entityId in groupEntities)
                    {
                        if (selectedObjects.Contains(entityId))
                        {
                            entityToGroup[entityId] = entry.Value;
                        }
                    }
                }
            }
            catch
            {
                // If there's any error accessing groups, just return empty map
            }

            return entityToGroup;
        }

        private static List<List<ObjectId>> GroupEntitiesForAlignment(ObjectId[] selectedObjects, Dictionary<ObjectId, ObjectId> entityToGroup)
        {
            var alignmentUnits = new List<List<ObjectId>>();
            var processedEntities = new HashSet<ObjectId>();
            var processedGroups = new HashSet<ObjectId>();

            foreach (ObjectId objId in selectedObjects)
            {
                if (processedEntities.Contains(objId))
                    continue;

                // Check if this entity belongs to a group
                if (entityToGroup.TryGetValue(objId, out ObjectId groupId))
                {
                    // Skip if we've already processed this group
                    if (processedGroups.Contains(groupId))
                        continue;

                    // Collect all selected entities that belong to this group
                    var groupMembers = new List<ObjectId>();
                    foreach (ObjectId selectedObjId in selectedObjects)
                    {
                        if (entityToGroup.TryGetValue(selectedObjId, out ObjectId memberGroupId) && memberGroupId == groupId)
                        {
                            groupMembers.Add(selectedObjId);
                            processedEntities.Add(selectedObjId);
                        }
                    }

                    if (groupMembers.Count > 0)
                    {
                        alignmentUnits.Add(groupMembers);
                        processedGroups.Add(groupId);
                    }
                }
                else
                {
                    // Entity is not part of a group, treat it individually
                    alignmentUnits.Add(new List<ObjectId> { objId });
                    processedEntities.Add(objId);
                }
            }

            return alignmentUnits;
        }

        private static double CalculateAlignmentReference(List<AlignmentUnitData> unitData, AlignmentType alignmentType)
        {
            switch (alignmentType)
            {
                case AlignmentType.Top:
                    // Align to the highest top edge
                    return unitData.Max(e => e.Extents.MaxPoint.Y);

                case AlignmentType.Bottom:
                    // Align to the lowest bottom edge
                    return unitData.Min(e => e.Extents.MinPoint.Y);

                case AlignmentType.Left:
                    // Align to the leftmost left edge
                    return unitData.Min(e => e.Extents.MinPoint.X);

                case AlignmentType.Right:
                    // Align to the rightmost right edge
                    return unitData.Max(e => e.Extents.MaxPoint.X);

                case AlignmentType.CenterVertically:
                    // Align to the overall vertical center (center on vertical axis)
                    double minX = unitData.Min(e => e.Extents.MinPoint.X);
                    double maxX = unitData.Max(e => e.Extents.MaxPoint.X);
                    return (minX + maxX) / 2.0;

                case AlignmentType.CenterHorizontally:
                    // Align to the overall horizontal center (center on horizontal axis)
                    double minY = unitData.Min(e => e.Extents.MinPoint.Y);
                    double maxY = unitData.Max(e => e.Extents.MaxPoint.Y);
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

        private class AlignmentUnitData
        {
            public List<ObjectId> EntityIds { get; set; }
            public Extents3d Extents { get; set; }
        }
    }
}
