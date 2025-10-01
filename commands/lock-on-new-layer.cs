using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.LockOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class LockOnNewLayerCommand
    {
        [CommandMethod("lock-on-new-layer", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void LockOnNewLayer()
        {
            LockLayerUtils.LockOnNewLayer(AcadApp.DocumentManager.MdiActiveDocument);
        }

        public static class LockLayerUtils
    {
        public static void LockOnNewLayer(Document doc)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            // Get pickfirst set or prompt for selection
            var selResult = ed.SelectImplied();
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to lock: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects selected.\n");
                return;
            }

            var selectedObjects = selResult.Value.GetObjectIds();
            var lockLayers = new HashSet<string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Process selected entities - move to lock layers
                foreach (ObjectId objId in selectedObjects)
                {
                    var entity = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    if (entity == null) continue;

                    string originalLayerName = entity.Layer;
                    string newLayerName = originalLayerName + " - lock";

                    // Create lock layer if it doesn't exist
                    if (!layerTable.Has(newLayerName))
                    {
                        layerTable.UpgradeOpen();
                        var originalLayer = (LayerTableRecord)tr.GetObject(layerTable[originalLayerName], OpenMode.ForRead);

                        var newLayer = new LayerTableRecord();
                        newLayer.Name = newLayerName;
                        newLayer.Color = originalLayer.Color;
                        newLayer.LinetypeObjectId = originalLayer.LinetypeObjectId;
                        newLayer.LineWeight = originalLayer.LineWeight;
                        try
                        {
                            newLayer.PlotStyleName = originalLayer.PlotStyleName;
                        }
                        catch
                        {
                            // PlotStyleName may not be available in all drawing modes
                        }
                        newLayer.IsPlottable = originalLayer.IsPlottable;

                        layerTable.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                        layerTable.DowngradeOpen();
                    }

                    // Move entity to new lock layer
                    entity.Layer = newLayerName;
                    lockLayers.Add(newLayerName);
                }

                // Lock all the lock layers
                foreach (string lockLayerName in lockLayers)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerTable[lockLayerName], OpenMode.ForWrite);
                    layer.IsLocked = true;
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nLocked {selectedObjects.Length} entities on new lock layers.\n");
        }

        public static void UnlockOnNewLayer(Document doc)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            var lockLayers = new List<string>();

            // First, get all lock layers and unlock them temporarily
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    string layerName = layer.Name;

                    // Skip xref layers (they contain pipe character)
                    if (layerName.Contains("|")) continue;

                    if (layerName.Length > 7 && layerName.EndsWith(" - lock"))
                    {
                        lockLayers.Add(layerName);
                        layer.UpgradeOpen();
                        layer.IsLocked = false;
                    }
                }

                tr.Commit();
            }

            // Now get selection
            ed.WriteMessage("\nSelect entities to unlock (use Window/Crossing for locked entities): ");
            var selResult = ed.SelectImplied();
            if (selResult.Status == PromptStatus.Error)
            {
                selResult = ed.GetSelection();
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                // Re-lock the layers if no selection made
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    foreach (string lockLayerName in lockLayers)
                    {
                        if (layerTable.Has(lockLayerName))
                        {
                            var layer = (LayerTableRecord)tr.GetObject(layerTable[lockLayerName], OpenMode.ForWrite);
                            layer.IsLocked = true;
                        }
                    }
                    tr.Commit();
                }
                ed.WriteMessage("\nNo entities selected.\n");
                return;
            }

            var selectedObjects = selResult.Value.GetObjectIds();
            int processedCount = 0;
            var layersToCheck = new HashSet<string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Process each entity in selection
                foreach (ObjectId objId in selectedObjects)
                {
                    var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;

                    string layerName = entity.Layer;

                    // Check if this entity is on a lock layer
                    if (layerName.Length > 7 && layerName.EndsWith(" - lock"))
                    {
                        string originalLayerName = layerName.Substring(0, layerName.Length - 7);

                        // Check if original layer exists
                        if (layerTable.Has(originalLayerName))
                        {
                            // Unlock the original layer if it's locked
                            var originalLayer = (LayerTableRecord)tr.GetObject(layerTable[originalLayerName], OpenMode.ForWrite);
                            originalLayer.IsLocked = false;

                            // Move entity back to original layer
                            entity.UpgradeOpen();
                            entity.Layer = originalLayerName;
                            entity.DowngradeOpen();
                            processedCount++;

                            // Track this lock layer for cleanup
                            layersToCheck.Add(layerName);
                        }
                        else
                        {
                            ed.WriteMessage($"\nWarning: Original layer '{originalLayerName}' not found for entity on layer '{layerName}'");
                        }
                    }
                }

                // Re-lock any lock layers that still have entities, or delete empty ones
                foreach (string lockLayerName in lockLayers)
                {
                    // Skip xref layers (they contain pipe character)
                    if (lockLayerName.Contains("|")) continue;

                    if (!layerTable.Has(lockLayerName)) continue;

                    // Check if any entities exist on this layer
                    bool hasEntities = false;
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        foreach (ObjectId entityId in btr)
                        {
                            var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.Layer == lockLayerName)
                            {
                                hasEntities = true;
                                break;
                            }
                        }
                        if (hasEntities) break;
                    }

                    var layer = (LayerTableRecord)tr.GetObject(layerTable[lockLayerName], OpenMode.ForWrite);
                    if (hasEntities)
                    {
                        // Layer has entities, lock it again
                        layer.IsLocked = true;
                    }
                    else
                    {
                        // Layer is empty, delete it
                        try
                        {
                            layer.Erase();
                            ed.WriteMessage($"\nDeleted empty lock layer: {lockLayerName}");
                        }
                        catch
                        {
                            ed.WriteMessage($"\nCould not delete layer: {lockLayerName} (may have references)");
                        }
                    }
                }

                tr.Commit();
            }

            if (processedCount > 0)
                ed.WriteMessage($"\nUnlocked {processedCount} entities and moved to original layers.\n");
            else
                ed.WriteMessage("\nNo entities on lock layers found in selection.\n");
        }
    }
    }
}
