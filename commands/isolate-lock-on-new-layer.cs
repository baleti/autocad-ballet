using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.IsolateLockOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class IsolateLockOnNewLayerCommand
    {
        [CommandMethod("isolate-lock-on-new-layer", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void IsolateLockOnNewLayer()
        {
            IsolateLockUtils.IsolateLock(AcadApp.DocumentManager.MdiActiveDocument);
        }

        public static class IsolateLockUtils
    {
        public static void IsolateLock(Document doc)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            // Get pickfirst set or prompt for selection
            var selResult = ed.SelectImplied();
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to isolate: ";
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
            var processedLayers = new HashSet<string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Process selected entities - move to isolate layers
                foreach (ObjectId objId in selectedObjects)
                {
                    var entity = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    if (entity == null) continue;

                    string originalLayerName = entity.Layer;
                    string newLayerName = originalLayerName + " - isolate lock";

                    // Create isolate layer if it doesn't exist
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

                    // Move entity to new isolate layer
                    entity.Layer = newLayerName;
                    processedLayers.Add(newLayerName);
                }

                // Lock all layers except isolate layers
                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    if (!processedLayers.Contains(layer.Name))
                    {
                        layer.IsLocked = true;
                    }
                }

                tr.Commit();
            }

            ed.Regen();
            ed.WriteMessage($"\nIsolated {selectedObjects.Length} entities on new layers. All other layers locked.\n");
        }

        public static void UnisolateLock(Document doc)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            int processedCount = 0;
            var layersToDelete = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Set current layer to 0 to avoid issues when unlocking/unfreezing layers
                if (layerTable.Has("0"))
                {
                    db.Clayer = layerTable["0"];
                }

                // Unlock all layers first
                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    layer.IsLocked = false;
                }

                // Find all isolate lock layers
                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    string layerName = layer.Name;

                    if (layerName.Length > 15 && layerName.EndsWith(" - isolate lock"))
                    {
                        string originalLayerName = layerName.Substring(0, layerName.Length - 15);

                        // Check if original layer exists
                        if (layerTable.Has(originalLayerName))
                        {
                            // Ensure original layer is unlocked and unfrozen to allow entity assignment
                            // Skip if it's the current layer (can't modify frozen state of current layer)
                            var originalLayerId = layerTable[originalLayerName];
                            if (db.Clayer != originalLayerId)
                            {
                                var originalLayer = (LayerTableRecord)tr.GetObject(originalLayerId, OpenMode.ForWrite);
                                originalLayer.IsLocked = false;
                                originalLayer.IsFrozen = false;
                            }

                            // Move all entities back to original layer
                            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            foreach (ObjectId btrId in blockTable)
                            {
                                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                foreach (ObjectId entityId in btr)
                                {
                                    var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                                    if (entity != null && entity.Layer == layerName)
                                    {
                                        entity.UpgradeOpen();
                                        entity.Layer = originalLayerName;
                                        entity.DowngradeOpen();
                                        processedCount++;
                                    }
                                }
                            }

                            layersToDelete.Add(layerId);
                            ed.WriteMessage($"\nDeleted isolate layer: {layerName}");
                        }
                        else
                        {
                            ed.WriteMessage($"\nWarning: Original layer '{originalLayerName}' not found. Keeping isolate layer '{layerName}'.");
                        }
                    }
                }

                // Delete empty isolate layers
                layerTable.UpgradeOpen();
                foreach (var layerId in layersToDelete)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    try
                    {
                        layer.Erase();
                    }
                    catch
                    {
                        ed.WriteMessage($"\nCould not delete layer: {layer.Name} (may have references)");
                    }
                }

                tr.Commit();
            }

            ed.Regen();
            if (processedCount > 0)
                ed.WriteMessage($"\nUnisolated {processedCount} entities and moved to original layers. All layers unlocked.\n");
            else
                ed.WriteMessage("\nNo entities found on isolate layers. All layers unlocked.\n");
        }
    }
    }
}
