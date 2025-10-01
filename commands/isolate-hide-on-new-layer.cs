using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.IsolateHideOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class IsolateHideOnNewLayerCommand
    {
        [CommandMethod("isolate-hide-on-new-layer", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void IsolateHideOnNewLayer()
        {
            IsolateHideUtils.IsolateHide(AcadApp.DocumentManager.MdiActiveDocument);
        }

        public static class IsolateHideUtils
    {
        public static void IsolateHide(Document doc)
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
            string firstIsolateLayer = null;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Process selected entities - move to isolate layers
                foreach (ObjectId objId in selectedObjects)
                {
                    var entity = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    if (entity == null) continue;

                    string originalLayerName = entity.Layer;
                    string newLayerName = originalLayerName + " - isolate hide";

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

                    // Track processed layers
                    if (!processedLayers.Contains(newLayerName))
                    {
                        processedLayers.Add(newLayerName);
                        if (firstIsolateLayer == null)
                            firstIsolateLayer = newLayerName;
                    }
                }

                // Set one of the isolate layers as current so we can freeze others
                if (firstIsolateLayer != null)
                {
                    db.Clayer = layerTable[firstIsolateLayer];
                }

                // Freeze all layers except isolate layers
                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    if (!processedLayers.Contains(layer.Name))
                    {
                        layer.IsFrozen = true;
                    }
                }

                tr.Commit();
            }

            ed.Regen();
            ed.WriteMessage($"\nIsolated {selectedObjects.Length} entities on new layers. All other layers frozen.\n");
        }

        public static void UnisolateHide(Document doc)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            int processedCount = 0;
            var layersToDelete = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Unfreeze and set current layer to 0 to avoid issues when unfreezing other layers
                if (layerTable.Has("0"))
                {
                    var layer0 = (LayerTableRecord)tr.GetObject(layerTable["0"], OpenMode.ForWrite);
                    layer0.IsFrozen = false;
                    layer0.IsLocked = false;
                    db.Clayer = layerTable["0"];
                }

                // Find all isolate hide layers
                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    string layerName = layer.Name;

                    if (layerName.Length > 15 && layerName.EndsWith(" - isolate hide"))
                    {
                        string originalLayerName = layerName.Substring(0, layerName.Length - 15);

                        // Check if original layer exists
                        if (layerTable.Has(originalLayerName))
                        {
                            // Thaw and unlock the original layer to ensure entities can be moved to it
                            // Skip if it's the current layer (can't modify frozen state of current layer)
                            var originalLayerId = layerTable[originalLayerName];
                            if (db.Clayer != originalLayerId)
                            {
                                var originalLayer = (LayerTableRecord)tr.GetObject(originalLayerId, OpenMode.ForWrite);
                                originalLayer.IsFrozen = false;
                                originalLayer.IsLocked = false;
                            }

                            // Thaw the isolate layer to access entities
                            layer.UpgradeOpen();
                            layer.IsFrozen = false;
                            layer.DowngradeOpen();

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
                            ed.WriteMessage($"\nWarning: Original layer '{originalLayerName}' not found. Keeping isolate layer '{layerName}' thawed.");
                            layer.UpgradeOpen();
                            layer.IsFrozen = false;
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

                // Thaw all layers (skip erased layers and current layer)
                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    // Skip if layer is erased or is the current layer
                    if (!layer.IsErased && db.Clayer != layerId)
                    {
                        layer.IsFrozen = false;
                    }
                }

                tr.Commit();
            }

            ed.Regen();
            if (processedCount > 0)
                ed.WriteMessage($"\nUnisolated {processedCount} entities and moved to original layers. All layers thawed.\n");
            else
                ed.WriteMessage("\nNo entities found on isolate layers. All layers thawed.\n");
        }
    }
    }
}
