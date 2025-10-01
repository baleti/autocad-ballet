using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.HideOnNewLayerCommand))]

namespace AutoCADBallet
{
    public class HideOnNewLayerCommand
    {
        [CommandMethod("hide-on-new-layer", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void HideOnNewLayer()
        {
            HideLayerUtils.HideOnNewLayer(AcadApp.DocumentManager.MdiActiveDocument);
        }

        public static class HideLayerUtils
    {
        public static void HideOnNewLayer(Document doc)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            // Get pickfirst set or prompt for selection
            var selResult = ed.SelectImplied();
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to hide: ";
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
            var hideLayers = new HashSet<string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Process selected entities - move to hide layers
                foreach (ObjectId objId in selectedObjects)
                {
                    var entity = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    if (entity == null) continue;

                    string originalLayerName = entity.Layer;
                    string newLayerName = originalLayerName + " - hide";

                    // Create hide layer if it doesn't exist
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

                    // Move entity to new hide layer
                    entity.Layer = newLayerName;
                    hideLayers.Add(newLayerName);
                }

                // Freeze all the hide layers
                foreach (string hideLayerName in hideLayers)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerTable[hideLayerName], OpenMode.ForWrite);
                    layer.IsFrozen = true;
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nHidden {selectedObjects.Length} entities on new frozen layers.\n");
        }

        public static void UnhideOnNewLayer(Document doc)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            int processedCount = 0;
            var layersToDelete = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // Find all hide layers
                foreach (ObjectId layerId in layerTable)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    string layerName = layer.Name;

                    if (layerName.Length > 7 && layerName.EndsWith(" - hide"))
                    {
                        // Thaw the layer to access entities
                        layer.UpgradeOpen();
                        layer.IsFrozen = false;
                        layer.DowngradeOpen();

                        string originalLayerName = layerName.Substring(0, layerName.Length - 7);

                        // Check if original layer exists
                        if (layerTable.Has(originalLayerName))
                        {
                            // Unfreeze the original layer if it's frozen
                            var originalLayer = (LayerTableRecord)tr.GetObject(layerTable[originalLayerName], OpenMode.ForWrite);
                            originalLayer.IsFrozen = false;

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
                            ed.WriteMessage($"\nDeleted hide layer: {layerName}");
                        }
                        else
                        {
                            // Original layer doesn't exist, keep hide layer but freeze it
                            ed.WriteMessage($"\nWarning: Original layer '{originalLayerName}' not found. Keeping hide layer '{layerName}' frozen.");
                            layer.UpgradeOpen();
                            layer.IsFrozen = true;
                        }
                    }
                }

                // Delete empty hide layers
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

            if (processedCount > 0)
                ed.WriteMessage($"\nUnhidden {processedCount} entities and moved to original layers.\n");
            else
                ed.WriteMessage("\nNo entities found on hide layers.\n");
        }
    }
    }
}
