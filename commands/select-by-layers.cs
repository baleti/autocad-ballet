using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Register the command class
[assembly: CommandClass(typeof(SelectByLayers))]

public class SelectByLayers
{
    // Simple class to store entity references for process scope
    public class EntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string LayerName { get; set; }
        public string SpaceName { get; set; }
    }

    [CommandMethod("select-by-layers", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void SelectByLayersCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        var currentMode = SelectionScopeManager.CurrentScope;

        if (currentMode == SelectionScopeManager.SelectionScope.process)
        {
            HandleProcessMode(ed);
        }
        else
        {
            HandleNormalModes(db, ed, currentMode);
        }
    }

    private void HandleProcessMode(Editor ed)
    {
        ed.WriteMessage("\nProcess Mode: Gathering entities from all open documents...\n");

        // Gather from all open documents
        var docManager = AcadApp.DocumentManager;
        var allReferences = new List<EntityReference>();
        var layerGroups = new Dictionary<string, List<EntityReference>>();

        foreach (Autodesk.AutoCAD.ApplicationServices.Document doc in docManager)
        {
            string docPath = doc.Name;
            string docName = Path.GetFileName(docPath);

            ed.WriteMessage($"\nScanning: {docName}...");

            try
            {
                // Read from document database without switching
                var refs = GatherEntityReferencesFromDocument(doc.Database, docPath, docName);
                allReferences.AddRange(refs);

                // Group by layer name
                foreach (var entityRef in refs)
                {
                    if (!layerGroups.ContainsKey(entityRef.LayerName))
                        layerGroups[entityRef.LayerName] = new List<EntityReference>();

                    layerGroups[entityRef.LayerName].Add(entityRef);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
            }
        }

        if (layerGroups.Count == 0)
        {
            ed.WriteMessage("\nNo entities found across open documents.\n");
            return;
        }

        // Prepare summary for DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var layer in layerGroups.OrderBy(l => l.Key))
        {
            // Count entities per document
            var docCounts = layer.Value.GroupBy(e => e.DocumentName)
                                       .Select(g => $"{g.Key}: {g.Count()}")
                                       .ToList();

            entries.Add(new Dictionary<string, object>
            {
                { "Layer", layer.Key },
                { "Total Count", layer.Value.Count },
                { "Documents", string.Join(", ", docCounts) }
            });
        }

        var propertyNames = new List<string> { "Layer", "Total Count", "Documents" };

        // Show DataGrid for selection
        ed.WriteMessage("\nSelect layers to include in process selection...\n");
        var selectedLayers = CustomGUIs.DataGrid(entries, propertyNames, false);

        if (selectedLayers == null || selectedLayers.Count == 0)
        {
            ed.WriteMessage("\nNo layers selected.\n");
            return;
        }

        // Filter entities to only selected layers
        var selectedLayerNames = selectedLayers.Select(s => s["Layer"].ToString()).ToList();
        var selectedEntities = new List<EntityReference>();
        var layerCounts = new Dictionary<string, int>();

        foreach (var layerName in selectedLayerNames)
        {
            if (layerGroups.ContainsKey(layerName))
            {
                selectedEntities.AddRange(layerGroups[layerName]);
                layerCounts[layerName] = layerGroups[layerName].Count;
            }
        }

        // Convert to unified SelectionItem format and save using SelectionStorage
        var selectionItems = new List<AutoCADBallet.SelectionItem>();
        foreach (var entityRef in selectedEntities)
        {
            selectionItems.Add(new AutoCADBallet.SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null // Will be auto-generated
            });
        }

        try
        {
            // Save using unified selection storage
            AutoCADBallet.SelectionStorage.SaveSelection(selectionItems);

            ed.WriteMessage($"\nProcess selection saved to unified selection storage.\n");
            ed.WriteMessage($"\nSummary:\n");
            ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
            ed.WriteMessage($"  Layers: {selectedLayerNames.Count}\n");
            ed.WriteMessage($"  Documents: {selectedEntities.Select(e => e.DocumentName).Distinct().Count()}\n");

            foreach (var layer in layerCounts)
            {
                ed.WriteMessage($"    {layer.Key}: {layer.Value} entities\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
        }
    }

    private List<EntityReference> GatherEntityReferencesFromDocument(Database db, string docPath, string docName)
    {
        var references = new List<EntityReference>();

        // Important: Use a separate transaction for each external database
        using (var tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    string spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity == null) continue;

                            // Skip certain system entities
                            if (entity is Viewport && spaceName != "Model")
                            {
                                var vp = entity as Viewport;
                                if (vp.Number == 1) // Paper space viewport
                                    continue;
                            }

                            string layerName = GetEntityLayerName(entity);

                            var entityRef = new EntityReference
                            {
                                DocumentPath = docPath,
                                DocumentName = docName,
                                Handle = entity.Handle.ToString(),
                                LayerName = layerName,
                                SpaceName = spaceName
                            };

                            references.Add(entityRef);
                        }
                        catch
                        {
                            // Skip entities that can't be read
                            continue;
                        }
                    }
                }

                tr.Commit();
            }
            catch (System.Exception)
            {
                tr.Abort();
                throw;
            }
        }

        return references;
    }

    private void HandleNormalModes(Database db, Editor ed, SelectionScopeManager.SelectionScope scope)
    {
        // Gather entity layers based on current scope
        var layers = GatherEntityLayers(db, scope);

        if (layers.Count == 0)
        {
            ed.WriteMessage("\nNo entities found in current scope.\n");
            return;
        }

        // Prepare data for DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var layer in layers.OrderBy(l => l.Key))
        {
            entries.Add(new Dictionary<string, object>
            {
                { "Layer", layer.Key },
                { "Count", layer.Value.Count }
            });
        }

        var propertyNames = new List<string> { "Layer", "Count" };

        // Show DataGrid for selection
        ed.WriteMessage("\nSelect layers to include in selection...\n");
        var selectedLayers = CustomGUIs.DataGrid(entries, propertyNames, false);

        if (selectedLayers == null || selectedLayers.Count == 0)
        {
            ed.WriteMessage("\nNo layers selected.\n");
            return;
        }

        // Collect all ObjectIds for selected layers
        var allSelectedIds = new List<ObjectId>();
        foreach (var selected in selectedLayers)
        {
            string layerName = selected["Layer"].ToString();
            if (layers.ContainsKey(layerName))
            {
                allSelectedIds.AddRange(layers[layerName]);
            }
        }

        // Set selection using extension method
        if (allSelectedIds.Count > 0)
        {
            ed.SetImpliedSelectionEx(allSelectedIds.ToArray());
            ed.WriteMessage($"\n{allSelectedIds.Count} objects selected from {selectedLayers.Count} layers.\n");

            // Report details
            foreach (var selected in selectedLayers)
            {
                string layerName = selected["Layer"].ToString();
                int count = layers[layerName].Count;
                ed.WriteMessage($"  {layerName}: {count} objects\n");
            }
        }
        else
        {
            ed.WriteMessage("\nNo objects found for selected layers.\n");
        }
    }

    private Dictionary<string, List<ObjectId>> GatherEntityLayers(Database db, SelectionScopeManager.SelectionScope scope)
    {
        var layers = new Dictionary<string, List<ObjectId>>();

        switch (scope)
        {
            case SelectionScopeManager.SelectionScope.view:
                GatherFromCurrentSpace(db, layers);
                break;

            case SelectionScopeManager.SelectionScope.document:
                GatherFromEntireDrawing(db, layers);
                break;

            case SelectionScopeManager.SelectionScope.desktop:
            case SelectionScopeManager.SelectionScope.network:
                // For now, fall back to Document scope
                GatherFromEntireDrawing(db, layers);
                break;
        }

        return layers;
    }

    private void GatherFromCurrentSpace(Database db, Dictionary<string, List<ObjectId>> layers)
    {
        using (var tr = db.TransactionManager.StartTransaction())
        {
            var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

            foreach (ObjectId id in currentSpace)
            {
                var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (entity == null) continue;

                // Skip entities that are hidden (not visible) in current view
                if (!entity.Visible)
                    continue;

                string layerName = GetEntityLayerName(entity);

                if (!layers.ContainsKey(layerName))
                    layers[layerName] = new List<ObjectId>();

                layers[layerName].Add(id);
            }

            tr.Commit();
        }
    }

    private void GatherFromEntireDrawing(Database db, Dictionary<string, List<ObjectId>> layers)
    {
        using (var tr = db.TransactionManager.StartTransaction())
        {
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                string spaceName = layout.LayoutName;

                foreach (ObjectId id in btr)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity == null) continue;

                        // Skip certain system entities
                        if (entity is Viewport && spaceName != "Model")
                        {
                            var vp = entity as Viewport;
                            if (vp.Number == 1) // Paper space viewport
                                continue;
                        }

                        string layerName = GetEntityLayerName(entity);

                        if (!layers.ContainsKey(layerName))
                            layers[layerName] = new List<ObjectId>();

                        layers[layerName].Add(id);
                    }
                    catch
                    {
                        // Skip entities that can't be read
                        continue;
                    }
                }
            }

            tr.Commit();
        }
    }

    private string GetEntityLayerName(Entity entity)
    {
        try
        {
            // Get the layer name directly from the entity
            return entity.Layer;
        }
        catch
        {
            // Fallback for entities that might not have a layer property
            return "0"; // Default layer
        }
    }
}