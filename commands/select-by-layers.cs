using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    // Simple class to store entity references for layers
    public class LayerEntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string LayerName { get; set; }
        public string SpaceName { get; set; }
    }

    public static class SelectByLayers
    {
        public static void ExecuteViewScope(Editor ed, Database db)
        {
            ed.WriteMessage("\nView Mode: Gathering entities from current view/layout...\n");

            var layerGroups = new Dictionary<string, List<ObjectId>>();
            var spaceName = GetCurrentSpaceName(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in currentSpace)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity != null && !string.IsNullOrEmpty(entity.Layer))
                        {
                            if (!layerGroups.ContainsKey(entity.Layer))
                                layerGroups[entity.Layer] = new List<ObjectId>();

                            layerGroups[entity.Layer].Add(id);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                tr.Commit();
            }

            ShowSelectionDialogForView(ed, layerGroups, spaceName);
        }

        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            ed.WriteMessage("\nDocument Mode: Gathering entities from entire document...\n");

            var layerGroups = new Dictionary<string, List<LayerEntityReference>>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Get all layouts
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity != null && !string.IsNullOrEmpty(entity.Layer))
                            {
                                if (!layerGroups.ContainsKey(entity.Layer))
                                    layerGroups[entity.Layer] = new List<LayerEntityReference>();

                                layerGroups[entity.Layer].Add(new LayerEntityReference
                                {
                                    DocumentPath = doc.Name,
                                    DocumentName = Path.GetFileName(doc.Name),
                                    Handle = entity.Handle.ToString(),
                                    LayerName = entity.Layer,
                                    SpaceName = spaceName
                                });
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                tr.Commit();
            }

            ShowSelectionDialogForDocument(ed, layerGroups, doc.Name);
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            ed.WriteMessage("\nApplication Mode: Gathering entities from all open documents...\n");

            var docManager = AcadApp.DocumentManager;
            var allReferences = new List<LayerEntityReference>();
            var layerGroups = new Dictionary<string, List<LayerEntityReference>>();

            foreach (Autodesk.AutoCAD.ApplicationServices.Document doc in docManager)
            {
                string docPath = doc.Name;
                string docName = Path.GetFileName(docPath);

                ed.WriteMessage($"\nScanning: {docName}...");

                try
                {
                    var refs = GatherEntityReferencesFromDocument(doc.Database, docPath, docName);
                    allReferences.AddRange(refs);

                    foreach (var entityRef in refs)
                    {
                        if (!layerGroups.ContainsKey(entityRef.LayerName))
                            layerGroups[entityRef.LayerName] = new List<LayerEntityReference>();

                        layerGroups[entityRef.LayerName].Add(entityRef);
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
                }
            }

            ShowSelectionDialogForApplication(ed, layerGroups);
        }

        private static void ShowSelectionDialogForView(Editor ed, Dictionary<string, List<ObjectId>> layerGroups, string spaceName)
        {
            if (layerGroups.Count == 0)
            {
                ed.WriteMessage($"\nNo entities found in {spaceName}.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var layer in layerGroups.OrderBy(l => l.Key))
            {
                entries.Add(new Dictionary<string, object>
                {
                    { "Layer", layer.Key },
                    { "Count", layer.Value.Count },
                    { "Space", spaceName }
                });
            }

            var propertyNames = new List<string> { "Layer", "Count", "Space" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo layers selected.\n");
                return;
            }

            var selectedLayers = new HashSet<string>(chosenRows.Select(row => row["Layer"].ToString()));
            var selectedIds = new List<ObjectId>();

            foreach (var layer in selectedLayers)
            {
                if (layerGroups.ContainsKey(layer))
                {
                    selectedIds.AddRange(layerGroups[layer]);
                }
            }

            if (selectedIds.Count > 0)
            {
                ed.SetImpliedSelection(selectedIds.ToArray());
                ed.WriteMessage($"\nSelected {selectedIds.Count} entities from {selectedLayers.Count} layers in {spaceName}.\n");
            }
        }

        private static void ShowSelectionDialogForDocument(Editor ed, Dictionary<string, List<LayerEntityReference>> layerGroups, string documentName)
        {
            if (layerGroups.Count == 0)
            {
                ed.WriteMessage($"\nNo entities found in document.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var layer in layerGroups.OrderBy(l => l.Key))
            {
                var docCounts = layer.Value.GroupBy(e => e.SpaceName)
                                           .Select(g => $"{g.Key}: {g.Count()}")
                                           .ToList();

                entries.Add(new Dictionary<string, object>
                {
                    { "Layer", layer.Key },
                    { "Total Count", layer.Value.Count },
                    { "Layouts", string.Join(", ", docCounts) }
                });
            }

            var propertyNames = new List<string> { "Layer", "Total Count", "Layouts" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo layers selected.\n");
                return;
            }

            var selectedLayerNames = new HashSet<string>(chosenRows.Select(row => row["Layer"].ToString()));
            var selectedEntities = new List<LayerEntityReference>();

            foreach (var layerName in selectedLayerNames)
            {
                if (layerGroups.ContainsKey(layerName))
                {
                    selectedEntities.AddRange(layerGroups[layerName]);
                }
            }

            var selectionItems = selectedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            try
            {
                var docName = Path.GetFileName(documentName);
                SelectionStorage.SaveSelection(selectionItems, docName);
                ed.WriteMessage($"\nDocument selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
                ed.WriteMessage($"  Layers: {selectedLayerNames.Count}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }
        }

        private static void ShowSelectionDialogForApplication(Editor ed, Dictionary<string, List<LayerEntityReference>> layerGroups)
        {
            if (layerGroups.Count == 0)
            {
                ed.WriteMessage("\nNo entities found across open documents.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var layer in layerGroups.OrderBy(l => l.Key))
            {
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
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo layers selected.\n");
                return;
            }

            var selectedLayerNames = new HashSet<string>(chosenRows.Select(row => row["Layer"].ToString()));
            var selectedEntities = new List<LayerEntityReference>();

            foreach (var layerName in selectedLayerNames)
            {
                if (layerGroups.ContainsKey(layerName))
                {
                    selectedEntities.AddRange(layerGroups[layerName]);
                }
            }

            var selectionItems = selectedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            try
            {
                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage($"\nApplication selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
                ed.WriteMessage($"  Layers: {selectedLayerNames.Count}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }
        }

        private static List<LayerEntityReference> GatherEntityReferencesFromDocument(Database db, string docPath, string docName)
        {
            var references = new List<LayerEntityReference>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity != null && !string.IsNullOrEmpty(entity.Layer))
                            {
                                references.Add(new LayerEntityReference
                                {
                                    DocumentPath = docPath,
                                    DocumentName = docName,
                                    Handle = entity.Handle.ToString(),
                                    LayerName = entity.Layer,
                                    SpaceName = spaceName
                                });
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                tr.Commit();
            }

            return references;
        }

        private static string GetCurrentSpaceName(Database db)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    if (db.TileMode)
                    {
                        return "Model";
                    }
                    else
                    {
                        var layoutMgr = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
                        return layoutMgr.CurrentLayout;
                    }
                }
            }
            catch
            {
                return "Unknown";
            }
        }
    }
}
