using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

public class SelectSimilarByLayer
{
    public static void ExecuteViewScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        try
        {
            // Handle view scope: Use pickfirst set or prompt for selection
            var selResult = ed.SelectImplied();

            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect entities to get their layers: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo entities selected.\n");
                return;
            }

            // Collect layers from currently selected entities
            var layersToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Get layers from selected entities
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var objectId in selResult.Value.GetObjectIds())
                {
                    try
                    {
                        var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                        if (entity != null && !string.IsNullOrEmpty(entity.Layer))
                        {
                            layersToSelect.Add(entity.Layer);
                        }
                    }
                    catch
                    {
                        // Skip problematic entities
                        continue;
                    }
                }
                tr.Commit();
            }

            if (layersToSelect.Count == 0)
            {
                ed.WriteMessage("\nNo valid layers found in selected entities.\n");
                return;
            }

            ed.WriteMessage($"\nFound {layersToSelect.Count} unique layer(s): {string.Join(", ", layersToSelect)}\n");

            // View scope: Select from current space only
            var currentDocumentIds = new List<ObjectId>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                foreach (ObjectId id in currentSpace)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity != null && layersToSelect.Contains(entity.Layer))
                        {
                            currentDocumentIds.Add(id);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                tr.Commit();
            }

            // Set selection in AutoCAD
            if (currentDocumentIds.Count > 0)
            {
                ed.SetImpliedSelection(currentDocumentIds.ToArray());
                ed.WriteMessage($"\nSelected {currentDocumentIds.Count} entities on layer(s): {string.Join(", ", layersToSelect)}\n");
            }
            else
            {
                ed.WriteMessage($"\nNo entities found on layer(s): {string.Join(", ", layersToSelect)}\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
    }

    public static void ExecuteDocumentScope(Editor ed, Database db)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;

        try
        {
            var docName = Path.GetFileName(doc.Name);
            var storedSelection = SelectionStorage.LoadSelection(docName);
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage($"\nNo stored selection found for current document '{docName}'. Use commands like 'select-by-categories-in-document' first.\n");
                return;
            }

            // Collect layers from stored selection
            var layersToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Get layers from stored selection items in current document
            foreach (var item in storedSelection)
            {
                try
                {
                    var handle = Convert.ToInt64(item.Handle, 16);
                    var objectId = db.GetObjectId(false, new Handle(handle), 0);

                    if (objectId != ObjectId.Null)
                    {
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && !string.IsNullOrEmpty(entity.Layer))
                            {
                                layersToSelect.Add(entity.Layer);
                            }
                            tr.Commit();
                        }
                    }
                }
                catch
                {
                    // Skip problematic entities
                    continue;
                }
            }

            if (layersToSelect.Count == 0)
            {
                ed.WriteMessage("\nNo valid layers found in stored selection.\n");
                return;
            }

            ed.WriteMessage($"\nFound {layersToSelect.Count} unique layer(s): {string.Join(", ", layersToSelect)}\n");

            // Document scope: Select from all layouts in current document
            var selectedEntities = new List<SelectionItem>();
            var currentDocumentIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity != null && layersToSelect.Contains(entity.Layer))
                            {
                                currentDocumentIds.Add(id);
                                selectedEntities.Add(new SelectionItem
                                {
                                    DocumentPath = doc.Name,
                                    Handle = id.Handle.ToString()
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

            // Set selection
            if (currentDocumentIds.Count > 0)
            {
                ed.SetImpliedSelection(currentDocumentIds.ToArray());
                ed.WriteMessage($"\nSelected {currentDocumentIds.Count} entities on layer(s): {string.Join(", ", layersToSelect)} from current document\n");
            }

            // Save to selection storage
            if (selectedEntities.Count > 0)
            {
                SelectionStorage.SaveSelection(selectedEntities, docName);
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
    }

    public static void ExecuteApplicationScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        try
        {
            var storedSelection = SelectionStorage.LoadSelectionFromAllDocuments();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-categories-in-session' first.\n");
                return;
            }

            // Collect layers from stored selection
            var layersToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Get layers from stored selection items
            foreach (var item in storedSelection)
            {
                try
                {
                    // Check if this is from the current document
                    if (string.Equals(Path.GetFullPath(item.DocumentPath), Path.GetFullPath(doc.Name), StringComparison.OrdinalIgnoreCase))
                    {
                        // Current document - get entity directly
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (objectId != ObjectId.Null)
                        {
                            using (var tr = db.TransactionManager.StartTransaction())
                            {
                                var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                                if (entity != null && !string.IsNullOrEmpty(entity.Layer))
                                {
                                    layersToSelect.Add(entity.Layer);
                                }
                                tr.Commit();
                            }
                        }
                    }
                    else
                    {
                        // Different document - try to get layer from external document
                        var layerName = GetLayerFromExternalDocument(item.DocumentPath, item.Handle);
                        if (!string.IsNullOrEmpty(layerName))
                        {
                            layersToSelect.Add(layerName);
                        }
                    }
                }
                catch
                {
                    // Skip problematic entities
                    continue;
                }
            }

            if (layersToSelect.Count == 0)
            {
                ed.WriteMessage("\nNo valid layers found in stored selection.\n");
                return;
            }

            ed.WriteMessage($"\nFound {layersToSelect.Count} unique layer(s): {string.Join(", ", layersToSelect)}\n");

            // Application scope: Select from all documents in scope
            SelectEntitiesFromAllDocuments(layersToSelect, ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
    }

    private static string GetLayerFromExternalDocument(string documentPath, string handle)
    {
        try
        {
            var docs = AcadApp.DocumentManager;
            Document externalDoc = null;
            bool docWasAlreadyOpen = false;

            // Check if the document is already open
            foreach (Document openDoc in docs)
            {
                if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(documentPath), StringComparison.OrdinalIgnoreCase))
                {
                    externalDoc = openDoc;
                    docWasAlreadyOpen = true;
                    break;
                }
            }

            // If not already open, try to open it temporarily
            if (externalDoc == null && File.Exists(documentPath))
            {
                try
                {
                    externalDoc = docs.Open(documentPath, false); // Open read-only
                    docWasAlreadyOpen = false;
                }
                catch
                {
                    return null;
                }
            }

            // Get the layer name from the entity
            if (externalDoc != null)
            {
                try
                {
                    var handleValue = Convert.ToInt64(handle, 16);
                    var objectId = externalDoc.Database.GetObjectId(false, new Handle(handleValue), 0);

                    if (objectId != ObjectId.Null)
                    {
                        using (var tr = externalDoc.Database.TransactionManager.StartTransaction())
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null)
                            {
                                tr.Commit();
                                return entity.Layer;
                            }
                            tr.Commit();
                        }
                    }
                }
                finally
                {
                    // Close the document if we opened it temporarily
                    if (!docWasAlreadyOpen && externalDoc != null)
                    {
                        try
                        {
                            externalDoc.CloseAndDiscard();
                        }
                        catch
                        {
                            // Ignore close errors
                        }
                    }
                }
            }
        }
        catch
        {
            // If anything goes wrong, return null
        }

        return null;
    }

    private static void SelectEntitiesFromAllDocuments(HashSet<string> layersToSelect, Editor ed)
    {
        var docs = AcadApp.DocumentManager;
        var allSelectedEntities = new List<SelectionItem>();
        var currentDocumentIds = new List<ObjectId>();
        var currentDoc = docs.MdiActiveDocument;
        int totalCount = 0;

        // Process current document first
        using (var tr = currentDoc.Database.TransactionManager.StartTransaction())
        {
            var layoutDict = (DBDictionary)tr.GetObject(currentDoc.Database.LayoutDictionaryId, OpenMode.ForRead);
            foreach (DBDictionaryEntry entry in layoutDict)
            {
                var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                foreach (ObjectId id in btr)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity != null && layersToSelect.Contains(entity.Layer))
                        {
                            currentDocumentIds.Add(id);
                            allSelectedEntities.Add(new SelectionItem
                            {
                                DocumentPath = currentDoc.Name,
                                Handle = id.Handle.ToString()
                            });
                            totalCount++;
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

        // Process other open documents
        foreach (Document doc in docs)
        {
            if (doc == currentDoc) continue; // Already processed

            try
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var layoutDict = (DBDictionary)tr.GetObject(doc.Database.LayoutDictionaryId, OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                        foreach (ObjectId id in btr)
                        {
                            try
                            {
                                var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                if (entity != null && layersToSelect.Contains(entity.Layer))
                                {
                                    allSelectedEntities.Add(new SelectionItem
                                    {
                                        DocumentPath = doc.Name,
                                        Handle = id.Handle.ToString()
                                    });
                                    totalCount++;
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
            }
            catch
            {
                ed.WriteMessage($"\nWarning: Could not process document {doc.Name}\n");
                continue;
            }
        }

        // Set selection for current document entities
        if (currentDocumentIds.Count > 0)
        {
            ed.SetImpliedSelectionEx(currentDocumentIds.ToArray());
        }

        // Save all selected entities to selection storage
        if (allSelectedEntities.Count > 0)
        {
            SelectionStorage.SaveSelection(allSelectedEntities);
        }

        // Report results
        if (totalCount > 0)
        {
            var documentCount = allSelectedEntities.GroupBy(e => e.DocumentPath).Count();
            ed.WriteMessage($"\nSelected {totalCount} entities on layer(s): {string.Join(", ", layersToSelect)} from {documentCount} document(s)\n");

            if (currentDocumentIds.Count > 0)
            {
                ed.WriteMessage($"Current document: {currentDocumentIds.Count} entities selected\n");
            }

            var externalCount = totalCount - currentDocumentIds.Count;
            if (externalCount > 0)
            {
                ed.WriteMessage($"Other documents: {externalCount} entities stored in selection\n");
            }
        }
        else
        {
            ed.WriteMessage($"\nNo entities found on layer(s): {string.Join(", ", layersToSelect)}\n");
        }
    }
}