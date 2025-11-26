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

[assembly: CommandClass(typeof(FilterSelectionDocumentElements))]

/// <summary>
/// Command that always uses document scope for filtering selection, regardless of current selection scope setting
/// </summary>
public class FilterSelectionDocumentElements
{
    [CommandMethod("filter-selection-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void FilterSelectionDocumentCommand()
    {
        var command = new FilterSelectionDocumentImpl();
        command.Execute();
    }
}

public class FilterSelectionDocumentImpl : FilterElementsBase
{
    public override bool SpanAllScreens => false;
    public override bool UseSelectedElements => true; // Use stored selection from document scope
    public override bool IncludeProperties => true;
    public override SelectionScope Scope => SelectionScope.document;

    public override void Execute()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            // Load stored selection for document
            var docName = Path.GetFileName(doc.Name);
            var storedSelection = SelectionStorage.LoadSelection(docName);

            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found for current document. Use 'select-by-categories-in-document' first.\n");
                return;
            }

            // Collect entity data with optimized single transaction
            var entityData = new List<Dictionary<string, object>>();

            // Single transaction for all entities in the current document
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Build block hierarchy cache ONCE at the start (huge performance optimization)
                var blockHierarchyCache = FilterEntityDataHelper.BuildBlockHierarchyCache(db, tr);

                foreach (var item in storedSelection)
                {
                    try
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (objectId != ObjectId.Null)
                        {
                            var dbObject = tr.GetObject(objectId, OpenMode.ForRead);
                            if (dbObject != null)
                            {
                                var data = FilterEntityDataHelper.GetEntityDataDictionary(dbObject, item.DocumentPath, null, IncludeProperties, tr, blockHierarchyCache);
                                data["ObjectId"] = objectId;
                                entityData.Add(data);
                            }
                        }
                    }
                    catch (System.Exception)
                    {
                        // Skip problematic entities
                        continue;
                    }
                }
                tr.Commit();
            }

            if (entityData.Count == 0)
            {
                ed.WriteMessage("\nNo entities found in stored selection.\n");
                return;
            }

            // Process results with DataGrid
            ProcessSelectionResults(entityData);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in filter-selection-in-document: {ex.Message}\n");
        }
    }

    private void ProcessSelectionResults(List<Dictionary<string, object>> entityData)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        // Process parent block hierarchy into separate columns
        FilterEntityDataHelper.ProcessParentBlockColumns(entityData);

        // Add index to each entity for mapping back after user selection
        for (int i = 0; i < entityData.Count; i++)
        {
            entityData[i]["OriginalIndex"] = i;
        }

        // Get property names, excluding internal fields
        var propertyNames = entityData.First().Keys
            .Where(k => !k.EndsWith("ObjectId") && k != "OriginalIndex" && k != "_ParentBlocks")
            .ToList();

        // Reorder to put most useful columns first
        var orderedProps = new List<string> { "Name", "Contents", "Category", "Layer", "Layout", "DocumentName", "Color", "LineType", "Handle" };
        var remainingProps = propertyNames.Except(orderedProps);

        // Separate geometry properties, attributes and extension data
        var geometryProps = remainingProps.Where(p => FilterEntityDataHelper.IsGeometryProperty(p)).OrderBy(p => FilterEntityDataHelper.GetGeometryPropertyOrder(p));
        var attributeProps = remainingProps.Where(p => p.StartsWith("attr_")).OrderBy(p => p);
        var extensionProps = remainingProps.Where(p => p.StartsWith("xdata_") || p.StartsWith("ext_dict_")).OrderBy(p => p);
        var otherProps = remainingProps.Where(p => !p.StartsWith("attr_") && !p.StartsWith("xdata_") && !p.StartsWith("ext_dict_") && p != "DocumentPath" && p != "DisplayName" && !FilterEntityDataHelper.IsGeometryProperty(p)).OrderBy(p => p);
        var documentPathProp = propertyNames.Contains("DocumentPath") ? new[] { "DocumentPath" } : new string[0];

        propertyNames = orderedProps.Where(p => propertyNames.Contains(p))
            .Concat(geometryProps)
            .Concat(attributeProps)
            .Concat(extensionProps)
            .Concat(otherProps)
            .Concat(documentPathProp)
            .ToList();

        // Show DataGrid
        var chosenRows = CustomGUIs.DataGrid(entityData, propertyNames, SpanAllScreens, null);

        if (chosenRows.Count == 0)
        {
            ed.WriteMessage("\nNo entities selected.\n");
            return;
        }

        // Collect ObjectIds for selection
        var selectedIds = new List<ObjectId>();
        foreach (var row in chosenRows)
        {
            if (row.TryGetValue("ObjectId", out var objIdObj) && objIdObj is ObjectId objectId)
            {
                if (!objectId.IsNull && objectId.IsValid)
                {
                    selectedIds.Add(objectId);
                }
            }
        }

        // Set selection
        if (selectedIds.Count > 0)
        {
            try
            {
                ed.SetImpliedSelection(selectedIds.ToArray());
                ed.WriteMessage($"\n{selectedIds.Count} entities selected in current document.\n");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception acEx)
            {
                ed.WriteMessage($"\nError setting selection: {acEx.ErrorStatus} - {acEx.Message}\n");
            }
        }

        // Save the filtered selection back to document storage
        if (selectedIds.Count > 0)
        {
            var selectionItems = new List<SelectionItem>();
            foreach (var id in selectedIds)
            {
                selectionItems.Add(new SelectionItem
                {
                    DocumentPath = doc.Name,
                    Handle = id.Handle.ToString(),
                    SessionId = null
                });
            }

            var docName = Path.GetFileName(doc.Name);
            SelectionStorage.SaveSelection(selectionItems, docName);
            ed.WriteMessage("Filtered selection saved to document storage.\n");
        }
        else
        {
            // Clear selection if no entities were selected
            var docName = Path.GetFileName(doc.Name);
            SelectionStorage.SaveSelection(new List<SelectionItem>(), docName);
            ed.WriteMessage("Document selection cleared (no items selected in filter).\n");
        }
    }
}
