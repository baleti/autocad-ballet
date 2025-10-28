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

[assembly: CommandClass(typeof(FilterSelectionSessionElements))]

/// <summary>
/// Command that always uses session scope for filtering selection, regardless of current selection scope setting
/// </summary>
public class FilterSelectionSessionElements
{
    [CommandMethod("filter-selection-in-session", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void FilterSelectionSessionCommand()
    {
        var command = new FilterSelectionSessionImpl();
        command.Execute();
    }
}

public class FilterSelectionSessionImpl : FilterElementsBase
{
    public override bool SpanAllScreens => false;
    public override bool UseSelectedElements => true; // Use stored selection from session scope (all documents)
    public override bool IncludeProperties => true;
    public override SelectionScope Scope => SelectionScope.application;

    public override void Execute()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ObjectId[] originalSelection = null;

        try
        {
            // Force session scope behavior - load selection from all open documents
            var storedSelection = SelectionStorage.LoadSelectionFromAllDocuments();

            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-categories-in-session' first.\n");
                return;
            }

            // Create entity data from the stored selection using the helper method
            var entityData = new List<Dictionary<string, object>>();

            // Process stored selection items
            foreach (var item in storedSelection)
            {
                try
                {
                    // Find the document this entity belongs to (should already be open)
                    Document itemDoc = null;
                    foreach (Document openDoc in AcadApp.DocumentManager)
                    {
                        try
                        {
                            if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(item.DocumentPath), StringComparison.OrdinalIgnoreCase))
                            {
                                itemDoc = openDoc;
                                break;
                            }
                        }
                        catch
                        {
                            // Path comparison failed, try direct comparison
                            if (string.Equals(openDoc.Name, item.DocumentPath, StringComparison.OrdinalIgnoreCase))
                            {
                                itemDoc = openDoc;
                                break;
                            }
                        }
                    }

                    if (itemDoc != null)
                    {
                        // Document is open - get entity directly (no need to mark as external)
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = itemDoc.Database.GetObjectId(false, new Autodesk.AutoCAD.DatabaseServices.Handle(handle), 0);

                        if (objectId != Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                        {
                            using (var tr = itemDoc.Database.TransactionManager.StartTransaction())
                            {
                                var entity = tr.GetObject(objectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                if (entity != null)
                                {
                                    var data = FilterEntityDataHelper.GetEntityDataDictionary(entity, item.DocumentPath, null, IncludeProperties);

                                    // Only mark as external if it's not the current document
                                    if (itemDoc != doc)
                                    {
                                        data["IsExternal"] = true;
                                        data["DisplayName"] = $"External: {data["Name"]}";
                                    }
                                    else
                                    {
                                        data["ObjectId"] = objectId; // Store for selection in current doc
                                    }

                                    entityData.Add(data);
                                }
                                tr.Commit();
                            }
                        }
                    }
                    else
                    {
                        // Document is not open - skip it (session scope only works with open documents)
                        ed.WriteMessage($"\nWarning: Document '{Path.GetFileName(item.DocumentPath)}' is not open, skipping entity.\n");
                    }
                }
                catch
                {
                    // Skip problematic entities
                    continue;
                }
            }

            if (!entityData.Any())
            {
                ed.WriteMessage("\nNo valid entities found in stored selection.\n");
                return;
            }

            // Process the entity data using the standard filtering logic from base class
            ProcessSelectionResults(entityData, originalSelection, false);
        }
        catch (InvalidOperationException ex)
        {
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nUnexpected error: {ex.Message}\n");
        }
    }

    private void ProcessSelectionResults(List<Dictionary<string, object>> entityData, ObjectId[] originalSelection, bool editsWereApplied)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        // Add index to each entity for mapping back after user selection
        for (int i = 0; i < entityData.Count; i++)
        {
            entityData[i]["OriginalIndex"] = i;
        }

        // Get property names, excluding internal fields
        var propertyNames = entityData.First().Keys
            .Where(k => !k.EndsWith("ObjectId") && k != "OriginalIndex")
            .ToList();

        // Reorder to put most useful columns first
        var orderedProps = new List<string> { "Name", "Contents", "Category", "Layer", "Layout", "DocumentName", "Color", "LineType", "Handle" };
        var remainingProps = propertyNames.Except(orderedProps);

        // Separate geometry properties, attributes and extension data for better organization
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

        // Reset the edits flag at the start of each DataGrid session
        CustomGUIs.ResetEditsAppliedFlag();

        var chosenRows = CustomGUIs.DataGrid(entityData, propertyNames, SpanAllScreens, null);

        // Check if any edits were applied during DataGrid interaction
        editsWereApplied = CustomGUIs.HasPendingEdits() || CustomGUIs.WereEditsApplied();

        if (chosenRows.Count == 0)
        {
            ed.WriteMessage("\nNo entities selected.\n");
            if (editsWereApplied)
            {
                ed.WriteMessage("Entity modifications were applied. Selection not changed.\n");
            }
            return;
        }

        // Collect ObjectIds for selection and selected external entities
        var selectedIds = new List<Autodesk.AutoCAD.DatabaseServices.ObjectId>();
        var selectedExternalEntities = new List<Dictionary<string, object>>();

        foreach (var row in chosenRows)
        {
            // Check if this is an external entity
            if (row.TryGetValue("IsExternal", out var isExternal) && (bool)isExternal)
            {
                // Only add external entities that were actually selected by the user
                selectedExternalEntities.Add(row);
            }
            else if (row.TryGetValue("ObjectId", out var objIdObj) && objIdObj is Autodesk.AutoCAD.DatabaseServices.ObjectId objectId)
            {
                // Validate ObjectId before adding to avoid eInvalidInput errors
                if (!objectId.IsNull && objectId.IsValid)
                {
                    try
                    {
                        // Quick validation by checking if object exists in current database
                        using (var tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            var testObj = tr.GetObject(objectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, false);
                            if (testObj != null)
                            {
                                selectedIds.Add(objectId);
                            }
                            tr.Commit();
                        }
                    }
                    catch
                    {
                        // Skip invalid ObjectIds to prevent eInvalidInput errors
                        ed.WriteMessage($"\nSkipping invalid ObjectId: {objectId}\n");
                    }
                }
            }
        }

        // Set selection for current document entities (only if no edits were applied)
        if (selectedIds.Count > 0 && !editsWereApplied)
        {
            try
            {
                ed.SetImpliedSelection(selectedIds.ToArray());
                ed.WriteMessage($"\n{selectedIds.Count} entities selected in current document from session-wide stored selection.\n");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception acEx)
            {
                ed.WriteMessage($"\nError setting selection: {acEx.ErrorStatus} - {acEx.Message}\n");
            }
        }
        else if (selectedIds.Count > 0 && editsWereApplied)
        {
            ed.WriteMessage($"\nEntity modifications were applied. Selection not changed (would have selected {selectedIds.Count} entities).\n");
        }

        // Report selected external entities
        if (selectedExternalEntities.Count > 0)
        {
            ed.WriteMessage($"\nNote: {selectedExternalEntities.Count} external entities were selected but cannot be highlighted in current document:\n");
            foreach (var ext in selectedExternalEntities.Take(5)) // Show first 5
            {
                ed.WriteMessage($"  {ext["DocumentName"]} - Handle: {ext["Handle"]}\n");
            }
            if (selectedExternalEntities.Count > 5)
            {
                ed.WriteMessage($"  ... and {selectedExternalEntities.Count - 5} more\n");
            }
        }

        // Clear all existing selections and save only the filtered results (session scope behavior)
        ClearAllStoredSelections();

        if ((selectedIds.Count > 0 || selectedExternalEntities.Count > 0))
        {
            var selectionItems = new List<SelectionItem>();

            // Add current document entities
            foreach (var id in selectedIds)
            {
                selectionItems.Add(new SelectionItem
                {
                    DocumentPath = doc.Name,
                    Handle = id.Handle.ToString(),
                    SessionId = null // Session scope doesn't use session restrictions
                });
            }

            // Add selected external entities (from other documents)
            foreach (var ext in selectedExternalEntities)
            {
                selectionItems.Add(new SelectionItem
                {
                    DocumentPath = ext["DocumentPath"].ToString(),
                    Handle = ext["Handle"].ToString(),
                    SessionId = null
                });
            }

            SelectionStorage.SaveSelection(selectionItems); // Save to global storage for session scope
            ed.WriteMessage($"Filtered selection saved: {selectedIds.Count} current document entities + {selectedExternalEntities.Count} external entities.\n");
        }
        else
        {
            ed.WriteMessage("All selections cleared (no items selected in filter).\n");
        }
    }

    /// <summary>
    /// Clear all stored selections across all documents for session scope
    /// </summary>
    private void ClearAllStoredSelections()
    {
        try
        {
            // Clear all per-document selection files
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var selectionDir = Path.Combine(appDataPath, "autocad-ballet", "runtime", "selection");

            if (Directory.Exists(selectionDir))
            {
                foreach (var file in Directory.GetFiles(selectionDir))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        // Skip files that can't be deleted
                    }
                }
            }

            // Also clear legacy global file for backward compatibility
            var legacyFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "autocad-ballet", "selection");
            if (File.Exists(legacyFilePath))
            {
                try
                {
                    File.WriteAllLines(legacyFilePath, new string[0]);
                }
                catch
                {
                    // Skip if can't clear legacy file
                }
            }
        }
        catch
        {
            // If clearing fails, continue anyway - the save operation will overwrite
        }
    }
}