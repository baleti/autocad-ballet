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
    public override bool IncludeProperties => true; // Include geometry properties (with detailed diagnostics)
    public override SelectionScope Scope => SelectionScope.application;

    public override void Execute()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        ObjectId[] originalSelection = null;

        var totalTimer = System.Diagnostics.Stopwatch.StartNew();
        var timer = System.Diagnostics.Stopwatch.StartNew();

        // Create diagnostics file
        var diagPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "autocad-ballet", "runtime", "diagnostics");
        Directory.CreateDirectory(diagPath);
        var diagFile = Path.Combine(diagPath, $"filter-selection-in-session-{DateTime.Now:yyyyMMdd-HHmmss}.txt");
        var diagLog = new System.Text.StringBuilder();

        void LogDiag(string message)
        {
            diagLog.AppendLine(message);
            ed.WriteMessage($"{message}\n");
        }

        try
        {
            LogDiag("================================================================================");
            LogDiag($"FILTER-SELECTION-IN-SESSION DIAGNOSTICS - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            LogDiag("================================================================================");
            LogDiag("");
            LogDiag("[DIAG] Starting filter-selection-in-session");

            // Force session scope behavior - load selection from all open documents (not closed ones)
            timer.Restart();
            var storedSelection = SelectionStorage.LoadSelectionFromOpenDocuments();
            timer.Stop();
            LogDiag($"[DIAG] LoadSelectionFromOpenDocuments: {timer.ElapsedMilliseconds}ms, {storedSelection?.Count ?? 0} items");

            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-categories-in-session' first.\n");
                return;
            }

            // Create entity data from the stored selection using the helper method
            var entityData = new List<Dictionary<string, object>>();

            timer.Restart();
            var groupByDocTimer = new System.Diagnostics.Stopwatch();
            var findDocTimer = new System.Diagnostics.Stopwatch();
            var transactionTimer = new System.Diagnostics.Stopwatch();
            var getEntityDataTimer = new System.Diagnostics.Stopwatch();
            int itemsProcessed = 0;

            // Track per-document statistics
            var docStats = new Dictionary<string, (int count, long timeMs)>();

            LogDiag($"[DIAG] Starting entity collection loop for {storedSelection.Count} items");

            // Reset GetEntityDataDictionary diagnostics
            FilterEntityDataHelper.ResetDiagnostics();

            // OPTIMIZATION: Group items by document path first to avoid repeated document lookups
            groupByDocTimer.Start();
            var itemsByDocument = storedSelection.GroupBy(item => item.DocumentPath).ToDictionary(g => g.Key, g => g.ToList());
            groupByDocTimer.Stop();
            LogDiag($"[DIAG] Grouped {storedSelection.Count} items into {itemsByDocument.Count} documents in {groupByDocTimer.ElapsedMilliseconds}ms");

            // Build a map of open documents for fast lookup
            findDocTimer.Start();
            var openDocuments = new Dictionary<string, Document>(StringComparer.OrdinalIgnoreCase);
            foreach (Document openDoc in AcadApp.DocumentManager)
            {
                try
                {
                    var fullPath = Path.GetFullPath(openDoc.Name);
                    openDocuments[fullPath] = openDoc;
                    // Also add the raw name for fallback
                    if (!openDocuments.ContainsKey(openDoc.Name))
                    {
                        openDocuments[openDoc.Name] = openDoc;
                    }
                }
                catch
                {
                    // If GetFullPath fails, just use the name
                    openDocuments[openDoc.Name] = openDoc;
                }
            }
            findDocTimer.Stop();
            LogDiag($"[DIAG] Built open documents map in {findDocTimer.ElapsedMilliseconds}ms, {openDocuments.Count} documents open");

            // Process each document's items together
            int docIndex = 0;
            foreach (var docGroup in itemsByDocument)
            {
                docIndex++;
                var docPath = docGroup.Key;
                var items = docGroup.Value;
                var docTimer = System.Diagnostics.Stopwatch.StartNew();

                LogDiag($"[DIAG] Processing document {docIndex}/{itemsByDocument.Count}: {Path.GetFileName(docPath)} ({items.Count} items)");

                try
                {
                    // Find the document in our open documents map
                    Document itemDoc = null;

                    // Try full path first
                    try
                    {
                        var fullPath = Path.GetFullPath(docPath);
                        if (openDocuments.TryGetValue(fullPath, out itemDoc))
                        {
                            // Found it
                        }
                    }
                    catch { }

                    // Try raw path as fallback
                    if (itemDoc == null)
                    {
                        openDocuments.TryGetValue(docPath, out itemDoc);
                    }

                    if (itemDoc != null)
                    {
                        var docName = Path.GetFileName(itemDoc.Name);
                        int entitiesCollected = 0;

                        // Process all items for this document in a single transaction
                        transactionTimer.Start();
                        using (var tr = itemDoc.Database.TransactionManager.StartTransaction())
                        {
                            // Build block hierarchy cache ONCE per document (huge performance optimization)
                            var cacheTimer = System.Diagnostics.Stopwatch.StartNew();
                            var blockHierarchyCache = FilterEntityDataHelper.BuildBlockHierarchyCache(itemDoc.Database, tr);
                            cacheTimer.Stop();
                            LogDiag($"[DIAG]   - Built block hierarchy cache for {docName} in {cacheTimer.ElapsedMilliseconds}ms");

                            foreach (var item in items)
                            {
                                itemsProcessed++;

                                try
                                {
                                    var handle = Convert.ToInt64(item.Handle, 16);
                                    var objectId = itemDoc.Database.GetObjectId(false, new Autodesk.AutoCAD.DatabaseServices.Handle(handle), 0);

                                    if (objectId != Autodesk.AutoCAD.DatabaseServices.ObjectId.Null)
                                    {
                                        var entity = tr.GetObject(objectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                        if (entity != null)
                                        {
                                            getEntityDataTimer.Start();
                                            var data = FilterEntityDataHelper.GetEntityDataDictionary(entity, item.DocumentPath, null, IncludeProperties, tr, blockHierarchyCache);
                                            getEntityDataTimer.Stop();

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
                                            entitiesCollected++;
                                        }
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
                        transactionTimer.Stop();

                        docTimer.Stop();

                        // Track per-document stats
                        docStats[docName] = (entitiesCollected, docTimer.ElapsedMilliseconds);

                        LogDiag($"[DIAG]   - Collected {entitiesCollected} entities from {docName} in {docTimer.ElapsedMilliseconds}ms (avg {(entitiesCollected > 0 ? docTimer.ElapsedMilliseconds / entitiesCollected : 0)}ms per entity)");
                    }
                    else
                    {
                        // Document is not open - skip all items for this document
                        itemsProcessed += items.Count;
                        LogDiag($"[DIAG]   - Document not open, skipped {items.Count} items");
                    }
                }
                catch (System.Exception ex)
                {
                    LogDiag($"[DIAG]   - ERROR processing document: {ex.Message}");
                    // Skip this entire document
                    itemsProcessed += items.Count;
                    continue;
                }
            }
            timer.Stop();

            LogDiag($"");
            LogDiag($"[DIAG] ========== Entity Collection Summary ==========");
            LogDiag($"[DIAG] Total time: {timer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG] Items in stored selection: {storedSelection.Count}");
            LogDiag($"[DIAG] Items processed: {itemsProcessed}");
            LogDiag($"[DIAG] Documents processed: {itemsByDocument.Count}");
            LogDiag($"[DIAG] Entities collected: {entityData.Count}");
            LogDiag($"[DIAG] Time breakdown:");
            LogDiag($"[DIAG]   - Group by document: {groupByDocTimer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG]   - Build open documents map: {findDocTimer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG]   - Transactions: {transactionTimer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG]   - GetEntityDataDictionary: {getEntityDataTimer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG] ================================================");
            LogDiag($"");
            LogDiag($"[DIAG] ========== GetEntityDataDictionary Breakdown ==========");
            LogDiag(FilterEntityDataHelper.GetDiagnosticsSummary());
            LogDiag($"[DIAG] ============================================================");

            if (!entityData.Any())
            {
                ed.WriteMessage("\nNo valid entities found in stored selection.\n");
                return;
            }

            timer.Restart();
            // Process the entity data using the standard filtering logic from base class
            ProcessSelectionResults(entityData, originalSelection, false, diagLog);
            timer.Stop();
            LogDiag($"[DIAG] ProcessSelectionResults: {timer.ElapsedMilliseconds}ms");

            totalTimer.Stop();
            LogDiag($"[DIAG] TOTAL Execute() time: {totalTimer.ElapsedMilliseconds}ms");

            // Save diagnostics to file
            File.WriteAllText(diagFile, diagLog.ToString());
            ed.WriteMessage($"\n[DIAG] Diagnostics saved to: {diagFile}\n");
        }
        catch (InvalidOperationException ex)
        {
            LogDiag($"[ERROR] InvalidOperationException: {ex.Message}");
            LogDiag($"[ERROR] Stack trace: {ex.StackTrace}");
            File.WriteAllText(diagFile, diagLog.ToString());
            ed.WriteMessage($"\n[DIAG] Error diagnostics saved to: {diagFile}\n");
            ed.WriteMessage($"\nError: {ex.Message}\n");
        }
        catch (System.Exception ex)
        {
            LogDiag($"[ERROR] Exception: {ex.Message}");
            LogDiag($"[ERROR] Stack trace: {ex.StackTrace}");
            File.WriteAllText(diagFile, diagLog.ToString());
            ed.WriteMessage($"\n[DIAG] Error diagnostics saved to: {diagFile}\n");
            ed.WriteMessage($"\nUnexpected error: {ex.Message}\n");
        }
    }

    private void ProcessSelectionResults(List<Dictionary<string, object>> entityData, ObjectId[] originalSelection, bool editsWereApplied, System.Text.StringBuilder diagLog)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var timer = System.Diagnostics.Stopwatch.StartNew();

        void LogDiag(string message)
        {
            diagLog.AppendLine(message);
            ed.WriteMessage($"{message}\n");
        }

        LogDiag($"[DIAG] ProcessSelectionResults starting with {entityData.Count} entities");

        // Process parent block hierarchy into separate columns
        timer.Restart();
        FilterEntityDataHelper.ProcessParentBlockColumns(entityData);
        timer.Stop();
        LogDiag($"[DIAG]   - ProcessParentBlockColumns: {timer.ElapsedMilliseconds}ms");

        // Add index to each entity for mapping back after user selection
        timer.Restart();
        for (int i = 0; i < entityData.Count; i++)
        {
            entityData[i]["OriginalIndex"] = i;
        }
        timer.Stop();
        LogDiag($"[DIAG]   - Add indices: {timer.ElapsedMilliseconds}ms");

        // Get property names, excluding internal fields
        timer.Restart();
        var propertyNames = entityData.First().Keys
            .Where(k => !k.EndsWith("ObjectId") && k != "OriginalIndex" && k != "_ParentBlocks")
            .ToList();
        timer.Stop();
        LogDiag($"[DIAG]   - Get property names: {timer.ElapsedMilliseconds}ms, {propertyNames.Count} properties");

        // Reorder to put most useful columns first
        timer.Restart();
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
        timer.Stop();
        LogDiag($"[DIAG]   - Reorder properties: {timer.ElapsedMilliseconds}ms");

        // Reset the edits flag at the start of each DataGrid session
        CustomGUIs.ResetEditsAppliedFlag();

        LogDiag($"[DIAG]   - About to show DataGrid with {entityData.Count} rows, {propertyNames.Count} columns...");
        timer.Restart();
        var chosenRows = CustomGUIs.DataGrid(entityData, propertyNames, SpanAllScreens, null);
        timer.Stop();
        LogDiag($"[DIAG]   - DataGrid returned: {timer.ElapsedMilliseconds}ms, {chosenRows.Count} rows selected");

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
        timer.Restart();
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
        timer.Stop();
        LogDiag($"[DIAG]   - Collect ObjectIds: {timer.ElapsedMilliseconds}ms");

        // Set selection for current document entities (only if no edits were applied)
        timer.Restart();
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
        timer.Stop();
        LogDiag($"[DIAG]   - Set selection and report: {timer.ElapsedMilliseconds}ms");

        // Clear all existing selections and save only the filtered results (session scope behavior)
        timer.Restart();
        ClearAllStoredSelections();
        timer.Stop();
        LogDiag($"[DIAG]   - ClearAllStoredSelections: {timer.ElapsedMilliseconds}ms");

        timer.Restart();
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
        timer.Stop();
        LogDiag($"[DIAG]   - Save selection: {timer.ElapsedMilliseconds}ms");
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