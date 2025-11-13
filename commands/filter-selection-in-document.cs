using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
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
        var totalTimer = Stopwatch.StartNew();
        var diagnostics = new StringBuilder();

        diagnostics.AppendLine("================================================================================");
        diagnostics.AppendLine($"FILTER-SELECTION-IN-DOCUMENT DIAGNOSTICS - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        diagnostics.AppendLine("================================================================================");
        diagnostics.AppendLine();

        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        void LogDiag(string message)
        {
            diagnostics.AppendLine(message);
            ed.WriteMessage($"{message}\n");
        }

        try
        {
            LogDiag("[DIAG] Starting filter-selection-in-document");

            // Load stored selection for document
            var loadTimer = Stopwatch.StartNew();
            var docName = Path.GetFileName(doc.Name);
            var storedSelection = SelectionStorage.LoadSelection(docName);
            loadTimer.Stop();

            LogDiag($"[DIAG] LoadSelection: {loadTimer.ElapsedMilliseconds}ms, {storedSelection?.Count ?? 0} items");

            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found for current document. Use 'select-by-categories-in-document' first.\n");
                return;
            }

            LogDiag($"[DIAG] Starting entity collection loop for {storedSelection.Count} items");

            // Collect entity data with optimized single transaction
            var entityData = new List<Dictionary<string, object>>();
            var entityCollectionTimer = Stopwatch.StartNew();
            var getEntityDataTimer = new Stopwatch();
            var transactionTimer = new Stopwatch();

            // Reset GetEntityDataDictionary diagnostics
            FilterEntityDataHelper.ResetDiagnostics();

            // Single transaction for all entities in the current document
            int entitiesCollected = 0;
            transactionTimer.Start();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Build block hierarchy cache ONCE at the start (huge performance optimization)
                var cacheTimer = Stopwatch.StartNew();
                var blockHierarchyCache = FilterEntityDataHelper.BuildBlockHierarchyCache(db, tr);
                cacheTimer.Stop();
                LogDiag($"[DIAG] Built block hierarchy cache in {cacheTimer.ElapsedMilliseconds}ms");

                foreach (var item in storedSelection)
                {
                    try
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (objectId != ObjectId.Null)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null)
                            {
                                getEntityDataTimer.Start();
                                var data = FilterEntityDataHelper.GetEntityDataDictionary(entity, item.DocumentPath, null, IncludeProperties, tr, blockHierarchyCache);
                                getEntityDataTimer.Stop();

                                data["ObjectId"] = objectId;
                                entityData.Add(data);
                                entitiesCollected++;
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        LogDiag($"[DIAG]   - Error processing entity {item.Handle}: {ex.Message}");
                    }
                }
                tr.Commit();
            }
            transactionTimer.Stop();
            entityCollectionTimer.Stop();

            LogDiag($"");
            LogDiag($"[DIAG] ========== Entity Collection Summary ==========");
            LogDiag($"[DIAG] Total time: {entityCollectionTimer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG] Items in stored selection: {storedSelection.Count}");
            LogDiag($"[DIAG] Entities collected: {entityData.Count}");
            LogDiag($"[DIAG] Time breakdown:");
            LogDiag($"[DIAG]   - Transaction: {transactionTimer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG]   - GetEntityDataDictionary: {getEntityDataTimer.ElapsedMilliseconds}ms");
            LogDiag($"[DIAG] ================================================");
            LogDiag($"");

            if (entityData.Count == 0)
            {
                ed.WriteMessage("\nNo entities found in stored selection.\n");
                return;
            }

            // Get GetEntityDataDictionary breakdown
            LogDiag($"[DIAG] ========== GetEntityDataDictionary Breakdown ==========");
            LogDiag(FilterEntityDataHelper.GetDiagnosticsSummary());
            LogDiag($"[DIAG] ============================================================");

            // Process results with DataGrid
            LogDiag($"[DIAG] ProcessSelectionResults starting with {entityData.Count} entities");
            var processTimer = Stopwatch.StartNew();
            ProcessSelectionResults(entityData, diagnostics);
            processTimer.Stop();
            LogDiag($"[DIAG] ProcessSelectionResults: {processTimer.ElapsedMilliseconds}ms");

            totalTimer.Stop();
            LogDiag($"[DIAG] TOTAL Execute() time: {totalTimer.ElapsedMilliseconds}ms");
            LogDiag($"");

            // Save diagnostics to file
            var runtimeDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "autocad-ballet",
                "runtime",
                "diagnostics");
            Directory.CreateDirectory(runtimeDir);

            var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            var diagnosticsFile = Path.Combine(runtimeDir, $"filter-selection-in-document-{timestamp}.txt");
            File.WriteAllText(diagnosticsFile, diagnostics.ToString());

            ed.WriteMessage($"[DIAG] Diagnostics saved to: {diagnosticsFile}\n");
        }
        catch (System.Exception ex)
        {
            diagnostics.AppendLine($"[ERROR] Exception: {ex.Message}");
            diagnostics.AppendLine($"[ERROR] Stack trace: {ex.StackTrace}");
            ed.WriteMessage($"\nError in filter-selection-in-document: {ex.Message}\n");

            // Save diagnostics even on error
            try
            {
                var runtimeDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "autocad-ballet",
                    "runtime",
                    "diagnostics");
                Directory.CreateDirectory(runtimeDir);

                var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                var diagnosticsFile = Path.Combine(runtimeDir, $"filter-selection-in-document-error-{timestamp}.txt");
                File.WriteAllText(diagnosticsFile, diagnostics.ToString());
            }
            catch { }
        }
    }

    private void ProcessSelectionResults(List<Dictionary<string, object>> entityData, StringBuilder diagLog)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;
        var timer = Stopwatch.StartNew();

        void LogDiag(string message)
        {
            diagLog.AppendLine(message);
            ed.WriteMessage($"{message}\n");
        }

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
        timer.Stop();
        LogDiag($"[DIAG]   - Reorder properties: {timer.ElapsedMilliseconds}ms");

        // Show DataGrid
        LogDiag($"[DIAG]   - About to show DataGrid with {entityData.Count} rows, {propertyNames.Count} columns...");
        timer.Restart();
        var chosenRows = CustomGUIs.DataGrid(entityData, propertyNames, SpanAllScreens, null);
        timer.Stop();
        LogDiag($"[DIAG]   - DataGrid returned: {timer.ElapsedMilliseconds}ms, {chosenRows.Count} rows selected");

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
    }
}