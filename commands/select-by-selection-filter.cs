using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public class FilterEntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string SpaceName { get; set; }
    }

    public static class SelectBySelectionFilter
    {
        /// <summary>
        /// Saves any modified filters back to storage after DataGrid editing
        /// </summary>
        private static void SaveModifiedFilters(List<Dictionary<string, object>> originalEntries, List<StoredSelectionFilter> originalFilters, Editor ed)
        {
            // Check if any entries were modified (Query column was edited)
            bool anyModified = false;
            var updatedFilters = new List<StoredSelectionFilter>();

            foreach (var filter in originalFilters)
            {
                // Find matching entry
                var entry = originalEntries.FirstOrDefault(e =>
                    e.ContainsKey("Name") &&
                    e["Name"].ToString() == filter.Name);

                if (entry != null)
                {
                    string currentQuery = entry.ContainsKey("Query") ? entry["Query"].ToString() : "";

                    if (currentQuery != filter.QueryString)
                    {
                        // Query was modified
                        anyModified = true;
                        updatedFilters.Add(new StoredSelectionFilter
                        {
                            Name = filter.Name,
                            QueryString = currentQuery,
                            SourceDocumentPath = filter.SourceDocumentPath
                        });
                    }
                    else
                    {
                        // No change
                        updatedFilters.Add(filter);
                    }
                }
                else
                {
                    // Filter was deleted or not found, keep original
                    updatedFilters.Add(filter);
                }
            }

            if (anyModified)
            {
                try
                {
                    SelectionFilterStorage.SaveFilters(updatedFilters);
                    ed.WriteMessage("\nSelection filters updated with your edits.\n");
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\nError saving edited filters: {ex.Message}\n");
                }
            }
        }

        public static void ExecuteViewScope(Editor ed, Database db)
        {
            ed.WriteMessage("\nView Mode: Loading selection filters...\n");

            // Load saved filters
            var filters = SelectionFilterStorage.LoadFilters();

            if (filters.Count == 0)
            {
                ed.WriteMessage("\nNo selection filters found. Use 'save-selection-filter-from-selected-in-view' to create filters.\n");
                return;
            }

            // Show filter selection dialog
            var entries = filters.Select(f => new Dictionary<string, object>
            {
                { "Name", f.Name },
                { "Source Document", f.SourceDocumentPath ?? "Unknown" },
                { "Query", f.QueryString }
            }).ToList();

            var propertyNames = new List<string> { "Name", "Source Document", "Query" };
            var chosenRows = CustomGUIs.DataGrid(
                entries,
                propertyNames,
                spanAllScreens: false,
                initialSelectionIndices: null,
                onDeleteEntries: (entriesToDelete) =>
                {
                    // Confirm deletion
                    if (entriesToDelete.Count == 0)
                        return false;

                    string message = entriesToDelete.Count == 1
                        ? $"Delete filter '{entriesToDelete[0]["Name"]}'?"
                        : $"Delete {entriesToDelete.Count} filters?";

                    var result = System.Windows.Forms.MessageBox.Show(
                        message,
                        "Confirm Deletion",
                        System.Windows.Forms.MessageBoxButtons.OKCancel,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button1); // OK is default

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        // Delete filters
                        var namesToDelete = entriesToDelete
                            .Where(e => e.ContainsKey("Name"))
                            .Select(e => e["Name"].ToString())
                            .ToList();

                        SelectionFilterStorage.DeleteFilters(namesToDelete);
                        return true; // Remove from grid
                    }

                    return false; // Don't remove from grid
                });

            // Save any modified filters after DataGrid closes
            SaveModifiedFilters(entries, filters, ed);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo filters selected.\n");
                return;
            }

            // Apply selected filter to current view
            var selectedFilter = chosenRows[0]; // Single selection
            string filterName = selectedFilter["Name"].ToString();
            string queryString = selectedFilter["Query"].ToString();

            ed.WriteMessage($"\nApplying filter '{filterName}' to current view...\n");

            // Parse query once before loop (optimization!)
            var parsedQuery = ParseQuery(queryString);

            var matchedIds = new List<ObjectId>();
            var spaceName = GetCurrentSpaceName(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in currentSpace)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity != null && MatchesQueryParsed(entity, parsedQuery, tr))
                        {
                            matchedIds.Add(id);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                tr.Commit();
            }

            if (matchedIds.Count > 0)
            {
                ed.SetImpliedSelection(matchedIds.ToArray());
                ed.WriteMessage($"\nSelected {matchedIds.Count} entities matching filter '{filterName}' in {spaceName}.\n");
            }
            else
            {
                ed.WriteMessage($"\nNo entities found matching filter '{filterName}' in {spaceName}.\n");
            }
        }

        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            ed.WriteMessage("\nDocument Mode: Loading selection filters...\n");

            // Load saved filters
            var filters = SelectionFilterStorage.LoadFilters();
            if (filters.Count == 0)
            {
                ed.WriteMessage("\nNo selection filters found. Use 'save-selection-filter-from-selected-in-view' to create filters.\n");
                return;
            }

            // Show filter selection dialog
            var entries = filters.Select(f => new Dictionary<string, object>
            {
                { "Name", f.Name },
                { "Source Document", f.SourceDocumentPath ?? "Unknown" },
                { "Query", f.QueryString }
            }).ToList();

            var propertyNames = new List<string> { "Name", "Source Document", "Query" };
            var chosenRows = CustomGUIs.DataGrid(
                entries,
                propertyNames,
                spanAllScreens: false,
                initialSelectionIndices: null,
                onDeleteEntries: (entriesToDelete) =>
                {
                    // Confirm deletion
                    if (entriesToDelete.Count == 0)
                        return false;

                    string message = entriesToDelete.Count == 1
                        ? $"Delete filter '{entriesToDelete[0]["Name"]}'?"
                        : $"Delete {entriesToDelete.Count} filters?";

                    var result = System.Windows.Forms.MessageBox.Show(
                        message,
                        "Confirm Deletion",
                        System.Windows.Forms.MessageBoxButtons.OKCancel,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button1); // OK is default

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        // Delete filters
                        var namesToDelete = entriesToDelete
                            .Where(e => e.ContainsKey("Name"))
                            .Select(e => e["Name"].ToString())
                            .ToList();

                        SelectionFilterStorage.DeleteFilters(namesToDelete);
                        return true; // Remove from grid
                    }

                    return false; // Don't remove from grid
                });

            // Save any modified filters after DataGrid closes
            SaveModifiedFilters(entries, filters, ed);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo filters selected.\n");
                return;
            }

            // Apply selected filter to entire document
            var selectedFilter = chosenRows[0]; // Single selection
            string filterName = selectedFilter["Name"].ToString();
            string queryString = selectedFilter["Query"].ToString();

            ed.WriteMessage($"\nApplying filter '{filterName}' to entire document...\n");
            ed.WriteMessage($"Query: {queryString}\n");

            var matchedEntities = new List<FilterEntityReference>();

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
                            if (entity != null && MatchesQuery(entity, queryString, db))
                            {
                                matchedEntities.Add(new FilterEntityReference
                                {
                                    DocumentPath = doc.Name,
                                    DocumentName = Path.GetFileName(doc.Name),
                                    Handle = entity.Handle.ToString(),
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

            // Remove duplicates
            matchedEntities = matchedEntities
                .GroupBy(e => e.Handle)
                .Select(g => g.First())
                .ToList();

            // Save to selection storage
            var selectionItems = matchedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            try
            {
                var docName = Path.GetFileName(doc.Name);
                SelectionStorage.SaveSelection(selectionItems, docName);
                ed.WriteMessage($"\nDocument selection saved: {matchedEntities.Count} entities.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection
            var currentViewIds = new List<ObjectId>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var entityRef in matchedEntities)
                {
                    try
                    {
                        var handle = Convert.ToInt64(entityRef.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (!objectId.IsNull && objectId.IsValid)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.BlockId == db.CurrentSpaceId)
                            {
                                currentViewIds.Add(objectId);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                tr.Commit();
            }

            if (currentViewIds.Count > 0)
            {
                ed.SetImpliedSelection(currentViewIds.ToArray());
                ed.WriteMessage($"Selected {currentViewIds.Count} entities in current view.\n");
            }
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            ed.WriteMessage("\nApplication Mode: Loading selection filters...\n");

            // Load saved filters
            var filters = SelectionFilterStorage.LoadFilters();
            if (filters.Count == 0)
            {
                ed.WriteMessage("\nNo selection filters found. Use 'save-selection-filter-from-selected-in-view' to create filters.\n");
                return;
            }

            // Show filter selection dialog
            var entries = filters.Select(f => new Dictionary<string, object>
            {
                { "Name", f.Name },
                { "Source Document", f.SourceDocumentPath ?? "Unknown" },
                { "Query", f.QueryString }
            }).ToList();

            var propertyNames = new List<string> { "Name", "Source Document", "Query" };
            var chosenRows = CustomGUIs.DataGrid(
                entries,
                propertyNames,
                spanAllScreens: false,
                initialSelectionIndices: null,
                onDeleteEntries: (entriesToDelete) =>
                {
                    // Confirm deletion
                    if (entriesToDelete.Count == 0)
                        return false;

                    string message = entriesToDelete.Count == 1
                        ? $"Delete filter '{entriesToDelete[0]["Name"]}'?"
                        : $"Delete {entriesToDelete.Count} filters?";

                    var result = System.Windows.Forms.MessageBox.Show(
                        message,
                        "Confirm Deletion",
                        System.Windows.Forms.MessageBoxButtons.OKCancel,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button1); // OK is default

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        // Delete filters
                        var namesToDelete = entriesToDelete
                            .Where(e => e.ContainsKey("Name"))
                            .Select(e => e["Name"].ToString())
                            .ToList();

                        SelectionFilterStorage.DeleteFilters(namesToDelete);
                        return true; // Remove from grid
                    }

                    return false; // Don't remove from grid
                });

            // Save any modified filters after DataGrid closes
            SaveModifiedFilters(entries, filters, ed);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo filters selected.\n");
                return;
            }

            // Apply selected filter to all open documents
            var selectedFilter = chosenRows[0]; // Single selection
            string filterName = selectedFilter["Name"].ToString();
            string queryString = selectedFilter["Query"].ToString();

            ed.WriteMessage($"\nApplying filter '{filterName}' to all open documents...\n");
            ed.WriteMessage($"Query: {queryString}\n");

            var docManager = AcadApp.DocumentManager;
            var matchedEntities = new List<FilterEntityReference>();

            foreach (Autodesk.AutoCAD.ApplicationServices.Document doc in docManager)
            {
                string docPath = doc.Name;
                string docName = Path.GetFileName(docPath);

                ed.WriteMessage($"\nScanning: {docName}...");

                try
                {
                    var refs = GatherMatchingEntitiesFromDocument(doc.Database, docPath, docName, queryString);
                    matchedEntities.AddRange(refs);
                    ed.WriteMessage($" found {refs.Count} matching entities");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
                }
            }

            // Remove duplicates
            matchedEntities = matchedEntities
                .GroupBy(e => new { e.DocumentPath, e.Handle })
                .Select(g => g.First())
                .ToList();

            // Save to selection storage
            var selectionItems = matchedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            try
            {
                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage($"\n\nApplication selection saved: {matchedEntities.Count} entities.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection
            var currentDoc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = currentDoc.Database;
            var currentDocPath = Path.GetFullPath(currentDoc.Name);
            var currentViewIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var entityRef in matchedEntities)
                {
                    try
                    {
                        var entityDocPath = Path.GetFullPath(entityRef.DocumentPath);
                        if (!string.Equals(entityDocPath, currentDocPath, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var handle = Convert.ToInt64(entityRef.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (!objectId.IsNull && objectId.IsValid)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.BlockId == db.CurrentSpaceId)
                            {
                                currentViewIds.Add(objectId);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                tr.Commit();
            }

            if (currentViewIds.Count > 0)
            {
                ed.SetImpliedSelection(currentViewIds.ToArray());
                ed.WriteMessage($"Selected {currentViewIds.Count} entities in current view.\n");
            }
        }

        private static List<FilterEntityReference> GatherMatchingEntitiesFromDocument(Database db, string docPath, string docName, string queryString)
        {
            var references = new List<FilterEntityReference>();

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
                            if (entity != null && MatchesQuery(entity, queryString, db))
                            {
                                references.Add(new FilterEntityReference
                                {
                                    DocumentPath = docPath,
                                    DocumentName = docName,
                                    Handle = entity.Handle.ToString(),
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

        // Pre-parsed query structure for performance
        private class ParsedQuery
        {
            public List<List<KeyValuePair<string, string>>> OrClauses { get; set; }
        }

        /// <summary>
        /// Parses query string once into reusable structure (OPTIMIZED)
        /// </summary>
        private static ParsedQuery ParseQuery(string queryString)
        {
            var result = new ParsedQuery
            {
                OrClauses = new List<List<KeyValuePair<string, string>>>()
            };

            if (string.IsNullOrWhiteSpace(queryString))
                return result;

            // Split by || for OR clauses
            var orClauses = queryString.Split(new[] { "||" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var orClause in orClauses)
            {
                var conditions = ParseConditions(orClause.Trim());
                if (conditions.Count > 0)
                {
                    result.OrClauses.Add(conditions.ToList());
                }
            }

            return result;
        }

        /// <summary>
        /// Checks if an entity matches the pre-parsed query (OPTIMIZED)
        /// Uses passed transaction instead of opening new ones
        /// </summary>
        private static bool MatchesQueryParsed(Entity entity, ParsedQuery parsedQuery, Transaction tr)
        {
            if (parsedQuery == null || parsedQuery.OrClauses.Count == 0)
                return false;

            // Check each OR clause
            foreach (var andConditions in parsedQuery.OrClauses)
            {
                bool allMatch = true;

                // Check all AND conditions in this clause
                foreach (var condition in andConditions)
                {
                    if (!MatchesConditionOptimized(entity, condition.Key, condition.Value, tr))
                    {
                        allMatch = false;
                        break;
                    }
                }

                // If all AND conditions matched, return true (OR satisfied)
                if (allMatch)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Checks if entity matches a single condition (OPTIMIZED)
        /// Reuses transaction instead of opening new ones
        /// </summary>
        private static bool MatchesConditionOptimized(Entity entity, string columnName, string expectedValue, Transaction tr)
        {
            string actualValue = GetEntityPropertyValueOptimized(entity, columnName, tr);

            // Case-insensitive comparison
            return string.Equals(actualValue, expectedValue, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Checks if an entity matches the query string
        /// Query format: "$column:\"value\"" with || for OR and space for AND
        /// Example: $category:"MText" $layer:"Annotations" || $category:"Text"
        /// </summary>
        private static bool MatchesQuery(Entity entity, string queryString, Database db)
        {
            if (string.IsNullOrWhiteSpace(queryString))
                return false;

            // Split by || for OR clauses
            var orClauses = queryString.Split(new[] { "||" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var orClause in orClauses)
            {
                // Each OR clause is a set of AND conditions
                if (MatchesAndClause(entity, orClause.Trim(), db))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Checks if entity matches all conditions in an AND clause
        /// </summary>
        private static bool MatchesAndClause(Entity entity, string andClause, Database db)
        {
            // Parse conditions from AND clause
            // Pattern: $column:"value" or $"column name":"value"
            var conditions = ParseConditions(andClause);

            foreach (var condition in conditions)
            {
                if (!MatchesCondition(entity, condition.Key, condition.Value, db))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Parses conditions from a clause string
        /// Returns dictionary of column name -> value
        /// </summary>
        private static Dictionary<string, string> ParseConditions(string clause)
        {
            var conditions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Pattern: $"column name":"value" or $column:"value"
            // More permissive regex that handles both quoted and unquoted column names
            var pattern = @"\$(?:""([^""]+)""|([^\s:]+)):""([^""]*)""";
            var matches = Regex.Matches(clause, pattern);

            foreach (Match match in matches)
            {
                string columnName = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
                string value = match.Groups[3].Value;

                // Unescape doubled quotes (from our escape logic)
                value = value.Replace("\"\"", "\"");

                conditions[columnName] = value;
            }

            return conditions;
        }

        /// <summary>
        /// Checks if entity matches a single condition
        /// </summary>
        private static bool MatchesCondition(Entity entity, string columnName, string expectedValue, Database db)
        {
            string actualValue = GetEntityPropertyValue(entity, columnName, db);

            // Case-insensitive comparison
            return string.Equals(actualValue, expectedValue, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets the property value for an entity based on column name (OPTIMIZED)
        /// Reuses transaction instead of opening new ones
        /// </summary>
        private static string GetEntityPropertyValueOptimized(Entity entity, string columnName, Transaction tr)
        {
            switch (columnName.ToLowerInvariant())
            {
                case "category":
                    return GetEntityCategoryOptimized(entity, tr);

                case "layer":
                    return entity.Layer ?? "";

                case "name":
                    return GetEntityNameOptimized(entity, tr) ?? "";

                case "dynamic block name":
                    return GetDynamicBlockNameOptimized(entity, tr) ?? "";

                case "color":
                    return entity.Color.ToString() ?? "";

                case "linetype":
                    return entity.Linetype ?? "";

                default:
                    return "";
            }
        }

        /// <summary>
        /// Gets the property value for an entity based on column name
        /// </summary>
        private static string GetEntityPropertyValue(Entity entity, string columnName, Database db)
        {
            switch (columnName.ToLowerInvariant())
            {
                case "category":
                    return GetEntityCategory(entity);

                case "layer":
                    return entity.Layer ?? "";

                case "name":
                    return GetEntityName(entity) ?? "";

                case "dynamic block name":
                    return GetDynamicBlockName(entity) ?? "";

                case "color":
                    return entity.Color.ToString() ?? "";

                case "linetype":
                    return entity.Linetype ?? "";

                default:
                    return "";
            }
        }

        private static string GetEntityCategory(DBObject entity)
        {
            if (entity is BlockReference)
            {
                var br = entity as BlockReference;
                using (var tr = br.Database.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr.IsFromExternalReference)
                        return "XRef";
                    else if (btr.IsAnonymous)
                        return "Dynamic Block";
                    else
                        return "Block Reference";
                }
            }
            else if (entity is MText)
                return "MText";
            else if (entity is DBText)
                return "Text";
            else if (entity is Line)
                return "Line";
            else if (entity is Polyline)
                return "Polyline";
            else if (entity is Circle)
                return "Circle";
            else if (entity is Arc)
                return "Arc";
            // Add more as needed
            else
                return entity.GetType().Name;
        }

        private static string GetEntityName(Entity entity)
        {
            if (entity is BlockReference br)
            {
                using (var tr = br.Database.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    return btr?.Name ?? "";
                }
            }
            return "";
        }

        private static string GetDynamicBlockName(Entity entity)
        {
            if (entity is BlockReference blockRef)
            {
                using (var tr = blockRef.Database.TransactionManager.StartTransaction())
                {
                    var dynamicBlockTableRecordId = blockRef.DynamicBlockTableRecord;
                    if (dynamicBlockTableRecordId != ObjectId.Null && dynamicBlockTableRecordId != blockRef.BlockTableRecord)
                    {
                        var dynamicBtr = tr.GetObject(dynamicBlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                        return dynamicBtr?.Name ?? "";
                    }
                }
            }
            return "";
        }

        // Optimized versions that reuse transactions (MAJOR PERFORMANCE IMPROVEMENT)
        private static string GetEntityCategoryOptimized(DBObject entity, Transaction tr)
        {
            if (entity is BlockReference)
            {
                var br = entity as BlockReference;
                var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (btr.IsFromExternalReference)
                    return "XRef";
                else if (btr.IsAnonymous)
                    return "Dynamic Block";
                else
                    return "Block Reference";
            }
            else if (entity is MText)
                return "MText";
            else if (entity is DBText)
                return "Text";
            else if (entity is Line)
                return "Line";
            else if (entity is Polyline)
                return "Polyline";
            else if (entity is Circle)
                return "Circle";
            else if (entity is Arc)
                return "Arc";
            else
                return entity.GetType().Name;
        }

        private static string GetEntityNameOptimized(Entity entity, Transaction tr)
        {
            if (entity is BlockReference br)
            {
                var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                return btr?.Name ?? "";
            }
            return "";
        }

        private static string GetDynamicBlockNameOptimized(Entity entity, Transaction tr)
        {
            if (entity is BlockReference blockRef)
            {
                var dynamicBlockTableRecordId = blockRef.DynamicBlockTableRecord;
                if (dynamicBlockTableRecordId != ObjectId.Null && dynamicBlockTableRecordId != blockRef.BlockTableRecord)
                {
                    var dynamicBtr = tr.GetObject(dynamicBlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                    return dynamicBtr?.Name ?? "";
                }
            }
            return "";
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
