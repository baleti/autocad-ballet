using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class SaveSelectionFilterFromSelected
    {
        /// <summary>
        /// Saves any modified filters back to storage after DataGrid editing
        /// </summary>
        private static void SaveModifiedFiltersFromEntries(List<Dictionary<string, object>> entries, List<StoredSelectionFilter> originalFilters, Editor ed)
        {
            // Check if any entries were modified (Query column was edited)
            bool anyModified = false;
            var updatedFilters = new List<StoredSelectionFilter>();

            foreach (var filter in originalFilters)
            {
                // Find matching entry
                var entry = entries.FirstOrDefault(e =>
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

        /// <summary>
        /// Generates a query string from selected entities in the current view
        /// </summary>
        public static void ExecuteViewScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Get pickfirst set or prompt for selection
            var selResult = ed.SelectImplied();

            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to create filter from: ";
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

            var selectedIds = selResult.Value.GetObjectIds();

            // Analyze selected entities and generate query string
            string queryString = GenerateQueryFromEntities(selectedIds, db, ed);

            if (string.IsNullOrWhiteSpace(queryString))
            {
                ed.WriteMessage("\nCould not generate query from selection.\n");
                return;
            }

            ed.WriteMessage($"\nGenerated query: {queryString}\n");

            // Load existing filters for display
            var existingFilters = SelectionFilterStorage.LoadFilters();
            var entries = existingFilters.Select(f => new Dictionary<string, object>
            {
                { "Name", f.Name },
                { "Source Document", f.SourceDocumentPath ?? "Unknown" },
                { "Query", f.QueryString }
            }).ToList();

            var propertyNames = new List<string> { "Name", "Source Document", "Query" };

            ed.WriteMessage($"\nEnter a name for this filter in the search box and press Enter, or select an existing filter to replace...\n");

            // Keep a copy of the original filters list for comparison
            var originalFilters = existingFilters.ToList();

            // Use DataGrid with allowCreateFromSearch enabled and delete functionality
            var selectedFilters = CustomGUIs.DataGrid(
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

                    var result = MessageBox.Show(
                        message,
                        "Confirm Deletion",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1); // OK is default

                    if (result == DialogResult.OK)
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
                },
                allowCreateFromSearch: true);

            // Save any modified filters after DataGrid closes (if user edited any Query columns)
            SaveModifiedFiltersFromEntries(entries, originalFilters, ed);

            if (selectedFilters == null || selectedFilters.Count == 0)
            {
                ed.WriteMessage("\nCancelled.\n");
                return;
            }

            // Extract filter name - either from selected entry or from search text
            string filterName = null;

            if (selectedFilters[0].ContainsKey("__SEARCH_TEXT__"))
            {
                // User typed a new filter name
                filterName = (string)selectedFilters[0]["__SEARCH_TEXT__"];
            }
            else if (selectedFilters[0].ContainsKey("Name"))
            {
                // User selected existing filter to replace
                filterName = (string)selectedFilters[0]["Name"];
            }

            if (string.IsNullOrWhiteSpace(filterName))
            {
                ed.WriteMessage("\nNo valid filter name provided.\n");
                return;
            }

            // Save the filter with source document path
            try
            {
                string sourceDocPath = doc.Name;
                SelectionFilterStorage.AddOrUpdateFilter(filterName, queryString, sourceDocPath);
                ed.WriteMessage($"\nSelection filter '{filterName}' saved successfully.\n");
                ed.WriteMessage($"Source document: {sourceDocPath}\n");
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError saving filter: {ex.Message}\n");
            }
        }

        /// <summary>
        /// Generates a filter query string from selected entities
        /// Similar to select-similar logic but outputs query syntax
        /// </summary>
        private static string GenerateQueryFromEntities(ObjectId[] objectIds, Database db, Editor ed)
        {
            var queryParts = new HashSet<string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var id in objectIds)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity == null)
                            continue;

                        // Generate query parts for this entity
                        var parts = GenerateEntityQueryParts(entity);
                        foreach (var part in parts)
                        {
                            queryParts.Add(part);
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

            if (queryParts.Count == 0)
                return string.Empty;

            // Combine query parts with OR operator
            return string.Join(" || ", queryParts.OrderBy(p => p));
        }

        /// <summary>
        /// Generates query parts for a single entity
        /// Returns a combination of entity characteristics as query strings
        /// </summary>
        private static List<string> GenerateEntityQueryParts(Entity entity)
        {
            var parts = new List<string>();
            var conditions = new List<string>();

            // Get category
            string category = GetEntityCategory(entity);
            if (!string.IsNullOrEmpty(category))
            {
                conditions.Add($"$category:\"{EscapeQueryValue(category)}\"");
            }

            // Get layer
            string layer = entity.Layer;
            if (!string.IsNullOrEmpty(layer))
            {
                conditions.Add($"$layer:\"{EscapeQueryValue(layer)}\"");
            }

            // For BlockReferences, get block name and dynamic block name
            if (entity is BlockReference blockRef)
            {
                using (var tr = blockRef.Database.TransactionManager.StartTransaction())
                {
                    // Get regular block name
                    var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    string blockName = btr?.Name;
                    if (!string.IsNullOrEmpty(blockName))
                    {
                        conditions.Add($"$name:\"{EscapeQueryValue(blockName)}\"");
                    }

                    // Get dynamic block parent name if it's a dynamic block
                    var dynamicBlockTableRecordId = blockRef.DynamicBlockTableRecord;
                    if (dynamicBlockTableRecordId != ObjectId.Null && dynamicBlockTableRecordId != blockRef.BlockTableRecord)
                    {
                        var dynamicBtr = tr.GetObject(dynamicBlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                        string dynBlockName = dynamicBtr?.Name;
                        if (!string.IsNullOrEmpty(dynBlockName))
                        {
                            conditions.Add($"$\"dynamic block name\":\"{EscapeQueryValue(dynBlockName)}\"");
                        }
                    }
                }
            }

            // Combine all conditions for this entity with AND (space-separated)
            if (conditions.Count > 0)
            {
                parts.Add(string.Join(" ", conditions));
            }

            return parts;
        }

        /// <summary>
        /// Gets the entity category (same as FilterEntityDataHelper.GetEntityCategory)
        /// </summary>
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
            else if (entity is Dimension)
            {
                if (entity is AlignedDimension)
                    return "Aligned Dimension";
                else if (entity is RotatedDimension)
                    return "Linear Dimension";
                else if (entity is RadialDimension)
                    return "Radial Dimension";
                else if (entity is DiametricDimension)
                    return "Diametric Dimension";
                else if (entity is OrdinateDimension)
                    return "Ordinate Dimension";
                else if (entity is ArcDimension)
                    return "Arc Dimension";
                else if (entity is RadialDimensionLarge)
                    return "Jogged Dimension";
                else
                    return "Dimension";
            }
            else if (entity is MText)
                return "MText";
            else if (entity is DBText)
                return "Text";
            else if (entity is Polyline)
                return "Polyline";
            else if (entity is Polyline2d)
                return "Polyline2D";
            else if (entity is Polyline3d)
                return "Polyline3D";
            else if (entity is Line)
                return "Line";
            else if (entity is Circle)
                return "Circle";
            else if (entity is Arc)
                return "Arc";
            else if (entity is Ellipse)
                return "Ellipse";
            else if (entity is Spline)
                return "Spline";
            else if (entity is Hatch)
                return "Hatch";
            else if (entity is Solid)
                return "2D Solid";
            else if (entity is Leader)
                return "Leader";
            else if (entity is MLeader)
                return "Multileader";
            else if (entity is Table)
                return "Table";
            else if (entity is Viewport)
                return "Viewport";
            else if (entity is RasterImage)
                return "Raster Image";
            else if (entity is Wipeout)
                return "Wipeout";
            else if (entity is DBPoint)
                return "Point";
            else if (entity is Ray)
                return "Ray";
            else if (entity is Xline)
                return "Construction Line";
            else
            {
                return entity.GetType().Name.Replace("Autodesk.AutoCAD.", "");
            }
        }

        /// <summary>
        /// Escapes special characters in query values
        /// </summary>
        private static string EscapeQueryValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            // Escape double quotes by doubling them
            return value.Replace("\"", "\"\"");
        }
    }
}
