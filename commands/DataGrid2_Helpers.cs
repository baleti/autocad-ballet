using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using System.Globalization;
using System.Text;
using System;

// Alias to avoid naming conflicts
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Performance Caching Fields
    // ──────────────────────────────────────────────────────────────

    // Virtual mode caching
    private static List<Dictionary<string, object>> _cachedOriginalData;
    private static List<Dictionary<string, object>> _cachedFilteredData;
    private static DataGridView _currentGrid;

    // Search index cache
    private static Dictionary<string, Dictionary<int, string>> _searchIndexByColumn;
    private static Dictionary<int, string> _searchIndexAllColumns;

    // Column visibility cache
    private static HashSet<string> _lastVisibleColumns = new HashSet<string>();
    private static string _lastColumnVisibilityFilter = "";

    // Column ordering cache
    private static string _lastColumnOrderingFilter = "";

    // Edit mode state
    private static bool _isEditMode = false;
    private static Dictionary<string, object> _pendingCellEdits = new Dictionary<string, object>();
    private static List<DataGridViewCell> _selectedEditCells = new List<DataGridViewCell>();
    private static HashSet<Dictionary<string, object>> _modifiedEntries = new HashSet<Dictionary<string, object>>();

    // Selection anchor for Shift+Arrow behavior (Excel-like)
    private static DataGridViewCell _selectionAnchor = null;

    // ──────────────────────────────────────────────────────────────
    //  Helper types
    // ──────────────────────────────────────────────────────────────

    private struct ColumnValueFilter
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public string       Value;         // value to look for in the cell
        public bool         IsExclusion;   // true ⇒ "must NOT contain"
        public bool         IsGlobPattern; // true if value contains wildcards
        public bool         IsExactMatch;  // true if value should match exactly
        public bool         IsColumnExactMatch; // true if column should match exactly
    }

    private enum ComparisonOperator
    {
        GreaterThan,
        LessThan
    }

    private struct ComparisonFilter
    {
        public List<string> ColumnParts;       // column-header fragments to match (null = all columns)
        public ComparisonOperator Operator;    // > or <
        public double Value;                   // numeric value to compare against
        public bool IsExclusion;               // true ⇒ "must NOT match comparison"
    }

    /// <summary>
    /// Represents column ordering information
    /// </summary>
    private struct ColumnOrderInfo
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public int Position;               // desired position (1-based)
        public bool IsExactMatch;          // true if column should match exactly
    }

    /// <summary>
    /// Represents a group of filters that use AND logic internally.
    /// Multiple FilterGroups are combined with OR logic.
    /// </summary>
    private class FilterGroup
    {
        public List<List<string>> ColVisibilityFilters { get; set; }
        public List<ColumnValueFilter> ColValueFilters { get; set; }
        public List<string> GeneralFilters { get; set; }
        public List<ComparisonFilter> ComparisonFilters { get; set; }
        public List<string> GeneralGlobPatterns { get; set; } // New field for glob patterns
        public List<ColumnOrderInfo> ColumnOrdering { get; set; } // New field for column ordering
        public List<bool> ColVisibilityExactMatch { get; set; } // Track exact match for visibility
        public List<string> GeneralExactFilters { get; set; } // Exact match general filters

        public FilterGroup()
        {
            ColVisibilityFilters = new List<List<string>>();
            ColValueFilters = new List<ColumnValueFilter>();
            GeneralFilters = new List<string>();
            ComparisonFilters = new List<ComparisonFilter>();
            GeneralGlobPatterns = new List<string>(); // Initialize new field
            ColumnOrdering = new List<ColumnOrderInfo>(); // Initialize column ordering
            ColVisibilityExactMatch = new List<bool>(); // Initialize exact match tracking
            GeneralExactFilters = new List<string>(); // Initialize exact match filters
        }
    }

    /// <summary>
    /// Comparer for List<string> to use in HashSet
    /// </summary>
    private class ListStringComparer : IEqualityComparer<List<string>>
    {
        public bool Equals(List<string> x, List<string> y)
        {
            if (x == null && y == null) return true;
            if (x == null || y == null) return false;
            if (x.Count != y.Count) return false;

            return x.SequenceEqual(y);
        }

        public int GetHashCode(List<string> obj)
        {
            if (obj == null) return 0;

            int hash = 17;
            foreach (string s in obj)
            {
                hash = hash * 31 + (s != null ? s.GetHashCode() : 0);
            }
            return hash;
        }
    }

    private class SortCriteria
    {
        public string ColumnName { get; set; }
        public ListSortDirection Direction { get; set; }
    }

    /// <summary>
    /// Represents a cell edit operation
    /// </summary>
    private class CellEdit
    {
        public int RowIndex { get; set; }
        public string ColumnName { get; set; }
        public object OriginalValue { get; set; }
        public object NewValue { get; set; }
    }

    // ──────────────────────────────────────────────────────────────
    //  Utility Methods
    // ──────────────────────────────────────────────────────────────

    private static string StripQuotes(string s)
    {
        return s.StartsWith("\"") && s.EndsWith("\"") && s.Length > 1
            ? s.Substring(1, s.Length - 2)
            : s;
    }

    /// <summary>Check if a string contains glob wildcards</summary>
    private static bool ContainsGlobWildcards(string pattern)
    {
        return pattern != null && pattern.Contains("*");
    }

    /// <summary>Convert glob pattern to regex pattern</summary>
    private static string GlobToRegexPattern(string globPattern)
    {
        // Escape special regex characters except *
        string escaped = Regex.Escape(globPattern).Replace("\\*", ".*");
        return "^" + escaped + "$";
    }

    /// <summary>Check if a value matches a glob pattern</summary>
    private static bool MatchesGlobPattern(string value, string pattern)
    {
        if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(pattern))
            return false;

        // Convert to lowercase for case-insensitive matching
        value = value.ToLowerInvariant();
        pattern = pattern.ToLowerInvariant();

        // If no wildcards, use simple contains (backward compatibility)
        if (!pattern.Contains("*"))
            return value.Contains(pattern);

        // Convert glob to regex and match
        string regexPattern = GlobToRegexPattern(pattern);
        return Regex.IsMatch(value, regexPattern);
    }

    /// <summary>Build search index for fast filtering</summary>
    private static void BuildSearchIndex(List<Dictionary<string, object>> data, List<string> propertyNames)
    {
        _searchIndexByColumn = new Dictionary<string, Dictionary<int, string>>();
        _searchIndexAllColumns = new Dictionary<int, string>();

        // Initialize column indices
        foreach (string prop in propertyNames)
        {
            _searchIndexByColumn[prop] = new Dictionary<int, string>();
        }

        // Build indices
        for (int i = 0; i < data.Count; i++)
        {
            var entry = data[i];
            var allValuesBuilder = new System.Text.StringBuilder();

            foreach (string prop in propertyNames)
            {
                object value;
                if (entry.TryGetValue(prop, out value) && value != null)
                {
                    string strVal = value.ToString().ToLowerInvariant();
                    _searchIndexByColumn[prop][i] = strVal;
                    allValuesBuilder.Append(strVal).Append(" ");
                }
            }

            _searchIndexAllColumns[i] = allValuesBuilder.ToString();
        }
    }

    // ──────────────────────────────────────────────────────────────
    //  Edit Mode Helper Methods
    // ──────────────────────────────────────────────────────────────

    /// <summary>Reset edit mode state</summary>
    private static void ResetEditMode()
    {
        _isEditMode = false;
        _pendingCellEdits.Clear();
        _selectedEditCells.Clear();
        _modifiedEntries.Clear();
        _selectionAnchor = null;
    }

    /// <summary>Check if a column is editable</summary>
    private static bool IsColumnEditable(string columnName)
    {
        string lowerName = columnName.ToLowerInvariant();

        // Editable columns
        switch (lowerName)
        {
            case "name":           // Entity names, layout names, etc.
            case "contents":       // Text contents of MText, DBText, and dimensions
            case "layer":          // Layer assignment
            case "color":          // Color property
            case "linetype":       // Linetype assignment
            case "layout":         // Layout names (for switch-view command)
            // Plot settings columns for layouts (commonly editable ones)
            case "papersize":      // Paper size
            case "plotstyletable": // Plot style table
            case "plotrotation":   // Drawing orientation
            case "plotconfigurationname": // Plot device/printer
            case "plotscale":      // Plot scale
            case "plottype":       // Plot type
            case "plotcentered":   // Plot centered
            // Geometry properties (position, scale, rotation)
            case "centerx":        // Center X coordinate
            case "centery":        // Center Y coordinate
            case "centerz":        // Center Z coordinate
            case "scalex":         // X scale factor
            case "scaley":         // Y scale factor
            case "scalez":         // Z scale factor
            case "rotation":       // Rotation angle
            case "width":          // Width (for rectangles, text boxes, etc.)
            case "height":         // Height (for rectangles, text boxes, etc.)
            case "radius":         // Radius (for circles)
            case "textheight":     // Text height
            case "widthfactor":    // Width factor for text
                return true;
        }

        // Block attributes are editable (start with "attr_")
        if (lowerName.StartsWith("attr_"))
            return true;

        // XData is editable (start with "xdata_")
        if (lowerName.StartsWith("xdata_"))
            return true;

        // Extension dictionary data is editable (start with "ext_dict_")
        if (lowerName.StartsWith("ext_dict_"))
            return true;

        // Read-only columns
        return false;
    }

    /// <summary>Toggle edit mode on/off</summary>
    private static void ToggleEditMode(DataGridView grid)
    {
        if (_isEditMode)
        {
            // Switch back to row selection mode
            _isEditMode = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = true;
            _selectedEditCells.Clear();

            // Reset column styles to normal
            UpdateColumnEditableStyles(grid, false);
        }
        else
        {
            // Switch to cell selection mode
            _isEditMode = true;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = true;
            grid.ClearSelection();

            // Update column styles to show editable/non-editable
            UpdateColumnEditableStyles(grid, true);
        }
    }

    /// <summary>Update column styles to indicate editable vs non-editable columns</summary>
    private static void UpdateColumnEditableStyles(DataGridView grid, bool editModeActive)
    {
        foreach (DataGridViewColumn column in grid.Columns)
        {
            if (editModeActive)
            {
                if (IsColumnEditable(column.Name))
                {
                    // Editable - normal appearance
                    column.HeaderCell.Style.BackColor = Color.LightGreen;
                    column.HeaderCell.Style.ForeColor = Color.Black;
                    column.DefaultCellStyle.BackColor = Color.White;
                    column.DefaultCellStyle.ForeColor = Color.Black;
                }
                else
                {
                    // Non-editable - greyed out
                    column.HeaderCell.Style.BackColor = Color.LightGray;
                    column.HeaderCell.Style.ForeColor = Color.DarkGray;
                    column.DefaultCellStyle.BackColor = Color.WhiteSmoke;
                    column.DefaultCellStyle.ForeColor = Color.Gray;
                }
            }
            else
            {
                // Reset to default appearance
                column.HeaderCell.Style.BackColor = Color.Empty;
                column.HeaderCell.Style.ForeColor = Color.Empty;
                column.DefaultCellStyle.BackColor = Color.Empty;
                column.DefaultCellStyle.ForeColor = Color.Empty;
            }
        }
    }

    /// <summary>Apply pending edits to the underlying data and track modified entries</summary>
    private static void ApplyPendingEdits()
    {
        foreach (var kvp in _pendingCellEdits)
        {
            string[] parts = kvp.Key.Split('|');
            if (parts.Length == 2)
            {
                int rowIndex = int.Parse(parts[0]);
                string columnName = parts[1];

                if (rowIndex >= 0 && rowIndex < _cachedFilteredData.Count)
                {
                    var entry = _cachedFilteredData[rowIndex];
                    entry[columnName] = kvp.Value;
                    _modifiedEntries.Add(entry);
                }
            }
        }
        _pendingCellEdits.Clear();
    }

    /// <summary>Get a unique key for a cell</summary>
    private static string GetCellKey(int rowIndex, string columnName)
    {
        return $"{rowIndex}|{columnName}";
    }

    /// <summary>Show advanced edit prompt for selected cells with find/replace, patterns, and math operations</summary>
    private static void ShowCellEditPrompt(DataGridView grid)
    {
        if (_selectedEditCells.Count == 0) return;

        // Filter out non-editable cells
        var editableCells = _selectedEditCells.Where(cell =>
            IsColumnEditable(grid.Columns[cell.ColumnIndex].Name)).ToList();

        if (editableCells.Count == 0)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            ed.WriteMessage("\nNo editable cells selected. Editable columns are highlighted in green.");
            return;
        }

        // Use only the editable cells
        var originalSelectedCells = _selectedEditCells.ToList();
        _selectedEditCells.Clear();
        _selectedEditCells.AddRange(editableCells);

        // Collect current values and corresponding data rows for advanced processing
        var currentValues = new List<string>();
        var dataRows = new List<Dictionary<string, object>>();

        foreach (var cell in _selectedEditCells)
        {
            if (cell.RowIndex < _cachedFilteredData.Count)
            {
                var entry = _cachedFilteredData[cell.RowIndex];
                string columnName = grid.Columns[cell.ColumnIndex].Name;
                string currentValue = "";

                if (entry.ContainsKey(columnName) && entry[columnName] != null)
                {
                    currentValue = entry[columnName].ToString();
                }

                currentValues.Add(currentValue);
                dataRows.Add(entry); // Full row data for pattern references
            }
            else
            {
                currentValues.Add("");
                dataRows.Add(new Dictionary<string, object>());
            }
        }

        // Show the advanced edit dialog
        using (var advancedForm = new AutoCADCommands.AdvancedEditDialog(currentValues, dataRows, "Advanced Cell Editor"))
        {
            if (advancedForm.ShowDialog() == WinForms.DialogResult.OK)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                ed.WriteMessage($"\nApplying advanced edits to {_selectedEditCells.Count} cells");

                // Apply transformations to each cell
                for (int i = 0; i < _selectedEditCells.Count; i++)
                {
                    var cell = _selectedEditCells[i];
                    string originalValue = i < currentValues.Count ? currentValues[i] : "";
                    var dataRow = i < dataRows.Count ? dataRows[i] : null;

                    // Apply the transformation using the dialog's built-in transformation
                    string newValue = TransformValue(originalValue, advancedForm, dataRow);

                    if (newValue != originalValue)
                    {
                        string columnName = grid.Columns[cell.ColumnIndex].Name;
                        string cellKey = GetCellKey(cell.RowIndex, columnName);

                        ed.WriteMessage($"\nCell [{cell.RowIndex}, {columnName}]: '{originalValue}' → '{newValue}'");

                        // Store the pending edit
                        _pendingCellEdits[cellKey] = newValue;

                        // Update the display data immediately for visual feedback
                        if (cell.RowIndex < _cachedFilteredData.Count)
                        {
                            var entry = _cachedFilteredData[cell.RowIndex];
                            entry[columnName] = newValue;
                            _modifiedEntries.Add(entry);
                        }
                    }
                }

                // Refresh grid to show changes
                grid.Invalidate();
                ed.WriteMessage($"\nAdvanced edit complete. Total pending edits: {_pendingCellEdits.Count}");
            }
            else
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                ed.WriteMessage("\nAdvanced edit cancelled");
            }
        }

        // Restore original selection
        _selectedEditCells.Clear();
        _selectedEditCells.AddRange(originalSelectedCells);
    }

    /// <summary>Bridge method to use AdvancedEditDialog with corrected precedence logic</summary>
    private static string TransformValue(string originalValue, AutoCADCommands.AdvancedEditDialog dialog, Dictionary<string, object> dataRow)
    {
        // Use the same transformation logic as AdvancedEditDialog.TransformValue
        // Find/Replace and Math operations take precedence over Pattern operations
        string value = originalValue;

        // 1. Pattern transformation (applied first as base)
        if (!string.IsNullOrEmpty(dialog.PatternText))
        {
            value = dialog.PatternText;

            // Replace {} with original value
            value = value.Replace("{}", originalValue);

            // Replace column references if dataRow is available
            if (dataRow != null)
            {
                foreach (var kvp in dataRow)
                {
                    string columnValue = kvp.Value?.ToString() ?? "";
                    // Replace both quoted and unquoted column references
                    value = value.Replace($"$\"{kvp.Key}\"", columnValue);
                    value = value.Replace($"${kvp.Key}", columnValue);
                }
            }
        }

        // 2. Find/Replace transformation (takes precedence - applied to pattern result)
        if (!string.IsNullOrEmpty(dialog.FindText))
        {
            value = value.Replace(dialog.FindText, dialog.ReplaceText ?? "");
        }

        // 3. Math operation (takes precedence - applied to result of pattern + find/replace, if result is numeric)
        if (!string.IsNullOrEmpty(dialog.MathOperationText))
        {
            if (double.TryParse(value, out double numericValue))
            {
                double result = DataRenamerHelper.ApplyMathOperation(numericValue, dialog.MathOperationText);
                value = result.ToString();
            }
        }

        return value;
    }


    /// <summary>Apply a value to all selected cells in edit mode</summary>
    private static void ApplyValueToSelectedCells(DataGridView grid, string newValue)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        foreach (var cell in _selectedEditCells)
        {
            if (cell.RowIndex >= 0 && cell.RowIndex < _cachedFilteredData.Count)
            {
                string columnName = grid.Columns[cell.ColumnIndex].Name;
                string cellKey = GetCellKey(cell.RowIndex, columnName);

                ed.WriteMessage($"\nStoring edit: Row {cell.RowIndex}, Column '{columnName}', Value '{newValue}'");

                // Store the pending edit
                _pendingCellEdits[cellKey] = newValue;

                // Update the display data immediately for visual feedback
                var entry = _cachedFilteredData[cell.RowIndex];
                entry[columnName] = newValue;
                _modifiedEntries.Add(entry);
            }
        }
    }

    /// <summary>Apply pending cell edits to actual AutoCAD entities</summary>
    public static void ApplyCellEditsToEntities()
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        ed.WriteMessage($"\n=== ApplyCellEditsToEntities START ===");
        ed.WriteMessage($"\nPending edits count: {_pendingCellEdits.Count}");

        if (_pendingCellEdits.Count == 0)
        {
            ed.WriteMessage("\nNo pending edits - returning early");
            return;
        }

        try
        {
            // Group edits by document path for cross-document support
            var editsByDocument = GroupEditsByDocument();
            ed.WriteMessage($"\nEdits grouped by {editsByDocument.Count} document(s)");

            int totalApplied = 0;
            int totalProcessed = 0;

            // Process edits for current document first
            string currentDocPath = System.IO.Path.GetFullPath(doc.Name);
            if (editsByDocument.ContainsKey(currentDocPath))
            {
                var currentDocEdits = editsByDocument[currentDocPath];
                ed.WriteMessage($"\nProcessing {currentDocEdits.Count} edits for current document");

                int applied = ApplyEditsToDocument(doc, currentDocEdits);
                totalApplied += applied;
                totalProcessed += currentDocEdits.Count;

                editsByDocument.Remove(currentDocPath);
            }

            // Process edits for external documents
            foreach (var docEdits in editsByDocument)
            {
                string externalDocPath = docEdits.Key;
                var edits = docEdits.Value;

                ed.WriteMessage($"\nProcessing {edits.Count} edits for external document: {System.IO.Path.GetFileName(externalDocPath)}");

                int applied = ApplyEditsToExternalDocument(externalDocPath, edits);
                totalApplied += applied;
                totalProcessed += edits.Count;
            }

            ed.WriteMessage($"\nFinal result: Applied {totalApplied} entity modifications out of {totalProcessed} processed.");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERROR in ApplyCellEditsToEntities: {ex.Message}");
            ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
        }
        finally
        {
            ed.WriteMessage($"\nClearing {_pendingCellEdits.Count} pending edits...");
            _pendingCellEdits.Clear();
            ed.WriteMessage($"\n=== ApplyCellEditsToEntities END ===");
        }
    }

    /// <summary>Group pending edits by document path for cross-document processing</summary>
    private static Dictionary<string, List<PendingEdit>> GroupEditsByDocument()
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        var editsByDocument = new Dictionary<string, List<PendingEdit>>(StringComparer.OrdinalIgnoreCase);

        foreach (var kvp in _pendingCellEdits)
        {
            string[] parts = kvp.Key.Split('|');
            if (parts.Length == 2)
            {
                int rowIndex = int.Parse(parts[0]);
                string columnName = parts[1];
                string newValue = kvp.Value?.ToString() ?? "";

                if (rowIndex >= 0 && rowIndex < _cachedFilteredData.Count)
                {
                    var entry = _cachedFilteredData[rowIndex];

                    // Determine document path for this edit
                    string documentPath;
                    if (entry.TryGetValue("DocumentPath", out var docPathObj))
                    {
                        documentPath = System.IO.Path.GetFullPath(docPathObj.ToString());
                    }
                    else
                    {
                        // Fallback to current document
                        documentPath = System.IO.Path.GetFullPath(doc.Name);
                    }

                    // Create pending edit object
                    var pendingEdit = new PendingEdit
                    {
                        RowIndex = rowIndex,
                        ColumnName = columnName,
                        NewValue = newValue,
                        Entry = entry
                    };

                    // Add to appropriate document group
                    if (!editsByDocument.ContainsKey(documentPath))
                    {
                        editsByDocument[documentPath] = new List<PendingEdit>();
                    }
                    editsByDocument[documentPath].Add(pendingEdit);

                    ed.WriteMessage($"\nGrouped edit for {System.IO.Path.GetFileName(documentPath)}: {columnName} = '{newValue}'");
                }
            }
        }

        return editsByDocument;
    }

    /// <summary>Apply edits to entities in the current document</summary>
    private static int ApplyEditsToDocument(Autodesk.AutoCAD.ApplicationServices.Document doc, List<PendingEdit> edits)
    {
        var ed = doc.Editor;
        var db = doc.Database;
        int appliedCount = 0;

        ed.WriteMessage("\nStarting transaction for current document...");
        using (DocumentLock docLock = doc.LockDocument())
        {
            using (var tr = db.TransactionManager.StartTransaction())
        {
            foreach (var edit in edits)
            {
                try
                {
                    // Get the ObjectId for this entity (may need to reconstruct from Handle for external entities)
                    ObjectId objectId = ObjectId.Null;

                    if (edit.Entry.TryGetValue("ObjectId", out var objIdValue) && objIdValue is ObjectId validObjectId)
                    {
                        objectId = validObjectId;
                    }
                    else if (edit.Entry.TryGetValue("Handle", out var handleValue))
                    {
                        // Try to reconstruct ObjectId from Handle (for external entities)
                        try
                        {
                            var handle = Convert.ToInt64(handleValue.ToString(), 16);
                            objectId = doc.Database.GetObjectId(false, new Autodesk.AutoCAD.DatabaseServices.Handle(handle), 0);
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nFailed to reconstruct ObjectId from Handle {handleValue}: {ex.Message}");
                        }
                    }

                    if (objectId != ObjectId.Null)
                    {
                        ed.WriteMessage($"\nApplying edit: {edit.ColumnName} = '{edit.NewValue}' to ObjectId {objectId}");

                        var dbObject = tr.GetObject(objectId, OpenMode.ForWrite);
                        if (dbObject != null)
                        {
                            ApplyEditToDBObject(dbObject, edit.ColumnName, edit.NewValue, tr);
                            appliedCount++;
                            ed.WriteMessage($"\nEdit applied successfully");
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\nSkipping edit - no valid ObjectId or Handle found");
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError applying edit to {edit.ColumnName}: {ex.Message}");
                    continue;
                }
            }

            ed.WriteMessage($"\nCommitting transaction with {appliedCount} modifications...");
            tr.Commit();
            ed.WriteMessage($"\nTransaction committed successfully!");
            }
        }

        return appliedCount;
    }

    /// <summary>Apply edits to entities in an external document without switching active document</summary>
    private static int ApplyEditsToExternalDocument(string documentPath, List<PendingEdit> edits)
    {
        var docs = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
        var currentDoc = docs.MdiActiveDocument;
        var ed = currentDoc.Editor;

        int appliedCount = 0;

        try
        {
            // Check if the external document is already open
            Autodesk.AutoCAD.ApplicationServices.Document targetDoc = null;

            foreach (Autodesk.AutoCAD.ApplicationServices.Document openDoc in docs)
            {
                if (string.Equals(System.IO.Path.GetFullPath(openDoc.Name), documentPath, StringComparison.OrdinalIgnoreCase))
                {
                    targetDoc = openDoc;
                    break;
                }
            }

            if (targetDoc == null)
            {
                ed.WriteMessage($"\nCannot access external document: {documentPath} (document not open)");
                return 0;
            }

            ed.WriteMessage($"\nProcessing edits in external document: {System.IO.Path.GetFileName(documentPath)}");

            // Apply edits directly to the external document without switching active document
            // Use document locking for reliable operations
            using (var docLock = targetDoc.LockDocument())
            {
                var targetDb = targetDoc.Database;
                var targetEditor = targetDoc.Editor;

                ed.WriteMessage("\nStarting transaction for external document...");
                using (var tr = targetDb.TransactionManager.StartTransaction())
                {
                    foreach (var edit in edits)
                    {
                        try
                        {
                            // Get the ObjectId for this entity in the target document
                            ObjectId objectId = ObjectId.Null;

                            if (edit.Entry.TryGetValue("Handle", out var handleValue))
                            {
                                // Reconstruct ObjectId from Handle in the target document
                                try
                                {
                                    var handle = Convert.ToInt64(handleValue.ToString(), 16);
                                    objectId = targetDb.GetObjectId(false, new Autodesk.AutoCAD.DatabaseServices.Handle(handle), 0);
                                }
                                catch (System.Exception ex)
                                {
                                    ed.WriteMessage($"\nFailed to reconstruct ObjectId from Handle {handleValue} in external document: {ex.Message}");
                                    continue;
                                }
                            }

                            if (objectId != ObjectId.Null)
                            {
                                ed.WriteMessage($"\nApplying edit to external document: {edit.ColumnName} = '{edit.NewValue}' to ObjectId {objectId}");

                                var dbObject = tr.GetObject(objectId, OpenMode.ForWrite);
                                if (dbObject != null)
                                {
                                    ApplyEditToDBObjectInExternalDocument(dbObject, edit.ColumnName, edit.NewValue, tr, targetEditor);
                                    appliedCount++;
                                    ed.WriteMessage($"\nExternal edit applied successfully");
                                }
                            }
                            else
                            {
                                ed.WriteMessage($"\nSkipping external edit - no valid Handle found");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError applying external edit to {edit.ColumnName}: {ex.Message}");
                            continue;
                        }
                    }

                    ed.WriteMessage($"\nCommitting external transaction with {appliedCount} modifications...");
                    tr.Commit();
                    ed.WriteMessage($"\nExternal transaction committed successfully!");
                }
            }

            ed.WriteMessage($"\nCompleted processing {appliedCount} edits in external document");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError processing external document {System.IO.Path.GetFileName(documentPath)}: {ex.Message}");
            ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            return 0;
        }

        return appliedCount;
    }

    /// <summary>Apply a single edit to a DBObject in an external document</summary>
    private static void ApplyEditToDBObjectInExternalDocument(DBObject dbObject, string columnName, string newValue, Transaction tr, Autodesk.AutoCAD.EditorInput.Editor targetEditor)
    {
        var currentDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = currentDoc.Editor;

        ed.WriteMessage($"\n  >> ApplyEditToDBObjectInExternalDocument: DBObject={dbObject.GetType().Name}, Column='{columnName}', Value='{newValue}'");

        try
        {
            switch (columnName.ToLowerInvariant())
            {
                case "contents":
                    ed.WriteMessage($"\n  >> Setting contents/text to '{newValue}' in external document");
                    // Handle contents changes for text entities
                    if (dbObject is MText mtextExt)
                    {
                        ed.WriteMessage($"\n  >> DBObject is MText, setting Contents in external document");
                        mtextExt.Contents = newValue;
                        ed.WriteMessage($"\n  >> MText Contents set successfully in external document");
                    }
                    else if (dbObject is DBText textExt)
                    {
                        ed.WriteMessage($"\n  >> DBObject is DBText, setting TextString in external document");
                        textExt.TextString = newValue;
                        ed.WriteMessage($"\n  >> DBText TextString set successfully in external document");
                    }
                    else if (dbObject is Dimension dimExt)
                    {
                        ed.WriteMessage($"\n  >> DBObject is Dimension, setting DimensionText in external document");
                        dimExt.DimensionText = newValue;
                        ed.WriteMessage($"\n  >> Dimension text set successfully in external document");
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject type not supported for contents editing in external document ({dbObject.GetType().Name})");
                    }
                    break;

                case "layout":
                case "name":
                    ed.WriteMessage($"\n  >> Setting name/text to '{newValue}' in external document");
                    // Handle name changes for different object types
                    if (dbObject is Layout layout)
                    {
                        ed.WriteMessage($"\n  >> DBObject is Layout, attempting to set LayoutName in external document");

                        try
                        {
                            // Check if this is the Model layout (cannot be renamed)
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                            {
                                ed.WriteMessage($"\n  >> Cannot rename Model space layout - this is not allowed in AutoCAD");
                                return; // Skip this edit
                            }

                            // Check if the layout name would conflict with existing layouts
                            var layoutDict = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(dbObject.Database.LayoutDictionaryId, OpenMode.ForRead);
                            if (layoutDict.Contains(newValue))
                            {
                                // Generate unique name if there's a conflict
                                int counter = 1;
                                string uniqueName = newValue;
                                while (layoutDict.Contains(uniqueName))
                                {
                                    uniqueName = $"{newValue}_{counter}";
                                    counter++;
                                }
                                newValue = uniqueName;
                                ed.WriteMessage($"\n  >> Layout name conflict resolved, using: '{uniqueName}'");
                            }

                            layout.LayoutName = newValue;
                            ed.WriteMessage($"\n  >> Layout name set successfully in external document");
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            ed.WriteMessage($"\n  >> AutoCAD Runtime Exception setting external layout name: {acEx.ErrorStatus} - {acEx.Message}");

                            if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotApplicable)
                            {
                                ed.WriteMessage($"\n  >> External layout name change not applicable - may be a special layout like Model space");
                                return; // Skip this edit and continue
                            }

                            throw; // Re-throw other exceptions for debugging
                        }
                    }
                    else if (dbObject is MText mtextExt2)
                    {
                        ed.WriteMessage($"\n  >> DBObject is MText, setting Contents in external document");
                        mtextExt2.Contents = newValue;
                        ed.WriteMessage($"\n  >> MText Contents set successfully in external document");
                    }
                    else if (dbObject is DBText textExt2)
                    {
                        ed.WriteMessage($"\n  >> DBObject is DBText, setting TextString in external document");
                        textExt2.TextString = newValue;
                        ed.WriteMessage($"\n  >> DBText TextString set successfully in external document");
                    }
                    else if (dbObject is BlockReference blockRef)
                    {
                        ed.WriteMessage($"\n  >> DBObject is BlockReference, renaming the block definition in external document");
                        try
                        {
                            // For block references, we need to rename the block definition (BlockTableRecord)
                            var blockTableRecord = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            if (blockTableRecord != null)
                            {
                                string oldName = blockTableRecord.Name;
                                ed.WriteMessage($"\n  >> Renaming block definition from '{oldName}' to '{newValue}' in external document");

                                // Check if the new name would conflict with existing blocks
                                var blockTable = (Autodesk.AutoCAD.DatabaseServices.BlockTable)tr.GetObject(dbObject.Database.BlockTableId, OpenMode.ForRead);
                                if (blockTable.Has(newValue))
                                {
                                    // Generate unique name if there's a conflict
                                    int counter = 1;
                                    string uniqueName = newValue;
                                    while (blockTable.Has(uniqueName))
                                    {
                                        uniqueName = $"{newValue}_{counter}";
                                        counter++;
                                    }
                                    newValue = uniqueName;
                                    ed.WriteMessage($"\n  >> Block name conflict resolved in external document, using: '{uniqueName}'");
                                }

                                blockTableRecord.Name = newValue;
                                ed.WriteMessage($"\n  >> Block definition name set successfully in external document");
                            }
                            else
                            {
                                ed.WriteMessage($"\n  >> Failed to get BlockTableRecord for block reference in external document");
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            ed.WriteMessage($"\n  >> AutoCAD Runtime Exception renaming block in external document: {acEx.ErrorStatus} - {acEx.Message}");
                            throw;
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> Unsupported object type for name editing in external document: {dbObject.GetType().Name}");
                    }
                    break;

                case "layer":
                    if (dbObject is Entity entity)
                    {
                        ed.WriteMessage($"\n  >> Setting layer to '{newValue}' in external document");
                        entity.Layer = newValue;
                        ed.WriteMessage($"\n  >> Layer set successfully in external document");
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot set layer in external document");
                    }
                    break;

                // Geometry properties for external documents
                case "centerx":
                case "centery":
                case "centerz":
                case "scalex":
                case "scaley":
                case "scalez":
                case "rotation":
                case "width":
                case "height":
                case "radius":
                case "textheight":
                case "widthfactor":
                    if (dbObject is Entity externalGeometryEntity)
                    {
                        ed.WriteMessage($"\n  >> Setting geometry property '{columnName}' to '{newValue}' in external document");
                        ApplyGeometryPropertyEdit(externalGeometryEntity, columnName, newValue, tr);
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot edit geometry properties in external document");
                    }
                    break;

                default:
                    // Handle block attributes (columns starting with "attr_")
                    if (columnName.StartsWith("attr_", StringComparison.OrdinalIgnoreCase))
                    {
                        if (dbObject is BlockReference blockRef)
                        {
                            string attributeTag = columnName.Substring(5); // Remove "attr_" prefix
                            ed.WriteMessage($"\n  >> DBObject is BlockReference, setting attribute '{attributeTag}' to '{newValue}' in external document");
                            ed.WriteMessage($"\n  >> Block has {blockRef.AttributeCollection.Count} attributes in external document");

                            bool attributeFound = false;
                            foreach (ObjectId attId in blockRef.AttributeCollection)
                            {
                                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                if (attRef != null)
                                {
                                    ed.WriteMessage($"\n  >> Checking external attribute: '{attRef.Tag}' (current value: '{attRef.TextString}')");
                                    if (string.Equals(attRef.Tag, attributeTag, StringComparison.OrdinalIgnoreCase))
                                    {
                                        ed.WriteMessage($"\n  >> Matched! Changing external attribute '{attRef.Tag}' from '{attRef.TextString}' to '{newValue}'");
                                        attRef.TextString = newValue;
                                        ed.WriteMessage($"\n  >> Block attribute '{attributeTag}' set successfully in external document");
                                        attributeFound = true;
                                        break;
                                    }
                                }
                            }

                            if (!attributeFound)
                            {
                                ed.WriteMessage($"\n  >> Block attribute '{attributeTag}' not found in external document");
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> DBObject is not a BlockReference, cannot set attribute '{columnName}' in external document");
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> Column '{columnName}' not supported for external document editing");
                    }
                    break;
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in ApplyEditToDBObjectInExternalDocument: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
            throw; // Re-throw to be handled by caller
        }
    }

    /// <summary>Data structure for pending edits</summary>
    private class PendingEdit
    {
        public int RowIndex { get; set; }
        public string ColumnName { get; set; }
        public string NewValue { get; set; }
        public Dictionary<string, object> Entry { get; set; }
    }

    /// <summary>Apply a single edit to a DBObject (entity or layout) based on column name</summary>
    private static void ApplyEditToDBObject(DBObject dbObject, string columnName, string newValue, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        ed.WriteMessage($"\n  >> ApplyEditToDBObject: DBObject={dbObject.GetType().Name}, Column='{columnName}', Value='{newValue}'");

        try
        {
            switch (columnName.ToLowerInvariant())
            {
                case "layer":
                    if (dbObject is Entity entity)
                    {
                        ed.WriteMessage($"\n  >> Setting layer to '{newValue}'");
                        entity.Layer = newValue;
                        ed.WriteMessage($"\n  >> Layer set successfully");
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot set layer");
                    }
                    break;

                case "color":
                    if (dbObject is Entity entity2)
                    {
                        ed.WriteMessage($"\n  >> Parsing color '{newValue}'");
                        // Try to parse color - this is simplified
                        if (int.TryParse(newValue, out int colorIndex))
                        {
                            ed.WriteMessage($"\n  >> Setting color index to {colorIndex}");
                            entity2.ColorIndex = colorIndex;
                            ed.WriteMessage($"\n  >> Color set successfully");
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> Failed to parse color '{newValue}' as integer");
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot set color");
                    }
                    break;

                case "linetype":
                    if (dbObject is Entity entity3)
                    {
                        ed.WriteMessage($"\n  >> Setting linetype to '{newValue}'");
                        entity3.Linetype = newValue;
                        ed.WriteMessage($"\n  >> Linetype set successfully");
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot set linetype");
                    }
                    break;

                case "contents":
                    ed.WriteMessage($"\n  >> Setting contents/text to '{newValue}'");
                    // Handle contents changes for text entities
                    if (dbObject is MText mtextContents)
                    {
                        ed.WriteMessage($"\n  >> DBObject is MText, setting Contents");
                        mtextContents.Contents = newValue;
                        ed.WriteMessage($"\n  >> MText Contents set successfully");
                    }
                    else if (dbObject is DBText textContents)
                    {
                        ed.WriteMessage($"\n  >> DBObject is DBText, setting TextString");
                        textContents.TextString = newValue;
                        ed.WriteMessage($"\n  >> DBText TextString set successfully");
                    }
                    else if (dbObject is Dimension dimContents)
                    {
                        ed.WriteMessage($"\n  >> DBObject is Dimension, setting DimensionText");
                        dimContents.DimensionText = newValue;
                        ed.WriteMessage($"\n  >> Dimension text set successfully");
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject type not supported for contents editing ({dbObject.GetType().Name})");
                    }
                    break;

                case "layout":
                case "name":
                    ed.WriteMessage($"\n  >> Setting name/text to '{newValue}'");
                    // Handle name changes for text entities, layouts, and other named objects
                    if (dbObject is MText mtext)
                    {
                        ed.WriteMessage($"\n  >> DBObject is MText, setting Contents");
                        mtext.Contents = newValue;
                        ed.WriteMessage($"\n  >> MText Contents set successfully");
                    }
                    else if (dbObject is DBText text)
                    {
                        ed.WriteMessage($"\n  >> DBObject is DBText, setting TextString");
                        text.TextString = newValue;
                        ed.WriteMessage($"\n  >> DBText TextString set successfully");
                    }
                    else if (dbObject is Layout layout)
                    {
                        ed.WriteMessage($"\n  >> DBObject is Layout, attempting to set LayoutName");

                        try
                        {
                            // Check if this is the Model layout (cannot be renamed)
                            if (string.Equals(layout.LayoutName, "Model", StringComparison.OrdinalIgnoreCase))
                            {
                                ed.WriteMessage($"\n  >> Cannot rename Model space layout - this is not allowed in AutoCAD");
                                return; // Skip this edit
                            }

                            // Check if the layout name would conflict with existing layouts
                            var layoutDict = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(dbObject.Database.LayoutDictionaryId, OpenMode.ForRead);
                            if (layoutDict.Contains(newValue))
                            {
                                // Generate unique name if there's a conflict
                                int counter = 1;
                                string uniqueName = newValue;
                                while (layoutDict.Contains(uniqueName))
                                {
                                    uniqueName = $"{newValue}_{counter}";
                                    counter++;
                                }
                                newValue = uniqueName;
                                ed.WriteMessage($"\n  >> Layout name conflict resolved, using: '{uniqueName}'");
                            }

                            layout.LayoutName = newValue;
                            ed.WriteMessage($"\n  >> Layout name set successfully");
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            ed.WriteMessage($"\n  >> AutoCAD Runtime Exception setting layout name: {acEx.ErrorStatus} - {acEx.Message}");

                            if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotApplicable)
                            {
                                ed.WriteMessage($"\n  >> Layout name change not applicable - may be a special layout like Model space");
                                return; // Skip this edit and continue
                            }

                            throw; // Re-throw other exceptions for debugging
                        }
                    }
                    else if (dbObject is LayerTableRecord layer)
                    {
                        ed.WriteMessage($"\n  >> DBObject is LayerTableRecord, setting Name");
                        layer.Name = newValue;
                        ed.WriteMessage($"\n  >> Layer name set successfully");
                    }
                    else if (dbObject is BlockTableRecord btr)
                    {
                        ed.WriteMessage($"\n  >> DBObject is BlockTableRecord, setting Name");
                        btr.Name = newValue;
                        ed.WriteMessage($"\n  >> Block name set successfully");
                    }
                    else if (dbObject is TextStyleTableRecord textStyle)
                    {
                        ed.WriteMessage($"\n  >> DBObject is TextStyleTableRecord, setting Name");
                        textStyle.Name = newValue;
                        ed.WriteMessage($"\n  >> Text style name set successfully");
                    }
                    else if (dbObject is LinetypeTableRecord linetype)
                    {
                        ed.WriteMessage($"\n  >> DBObject is LinetypeTableRecord, setting Name");
                        linetype.Name = newValue;
                        ed.WriteMessage($"\n  >> Linetype name set successfully");
                    }
                    else if (dbObject is DimStyleTableRecord dimStyle)
                    {
                        ed.WriteMessage($"\n  >> DBObject is DimStyleTableRecord, setting Name");
                        dimStyle.Name = newValue;
                        ed.WriteMessage($"\n  >> Dimension style name set successfully");
                    }
                    else if (dbObject is UcsTableRecord ucs)
                    {
                        ed.WriteMessage($"\n  >> DBObject is UcsTableRecord, setting Name");
                        ucs.Name = newValue;
                        ed.WriteMessage($"\n  >> UCS name set successfully");
                    }
                    else if (dbObject is BlockReference blockRef)
                    {
                        ed.WriteMessage($"\n  >> DBObject is BlockReference, renaming the block definition");
                        try
                        {
                            // For block references, we need to rename the block definition (BlockTableRecord)
                            var blockTableRecord = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            if (blockTableRecord != null)
                            {
                                string oldName = blockTableRecord.Name;
                                ed.WriteMessage($"\n  >> Renaming block definition from '{oldName}' to '{newValue}'");

                                // Check if the new name would conflict with existing blocks
                                var blockTable = (Autodesk.AutoCAD.DatabaseServices.BlockTable)tr.GetObject(dbObject.Database.BlockTableId, OpenMode.ForRead);
                                if (blockTable.Has(newValue))
                                {
                                    // Generate unique name if there's a conflict
                                    int counter = 1;
                                    string uniqueName = newValue;
                                    while (blockTable.Has(uniqueName))
                                    {
                                        uniqueName = $"{newValue}_{counter}";
                                        counter++;
                                    }
                                    newValue = uniqueName;
                                    ed.WriteMessage($"\n  >> Block name conflict resolved, using: '{uniqueName}'");
                                }

                                blockTableRecord.Name = newValue;
                                ed.WriteMessage($"\n  >> Block definition name set successfully");
                            }
                            else
                            {
                                ed.WriteMessage($"\n  >> Failed to get BlockTableRecord for block reference");
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                        {
                            ed.WriteMessage($"\n  >> AutoCAD Runtime Exception renaming block: {acEx.ErrorStatus} - {acEx.Message}");
                            throw;
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject type not supported for name editing ({dbObject.GetType().Name})");
                    }
                    break;

                // Plot settings columns
                case "papersize":
                case "plotstyletable":
                case "plotrotation":
                case "plotconfigurationname":
                case "plotscale":
                case "plottype":
                case "plotcentered":
                    if (dbObject is Layout layoutForPlotSettings)
                    {
                        ed.WriteMessage($"\n  >> Setting plot setting '{columnName}' to '{newValue}' for layout '{layoutForPlotSettings.LayoutName}'");
                        ApplyPlotSettingEdit(layoutForPlotSettings, columnName, newValue, tr);
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject is not a Layout, cannot edit plot settings");
                    }
                    break;

                // Geometry properties
                case "centerx":
                case "centery":
                case "centerz":
                case "scalex":
                case "scaley":
                case "scalez":
                case "rotation":
                case "width":
                case "height":
                case "radius":
                case "textheight":
                case "widthfactor":
                    if (dbObject is Entity geometryEntity)
                    {
                        ed.WriteMessage($"\n  >> Setting geometry property '{columnName}' to '{newValue}'");
                        ApplyGeometryPropertyEdit(geometryEntity, columnName, newValue, tr);
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot edit geometry properties");
                    }
                    break;

                default:
                    // Handle block attributes, xdata, and extension dictionary data
                    if (columnName.StartsWith("attr_"))
                    {
                        ed.WriteMessage($"\n  >> Processing block attribute: '{columnName}'");
                        // Block attribute
                        if (dbObject is Entity entity4)
                        {
                            ApplyBlockAttributeEdit(entity4, columnName, newValue, tr);
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot edit block attributes");
                        }
                    }
                    else if (columnName.StartsWith("xdata_"))
                    {
                        ed.WriteMessage($"\n  >> Processing XData: '{columnName}'");
                        // XData
                        if (dbObject is Entity entity5)
                        {
                            ApplyXDataEdit(entity5, columnName, newValue);
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot edit XData");
                        }
                    }
                    else if (columnName.StartsWith("ext_dict_"))
                    {
                        ed.WriteMessage($"\n  >> Processing extension dictionary: '{columnName}'");
                        // Extension dictionary
                        if (dbObject is Entity entity6)
                        {
                            ApplyExtensionDictEdit(entity6, columnName, newValue, tr);
                        }
                        else
                        {
                            ed.WriteMessage($"\n  >> DBObject is not an Entity, cannot edit extension dictionary");
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\n  >> Unknown column type: '{columnName}' - skipping");
                    }
                    break;
            }
            ed.WriteMessage($"\n  >> ApplyEditToDBObject completed successfully");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> ERROR in ApplyEditToDBObject: {ex.Message}");
            ed.WriteMessage($"\n  >> Stack trace: {ex.StackTrace}");
            throw; // Re-throw so parent can catch it
        }
    }

    /// <summary>Apply edit to block attribute</summary>
    private static void ApplyBlockAttributeEdit(Entity entity, string columnName, string newValue, Transaction tr)
    {
        if (entity is BlockReference blockRef)
        {
            string attributeTag = columnName.Substring(5); // Remove "attr_" prefix, keep original case
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            ed.WriteMessage($"\nTrying to edit attribute: columnName='{columnName}', attributeTag='{attributeTag}', newValue='{newValue}'");
            ed.WriteMessage($"\nBlock has {blockRef.AttributeCollection.Count} attributes");

            bool found = false;
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef != null)
                {
                    ed.WriteMessage($"\nChecking attribute: '{attRef.Tag}' (current value: '{attRef.TextString}')");
                    if (attRef.Tag.ToLowerInvariant() == attributeTag.ToLowerInvariant())
                    {
                        ed.WriteMessage($"\nMatched! Changing '{attRef.Tag}' from '{attRef.TextString}' to '{newValue}'");
                        attRef.TextString = newValue;
                        found = true;
                        break;
                    }
                }
            }

            if (!found)
            {
                ed.WriteMessage($"\nAttribute '{attributeTag}' not found in block!");
            }
        }
    }

    /// <summary>Apply edit to XData</summary>
    private static void ApplyXDataEdit(Entity entity, string columnName, string newValue)
    {
        // This is simplified - a full implementation would handle different XData types
        string appName = columnName.Substring(6); // Remove "xdata_" prefix

        var rb = new ResultBuffer(
            new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, newValue)
        );
        entity.XData = rb;
        rb.Dispose();
    }

    /// <summary>Apply edit to extension dictionary</summary>
    private static void ApplyExtensionDictEdit(Entity entity, string columnName, string newValue, Transaction tr)
    {
        // Create extension dictionary if it doesn't exist
        if (entity.ExtensionDictionary == ObjectId.Null)
        {
            entity.CreateExtensionDictionary();
        }

        string key = columnName.Substring(9); // Remove "ext_dict_" prefix
        var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;

        // This is simplified - a full implementation would handle different dictionary entry types
        var xrec = new Xrecord();
        xrec.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, newValue));

        if (extDict.Contains(key))
        {
            extDict.SetAt(key, xrec);
        }
        else
        {
            extDict.SetAt(key, xrec);
        }

        tr.AddNewlyCreatedDBObject(xrec, true);
    }

    /// <summary>Apply edit to plot settings for Layout entities</summary>
    private static void ApplyPlotSettingEdit(Layout layout, string columnName, string newValue, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            // Get the PlotSettings object from the Layout
            var plotSettings = tr.GetObject(layout.ObjectId, OpenMode.ForWrite) as PlotSettings;
            if (plotSettings == null)
            {
                ed.WriteMessage($"\n  >> Failed to get PlotSettings from layout");
                return;
            }

            switch (columnName.ToLowerInvariant())
            {
                case "papersize":
                    ed.WriteMessage($"\n  >> Paper size editing requires careful validation with plot devices");
                    ed.WriteMessage($"\n  >> For safe paper size changes, use AutoCAD's Page Setup Manager");
                    ed.WriteMessage($"\n  >> API method: SetPlotConfiguration() with proper device validation");
                    break;

                case "plotstyletable":
                    ed.WriteMessage($"\n  >> Plot style table editing requires validation of available CTB/STB files");
                    ed.WriteMessage($"\n  >> For safe plot style changes, use AutoCAD's Page Setup Manager");
                    ed.WriteMessage($"\n  >> API method: SetPlotConfiguration() with style table validation");
                    break;

                case "plotrotation":
                    ed.WriteMessage($"\n  >> Plot rotation editing requires recreation of the full plot configuration");
                    ed.WriteMessage($"\n  >> For rotation changes, use AutoCAD's Page Setup Manager");
                    ed.WriteMessage($"\n  >> API method: Copy entire PlotSettings with new rotation");
                    break;

                case "plotconfigurationname":
                    ed.WriteMessage($"\n  >> Plot device editing requires validation of available system printers/plotters");
                    ed.WriteMessage($"\n  >> For safe device changes, use AutoCAD's Page Setup Manager");
                    ed.WriteMessage($"\n  >> API method: SetPlotConfiguration() with device validation");
                    break;

                case "plotscale":
                    ed.WriteMessage($"\n  >> Plot scale can potentially be edited, but requires careful handling");
                    ed.WriteMessage($"\n  >> For scale changes, use AutoCAD's Page Setup Manager");
                    ed.WriteMessage($"\n  >> API method: SetCustomPrintScale() or standard scale methods");
                    break;

                case "plottype":
                    ed.WriteMessage($"\n  >> Plot type editing requires validation of plot area settings");
                    ed.WriteMessage($"\n  >> For plot type changes, use AutoCAD's Page Setup Manager");
                    ed.WriteMessage($"\n  >> API method: SetPlotType() with proper window/view validation");
                    break;

                case "plotcentered":
                    ed.WriteMessage($"\n  >> Plot centered setting can potentially be edited");
                    ed.WriteMessage($"\n  >> For centering changes, use AutoCAD's Page Setup Manager");
                    ed.WriteMessage($"\n  >> API method: SetPlotCentered()");
                    break;

                default:
                    ed.WriteMessage($"\n  >> Unknown plot setting: '{columnName}'");
                    break;
            }

            ed.WriteMessage($"\n  >> Plot settings are complex and require proper validation");
            ed.WriteMessage($"\n  >> The columns are displayed for viewing and comparison purposes");
            ed.WriteMessage($"\n  >> Use AutoCAD's built-in Page Setup Manager for safe editing");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> Error in plot setting edit handler: {ex.Message}");
            throw;
        }
    }

    /// <summary>Apply edit to geometry properties for Entity objects</summary>
    private static void ApplyGeometryPropertyEdit(Entity entity, string columnName, string newValue, Transaction tr)
    {
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            // Parse the numeric value
            if (!double.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
            {
                ed.WriteMessage($"\n  >> Failed to parse '{newValue}' as a number for geometry property '{columnName}'");
                return;
            }

            string lowerColumnName = columnName.ToLowerInvariant();

            // Apply geometry property edits based on entity type and property
            if (entity is Circle circle)
            {
                ApplyCircleGeometryEdit(circle, lowerColumnName, numericValue);
            }
            else if (entity is Arc arc)
            {
                ApplyArcGeometryEdit(arc, lowerColumnName, numericValue);
            }
            else if (entity is Line line)
            {
                ApplyLineGeometryEdit(line, lowerColumnName, numericValue);
            }
            else if (entity is Polyline polyline)
            {
                ApplyPolylineGeometryEdit(polyline, lowerColumnName, numericValue);
            }
            else if (entity is Ellipse ellipse)
            {
                ApplyEllipseGeometryEdit(ellipse, lowerColumnName, numericValue);
            }
            else if (entity is BlockReference blockRef)
            {
                ApplyBlockReferenceGeometryEdit(blockRef, lowerColumnName, numericValue);
            }
            else if (entity is DBText dbText)
            {
                ApplyDBTextGeometryEdit(dbText, lowerColumnName, numericValue);
            }
            else if (entity is MText mText)
            {
                ApplyMTextGeometryEdit(mText, lowerColumnName, numericValue);
            }
            else
            {
                ed.WriteMessage($"\n  >> Geometry property editing not implemented for entity type: {entity.GetType().Name}");
            }

            ed.WriteMessage($"\n  >> Geometry property '{columnName}' set to {numericValue} successfully");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n  >> Error setting geometry property '{columnName}' to '{newValue}': {ex.Message}");
            throw;
        }
    }

    /// <summary>Apply geometry edits to Circle entities</summary>
    private static void ApplyCircleGeometryEdit(Circle circle, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
                circle.Center = new Point3d(value, circle.Center.Y, circle.Center.Z);
                break;
            case "centery":
                circle.Center = new Point3d(circle.Center.X, value, circle.Center.Z);
                break;
            case "centerz":
                circle.Center = new Point3d(circle.Center.X, circle.Center.Y, value);
                break;
            case "radius":
                if (value > 0) circle.Radius = value;
                break;
        }
    }

    /// <summary>Apply geometry edits to Arc entities</summary>
    private static void ApplyArcGeometryEdit(Arc arc, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
                arc.Center = new Point3d(value, arc.Center.Y, arc.Center.Z);
                break;
            case "centery":
                arc.Center = new Point3d(arc.Center.X, value, arc.Center.Z);
                break;
            case "centerz":
                arc.Center = new Point3d(arc.Center.X, arc.Center.Y, value);
                break;
            case "radius":
                if (value > 0) arc.Radius = value;
                break;
        }
    }

    /// <summary>Apply geometry edits to Line entities</summary>
    private static void ApplyLineGeometryEdit(Line line, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
                var currentCenter = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var offset = new Vector3d(value - currentCenter.X, 0, 0);
                line.StartPoint = line.StartPoint + offset;
                line.EndPoint = line.EndPoint + offset;
                break;
            case "centery":
                var currentCenterY = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var offsetY = new Vector3d(0, value - currentCenterY.Y, 0);
                line.StartPoint = line.StartPoint + offsetY;
                line.EndPoint = line.EndPoint + offsetY;
                break;
            case "centerz":
                var currentCenterZ = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var offsetZ = new Vector3d(0, 0, value - currentCenterZ.Z);
                line.StartPoint = line.StartPoint + offsetZ;
                line.EndPoint = line.EndPoint + offsetZ;
                break;
            case "rotation":
                // Convert degrees to radians for rotation
                var angleRadians = value * Math.PI / 180.0;
                var center = line.StartPoint + (line.EndPoint - line.StartPoint) * 0.5;
                var transform = Matrix3d.Rotation(angleRadians, Vector3d.ZAxis, center);
                line.TransformBy(transform);
                break;
        }
    }

    /// <summary>Apply geometry edits to Polyline entities</summary>
    private static void ApplyPolylineGeometryEdit(Polyline polyline, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
            case "centery":
            case "centerz":
                var bounds = polyline.GeometricExtents;
                var currentCenter = bounds.MinPoint + (bounds.MaxPoint - bounds.MinPoint) * 0.5;
                Vector3d offset = new Vector3d(0, 0, 0);

                if (columnName == "centerx")
                    offset = new Vector3d(value - currentCenter.X, 0, 0);
                else if (columnName == "centery")
                    offset = new Vector3d(0, value - currentCenter.Y, 0);
                else if (columnName == "centerz")
                    offset = new Vector3d(0, 0, value - currentCenter.Z);

                polyline.TransformBy(Matrix3d.Displacement(offset));
                break;
        }
    }

    /// <summary>Apply geometry edits to Ellipse entities</summary>
    private static void ApplyEllipseGeometryEdit(Ellipse ellipse, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
                ellipse.Center = new Point3d(value, ellipse.Center.Y, ellipse.Center.Z);
                break;
            case "centery":
                ellipse.Center = new Point3d(ellipse.Center.X, value, ellipse.Center.Z);
                break;
            case "centerz":
                ellipse.Center = new Point3d(ellipse.Center.X, ellipse.Center.Y, value);
                break;
        }
    }

    /// <summary>Apply geometry edits to BlockReference entities</summary>
    private static void ApplyBlockReferenceGeometryEdit(BlockReference blockRef, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
                blockRef.Position = new Point3d(value, blockRef.Position.Y, blockRef.Position.Z);
                break;
            case "centery":
                blockRef.Position = new Point3d(blockRef.Position.X, value, blockRef.Position.Z);
                break;
            case "centerz":
                blockRef.Position = new Point3d(blockRef.Position.X, blockRef.Position.Y, value);
                break;
            case "scalex":
                if (value > 0)
                {
                    var newScale = new Scale3d(value, blockRef.ScaleFactors.Y, blockRef.ScaleFactors.Z);
                    blockRef.ScaleFactors = newScale;
                }
                break;
            case "scaley":
                if (value > 0)
                {
                    var newScale = new Scale3d(blockRef.ScaleFactors.X, value, blockRef.ScaleFactors.Z);
                    blockRef.ScaleFactors = newScale;
                }
                break;
            case "scalez":
                if (value > 0)
                {
                    var newScale = new Scale3d(blockRef.ScaleFactors.X, blockRef.ScaleFactors.Y, value);
                    blockRef.ScaleFactors = newScale;
                }
                break;
            case "rotation":
                // Convert degrees to radians
                blockRef.Rotation = value * Math.PI / 180.0;
                break;
        }
    }

    /// <summary>Apply geometry edits to DBText entities</summary>
    private static void ApplyDBTextGeometryEdit(DBText dbText, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
                dbText.Position = new Point3d(value, dbText.Position.Y, dbText.Position.Z);
                break;
            case "centery":
                dbText.Position = new Point3d(dbText.Position.X, value, dbText.Position.Z);
                break;
            case "centerz":
                dbText.Position = new Point3d(dbText.Position.X, dbText.Position.Y, value);
                break;
            case "textheight":
                if (value > 0) dbText.Height = value;
                break;
            case "widthfactor":
                if (value > 0) dbText.WidthFactor = value;
                break;
            case "rotation":
                // Convert degrees to radians
                dbText.Rotation = value * Math.PI / 180.0;
                break;
        }
    }

    /// <summary>Apply geometry edits to MText entities</summary>
    private static void ApplyMTextGeometryEdit(MText mText, string columnName, double value)
    {
        switch (columnName)
        {
            case "centerx":
                mText.Location = new Point3d(value, mText.Location.Y, mText.Location.Z);
                break;
            case "centery":
                mText.Location = new Point3d(mText.Location.X, value, mText.Location.Z);
                break;
            case "centerz":
                mText.Location = new Point3d(mText.Location.X, mText.Location.Y, value);
                break;
            case "textheight":
                if (value > 0) mText.TextHeight = value;
                break;
            case "width":
                if (value > 0) mText.Width = value;
                break;
            case "rotation":
                // Convert degrees to radians
                mText.Rotation = value * Math.PI / 180.0;
                break;
        }
    }

    // ──────────────────────────────────────────────────────────────
    //  Excel-like Selection and Navigation Helper Methods
    // ──────────────────────────────────────────────────────────────

    /// <summary>Select entire rows of currently selected cells (Excel Shift+Space behavior)</summary>
    private static void SelectRowsOfSelectedCells(DataGridView grid)
    {
        if (_selectedEditCells.Count == 0) return;

        // Get unique row indices from selected cells
        var rowIndices = _selectedEditCells.Select(cell => cell.RowIndex).Distinct().ToList();

        // Clear current selection and select entire rows (only editable columns)
        grid.ClearSelection();

        foreach (int rowIndex in rowIndices)
        {
            if (rowIndex >= 0 && rowIndex < grid.Rows.Count)
            {
                foreach (DataGridViewCell cell in grid.Rows[rowIndex].Cells)
                {
                    if (cell.Visible && IsColumnEditable(grid.Columns[cell.ColumnIndex].Name))
                    {
                        cell.Selected = true;
                    }
                }
            }
        }
    }

    /// <summary>Select entire columns of currently selected cells (Excel Ctrl+Space behavior)</summary>
    private static void SelectColumnsOfSelectedCells(DataGridView grid)
    {
        if (_selectedEditCells.Count == 0) return;

        // Get unique column indices from selected cells (only editable ones)
        var columnIndices = _selectedEditCells.Select(cell => cell.ColumnIndex).Distinct()
            .Where(colIndex => IsColumnEditable(grid.Columns[colIndex].Name)).ToList();

        if (columnIndices.Count == 0) return;

        // Clear current selection and select entire columns
        grid.ClearSelection();

        foreach (int columnIndex in columnIndices)
        {
            if (columnIndex >= 0 && columnIndex < grid.Columns.Count)
            {
                for (int rowIndex = 0; rowIndex < grid.Rows.Count; rowIndex++)
                {
                    if (rowIndex < grid.Rows.Count && grid.Rows[rowIndex].Cells[columnIndex].Visible)
                    {
                        grid.Rows[rowIndex].Cells[columnIndex].Selected = true;
                    }
                }
            }
        }
    }

    /// <summary>Move current cell with arrow keys in edit mode, skipping non-editable columns</summary>
    private static void MoveCellWithArrows(DataGridView grid, Keys keyCode)
    {
        if (grid.CurrentCell == null) return;

        int currentRow = grid.CurrentCell.RowIndex;
        int currentCol = grid.CurrentCell.ColumnIndex;
        int newRow = currentRow;
        int newCol = currentCol;

        switch (keyCode)
        {
            case Keys.Up:
                newRow = Math.Max(0, currentRow - 1);
                break;
            case Keys.Down:
                newRow = Math.Min(grid.Rows.Count - 1, currentRow + 1);
                break;
            case Keys.Left:
                newCol = FindNextEditableColumn(grid, currentCol, -1);
                break;
            case Keys.Right:
                newCol = FindNextEditableColumn(grid, currentCol, 1);
                break;
        }

        if (newRow != currentRow || newCol != currentCol)
        {
            // Clear current selection and move to new cell
            grid.ClearSelection();
            try
            {
                grid.CurrentCell = grid.Rows[newRow].Cells[newCol];
                grid.Rows[newRow].Cells[newCol].Selected = true;
                // Set the new anchor point for future Shift+Arrow operations
                _selectionAnchor = grid.Rows[newRow].Cells[newCol];
            }
            catch (System.Exception)
            {
                // Ignore errors in virtual mode
            }
        }
    }

    /// <summary>Extend selection with Ctrl+Arrow or Shift+Arrow keys in edit mode</summary>
    private static void ExtendSelectionWithArrows(DataGridView grid, Keys keyCode, bool isShiftKey)
    {
        if (grid.CurrentCell == null) return;

        int currentRow = grid.CurrentCell.RowIndex;
        int currentCol = grid.CurrentCell.ColumnIndex;
        int newRow = currentRow;
        int newCol = currentCol;

        switch (keyCode)
        {
            case Keys.Up:
                newRow = Math.Max(0, currentRow - 1);
                break;
            case Keys.Down:
                newRow = Math.Min(grid.Rows.Count - 1, currentRow + 1);
                break;
            case Keys.Left:
                newCol = FindNextEditableColumn(grid, currentCol, -1);
                break;
            case Keys.Right:
                newCol = FindNextEditableColumn(grid, currentCol, 1);
                break;
        }

        if (newRow != currentRow || newCol != currentCol)
        {
            try
            {
                if (isShiftKey)
                {
                    // Shift+Arrow: Excel-like behavior - extend selection from anchor point
                    if (_selectionAnchor == null)
                    {
                        // Set current cell as anchor if none exists
                        _selectionAnchor = grid.CurrentCell;
                    }

                    // Clear current selection
                    grid.ClearSelection();

                    // Select all cells in the rectangular area from anchor to new position
                    SelectRectangularArea(grid, _selectionAnchor.RowIndex, _selectionAnchor.ColumnIndex,
                                        newRow, newCol);

                    // Set current cell to new position
                    grid.CurrentCell = grid.Rows[newRow].Cells[newCol];
                }
                else
                {
                    // Ctrl+Arrow: Original behavior - add to selection
                    grid.Rows[newRow].Cells[newCol].Selected = true;
                    grid.CurrentCell = grid.Rows[newRow].Cells[newCol];
                }
            }
            catch (System.Exception)
            {
                // Ignore errors in virtual mode
            }
        }
    }

    /// <summary>Find the next editable column in the specified direction</summary>
    private static int FindNextEditableColumn(DataGridView grid, int currentCol, int direction)
    {
        int newCol = currentCol + direction;

        // Keep moving in the specified direction until we find an editable column or hit the boundary
        while (newCol >= 0 && newCol < grid.Columns.Count)
        {
            string columnName = grid.Columns[newCol].Name;
            if (IsColumnEditable(columnName))
            {
                return newCol;
            }
            newCol += direction;
        }

        // If we can't find an editable column in that direction, stay in current column
        return currentCol;
    }

    /// <summary>Select all cells in a rectangular area (Excel-like selection)</summary>
    private static void SelectRectangularArea(DataGridView grid, int anchorRow, int anchorCol, int targetRow, int targetCol)
    {
        // Determine the bounds of the selection rectangle
        int startRow = Math.Min(anchorRow, targetRow);
        int endRow = Math.Max(anchorRow, targetRow);
        int startCol = Math.Min(anchorCol, targetCol);
        int endCol = Math.Max(anchorCol, targetCol);

        // Select all cells in the rectangle, but only in editable columns
        for (int row = startRow; row <= endRow; row++)
        {
            if (row >= 0 && row < grid.Rows.Count)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    if (col >= 0 && col < grid.Columns.Count)
                    {
                        string columnName = grid.Columns[col].Name;
                        if (IsColumnEditable(columnName))
                        {
                            try
                            {
                                grid.Rows[row].Cells[col].Selected = true;
                            }
                            catch (System.Exception)
                            {
                                // Ignore errors in virtual mode
                            }
                        }
                    }
                }
            }
        }
    }

    /// <summary>Set the selection anchor for Shift+Arrow operations</summary>
    public static void SetSelectionAnchor(DataGridViewCell cell)
    {
        _selectionAnchor = cell;
    }

    // ──────────────────────────────────────────────────────────────
    //  Integrated Rename Data Helper Classes (from rename-data-of-selection)
    // ──────────────────────────────────────────────────────────────

    /// <summary>Helper class for math operations and data transformations</summary>
    public static class DataRenamerHelper
    {
        #region Math helpers ----------------------------------------------------------------------

        /// <summary>Interprets <paramref name="mathOp"/> ("2x", "x/2", "x+1", …) and applies it to <paramref name="x"/>.</summary>
        public static double ApplyMathOperation(double x, string mathOp)
        {
            if (string.IsNullOrWhiteSpace(mathOp))
                return x;

            mathOp = mathOp.Replace(" ", string.Empty);

            // identity / negate
            if (mathOp.Equals("x", StringComparison.OrdinalIgnoreCase)) return x;
            if (mathOp.Equals("-x", StringComparison.OrdinalIgnoreCase)) return -x;

            // "2x" (multiplier before x)
            if (!mathOp.StartsWith("x", StringComparison.OrdinalIgnoreCase) &&
                 mathOp.EndsWith("x", StringComparison.OrdinalIgnoreCase))
            {
                string multStr = mathOp.Substring(0, mathOp.Length - 1);
                if (double.TryParse(multStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double mult))
                    return mult * x;
            }

            // operations after x
            if (mathOp.StartsWith("x", StringComparison.OrdinalIgnoreCase) && mathOp.Length >= 3)
            {
                char op = mathOp[1];
                string num = mathOp.Substring(2);
                if (double.TryParse(num, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                {
                    switch (op)
                    {
                        case '+': return x + n;
                        case '-': return x - n;
                        case '*': return x * n;
                        case '/': return Math.Abs(n) < double.Epsilon ? x : x / n;
                    }
                }
            }
            return x;   // unrecognised -> unchanged
        }

        /// <summary>Applies <see cref="ApplyMathOperation"/> to every numeric token in <paramref name="input"/>.</summary>
        public static string ApplyMathToNumbersInString(string input, string mathOp)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrWhiteSpace(mathOp))
                return input;

            return Regex.Replace(
                input,
                @"-?\d+(?:\.\d+)?",                              // signed integers/decimals
                m =>
                {
                    if (!double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                        return m.Value;                           // shouldn't happen

                    double res = ApplyMathOperation(n, mathOp);

                    // keep integer look-and-feel where possible
                    return Math.Abs(res % 1) < 1e-10
                         ? Math.Round(res).ToString(CultureInfo.InvariantCulture)
                         : res.ToString(CultureInfo.InvariantCulture);
                });
        }

        #endregion

        #region Pattern parsing helper ------------------------------------------------------------

        /// <summary>
        /// Parses pattern string and replaces data references.
        /// Supports: {} for current value, $"DataName" or $DataName for other data.
        /// Also supports pseudo data on entities: "Layer", "Color", "LineType", "Handle".
        /// </summary>
        public static string ParsePatternWithDataReferences(string pattern, string currentValue, Dictionary<string, object> dataRow)
        {
            if (string.IsNullOrEmpty(pattern))
                return currentValue;

            // First replace {} with current value
            string result = pattern.Replace("{}", currentValue);

            // Regex to match $"Data Name" or $DataNameWithoutSpaces
            var regex = new Regex(@"\$""([^""]+)""|(?<!\$)\$(\w+)");

            result = regex.Replace(result, match =>
            {
                string dataName = !string.IsNullOrEmpty(match.Groups[1].Value)
                    ? match.Groups[1].Value
                    : match.Groups[2].Value;

                // Try getting data value from the row
                string dataValue = GetDataValueFromRow(dataRow, dataName);
                return !string.IsNullOrEmpty(dataValue) ? dataValue : match.Value;
            });

            return result;
        }

        /// <summary>Gets data value from the DataGrid row</summary>
        private static string GetDataValueFromRow(Dictionary<string, object> dataRow, string dataName)
        {
            if (dataRow == null || string.IsNullOrEmpty(dataName))
                return string.Empty;

            // Try exact match first
            if (dataRow.TryGetValue(dataName, out object value) && value != null)
                return value.ToString();

            // Try case-insensitive match
            foreach (var kvp in dataRow)
            {
                if (string.Equals(kvp.Key, dataName, StringComparison.OrdinalIgnoreCase))
                    return kvp.Value?.ToString() ?? string.Empty;
            }

            return string.Empty;
        }

        #endregion

        #region Value transformation --------------------------------------------------------------

        /// <summary>Transform a value using the rename form settings</summary>
        public static string TransformValue(string original, AdvancedRenameForm form, Dictionary<string, object> dataRow = null)
        {
            string value = original;

            // 1. Find / Replace
            if (!string.IsNullOrEmpty(form.FindText))
                value = value.Replace(form.FindText, form.ReplaceText);
            else if (!string.IsNullOrEmpty(form.ReplaceText))
                value = form.ReplaceText;

            // 2. Pattern (with data reference support)
            if (!string.IsNullOrEmpty(form.PatternText))
            {
                if (dataRow != null)
                {
                    value = ParsePatternWithDataReferences(form.PatternText, value, dataRow);
                }
                else
                {
                    // Fallback for preview mode (no dataRow available)
                    value = form.PatternText.Replace("{}", value);
                }
            }

            // 3. Math
            if (!string.IsNullOrEmpty(form.MathOperationText))
            {
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                    value = ApplyMathOperation(n, form.MathOperationText).ToString(CultureInfo.InvariantCulture);
                else
                    value = ApplyMathToNumbersInString(value, form.MathOperationText);
            }
            return value;
        }

        #endregion
    }

    /// <summary>Advanced rename form with find/replace, patterns, math operations, and live preview</summary>
    public class AdvancedRenameForm : WinForms.Form
    {
        private readonly List<string> _originalValues;
        private readonly List<Dictionary<string, object>> _dataRows;

        private WinForms.TextBox _txtFind;
        private WinForms.TextBox _txtReplace;
        private WinForms.TextBox _txtPattern;
        private WinForms.TextBox _txtMath;
        private WinForms.RichTextBox _rtbBefore;
        private WinForms.RichTextBox _rtbAfter;

        #region Exposed properties

        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string PatternText => _txtPattern.Text;
        public string MathOperationText => _txtMath.Text;

        #endregion

        public AdvancedRenameForm(List<string> originalValues, List<Dictionary<string, object>> dataRows = null)
        {
            _originalValues = originalValues ?? new List<string>();
            _dataRows = dataRows ?? new List<Dictionary<string, object>>();
            BuildUI();
            LoadCurrentValues();
            InitializePatternValue();
        }

        private void BuildUI()
        {
            Text = "Advanced Cell Editor";
            Font = new Drawing.Font("Segoe UI", 9);
            MinimumSize = new Drawing.Size(520, 480);
            Size = new Drawing.Size(640, 560);
            KeyPreview = true;
            KeyDown += (s, e) => { if (e.KeyCode == WinForms.Keys.Escape) Close(); };

            // === layout ========================================================================
            var grid = new WinForms.TableLayoutPanel
            {
                Dock = WinForms.DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 8,
                Padding = new WinForms.Padding(8)
            };
            grid.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 90));
            grid.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100));
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 0: Pattern
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 22)); // Row 1: Pattern hint
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 2: Find
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 36)); // Row 3: Replace (increased for spacing)
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 4: Math
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 22)); // Row 5: Math hint
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));

            // Pattern (moved to top)
            grid.Controls.Add(MakeLabel("Pattern:"), 0, 0);
            _txtPattern = MakeTextBox("{}");   // default
            grid.Controls.Add(_txtPattern, 1, 0);

            grid.Controls.Add(MakeHint("Use {} for current value. Use $\"Column Name\" or $ColumnName to reference other columns."), 1, 1);

            // Find / Replace
            grid.Controls.Add(MakeLabel("Find:"), 0, 2);
            _txtFind = MakeTextBox();
            grid.Controls.Add(_txtFind, 1, 2);

            grid.Controls.Add(MakeLabel("Replace:"), 0, 3);
            _txtReplace = MakeTextBox();
            grid.Controls.Add(_txtReplace, 1, 3);

            // Math
            grid.Controls.Add(MakeLabel("Math:"), 0, 4);
            _txtMath = MakeTextBox();
            grid.Controls.Add(_txtMath, 1, 4);

            grid.Controls.Add(MakeHint("Use x to represent current value (e.g. 2x, x/2, x+3, -x)."), 1, 5);

            // Before / After preview
            _rtbBefore = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            var grpBefore = new WinForms.GroupBox { Text = "Current Values", Dock = WinForms.DockStyle.Fill };
            grpBefore.Controls.Add(_rtbBefore);
            grid.Controls.Add(grpBefore, 0, 6);
            grid.SetColumnSpan(grpBefore, 2);

            _rtbAfter = new WinForms.RichTextBox { ReadOnly = true, Dock = WinForms.DockStyle.Fill };
            var grpAfter = new WinForms.GroupBox { Text = "Preview", Dock = WinForms.DockStyle.Fill };
            grpAfter.Controls.Add(_rtbAfter);
            grid.Controls.Add(grpAfter, 0, 7);
            grid.SetColumnSpan(grpAfter, 2);

            // buttons
            var buttons = new WinForms.FlowLayoutPanel
            {
                FlowDirection = WinForms.FlowDirection.RightToLeft,
                Dock = WinForms.DockStyle.Bottom,
                Padding = new WinForms.Padding(8)
            };

            var btnOK = new WinForms.Button { Text = "OK", DialogResult = WinForms.DialogResult.OK };
            var btnCancel = new WinForms.Button { Text = "Cancel", DialogResult = WinForms.DialogResult.Cancel };
            buttons.Controls.Add(btnOK);
            buttons.Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            // events
            _txtFind.TextChanged += (s, e) => RefreshPreview();
            _txtReplace.TextChanged += (s, e) => RefreshPreview();
            _txtPattern.TextChanged += (s, e) => RefreshPreview();
            _txtMath.TextChanged += (s, e) => RefreshPreview();

            Controls.Add(grid);
            Controls.Add(buttons);
        }

        private static WinForms.Label MakeLabel(string txt) =>
            new WinForms.Label
            {
                Text = txt,
                TextAlign = Drawing.ContentAlignment.MiddleRight,
                Dock = WinForms.DockStyle.Fill
            };

        private static WinForms.TextBox MakeTextBox(string initial = "") =>
            new WinForms.TextBox { Text = initial, Dock = WinForms.DockStyle.Fill };

        private static WinForms.Label MakeHint(string txt) =>
            new WinForms.Label
            {
                Text = txt,
                Dock = WinForms.DockStyle.Fill,
                ForeColor = Drawing.Color.Gray,
                Font = new Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Italic)
            };

        private void LoadCurrentValues()
        {
            foreach (var value in _originalValues)
                _rtbBefore.AppendText(value + Environment.NewLine);

            RefreshPreview();
        }

        private void RefreshPreview()
        {
            _rtbAfter.Clear();
            for (int i = 0; i < _originalValues.Count; i++)
            {
                string originalValue = _originalValues[i];
                Dictionary<string, object> dataRow = i < _dataRows.Count ? _dataRows[i] : null;

                string transformedValue = DataRenamerHelper.TransformValue(originalValue, this, dataRow);
                _rtbAfter.AppendText(transformedValue + Environment.NewLine);
            }
        }

        /// <summary>Initialize pattern value - replace {} with actual value if all selected values are the same</summary>
        private void InitializePatternValue()
        {
            // Check if all original values are the same (or only one value)
            if (_originalValues.Count <= 1 || _originalValues.All(v => v == _originalValues[0]))
            {
                // All values are the same, replace {} with the actual value
                string actualValue = _originalValues.Count > 0 ? _originalValues[0] : "";
                _txtPattern.Text = _txtPattern.Text.Replace("{}", actualValue);

                // Set focus to pattern field since it's first in the list
                _txtPattern.Focus();
                _txtPattern.SelectAll();
            }
            else
            {
                // Values are different, keep {} symbol
                // Set focus to pattern field
                _txtPattern.Focus();
                _txtPattern.SelectAll();
            }
        }
    }
}
