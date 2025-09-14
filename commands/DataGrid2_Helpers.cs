using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

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
    }

    /// <summary>Check if a column is editable</summary>
    private static bool IsColumnEditable(string columnName)
    {
        string lowerName = columnName.ToLowerInvariant();

        // Editable columns
        switch (lowerName)
        {
            case "name":           // Entity names, layout names, etc.
            case "layer":          // Layer assignment
            case "color":          // Color property
            case "linetype":       // Linetype assignment
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

    /// <summary>Show text edit prompt for selected cells (Excel-like functionality)</summary>
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

        // Get the value from the first selected cell
        var firstCell = _selectedEditCells[0];
        string currentValue = "";

        if (firstCell.RowIndex < _cachedFilteredData.Count)
        {
            var entry = _cachedFilteredData[firstCell.RowIndex];
            string columnName = grid.Columns[firstCell.ColumnIndex].Name;
            if (entry.ContainsKey(columnName) && entry[columnName] != null)
            {
                currentValue = entry[columnName].ToString();
            }
        }

        // Create edit form (Excel-like input box)
        using (Form editForm = new Form())
        {
            editForm.Text = "Edit Cell Value";
            editForm.Size = new Size(400, 150);
            editForm.StartPosition = FormStartPosition.CenterParent;
            editForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            editForm.MaximizeBox = false;
            editForm.MinimizeBox = false;

            Label label = new Label()
            {
                Text = $"Edit value for {_selectedEditCells.Count} cell(s):",
                Location = new Point(10, 15),
                Size = new Size(350, 20)
            };

            TextBox textBox = new TextBox()
            {
                Text = currentValue,
                Location = new Point(10, 40),
                Size = new Size(360, 25),
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };

            Button okButton = new Button()
            {
                Text = "OK",
                Location = new Point(215, 80),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };

            Button cancelButton = new Button()
            {
                Text = "Cancel",
                Location = new Point(295, 80),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            editForm.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
            editForm.AcceptButton = okButton;
            editForm.CancelButton = cancelButton;

            // Handle Enter key for multi-line editing or confirmation
            textBox.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift)
                {
                    editForm.DialogResult = DialogResult.OK;
                    editForm.Close();
                    e.Handled = true;
                }
            };

            textBox.SelectAll();
            textBox.Focus();

            if (editForm.ShowDialog() == DialogResult.OK)
            {
                // Apply the new value to all selected cells
                string newValue = textBox.Text;
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                ed.WriteMessage($"\nApplying edit: '{newValue}' to {_selectedEditCells.Count} editable cells");
                ApplyValueToSelectedCells(grid, newValue);

                // Refresh grid to show changes
                grid.Invalidate();
                ed.WriteMessage($"\nTotal pending edits now: {_pendingCellEdits.Count}");
            }
            else
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                ed.WriteMessage("\nEdit cancelled");
            }

            // Restore original selection
            _selectedEditCells.Clear();
            _selectedEditCells.AddRange(originalSelectedCells);
        }
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
                            }

                            throw; // Re-throw for debugging
                        }
                    }
                    else if (dbObject is MText mtext)
                    {
                        ed.WriteMessage($"\n  >> DBObject is MText, setting Contents in external document");
                        mtext.Contents = newValue;
                        ed.WriteMessage($"\n  >> MText Contents set successfully in external document");
                    }
                    else if (dbObject is DBText text)
                    {
                        ed.WriteMessage($"\n  >> DBObject is DBText, setting TextString in external document");
                        text.TextString = newValue;
                        ed.WriteMessage($"\n  >> DBText TextString set successfully in external document");
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
                            }

                            throw; // Re-throw for debugging
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
                    else
                    {
                        ed.WriteMessage($"\n  >> DBObject type not supported for name editing ({dbObject.GetType().Name})");
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
}
