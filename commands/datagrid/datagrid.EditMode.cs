using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using System.Windows.Forms;
using System.Drawing;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

public partial class CustomGUIs
{
    // Edit mode state extracted from DataGrid2_Helpers.cs
    private static bool _isEditMode = false;
    // CRITICAL: Store edits by entity identifier (Handle@DocumentPath|ColumnName) instead of row index
    // This ensures edits stay with the correct entity even when filtering changes row order
    private static Dictionary<string, object> _pendingCellEdits = new Dictionary<string, object>();
    private static List<DataGridViewCell> _selectedEditCells = new List<DataGridViewCell>();
    private static HashSet<Dictionary<string, object>> _modifiedEntries = new HashSet<Dictionary<string, object>>();

    // Track if edits were applied in the current session
    private static bool _editsWereApplied = false;

    /// <summary>Check if there are pending edits waiting to be applied</summary>
    public static bool HasPendingEdits() => _pendingCellEdits.Count > 0;

    /// <summary>Check if edits were applied in the current session</summary>
    public static bool WereEditsApplied() => _editsWereApplied;

    /// <summary>Reset the edits applied flag (called at start of new DataGrid session)</summary>
    public static void ResetEditsAppliedFlag() => _editsWereApplied = false;

    // Selection anchor for Shift+Arrow behavior (Excel-like)
    private static DataGridViewCell _selectionAnchor = null;

    // Edit Mode Helper Methods
    private static void ResetEditMode()
    {
        _isEditMode = false;
        _pendingCellEdits.Clear();
        _selectedEditCells.Clear();
        _modifiedEntries.Clear();
        _selectionAnchor = null;
    }

    /// <summary>Create edit key using internal ID instead of row index</summary>
    private static string GetEditKey(Dictionary<string, object> entry, string columnName)
    {
        // Use the stable internal ID assigned to each entry
        // This is independent of row position, data content, or specific columns
        long internalId = CustomGUIs.GetInternalId(entry);
        return $"{internalId}|{columnName}";
    }

    private static bool IsColumnEditable(string columnName)
    {
        string lowerName = columnName.ToLowerInvariant();
        switch (lowerName)
        {
            case "name":
            case "dynamicblockname":
            case "contents":
            case "value":
            case "layer":
            case "color":
            case "linetype":
            case "layout":
            case "papersize":
            case "plotstyletable":
            case "plotrotation":
            case "plotconfigurationname":
            case "plotscale":
            case "plottype":
            case "plotcentered":
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
            case "query":
            case "tags":
            case "xrefpath":
            // Layer-specific properties
            case "isfrozen":
            case "islocked":
            case "isoff":
            case "isplottable":
            case "lineweight":
            case "transparency":
            case "description":
                return true;
        }
        if (lowerName.StartsWith("attr_")) return true;
        if (lowerName.StartsWith("xdata_")) return true;
        if (lowerName.StartsWith("ext_dict_")) return true;
        if (lowerName.StartsWith("tag_")) return true;
        return false;
    }

    private static void ToggleEditMode(DataGridView grid)
    {
        if (_isEditMode)
        {
            _isEditMode = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = true;
            _selectedEditCells.Clear();
            _selectionAnchor = null;
            UpdateColumnEditableStyles(grid, false);
        }
        else
        {
            _isEditMode = true;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = true;
            grid.ClearSelection();
            _selectionAnchor = null; // Reset anchor when entering edit mode
            UpdateColumnEditableStyles(grid, true);

            // Set initial anchor to current cell if one exists
            if (grid.CurrentCell != null && IsColumnEditable(grid.Columns[grid.CurrentCell.ColumnIndex].Name))
            {
                _selectionAnchor = grid.CurrentCell;
            }
        }
    }

    private static void UpdateColumnEditableStyles(DataGridView grid, bool editModeActive)
    {
        foreach (DataGridViewColumn column in grid.Columns)
        {
            if (editModeActive)
            {
                if (IsColumnEditable(column.Name))
                {
                    column.HeaderCell.Style.BackColor = Color.LightGreen;
                    column.HeaderCell.Style.ForeColor = Color.Black;
                    column.DefaultCellStyle.BackColor = Color.White;
                    column.DefaultCellStyle.ForeColor = Color.Black;
                }
                else
                {
                    column.HeaderCell.Style.BackColor = Color.LightGray;
                    column.HeaderCell.Style.ForeColor = Color.DarkGray;
                    column.DefaultCellStyle.BackColor = Color.WhiteSmoke;
                    column.DefaultCellStyle.ForeColor = Color.Gray;
                }
            }
            else
            {
                column.HeaderCell.Style.BackColor = Color.Empty;
                column.HeaderCell.Style.ForeColor = Color.Empty;
                column.DefaultCellStyle.BackColor = Color.Empty;
                column.DefaultCellStyle.ForeColor = Color.Empty;
            }
        }
    }

    private static void ApplyPendingEdits()
    {
        // Edits are already applied to entry dictionaries in real-time
        // This method is kept for backwards compatibility but does nothing
    }

    private static void ShowCellEditPrompt(DataGridView grid)
    {
        if (_selectedEditCells.Count == 0) return;

        var editableCells = _selectedEditCells.Where(cell => IsColumnEditable(grid.Columns[cell.ColumnIndex].Name)).ToList();
        if (editableCells.Count == 0)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            ed.WriteMessage("\nNo editable cells selected. Editable columns are highlighted in green.");
            return;
        }

        var originalSelectedCells = _selectedEditCells.ToList();
        _selectedEditCells.Clear();
        _selectedEditCells.AddRange(editableCells);

        var currentValues = new List<string>();
        var dataRows = new List<Dictionary<string, object>>();

        foreach (var cell in _selectedEditCells)
        {
            if (cell.RowIndex < _cachedFilteredData.Count)
            {
                var entry = _cachedFilteredData[cell.RowIndex];
                string columnName = grid.Columns[cell.ColumnIndex].Name;
                string currentValue = entry.ContainsKey(columnName) && entry[columnName] != null ? entry[columnName].ToString() : "";
                currentValues.Add(currentValue);
                dataRows.Add(entry);
            }
            else
            {
                currentValues.Add("");
                dataRows.Add(new Dictionary<string, object>());
            }
        }

        using (var advancedForm = new AutoCADCommands.AdvancedEditDialog(currentValues, dataRows, "Advanced Cell Editor"))
        {
            if (advancedForm.ShowDialog() == WinForms.DialogResult.OK)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                ed.WriteMessage($"\nApplying advanced edits to {_selectedEditCells.Count} cells");

                for (int i = 0; i < _selectedEditCells.Count; i++)
                {
                    var cell = _selectedEditCells[i];
                    string originalValue = i < currentValues.Count ? currentValues[i] : "";
                    var dataRow = i < dataRows.Count ? dataRows[i] : null;
                    string newValue = TransformValue(originalValue, advancedForm, dataRow);

                    if (newValue != originalValue)
                    {
                        string columnName = grid.Columns[cell.ColumnIndex].Name;

                        // Validate DynamicBlockName edits before applying
                        if (columnName.Equals("DynamicBlockName", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!ValidateDynamicBlockNameEdit(newValue, originalValue, dataRow, ed))
                            {
                                // Validation failed, skip this edit
                                continue;
                            }
                        }

                        if (cell.RowIndex < _cachedFilteredData.Count)
                        {
                            var entry = _cachedFilteredData[cell.RowIndex];
                            string editKey = GetEditKey(entry, columnName);

                            ed.WriteMessage($"\nCell [{cell.RowIndex}, {columnName}]: '{originalValue}' → '{newValue}'");
                            _pendingCellEdits[editKey] = newValue;
                            entry[columnName] = newValue;
                            _modifiedEntries.Add(entry);
                        }
                    }
                }

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

        _selectedEditCells.Clear();
        _selectedEditCells.AddRange(originalSelectedCells);
    }

    /// <summary>Bridge method to use AdvancedEditDialog with corrected precedence logic</summary>
    private static string TransformValue(string originalValue, AutoCADCommands.AdvancedEditDialog dialog, Dictionary<string, object> dataRow)
    {
        // Use the same transformation logic as AdvancedEditDialog.TransformValue
        // 1. Find/Replace
        string value = originalValue ?? string.Empty;
        if (!string.IsNullOrEmpty(dialog.FindText))
        {
            if (dialog.IsRegexMode)
            {
                try
                {
                    // Use regex mode for find/replace
                    value = Regex.Replace(value, dialog.FindText, dialog.ReplaceText ?? string.Empty);
                }
                catch (ArgumentException)
                {
                    // Invalid regex pattern - fallback to literal replacement
                    value = value.Replace(dialog.FindText, dialog.ReplaceText ?? string.Empty);
                }
            }
            else
            {
                // Use simple string replacement
                value = value.Replace(dialog.FindText, dialog.ReplaceText ?? string.Empty);
            }
        }
        else if (!string.IsNullOrEmpty(dialog.ReplaceText))
            value = dialog.ReplaceText;

        // 2. Pattern (with data reference support)
        if (!string.IsNullOrEmpty(dialog.PatternText))
        {
            if (dataRow != null)
                value = DataRenamerHelper.ParsePatternWithDataReferences(dialog.PatternText, value, dataRow);
            else
                value = dialog.PatternText.Replace("{}", value);
        }

        // 3. Math
        if (!string.IsNullOrEmpty(dialog.MathOperationText))
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                value = DataRenamerHelper.ApplyMathOperation(n, dialog.MathOperationText).ToString(CultureInfo.InvariantCulture);
            else
                value = DataRenamerHelper.ApplyMathToNumbersInString(value, dialog.MathOperationText);
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
                var entry = _cachedFilteredData[cell.RowIndex];
                string editKey = GetEditKey(entry, columnName);

                ed.WriteMessage($"\nStoring edit: Row {cell.RowIndex}, Column '{columnName}', Value '{newValue}'");
                _pendingCellEdits[editKey] = newValue;
                entry[columnName] = newValue;
                _modifiedEntries.Add(entry);
            }
        }
    }

    // Excel-like selection helpers
    private static void SelectRowsOfSelectedCells(DataGridView grid)
    {
        if (_selectedEditCells.Count == 0) return;
        var rowIndices = _selectedEditCells.Select(cell => cell.RowIndex).Distinct().ToList();
        grid.ClearSelection();
        foreach (int rowIndex in rowIndices)
        {
            if (rowIndex >= 0 && rowIndex < grid.Rows.Count)
            {
                foreach (DataGridViewCell cell in grid.Rows[rowIndex].Cells)
                    if (cell.Visible && IsColumnEditable(grid.Columns[cell.ColumnIndex].Name))
                        cell.Selected = true;
            }
        }
    }

    private static void SelectColumnsOfSelectedCells(DataGridView grid)
    {
        if (_selectedEditCells.Count == 0) return;
        var columnIndices = _selectedEditCells.Select(cell => cell.ColumnIndex).Distinct()
            .Where(colIndex => IsColumnEditable(grid.Columns[colIndex].Name)).ToList();
        if (columnIndices.Count == 0) return;
        grid.ClearSelection();
        foreach (int columnIndex in columnIndices)
        {
            if (columnIndex >= 0 && columnIndex < grid.Columns.Count)
            {
                for (int rowIndex = 0; rowIndex < grid.Rows.Count; rowIndex++)
                {
                    if (rowIndex < grid.Rows.Count && grid.Rows[rowIndex].Cells[columnIndex].Visible)
                        grid.Rows[rowIndex].Cells[columnIndex].Selected = true;
                }
            }
        }
    }

    private static void MoveCellWithArrows(DataGridView grid, Keys keyCode)
    {
        if (grid.CurrentCell == null) return;
        int currentRow = grid.CurrentCell.RowIndex;
        int currentCol = grid.CurrentCell.ColumnIndex;
        int newRow = currentRow;
        int newCol = currentCol;
        switch (keyCode)
        {
            case Keys.Up: newRow = Math.Max(0, currentRow - 1); break;
            case Keys.Down: newRow = Math.Min(grid.Rows.Count - 1, currentRow + 1); break;
            case Keys.Left: newCol = FindNextEditableColumn(grid, currentCol, -1); break;
            case Keys.Right: newCol = FindNextEditableColumn(grid, currentCol, 1); break;
        }
        if (newRow != currentRow || newCol != currentCol)
        {
            grid.ClearSelection();
            try
            {
                grid.CurrentCell = grid.Rows[newRow].Cells[newCol];
                grid.Rows[newRow].Cells[newCol].Selected = true;
                // Reset anchor when moving without Shift (Excel behavior)
                _selectionAnchor = grid.Rows[newRow].Cells[newCol];
            }
            catch (System.Exception) { }
        }
    }

    private static void ExtendSelectionWithArrows(DataGridView grid, Keys keyCode, bool isShiftKey)
    {
        if (grid.CurrentCell == null) return;
        int currentRow = grid.CurrentCell.RowIndex;
        int currentCol = grid.CurrentCell.ColumnIndex;
        int newRow = currentRow;
        int newCol = currentCol;
        switch (keyCode)
        {
            case Keys.Up: newRow = Math.Max(0, currentRow - 1); break;
            case Keys.Down: newRow = Math.Min(grid.Rows.Count - 1, currentRow + 1); break;
            case Keys.Left: newCol = FindNextEditableColumn(grid, currentCol, -1); break;
            case Keys.Right: newCol = FindNextEditableColumn(grid, currentCol, 1); break;
        }
        if (newRow != currentRow || newCol != currentCol)
        {
            try
            {
                if (isShiftKey)
                {
                    // Excel-like Shift+Arrow: Keep anchor fixed, extend/shrink selection to new position
                    if (_selectionAnchor == null) _selectionAnchor = grid.CurrentCell;

                    // Important: Set CurrentCell FIRST before clearing selection
                    // This prevents the CurrentCell from being automatically reset
                    grid.CurrentCell = grid.Rows[newRow].Cells[newCol];

                    // Now clear and rebuild selection
                    grid.ClearSelection();
                    SelectRectangularArea(grid, _selectionAnchor.RowIndex, _selectionAnchor.ColumnIndex, newRow, newCol);

                    // Don't update anchor - it stays fixed during Shift+Arrow operations
                }
                else
                {
                    // Ctrl+Arrow (original multi-select behavior): Add to selection without anchor
                    grid.Rows[newRow].Cells[newCol].Selected = true;
                    grid.CurrentCell = grid.Rows[newRow].Cells[newCol];
                    _selectionAnchor = grid.Rows[newRow].Cells[newCol];
                }
            }
            catch (System.Exception) { }
        }
    }

    private static int FindNextEditableColumn(DataGridView grid, int currentCol, int direction)
    {
        int newCol = currentCol + direction;
        while (newCol >= 0 && newCol < grid.Columns.Count)
        {
            DataGridViewColumn column = grid.Columns[newCol];
            string columnName = column.Name;
            // Check both if column is editable AND visible (important when columns are filtered)
            if (IsColumnEditable(columnName) && column.Visible) return newCol;
            newCol += direction;
        }
        return currentCol;
    }

    private static void SelectRectangularArea(DataGridView grid, int anchorRow, int anchorCol, int targetRow, int targetCol)
    {
        int startRow = Math.Min(anchorRow, targetRow);
        int endRow = Math.Max(anchorRow, targetRow);
        int startCol = Math.Min(anchorCol, targetCol);
        int endCol = Math.Max(anchorCol, targetCol);
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
                            try { grid.Rows[row].Cells[col].Selected = true; } catch { }
                        }
                    }
                }
            }
        }
    }

    public static void SetSelectionAnchor(DataGridViewCell cell) => _selectionAnchor = cell;

    /// <summary>Handle clipboard paste (Ctrl+V) in edit mode - supports multi-cell paste</summary>
    private static void HandleClipboardPaste(DataGridView grid)
    {
        if (!_isEditMode || grid.CurrentCell == null)
            return;

        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            // Get clipboard text
            if (!Clipboard.ContainsText())
            {
                ed.WriteMessage("\nClipboard does not contain text data.");
                return;
            }

            string clipboardText = Clipboard.GetText();
            if (string.IsNullOrEmpty(clipboardText))
            {
                ed.WriteMessage("\nClipboard is empty.");
                return;
            }

            // Parse clipboard data into 2D array (tab-separated columns, newline-separated rows)
            var rows = clipboardText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var clipboardData = new List<List<string>>();

            foreach (var row in rows)
            {
                if (string.IsNullOrEmpty(row) && clipboardData.Count > 0)
                    continue; // Skip trailing empty lines

                var columns = row.Split('\t');
                clipboardData.Add(new List<string>(columns));
            }

            if (clipboardData.Count == 0)
            {
                ed.WriteMessage("\nNo data to paste.");
                return;
            }

            // Check if clipboard contains a single value (1 row, 1 column)
            bool isSingleValue = clipboardData.Count == 1 && clipboardData[0].Count == 1;

            int pastedCells = 0;
            int skippedCells = 0;

            // Special case: Single value pasted to multiple selected cells
            if (isSingleValue && _selectedEditCells.Count > 1)
            {
                string singleValue = clipboardData[0][0];

                // Paste the single value to all selected editable cells
                foreach (var cell in _selectedEditCells)
                {
                    if (!IsColumnEditable(grid.Columns[cell.ColumnIndex].Name))
                    {
                        skippedCells++;
                        continue;
                    }

                    if (cell.RowIndex < _cachedFilteredData.Count)
                    {
                        string columnName = grid.Columns[cell.ColumnIndex].Name;
                        var entry = _cachedFilteredData[cell.RowIndex];
                        string editKey = GetEditKey(entry, columnName);

                        _pendingCellEdits[editKey] = singleValue;
                        entry[columnName] = singleValue;
                        _modifiedEntries.Add(entry);

                        pastedCells++;
                    }
                    else
                    {
                        skippedCells++;
                    }
                }
            }
            else
            {
                // Multi-cell paste: Use top-left cell of selection as starting point
                int startRow = grid.CurrentCell.RowIndex;
                int startCol = grid.CurrentCell.ColumnIndex;

                // If multiple cells are selected, find the top-left cell
                if (_selectedEditCells.Count > 1)
                {
                    startRow = _selectedEditCells.Min(cell => cell.RowIndex);
                    startCol = _selectedEditCells.Min(cell => cell.ColumnIndex);
                }

                // Find the first editable column at or after startCol
                while (startCol < grid.Columns.Count && !IsColumnEditable(grid.Columns[startCol].Name))
                {
                    startCol++;
                }

                if (startCol >= grid.Columns.Count)
                {
                    ed.WriteMessage("\nNo editable column found at cursor position.");
                    return;
                }

                // Paste data starting from top-left cell
                for (int clipRow = 0; clipRow < clipboardData.Count; clipRow++)
                {
                    int targetRow = startRow + clipRow;
                    if (targetRow >= grid.Rows.Count)
                        break; // Don't paste beyond available rows

                    var clipboardRow = clipboardData[clipRow];
                    int targetCol = startCol;

                    for (int clipCol = 0; clipCol < clipboardRow.Count; clipCol++)
                    {
                        // Find next editable column
                        while (targetCol < grid.Columns.Count && !IsColumnEditable(grid.Columns[targetCol].Name))
                        {
                            targetCol++;
                        }

                        if (targetCol >= grid.Columns.Count)
                            break; // No more editable columns

                        string columnName = grid.Columns[targetCol].Name;
                        string newValue = clipboardRow[clipCol];

                        // Apply the paste
                        if (targetRow < _cachedFilteredData.Count)
                        {
                            var entry = _cachedFilteredData[targetRow];
                            string editKey = GetEditKey(entry, columnName);

                            _pendingCellEdits[editKey] = newValue;
                            entry[columnName] = newValue;
                            _modifiedEntries.Add(entry);

                            pastedCells++;
                        }
                        else
                        {
                            skippedCells++;
                        }

                        targetCol++; // Move to next column
                    }
                }
            }

            // Update grid display
            grid.Invalidate();

            // Report results
            ed.WriteMessage($"\nPasted {pastedCells} cell(s)");
            if (skippedCells > 0)
                ed.WriteMessage($" ({skippedCells} cell(s) skipped - out of range or non-editable)");

            ed.WriteMessage($"\nTotal pending edits: {_pendingCellEdits.Count}");
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError pasting clipboard data: {ex.Message}");
        }
    }

    // Integrated rename helpers (used by TransformValue)
    public static class DataRenamerHelper
    {
        public static double ApplyMathOperation(double x, string mathOp)
        {
            if (string.IsNullOrWhiteSpace(mathOp)) return x;
            mathOp = mathOp.Replace(" ", string.Empty);
            if (mathOp.Equals("x", StringComparison.OrdinalIgnoreCase)) return x;
            if (mathOp.Equals("-x", StringComparison.OrdinalIgnoreCase)) return -x;
            if (!mathOp.StartsWith("x", StringComparison.OrdinalIgnoreCase) && mathOp.EndsWith("x", StringComparison.OrdinalIgnoreCase))
            {
                string multStr = mathOp.Substring(0, mathOp.Length - 1);
                if (double.TryParse(multStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double mult)) return mult * x;
            }
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
            return x;
        }

        public static string ApplyMathToNumbersInString(string input, string mathOp)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrWhiteSpace(mathOp)) return input;
            return Regex.Replace(input, @"-?\d+(?:\.\d+)?", m =>
            {
                if (!double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double n)) return m.Value;
                double res = ApplyMathOperation(n, mathOp);
                return Math.Abs(res % 1) < 1e-10
                    ? Math.Round(res).ToString(CultureInfo.InvariantCulture)
                    : res.ToString(CultureInfo.InvariantCulture);
            });
        }

        public static string ParsePatternWithDataReferences(string pattern, string currentValue, Dictionary<string, object> dataRow)
        {
            if (string.IsNullOrEmpty(pattern)) return currentValue;
            string result = pattern.Replace("{}", currentValue);
            // Regex pattern: $"Column Name" (quoted only, requires explicit spaces)
            var regex = new Regex(@"\$""([^""]+)""");
            result = regex.Replace(result, match =>
            {
                string dataName = match.Groups[1].Value;
                string dataValue = GetDataValueFromRow(dataRow, dataName);
                return !string.IsNullOrEmpty(dataValue) ? dataValue : match.Value;
            });
            return result;
        }

        private static string GetDataValueFromRow(Dictionary<string, object> row, string key)
        {
            if (row == null || string.IsNullOrEmpty(key)) return string.Empty;

            // Try exact match first
            if (row.TryGetValue(key, out var exactValue) && exactValue != null)
                return exactValue.ToString();

            // Try case-insensitive match
            var kvp = row.FirstOrDefault(k => string.Equals(k.Key, key, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(kvp.Key) && kvp.Value != null)
                return kvp.Value.ToString();

            // Try matching with spaces converted to underscores
            // This allows $"attr dated" to match "attr_dated" key in dictionary
            string keyWithUnderscores = key.Replace(" ", "_");
            if (keyWithUnderscores != key) // Only try if we actually replaced something
            {
                kvp = row.FirstOrDefault(k => string.Equals(k.Key, keyWithUnderscores, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(kvp.Key) && kvp.Value != null)
                    return kvp.Value.ToString();
            }

            // Try matching with underscores converted to spaces
            // This allows $"attr_dated" to match "attr dated" key in dictionary
            string keyWithSpaces = key.Replace("_", " ");
            if (keyWithSpaces != key) // Only try if we actually replaced something
            {
                kvp = row.FirstOrDefault(k => string.Equals(k.Key, keyWithSpaces, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(kvp.Key) && kvp.Value != null)
                    return kvp.Value.ToString();
            }

            // Try matching with normalized spaces (remove spaces from both)
            // This allows $"Entity Type" to match "EntityType" key in dictionary
            string keyWithoutSpaces = key.Replace(" ", "").Replace("_", "");
            kvp = row.FirstOrDefault(k =>
                string.Equals(
                    k.Key.Replace(" ", "").Replace("_", ""),
                    keyWithoutSpaces,
                    StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(kvp.Key) && kvp.Value != null)
                return kvp.Value.ToString();

            return string.Empty;
        }
    }

    /// <summary>
    /// Validates DynamicBlockName edits to ensure block exists before allowing the edit.
    /// Shows warning dialog if block doesn't exist.
    /// </summary>
    private static bool ValidateDynamicBlockNameEdit(string newValue, string originalValue, Dictionary<string, object> dataRow, Autodesk.AutoCAD.EditorInput.Editor ed)
    {
        if (dataRow == null) return true;

        // Get the entity's category to check if it's a block reference
        if (!dataRow.TryGetValue("Category", out var categoryObj)) return true;
        string category = categoryObj?.ToString() ?? "";

        // Only validate for block references
        if (!category.Contains("Block", StringComparison.OrdinalIgnoreCase)) return true;

        // Check if the new block name exists in the block table
        var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        using (var tr = db.TransactionManager.StartTransaction())
        {
            var blockTable = (Autodesk.AutoCAD.DatabaseServices.BlockTable)tr.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

            if (!blockTable.Has(newValue))
            {
                // Block doesn't exist - show warning dialog
                tr.Commit();

                System.Windows.Forms.MessageBox.Show(
                    $"Cannot create new dynamic block '{newValue}'.\n\n" +
                    $"The DynamicBlockName column only supports swapping to existing block definitions.\n\n" +
                    $"Current block: {originalValue}\n" +
                    $"Attempted new name: {newValue}\n\n" +
                    $"To use this feature:\n" +
                    $"• Enter the name of an existing block definition in your drawing\n" +
                    $"• The block reference will swap to that definition\n" +
                    $"• All matching attributes will be preserved\n\n" +
                    $"Note: Creating new dynamic blocks requires using the Block Editor in AutoCAD.",
                    "Dynamic Block Limitation",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);

                ed.WriteMessage($"\n  >> Edit rejected: Block '{newValue}' does not exist. Only swapping to existing blocks is supported.");
                return false;
            }

            tr.Commit();
        }

        return true;
    }
}

