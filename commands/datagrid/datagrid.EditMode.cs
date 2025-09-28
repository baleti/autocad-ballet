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

    private static bool IsColumnEditable(string columnName)
    {
        string lowerName = columnName.ToLowerInvariant();
        switch (lowerName)
        {
            case "name":
            case "contents":
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
                return true;
        }
        if (lowerName.StartsWith("attr_")) return true;
        if (lowerName.StartsWith("xdata_")) return true;
        if (lowerName.StartsWith("ext_dict_")) return true;
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
            UpdateColumnEditableStyles(grid, false);
        }
        else
        {
            _isEditMode = true;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = true;
            grid.ClearSelection();
            UpdateColumnEditableStyles(grid, true);
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

    private static string GetCellKey(int rowIndex, string columnName) => $"{rowIndex}|{columnName}";

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
                        string cellKey = GetCellKey(cell.RowIndex, columnName);
                        ed.WriteMessage($"\nCell [{cell.RowIndex}, {columnName}]: '{originalValue}' → '{newValue}'");
                        _pendingCellEdits[cellKey] = newValue;

                        if (cell.RowIndex < _cachedFilteredData.Count)
                        {
                            var entry = _cachedFilteredData[cell.RowIndex];
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
            value = value.Replace(dialog.FindText, dialog.ReplaceText ?? string.Empty);
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
                string cellKey = GetCellKey(cell.RowIndex, columnName);
                ed.WriteMessage($"\nStoring edit: Row {cell.RowIndex}, Column '{columnName}', Value '{newValue}'");
                _pendingCellEdits[cellKey] = newValue;
                var entry = _cachedFilteredData[cell.RowIndex];
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
                    if (_selectionAnchor == null) _selectionAnchor = grid.CurrentCell;
                    grid.ClearSelection();
                    SelectRectangularArea(grid, _selectionAnchor.RowIndex, _selectionAnchor.ColumnIndex, newRow, newCol);
                    grid.CurrentCell = grid.Rows[newRow].Cells[newCol];
                }
                else
                {
                    grid.Rows[newRow].Cells[newCol].Selected = true;
                    grid.CurrentCell = grid.Rows[newRow].Cells[newCol];
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
            string columnName = grid.Columns[newCol].Name;
            if (IsColumnEditable(columnName)) return newCol;
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
            var regex = new Regex(@"\$""([^""]+)""|(?<!\$)\$(\w+)");
            result = regex.Replace(result, match =>
            {
                string dataName = !string.IsNullOrEmpty(match.Groups[1].Value) ? match.Groups[1].Value : match.Groups[2].Value;
                string dataValue = GetDataValueFromRow(dataRow, dataName);
                return !string.IsNullOrEmpty(dataValue) ? dataValue : match.Value;
            });
            return result;
        }

        private static string GetDataValueFromRow(Dictionary<string, object> row, string key)
        {
            if (row == null || string.IsNullOrEmpty(key)) return string.Empty;
            if (row.TryGetValue(key, out var v) && v != null) return v.ToString();
            // Also try case-insensitive match
            var kv = row.FirstOrDefault(k => string.Equals(k.Key, key, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(kv.Key) && kv.Value != null) return kv.Value.ToString();
            return string.Empty;
        }
    }
}

