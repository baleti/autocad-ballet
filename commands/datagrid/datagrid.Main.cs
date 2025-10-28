using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

public partial class CustomGUIs
{
    // ──────────────────────────────────────────────────────────────
    //  Screen expansion state tracking
    // ──────────────────────────────────────────────────────────────
    private static int _currentScreenState = 0; // 0 = original, 1+ = expanded to N screens
    private static Rectangle _originalFormBounds;
    private static Screen _originalScreen;

    // ──────────────────────────────────────────────────────────────
    //  DataGrid sizing state tracking
    // ──────────────────────────────────────────────────────────────
    private static bool _initialSizingDone = false;

    // ──────────────────────────────────────────────────────────────
    //  Main DataGrid Method
    // ──────────────────────────────────────────────────────────────

    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null,
        Func<List<Dictionary<string, object>>, bool> onDeleteEntries = null,
        bool allowCreateFromSearch = false)
    {
        if (entries == null || propertyNames == null || propertyNames.Count == 0)
            return new List<Dictionary<string, object>>();

        // Allow empty entries when allowCreateFromSearch is enabled (user can type new values)
        if (entries.Count == 0 && !allowCreateFromSearch)
            return new List<Dictionary<string, object>>();

        // Clear any previous cached data
        _cachedOriginalData = entries;
        _cachedFilteredData = entries;
        _searchIndexByColumn = null;
        _searchIndexAllColumns = null;
        _lastVisibleColumns.Clear();
        _lastColumnVisibilityFilter = "";

        // Reset sizing and edit mode state
        _initialSizingDone = false;
        ResetEditMode();

        // Build search index upfront for performance
        BuildSearchIndex(entries, propertyNames);

        // State variables
        List<Dictionary<string, object>> selectedEntries = new List<Dictionary<string, object>>();
        List<Dictionary<string, object>> workingSet = new List<Dictionary<string, object>>(entries);
        List<SortCriteria> sortCriteria = new List<SortCriteria>();

        // Create form
        Form form = new Form
        {
            StartPosition = FormStartPosition.CenterScreen,
            Text = "Total Entries: " + entries.Count,
            BackColor = Color.White,
            ShowIcon = false
        };

        // Create DataGridView with virtual mode
        DataGridView grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            ReadOnly = true,
            AutoGenerateColumns = false,
            RowHeadersVisible = false,
            AllowUserToAddRows = false,
            AllowUserToResizeRows = false,
            BackgroundColor = Color.White,
            RowTemplate = { Height = 18 },
            VirtualMode = true,
            ScrollBars = ScrollBars.Both
        };

        // Disable built-in sorting
        grid.SortCompare += (sender, e) =>
        {
            e.Handled = true;
            e.SortResult = naturalComparer.Compare(e.CellValue1, e.CellValue2);
        };

        // Add columns
        foreach (string col in propertyNames)
        {
            var column = new DataGridViewTextBoxColumn
            {
                Name = col,
                HeaderText = FormatColumnHeader(col),
                DataPropertyName = col,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            grid.Columns.Add(column);
        }

        // Search box - keep original appearance
        TextBox searchBox = new TextBox { Dock = DockStyle.Top };

        // Set up virtual mode cell value handler
        grid.CellValueNeeded += (s, e) =>
        {
            if (e.RowIndex >= 0 && e.RowIndex < _cachedFilteredData.Count && e.ColumnIndex >= 0)
            {
                var row = _cachedFilteredData[e.RowIndex];
                string columnName = grid.Columns[e.ColumnIndex].Name;
                object value;
                e.Value = row.TryGetValue(columnName, out value) ? value : null;
            }
        };

        // Initialize grid with data
        grid.RowCount = workingSet.Count;
        
        // Preserve caller-provided row order on initial load
        // (Do not apply any default sort here; sorting is user-driven.)

        // Helper to get first visible column
        Func<int> GetFirstVisibleColumnIndex = () =>
        {
            foreach (DataGridViewColumn c in grid.Columns)
                if (c.Visible) return c.Index;
            return -1;
        };

        // Helper to update filtered grid
        Action UpdateFilteredGrid = () =>
        {
            // Use optimized filtering
            var filteredData = ApplyFilters(_cachedOriginalData, propertyNames, searchBox.Text, grid);

            // Apply sorting
            workingSet = filteredData;
            if (sortCriteria.Count > 0)
            {
                workingSet = ApplySorting(workingSet, sortCriteria);
            }

            // Update cached filtered data and virtual grid row count
            _cachedFilteredData = workingSet;
            grid.RowCount = 0; // Force refresh
            grid.RowCount = workingSet.Count;

            // Update form title
            form.Text = "Total Entries: " + workingSet.Count + " / " + entries.Count;

            // Auto-resize columns and form width only on initial load
            if (!_initialSizingDone)
            {
                if (grid.Columns.Count < 20)
                {
                    grid.AutoResizeColumns();
                }

                int reqWidth = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                              + SystemInformation.VerticalScrollBarWidth + 50;
                form.Width = Math.Min(reqWidth, Screen.PrimaryScreen.WorkingArea.Width - 20);

                _initialSizingDone = true;
            }
        };

        // Apply initial sorting and set up grid
        UpdateFilteredGrid();

        // Initial selection
        if (initialSelectionIndices != null && initialSelectionIndices.Count > 0)
        {
            int firstVisible = GetFirstVisibleColumnIndex();
            bool currentCellSet = false;
            
            foreach (int idx in initialSelectionIndices)
            {
                if (idx >= 0 && idx < grid.Rows.Count && firstVisible >= 0)
                {
                    grid.Rows[idx].Selected = true;
                    
                    // Only set CurrentCell once, for the first valid selection
                    if (!currentCellSet)
                    {
                        try
                        {
                            grid.CurrentCell = grid.Rows[idx].Cells[firstVisible];
                            currentCellSet = true;
                        }
                        catch (System.Exception)
                        {
                            // Ignore CurrentCell setting errors in virtual mode
                        }
                    }
                }
            }
        }

        // Setup delay timer for large datasets
        bool useDelay = entries.Count > 200;
        System.Windows.Forms.Timer delayTimer = new System.Windows.Forms.Timer { Interval = 200 };
        delayTimer.Tick += delegate { delayTimer.Stop(); UpdateFilteredGrid(); };
        form.FormClosed += delegate { delayTimer.Dispose(); };

        // Search box text changed
        searchBox.TextChanged += delegate
        {
            if (useDelay)
            {
                delayTimer.Stop();
                delayTimer.Start();
            }
            else
            {
                UpdateFilteredGrid();
            }
        };

        // Column header click for sorting
        grid.ColumnHeaderMouseClick += (s, e) =>
        {
            string colName = grid.Columns[e.ColumnIndex].Name;
            SortCriteria existing = sortCriteria.FirstOrDefault(sc => sc.ColumnName == colName);

            if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                if (existing != null) sortCriteria.Remove(existing);
            }
            else
            {
                if (existing != null)
                {
                    existing.Direction = existing.Direction == ListSortDirection.Ascending
                        ? ListSortDirection.Descending
                        : ListSortDirection.Ascending;
                    sortCriteria.Remove(existing);
                }
                else
                {
                    existing = new SortCriteria
                    {
                        ColumnName = colName,
                        Direction = ListSortDirection.Ascending
                    };
                }
                sortCriteria.Insert(0, existing);
                if (sortCriteria.Count > 3)
                    sortCriteria = sortCriteria.Take(3).ToList();
            }

            UpdateFilteredGrid();
        };

        // Finish selection helper
        Action FinishSelection = () =>
        {
            selectedEntries.Clear();
            if (_isEditMode)
            {
                // In edit mode, return only modified entries
                selectedEntries.AddRange(_modifiedEntries);
            }
            else
            {
                // In normal mode, return selected rows
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    if (row.Index < _cachedFilteredData.Count)
                    {
                        selectedEntries.Add(_cachedFilteredData[row.Index]);
                    }
                }
            }
            form.Close();
        };

        // Double-click to select
        grid.CellDoubleClick += (s, e) => FinishSelection();

        // Key handling - restore original behavior
        Action<KeyEventArgs, Control> HandleKeyDown = (e, sender) =>
        {
            if (e.KeyCode == Keys.F2)
            {
                if (_isEditMode && _selectedEditCells.Count > 0)
                {
                    // Enter cell edit mode - show text edit prompt
                    var doc = AcadApp.DocumentManager.MdiActiveDocument;
                    var ed = doc.Editor;
                    ed.WriteMessage($"\nOpening edit prompt for {_selectedEditCells.Count} selected cells");
                    ShowCellEditPrompt(grid);
                }
                else
                {
                    // Toggle edit mode
                    ToggleEditMode(grid);
                    var doc = AcadApp.DocumentManager.MdiActiveDocument;
                    var ed = doc.Editor;
                    ed.WriteMessage($"\nToggled edit mode. Now in {(_isEditMode ? "EDIT" : "NORMAL")} mode");
                    form.Text = "Total Entries: " + workingSet.Count + " / " + entries.Count +
                               (_isEditMode ? " [EDIT MODE]" : "");
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                if (_isEditMode)
                {
                    // In edit mode, first press of escape cancels edit mode and returns to normal mode
                    ToggleEditMode(grid);
                    var doc = AcadApp.DocumentManager.MdiActiveDocument;
                    var ed = doc.Editor;
                    ed.WriteMessage($"\nExited edit mode. Now in NORMAL mode");
                    form.Text = "Total Entries: " + workingSet.Count + " / " + entries.Count;
                    e.Handled = true;
                }
                else
                {
                    // In normal mode, escape cancels entire operation - no changes applied
                    selectedEntries.Clear();
                    form.Close();
                }
            }
            else if (e.KeyCode == Keys.Enter)
            {
                if (_isEditMode && _pendingCellEdits.Count > 0)
                {
                    // Apply pending edits to actual AutoCAD entities and return modified entries
                    var doc = AcadApp.DocumentManager.MdiActiveDocument;
                    var ed = doc.Editor;
                    ed.WriteMessage($"\nApplying {_pendingCellEdits.Count} edits to AutoCAD entities");
                    ApplyCellEditsToEntities();
                    selectedEntries.Clear();
                    selectedEntries.AddRange(_modifiedEntries);
                    form.Close();
                }
                else if (allowCreateFromSearch && grid.SelectedRows.Count == 0 && !string.IsNullOrWhiteSpace(searchBox.Text))
                {
                    // No selection but search text exists - return search text as new entry
                    var newEntry = new Dictionary<string, object>();
                    newEntry["__SEARCH_TEXT__"] = searchBox.Text.Trim();
                    selectedEntries.Clear();
                    selectedEntries.Add(newEntry);
                    form.Close();
                }
                else
                {
                    FinishSelection();
                }
            }
            else if (e.KeyCode == Keys.Tab && sender == grid && !e.Shift)
            {
                searchBox.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Tab && sender == searchBox && !e.Shift)
            {
                grid.Focus();
                e.Handled = true;
            }
            else if ((e.KeyCode == Keys.Down || e.KeyCode == Keys.Up) && sender == searchBox)
            {
                if (grid.Rows.Count > 0)
                {
                    grid.Focus();
                    int newIdx = 0;
                    if (grid.SelectedRows.Count > 0)
                    {
                        int curIdx = grid.SelectedRows[0].Index;
                        newIdx = e.KeyCode == Keys.Down
                            ? Math.Min(curIdx + 1, grid.Rows.Count - 1)
                            : Math.Max(curIdx - 1, 0);
                    }
                    int firstVisible = GetFirstVisibleColumnIndex();
                    if (firstVisible >= 0)
                    {
                        grid.ClearSelection();
                        grid.Rows[newIdx].Selected = true;
                        grid.CurrentCell = grid.Rows[newIdx].Cells[firstVisible];
                    }
                    e.Handled = true;
                }
            }
            else if (!_isEditMode && (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) && (e.Shift && sender == grid))
            {
                // Only scroll when NOT in edit mode
                int offset = e.KeyCode == Keys.Right ? 1000 : -1000;
                grid.HorizontalScrollingOffset += offset;
                e.Handled = true;
            }
            else if (!_isEditMode && e.KeyCode == Keys.Right && sender == grid)
            {
                // Only scroll when NOT in edit mode
                grid.HorizontalScrollingOffset += 50;
                e.Handled = true;
            }
            else if (!_isEditMode && e.KeyCode == Keys.Left && sender == grid)
            {
                // Only scroll when NOT in edit mode
                grid.HorizontalScrollingOffset = Math.Max(grid.HorizontalScrollingOffset - 50, 0);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Space && sender == searchBox && string.IsNullOrEmpty(searchBox.Text))
            {
                if (grid.Rows.Count > 0)
                {
                    int currentIdx = 0;
                    if (grid.SelectedRows.Count > 0)
                    {
                        currentIdx = grid.SelectedRows[0].Index;
                    }

                    int targetIdx;
                    if (e.Shift)
                    {
                        // Shift+Space: move up one row
                        targetIdx = Math.Max(currentIdx - 1, 0);
                    }
                    else
                    {
                        // Space: move down one row
                        targetIdx = Math.Min(currentIdx + 1, grid.Rows.Count - 1);
                    }

                    int firstVisible = GetFirstVisibleColumnIndex();
                    if (firstVisible >= 0)
                    {
                        grid.ClearSelection();
                        grid.Rows[targetIdx].Selected = true;
                        grid.CurrentCell = grid.Rows[targetIdx].Cells[firstVisible];
                        FinishSelection();
                    }
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.D && e.Alt)
            {
                searchBox.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.S && e.Alt)
            {
                // Alt+S: Cycle through screen expansion states
                CycleScreenExpansion(form);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.E && e.Control)
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string storagePath = Path.Combine(appData, "autocad-ballet", "DataGrid-last-export-location");
                string initialPath = "";
                if (File.Exists(storagePath))
                {
                    initialPath = File.ReadAllText(storagePath).Trim();
                }
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "CSV Files|*.csv";
                sfd.Title = "Export DataGrid to CSV";
                sfd.DefaultExt = "csv";
                if (!string.IsNullOrEmpty(initialPath))
                {
                    string dir = Path.GetDirectoryName(initialPath);
                    if (Directory.Exists(dir))
                    {
                        sfd.InitialDirectory = dir;
                        sfd.FileName = Path.GetFileName(initialPath);
                    }
                }
                else
                {
                    sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    sfd.FileName = "DataGridExport.csv";
                }
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportToCsv(grid, _cachedFilteredData, sfd.FileName);
                        Directory.CreateDirectory(Path.GetDirectoryName(storagePath));
                        File.WriteAllText(storagePath, sfd.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error exporting: " + ex.Message);
                    }
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Delete && !_isEditMode && sender == grid && onDeleteEntries != null)
            {
                // Delete key: Invoke callback if provided
                if (grid.SelectedRows.Count > 0)
                {
                    // Show confirmation dialog
                    var result = MessageBox.Show(
                        $"Delete {grid.SelectedRows.Count} selected entry(ies)?",
                        "Confirm Deletion",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button2); // Default to Cancel for safety

                    if (result == DialogResult.OK)
                    {
                        // Collect entries to delete
                        var entriesToDelete = new List<Dictionary<string, object>>();

                        foreach (DataGridViewRow row in grid.SelectedRows)
                        {
                            if (row.Index < _cachedFilteredData.Count)
                            {
                                entriesToDelete.Add(_cachedFilteredData[row.Index]);
                            }
                        }

                        // Invoke callback to handle deletion
                        bool success = onDeleteEntries(entriesToDelete);

                        if (success)
                        {
                            // Remove entries from cached data
                            foreach (var entry in entriesToDelete)
                            {
                                _cachedOriginalData.Remove(entry);
                                _cachedFilteredData.Remove(entry);
                            }

                            // Rebuild search index after deletion
                            BuildSearchIndex(_cachedOriginalData, propertyNames);

                            // Update grid
                            UpdateFilteredGrid();
                        }
                    }
                }
                e.Handled = true;
            }
            else if (_isEditMode && sender == grid)
            {
                // Excel-like shortcuts and navigation in edit mode
                if (e.KeyCode == Keys.Space && e.Shift)
                {
                    // Shift+Space: Select entire rows of currently selected cells
                    SelectRowsOfSelectedCells(grid);
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.Space && e.Control)
                {
                    // Ctrl+Space: Select entire columns of currently selected cells
                    SelectColumnsOfSelectedCells(grid);
                    e.Handled = true;
                }
                else if ((e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
                {
                    if (e.Shift && grid.CurrentCell != null)
                    {
                        // Shift+Arrow: Extend selection (Excel-like)
                        ExtendSelectionWithArrows(grid, e.KeyCode, true);
                        e.Handled = true;
                    }
                    else if (e.Control && grid.CurrentCell != null)
                    {
                        // Ctrl+Arrow: Extend selection (original behavior)
                        ExtendSelectionWithArrows(grid, e.KeyCode, false);
                        e.Handled = true;
                    }
                    else if (grid.CurrentCell != null)
                    {
                        // Arrow keys: Move current cell
                        MoveCellWithArrows(grid, e.KeyCode);
                        e.Handled = true;
                    }
                }
            }
        };

        grid.KeyDown += (s, e) => HandleKeyDown(e, grid);
        searchBox.KeyDown += (s, e) => HandleKeyDown(e, searchBox);

        // Handle cell selection changes in edit mode
        grid.SelectionChanged += (s, e) =>
        {
            if (_isEditMode && grid.SelectionMode == DataGridViewSelectionMode.CellSelect)
            {
                _selectedEditCells.Clear();
                foreach (DataGridViewCell cell in grid.SelectedCells)
                {
                    if (cell.RowIndex >= 0 && cell.ColumnIndex >= 0)
                    {
                        _selectedEditCells.Add(cell);
                    }
                }
            }
        };

        // Handle mouse clicks to set selection anchor in edit mode
        grid.CellMouseDown += (s, e) =>
        {
            if (_isEditMode && e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // Set the clicked cell as the new anchor for Shift+Arrow operations
                try
                {
                    SetSelectionAnchor(grid.Rows[e.RowIndex].Cells[e.ColumnIndex]);
                }
                catch (System.Exception)
                {
                    // Ignore errors in virtual mode
                }
            }
        };

        // Form load - restore original sizing logic
        form.Load += delegate
        {
            grid.AutoResizeColumns();

            int padding = 20;
            int rowsHeight = grid.Rows.GetRowsHeight(DataGridViewElementStates.Visible);
            int reqHeight = rowsHeight + grid.ColumnHeadersHeight +
                            2 * grid.RowTemplate.Height +
                            SystemInformation.HorizontalScrollBarHeight + 30;

            int availHeight = Screen.PrimaryScreen.WorkingArea.Height - padding * 2;
            form.Height = Math.Min(reqHeight, availHeight);

            if (spanAllScreens)
            {
                _currentScreenState = Screen.AllScreens.Length;
                form.Width = Screen.AllScreens.Sum(s => s.WorkingArea.Width);
                form.Location = new Point(
                    Screen.AllScreens.Min(s => s.Bounds.X),
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }
            else
            {
                _currentScreenState = 0;
                int reqWidth = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                              + SystemInformation.VerticalScrollBarWidth + 43;
                form.Width = Math.Min(reqWidth, Screen.PrimaryScreen.WorkingArea.Width - padding * 2);
                form.Location = new Point(
                    (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2);
            }

            // Store original bounds and screen for cycling
            _originalFormBounds = form.Bounds;
            _originalScreen = Screen.FromControl(form);
        };

        // Add controls and show
        form.Controls.Add(grid);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        return selectedEntries;
    }

    private static void ExportToCsv(DataGridView grid, List<Dictionary<string, object>> data, string filePath)
    {
        var visibleColumns = grid.Columns.Cast<DataGridViewColumn>()
            .Where(c => c.Visible)
            .OrderBy(c => c.DisplayIndex)
            .ToList();

        using (var writer = new StreamWriter(filePath))
        {
            // Write header
            writer.WriteLine(string.Join(",", visibleColumns.Select(c => CsvQuote(c.HeaderText))));

            // Write rows
            foreach (var row in data)
            {
                var values = visibleColumns.Select(c => CsvQuote(row.ContainsKey(c.Name) ? row[c.Name]?.ToString() ?? "" : ""));
                writer.WriteLine(string.Join(",", values));
            }
        }
    }

    private static string CsvQuote(string value)
    {
        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    private static string FormatColumnHeader(string columnName)
    {
        if (string.IsNullOrEmpty(columnName))
            return columnName;

        var result = new StringBuilder();

        for (int i = 0; i < columnName.Length; i++)
        {
            char c = columnName[i];

            // Replace underscores with spaces
            if (c == '_')
            {
                result.Append(' ');
            }
            // Add space before uppercase letters (except at start)
            else if (i > 0 && char.IsUpper(c) && !char.IsUpper(columnName[i - 1]))
            {
                result.Append(' ');
                result.Append(char.ToLower(c));
            }
            // Convert to lowercase
            else
            {
                result.Append(char.ToLower(c));
            }
        }

        return result.ToString();
    }

    private static void CycleScreenExpansion(Form form)
    {
        Screen[] allScreens = Screen.AllScreens.OrderBy(s => s.Bounds.X).ToArray();
        Screen currentScreen = Screen.FromControl(form);

        // Set window state to normal first (in case it was maximized)
        form.WindowState = FormWindowState.Normal;

        if (allScreens.Length <= 1)
        {
            // Only one screen available, toggle between original size and maximized
            if (_currentScreenState == 0)
            {
                form.WindowState = FormWindowState.Maximized;
                _currentScreenState = 1;
            }
            else
            {
                form.Bounds = _originalFormBounds;
                _currentScreenState = 0;
            }
            return;
        }

        // Multi-screen cycling logic
        if (_currentScreenState == 0)
        {
            // State 0: Original -> State 1: Expand to screen on right
            int currentIndex = Array.IndexOf(allScreens, currentScreen);
            int rightIndex = (currentIndex + 1) % allScreens.Length;
            Screen rightScreen = allScreens[rightIndex];

            // Calculate combined bounds of current and right screen
            int leftX = Math.Min(currentScreen.WorkingArea.X, rightScreen.WorkingArea.X);
            int rightX = Math.Max(currentScreen.WorkingArea.Right, rightScreen.WorkingArea.Right);
            int topY = Math.Min(currentScreen.WorkingArea.Y, rightScreen.WorkingArea.Y);
            int maxHeight = Math.Max(currentScreen.WorkingArea.Height, rightScreen.WorkingArea.Height);

            form.Location = new Point(leftX, topY);
            form.Width = rightX - leftX;
            form.Height = Math.Min(form.Height, maxHeight);
            _currentScreenState = 1;
        }
        else if (_currentScreenState == 1)
        {
            // State 1: Right expansion -> State 2: Expand to screen on left
            int currentIndex = Array.IndexOf(allScreens, _originalScreen);
            int leftIndex = currentIndex == 0 ? allScreens.Length - 1 : currentIndex - 1;
            Screen leftScreen = allScreens[leftIndex];

            // Calculate combined bounds of original and left screen
            int leftX = Math.Min(_originalScreen.WorkingArea.X, leftScreen.WorkingArea.X);
            int rightX = Math.Max(_originalScreen.WorkingArea.Right, leftScreen.WorkingArea.Right);
            int topY = Math.Min(_originalScreen.WorkingArea.Y, leftScreen.WorkingArea.Y);
            int maxHeight = Math.Max(_originalScreen.WorkingArea.Height, leftScreen.WorkingArea.Height);

            form.Location = new Point(leftX, topY);
            form.Width = rightX - leftX;
            form.Height = Math.Min(form.Height, maxHeight);
            _currentScreenState = 2;
        }
        else if (_currentScreenState < allScreens.Length)
        {
            // Continue expanding to include more screens until all are covered
            _currentScreenState++;

            // Calculate bounds for first N screens
            var screensToInclude = allScreens.Take(_currentScreenState).ToArray();
            int leftX = screensToInclude.Min(s => s.WorkingArea.X);
            int rightX = screensToInclude.Max(s => s.WorkingArea.Right);
            int topY = screensToInclude.Min(s => s.WorkingArea.Y);
            int maxHeight = screensToInclude.Max(s => s.WorkingArea.Height);

            form.Location = new Point(leftX, topY);
            form.Width = rightX - leftX;
            form.Height = Math.Min(form.Height, maxHeight);
        }
        else
        {
            // All screens covered -> Reset to original
            form.Bounds = _originalFormBounds;
            _currentScreenState = 0;
        }

        // Bring form to front
        form.BringToFront();
        form.Focus();
    }
}
