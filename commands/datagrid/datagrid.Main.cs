using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    //  Internal ID tracking for stable edit tracking
    // ──────────────────────────────────────────────────────────────

    private static long _nextInternalId = 1;
    private const string INTERNAL_ID_KEY = "__DATAGRID_INTERNAL_ID__";

    /// <summary>
    /// Assigns unique internal IDs to all entries that don't already have one.
    /// This provides a stable identifier for edit tracking that is independent of:
    /// - Row position (which changes with filtering/sorting)
    /// - Data content (which can be edited)
    /// - Specific columns (like Handle or DocumentPath which may not exist)
    /// </summary>
    private static void AssignInternalIdsToEntries(List<Dictionary<string, object>> entries)
    {
        foreach (var entry in entries)
        {
            if (!entry.ContainsKey(INTERNAL_ID_KEY))
            {
                entry[INTERNAL_ID_KEY] = _nextInternalId++;
            }
        }
    }

    /// <summary>
    /// Gets the internal ID for an entry, or creates one if it doesn't exist
    /// </summary>
    public static long GetInternalId(Dictionary<string, object> entry)
    {
        if (entry.ContainsKey(INTERNAL_ID_KEY))
        {
            return Convert.ToInt64(entry[INTERNAL_ID_KEY]);
        }

        // Assign a new ID if one doesn't exist
        long newId = _nextInternalId++;
        entry[INTERNAL_ID_KEY] = newId;
        return newId;
    }

    // ──────────────────────────────────────────────────────────────
    //  Command name inference from call stack
    // ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Infers the AutoCAD command name from the call stack by looking for methods
    /// with [CommandMethod] attributes or recognizable command patterns.
    /// </summary>
    private static string InferCommandNameFromCallStack()
    {
        try
        {
            var stackTrace = new StackTrace();
            var frames = stackTrace.GetFrames();

            if (frames == null) return null;

            // Look through the call stack for methods that look like AutoCAD commands
            for (int i = 0; i < frames.Length; i++)
            {
                var method = frames[i].GetMethod();
                if (method == null) continue;

                var declaringType = method.DeclaringType;
                if (declaringType == null) continue;

                // Skip our own DataGrid-related methods
                if (declaringType.Name == "CustomGUIs" ||
                    declaringType.Name.Contains("DataGrid"))
                    continue;

                // Look for methods that might be AutoCAD commands:
                // 1. Class names ending with "Command"
                // 2. Method names matching common command patterns
                // 3. Namespace AutoCADBallet

                if (declaringType.Namespace == "AutoCADBallet" ||
                    declaringType.Name.EndsWith("Command") ||
                    method.Name.Contains("Execute"))
                {
                    // Try to extract command name from class name first
                    string className = declaringType.Name;

                    // Remove "Command" suffix if present
                    if (className.EndsWith("Command"))
                    {
                        className = className.Substring(0, className.Length - "Command".Length);
                    }

                    // Convert PascalCase to kebab-case
                    string commandName = ConvertToKebabCase(className);

                    // If we got a reasonable command name, return it
                    if (!string.IsNullOrWhiteSpace(commandName) &&
                        commandName.Length > 2 &&
                        commandName != "custom-gu-is")
                    {
                        return commandName;
                    }
                }
            }
        }
        catch
        {
            // If stack trace inspection fails, silently return null
        }

        return null;
    }

    /// <summary>
    /// Converts PascalCase to kebab-case (e.g., "SwitchView" -> "switch-view")
    /// </summary>
    private static string ConvertToKebabCase(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return input;

        // Insert hyphen before uppercase letters (except first character)
        var kebabCase = Regex.Replace(input, "(?<!^)([A-Z])", "-$1");

        // Convert to lowercase
        return kebabCase.ToLower();
    }

    // ──────────────────────────────────────────────────────────────
    //  Main DataGrid Method
    // ──────────────────────────────────────────────────────────────

    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null,
        Func<List<Dictionary<string, object>>, bool> onDeleteEntries = null,
        bool allowCreateFromSearch = false,
        string commandName = null)
    {
        if (entries == null || propertyNames == null || propertyNames.Count == 0)
            return new List<Dictionary<string, object>>();

        // Allow empty entries when allowCreateFromSearch is enabled (user can type new values)
        if (entries.Count == 0 && !allowCreateFromSearch)
            return new List<Dictionary<string, object>>();

        // Auto-infer command name from call stack if not provided
        if (string.IsNullOrWhiteSpace(commandName))
        {
            commandName = InferCommandNameFromCallStack();
        }

        // CRITICAL: Assign unique internal IDs to all entries for stable edit tracking
        // This ensures edits are correctly applied even when filters change row order
        AssignInternalIdsToEntries(entries);

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
            Text = "Selected: 0, Filtered: " + entries.Count + ", Total: " + entries.Count,
            BackColor = Color.White,
            ShowIcon = false,
            KeyPreview = true  // Allow form to intercept keys before controls
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
            ScrollBars = ScrollBars.Both,
            StandardTab = true,  // Allow Tab key to be handled by form instead of DataGridView
            ShowCellToolTips = false  // Disable tooltips on hover
        };

        // Disable built-in sorting
        grid.SortCompare += (sender, e) =>
        {
            e.Handled = true;
            e.SortResult = naturalComparer.Compare(e.CellValue1, e.CellValue2);
        };

        // Add columns (skip internal ID column)
        foreach (string col in propertyNames)
        {
            // Skip the internal tracking ID - never display it to users
            if (col == INTERNAL_ID_KEY)
                continue;

            var column = new DataGridViewTextBoxColumn
            {
                Name = col,
                HeaderText = FormatColumnHeader(col),
                DataPropertyName = col,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            grid.Columns.Add(column);
        }

        // Search box with dropdown button container
        Panel searchPanel = new Panel { Dock = DockStyle.Top, Height = 21 };

        TextBox searchBox = new TextBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle
        };

        // Dropdown button for search history
        Button dropdownButton = new Button
        {
            Dock = DockStyle.Right,
            Width = 20,
            Height = 21,
            Text = "▾",
            FlatStyle = FlatStyle.Flat,
            TabStop = false,
            Cursor = Cursors.Hand,
            ForeColor = Color.DimGray,
            BackColor = Color.White,
            Font = new Font("Arial", 9, FontStyle.Regular),
            Margin = new Padding(2, 0, 0, 0)
        };

        dropdownButton.FlatAppearance.BorderSize = 1;
        dropdownButton.FlatAppearance.BorderColor = SystemColors.ControlDark;

        // Add tooltip to dropdown button
        ToolTip dropdownTooltip = new ToolTip();
        dropdownTooltip.SetToolTip(dropdownButton, "Show search history (F4)");

        // Only show dropdown button if commandName is provided
        if (!string.IsNullOrWhiteSpace(commandName))
        {
            searchPanel.Controls.Add(searchBox);
            searchPanel.Controls.Add(dropdownButton);
        }
        else
        {
            // No command name, just use searchBox without button
            searchBox.Dock = DockStyle.Top;
        }

        // Track last search text for history recording
        string lastSearchTextForHistory = "";

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

        // Helper to update form title
        Action UpdateFormTitle = () =>
        {
            int selectedCount = grid.SelectedRows.Count;
            int filteredCount = workingSet.Count;
            int totalCount = entries.Count;
            string editModeIndicator = _isEditMode ? " [EDIT MODE]" : "";
            form.Text = "Selected: " + selectedCount + ", Filtered: " + filteredCount + ", Total: " + totalCount + editModeIndicator;
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
            UpdateFormTitle();

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

        // Record search query when focus leaves searchBox (user is done typing)
        searchBox.LostFocus += delegate
        {
            if (!string.IsNullOrWhiteSpace(commandName) &&
                !string.IsNullOrWhiteSpace(searchBox.Text) &&
                searchBox.Text.Trim() != lastSearchTextForHistory)
            {
                SearchQueryHistory.RecordQuery(commandName, searchBox.Text.Trim());
                lastSearchTextForHistory = searchBox.Text.Trim();
            }
        };

        // Dropdown button click - show search history
        if (!string.IsNullOrWhiteSpace(commandName))
        {
            dropdownButton.Click += delegate
            {
                ShowSearchHistoryDropdown(searchBox, commandName);
            };
        }

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
            if (e.KeyCode == Keys.F4 && sender == searchBox && !string.IsNullOrWhiteSpace(commandName))
            {
                // Show search history dropdown
                ShowSearchHistoryDropdown(searchBox, commandName);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.V && e.Control && _isEditMode && sender == grid)
            {
                // Ctrl+V in edit mode: Paste clipboard values into selected cells
                HandleClipboardPaste(grid);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F2)
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
                    UpdateFormTitle();
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
                    UpdateFormTitle();
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
                // Delete key: Invoke callback if provided (callback handles confirmation)
                if (grid.SelectedRows.Count > 0)
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

                    // Invoke callback to handle deletion (callback shows its own confirmation dialog)
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

        // PreviewKeyDown to mark arrow keys as input keys in edit mode
        grid.PreviewKeyDown += (s, e) =>
        {
            if (_isEditMode && e.Shift && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
            {
                e.IsInputKey = true;  // Mark as input key so it reaches KeyDown handler
            }
        };

        grid.KeyDown += (s, e) => HandleKeyDown(e, grid);
        searchBox.KeyDown += (s, e) => HandleKeyDown(e, searchBox);

        // Form-level key handling to intercept Shift+Arrow in edit mode before DataGridView processes it
        form.KeyDown += (s, e) =>
        {
            // Only handle when grid has focus and we're in edit mode
            if (_isEditMode && grid.Focused && grid.CurrentCell != null)
            {
                if ((e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right) && e.Shift)
                {
                    // Shift+Arrow: Extend selection (Excel-like)
                    ExtendSelectionWithArrows(grid, e.KeyCode, true);
                    e.Handled = true;
                    e.SuppressKeyPress = true;  // Prevent further processing
                }
            }
        };

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

            // Update title to reflect current selection count
            UpdateFormTitle();
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
        if (!string.IsNullOrWhiteSpace(commandName))
        {
            form.Controls.Add(searchPanel);
        }
        else
        {
            form.Controls.Add(searchBox);
        }
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

    private static void ShowSearchHistoryDropdown(TextBox searchBox, string commandName)
    {
        // Get search history for this command
        var history = SearchQueryHistory.GetQueryHistory(commandName);

        if (history.Count == 0)
        {
            // No history to show
            return;
        }

        // Reverse to show most recent first
        history.Reverse();

        // Create a ListBox to show history
        ListBox historyList = new ListBox
        {
            Width = searchBox.Width,
            Height = Math.Min(history.Count * 20 + 4, 200), // Limit height
            BorderStyle = BorderStyle.FixedSingle,
            Font = searchBox.Font
        };

        // Add items to list
        foreach (string query in history)
        {
            historyList.Items.Add(query);
        }

        // Create a form to host the listbox (acts as dropdown)
        Form dropdown = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            StartPosition = FormStartPosition.Manual,
            ShowInTaskbar = false,
            TopMost = true,
            Width = historyList.Width,
            Height = historyList.Height
        };

        dropdown.Controls.Add(historyList);
        historyList.Dock = DockStyle.Fill;

        // Position below search box
        var searchBoxLocation = searchBox.PointToScreen(new Point(0, searchBox.Height));
        dropdown.Location = searchBoxLocation;

        // Handle selection
        historyList.Click += (s, e) =>
        {
            if (historyList.SelectedIndex >= 0)
            {
                searchBox.Text = historyList.SelectedItem.ToString();
                searchBox.SelectionStart = searchBox.Text.Length;
                dropdown.Close();
                searchBox.Focus();
            }
        };

        // Handle keyboard navigation
        historyList.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter && historyList.SelectedIndex >= 0)
            {
                searchBox.Text = historyList.SelectedItem.ToString();
                searchBox.SelectionStart = searchBox.Text.Length;
                dropdown.Close();
                searchBox.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                dropdown.Close();
                searchBox.Focus();
                e.Handled = true;
            }
        };

        // Close when focus is lost
        dropdown.Deactivate += (s, e) => dropdown.Close();

        // Select first item by default
        if (historyList.Items.Count > 0)
        {
            historyList.SelectedIndex = 0;
        }

        // Show dropdown
        dropdown.Show();
        historyList.Focus();
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
