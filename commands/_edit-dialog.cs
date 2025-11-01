using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace AutoCADCommands
{
    /// <summary>
    /// Reusable advanced text editing dialog with pattern/find/replace functionality.
    /// Supports both cell editing (for DataGrid) and text entity editing.
    /// </summary>
    public class AdvancedEditDialog : WinForms.Form
    {
        private readonly List<string> _originalValues;
        private readonly List<Dictionary<string, object>> _dataRows;

        private WinForms.TextBox _txtPattern;
        private WinForms.TextBox _txtFind;
        private WinForms.TextBox _txtReplace;
        private WinForms.TextBox _txtMath;
        private WinForms.RichTextBox _rtbBefore;
        private WinForms.RichTextBox _rtbAfter;
        private WinForms.Button _btnRegexToggle;
        private WinForms.Panel _findBorderPanel;
        private WinForms.Panel _findInnerPanel;
        private WinForms.ToolTip _toolTip;
        private bool _isRegexMode = false;

        #region Exposed properties

        public string PatternText => _txtPattern.Text;
        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string MathOperationText => _txtMath.Text;
        public bool IsRegexMode => _isRegexMode;

        #endregion

        public AdvancedEditDialog(List<string> originalValues, List<Dictionary<string, object>> dataRows = null, string dialogTitle = "Advanced Editor")
        {
            _originalValues = originalValues ?? new List<string>();
            _dataRows = dataRows ?? new List<Dictionary<string, object>>();

            Text = dialogTitle;
            BuildUI();
            LoadCurrentValues();
            InitializePatternValue();

            // Ensure focus is set when the dialog is shown
            this.Shown += (s, e) =>
            {
                _txtPattern.Focus();
                _txtPattern.SelectAll();
            };
        }

        private void BuildUI()
        {
            Font = new Drawing.Font("Segoe UI", 9);
            MinimumSize = new Drawing.Size(520, 480);
            Size = new Drawing.Size(640, 560);
            KeyPreview = true;
            KeyDown += (s, e) => { if (e.KeyCode == WinForms.Keys.Escape) Close(); };

            // Initialize tooltip
            _toolTip = new WinForms.ToolTip
            {
                AutoPopDelay = 5000,
                InitialDelay = 500,
                ReshowDelay = 100
            };

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
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 3: Replace
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 4: Math
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 22)); // Row 5: Math hint
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));

            // Pattern (moved to top)
            grid.Controls.Add(MakeLabel("Pattern:"), 0, 0);
            var patternBorderPanel = MakeBorderedTextBox(out _txtPattern, "{}");
            grid.Controls.Add(patternBorderPanel, 1, 0);

            grid.Controls.Add(MakeHint("Use {} for current value. Use $\"Column Name\" or $ColumnName to reference other columns."), 1, 1);

            // Find / Replace
            grid.Controls.Add(MakeLabel("Find:"), 0, 2);

            // Create a container for Find field with regex toggle button
            var findContainer = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                Padding = new WinForms.Padding(0)
            };

            // Create a border panel for custom focus border effect
            _findBorderPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                Padding = new WinForms.Padding(1),
                BackColor = Drawing.SystemColors.ControlDark
            };

            _txtFind = new WinForms.TextBox
            {
                Dock = WinForms.DockStyle.Fill,
                BorderStyle = WinForms.BorderStyle.None,
                Font = new Drawing.Font("Segoe UI", 9)
            };
            _txtFind.KeyDown += OnFindKeyDown;
            _txtFind.Enter += OnFindEnter;
            _txtFind.Leave += OnFindLeave;

            // Inner panel for TextBox to add padding
            _findInnerPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                Padding = new WinForms.Padding(3, 2, 3, 2),
                BackColor = Drawing.SystemColors.Window
            };
            _findInnerPanel.Controls.Add(_txtFind);
            _findBorderPanel.Controls.Add(_findInnerPanel);

            _btnRegexToggle = new WinForms.Button
            {
                Text = ".*",
                Width = 26,
                Height = 18,
                Dock = WinForms.DockStyle.Right,
                FlatStyle = WinForms.FlatStyle.Flat,
                BackColor = Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = Drawing.Color.Gray,
                Font = new Drawing.Font("Consolas", 8, Drawing.FontStyle.Bold),
                Cursor = WinForms.Cursors.Hand,
                TabStop = false,
                Margin = new WinForms.Padding(4, 0, 0, 0)
            };
            _btnRegexToggle.FlatAppearance.BorderColor = Drawing.Color.FromArgb(200, 200, 200);
            _btnRegexToggle.Click += OnRegexToggleClick;

            // Set tooltip for regex button
            _toolTip.SetToolTip(_btnRegexToggle, "Toggle regex mode (Ctrl+R)\nWhen enabled, Find field uses regular expressions");

            // Add spacing panel between find field and button
            var spacingPanel = new WinForms.Panel
            {
                Width = 4,
                Dock = WinForms.DockStyle.Right
            };

            findContainer.Controls.Add(_findBorderPanel);
            findContainer.Controls.Add(spacingPanel);
            findContainer.Controls.Add(_btnRegexToggle);

            grid.Controls.Add(findContainer, 1, 2);

            grid.Controls.Add(MakeLabel("Replace:"), 0, 3);
            var replaceBorderPanel = MakeBorderedTextBox(out _txtReplace);
            grid.Controls.Add(replaceBorderPanel, 1, 3);

            // Math
            grid.Controls.Add(MakeLabel("Math:"), 0, 4);
            var mathBorderPanel = MakeBorderedTextBox(out _txtMath);
            grid.Controls.Add(mathBorderPanel, 1, 4);

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

            buttons.Controls.Add(btnCancel);
            buttons.Controls.Add(btnOK);

            Controls.Add(buttons);
            Controls.Add(grid);

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            // events
            _txtFind.TextChanged += (s, e) => RefreshPreview();
            _txtReplace.TextChanged += (s, e) => RefreshPreview();
            _txtPattern.TextChanged += (s, e) => RefreshPreview();
            _txtMath.TextChanged += (s, e) => RefreshPreview();
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

        /// <summary>Create a bordered textbox with custom focus border effect</summary>
        private WinForms.Panel MakeBorderedTextBox(out WinForms.TextBox textBox, string initial = "")
        {
            // Create border panel
            var borderPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                Padding = new WinForms.Padding(1),
                BackColor = Drawing.SystemColors.ControlDark
            };

            // Create textbox
            textBox = new WinForms.TextBox
            {
                Text = initial,
                Dock = WinForms.DockStyle.Fill,
                BorderStyle = WinForms.BorderStyle.None,
                Font = new Drawing.Font("Segoe UI", 9)
            };

            // Inner panel for padding
            var innerPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                Padding = new WinForms.Padding(3, 2, 3, 2),
                BackColor = Drawing.SystemColors.Window
            };

            // Wire up focus events
            textBox.Enter += (s, e) => borderPanel.BackColor = Drawing.Color.FromArgb(0, 120, 215); // Blue focus
            textBox.Leave += (s, e) => borderPanel.BackColor = Drawing.SystemColors.ControlDark; // Gray unfocused

            innerPanel.Controls.Add(textBox);
            borderPanel.Controls.Add(innerPanel);

            return borderPanel;
        }

        private static WinForms.Label MakeHint(string txt) =>
            new WinForms.Label
            {
                Text = txt,
                Dock = WinForms.DockStyle.Fill,
                ForeColor = Drawing.Color.Gray,
                Font = new Drawing.Font("Segoe UI", 8, Drawing.FontStyle.Italic)
            };

        private void OnFindKeyDown(object sender, WinForms.KeyEventArgs e)
        {
            // Toggle regex mode with Ctrl+R
            if (e.Control && e.KeyCode == WinForms.Keys.R)
            {
                ToggleRegexMode();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void OnFindEnter(object sender, EventArgs e)
        {
            UpdateFindBorderColor(true);
        }

        private void OnFindLeave(object sender, EventArgs e)
        {
            UpdateFindBorderColor(false);
        }

        private void OnRegexToggleClick(object sender, EventArgs e)
        {
            ToggleRegexMode();
        }

        private void ToggleRegexMode()
        {
            _isRegexMode = !_isRegexMode;
            UpdateRegexButtonAppearance();
            UpdateFindBorderColor(_txtFind.Focused);
            RefreshPreview();
        }

        private void UpdateRegexButtonAppearance()
        {
            if (_isRegexMode)
            {
                _btnRegexToggle.BackColor = Drawing.Color.FromArgb(147, 112, 219); // Medium purple
                _btnRegexToggle.ForeColor = Drawing.Color.White;
                _findInnerPanel.BackColor = Drawing.Color.FromArgb(252, 250, 255); // Very subtle purple hue
            }
            else
            {
                _btnRegexToggle.BackColor = Drawing.Color.FromArgb(240, 240, 240);
                _btnRegexToggle.ForeColor = Drawing.Color.Gray;
                _findInnerPanel.BackColor = Drawing.SystemColors.Window; // White (default)
            }
        }

        private void UpdateFindBorderColor(bool hasFocus)
        {
            if (hasFocus)
            {
                // Purple border when in regex mode and focused, blue otherwise
                _findBorderPanel.BackColor = _isRegexMode
                    ? Drawing.Color.FromArgb(147, 112, 219)  // Medium purple
                    : Drawing.Color.FromArgb(0, 120, 215);   // Blue (Windows accent)
            }
            else
            {
                // Gray border when not focused
                _findBorderPanel.BackColor = Drawing.SystemColors.ControlDark;
            }
        }

        private void LoadCurrentValues()
        {
            _rtbBefore.Clear();
            foreach (var value in _originalValues)
            {
                // Escape newlines for display
                string displayValue = EscapeForDisplay(value);
                _rtbBefore.AppendText(displayValue + Environment.NewLine);
            }
            RefreshPreview();
        }

        /// <summary>Initialize pattern value - replace {} with actual value if all selected values are the same</summary>
        private void InitializePatternValue()
        {
            // Check if all original values are the same (or only one value)
            if (_originalValues.Count <= 1 || _originalValues.All(v => v == _originalValues[0]))
            {
                // All values are the same, replace {} with the actual value
                string actualValue = _originalValues.Count > 0 ? EscapeForDisplay(_originalValues[0]) : "";
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

        private void RefreshPreview()
        {
            try
            {
                _rtbAfter.Clear();
                for (int i = 0; i < _originalValues.Count; i++)
                {
                    string originalValue = _originalValues[i];
                    var dataRow = i < _dataRows.Count ? _dataRows[i] : null;

                    // Apply the transformation using the same logic as DataRenamerHelper
                    string transformedValue = TransformValue(originalValue, dataRow);

                    // Escape newlines for display in preview
                    string displayValue = EscapeForDisplay(transformedValue);
                    _rtbAfter.AppendText(displayValue + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                _rtbAfter.Text = "Error in preview: " + ex.Message;
            }
        }

        /// <summary>
        /// Transform a value using pattern/find/replace/math operations.
        /// Find/Replace and Math operations take precedence and are applied to the pattern result.
        /// This is a simplified version of the DataRenamerHelper.TransformValue logic.
        /// </summary>
        private string TransformValue(string originalValue, Dictionary<string, object> dataRow)
        {
            string result = originalValue;

            // 1. Pattern transformation (applied first as base)
            if (!string.IsNullOrEmpty(_txtPattern.Text))
            {
                string patternResult = UnescapeFromDisplay(_txtPattern.Text);

                // Replace {} with the original value
                patternResult = patternResult.Replace("{}", originalValue);

                // Replace column references if dataRow is available
                if (dataRow != null)
                {
                    foreach (var kvp in dataRow)
                    {
                        string columnValue = kvp.Value?.ToString() ?? "";
                        // Replace both quoted and unquoted column references
                        patternResult = patternResult.Replace($"$\"{kvp.Key}\"", columnValue);
                        patternResult = patternResult.Replace($"${kvp.Key}", columnValue);
                    }
                }

                result = patternResult;
            }

            // 2. Find/Replace transformation (takes precedence - applied to pattern result)
            if (!string.IsNullOrEmpty(_txtFind.Text))
            {
                string findText = UnescapeFromDisplay(_txtFind.Text);
                string replaceText = UnescapeFromDisplay(_txtReplace.Text ?? "");

                if (_isRegexMode)
                {
                    try
                    {
                        // Use regex mode for find/replace
                        result = Regex.Replace(result, findText, replaceText);
                    }
                    catch (ArgumentException)
                    {
                        // Invalid regex pattern - fallback to literal replacement
                        result = result.Replace(findText, replaceText);
                    }
                }
                else
                {
                    // Use simple string replacement
                    result = result.Replace(findText, replaceText);
                }
            }

            // 3. Math transformation (takes precedence - applied to result of pattern + find/replace)
            if (!string.IsNullOrEmpty(_txtMath.Text))
            {
                result = ApplyMathToAlphanumeric(result, _txtMath.Text);
            }

            return result;
        }

        /// <summary>
        /// Apply math operations to alphanumeric strings by extracting and modifying numeric parts.
        /// Examples: "W1" + "x+3" → "W4", "Room12" + "x*2" → "Room24"
        /// </summary>
        private string ApplyMathToAlphanumeric(string input, string mathExpression)
        {
            try
            {
                // First try pure numeric approach (existing behavior)
                if (double.TryParse(input, out double numericValue))
                {
                    string expr = mathExpression.Replace("x", numericValue.ToString());
                    var dataTable = new System.Data.DataTable();
                    var computedValue = dataTable.Compute(expr, null);
                    if (computedValue != DBNull.Value)
                    {
                        return computedValue.ToString();
                    }
                    return input;
                }

                // Handle alphanumeric strings - extract numeric parts
                var matches = System.Text.RegularExpressions.Regex.Matches(input, @"\d+");
                if (matches.Count == 0)
                {
                    // No numbers found, return original
                    return input;
                }

                // Apply math to each numeric part
                string result = input;
                for (int i = matches.Count - 1; i >= 0; i--) // Process in reverse to maintain positions
                {
                    var match = matches[i];
                    if (double.TryParse(match.Value, out double numberValue))
                    {
                        string expr = mathExpression.Replace("x", numberValue.ToString());
                        var dataTable = new System.Data.DataTable();
                        var computedValue = dataTable.Compute(expr, null);
                        if (computedValue != DBNull.Value)
                        {
                            // Format the result to remove unnecessary decimal places for integers
                            string newNumberStr = FormatNumericResult(computedValue);
                            result = result.Substring(0, match.Index) + newNumberStr + result.Substring(match.Index + match.Length);
                        }
                    }
                }
                return result;
            }
            catch
            {
                return input; // If anything fails, return original value
            }
        }

        /// <summary>
        /// Format numeric result to avoid unnecessary decimal places for whole numbers
        /// </summary>
        private string FormatNumericResult(object computedValue)
        {
            if (computedValue is double doubleVal)
            {
                // If it's a whole number, format as integer
                if (doubleVal == Math.Floor(doubleVal))
                {
                    return ((long)doubleVal).ToString();
                }
                return doubleVal.ToString();
            }
            return computedValue.ToString();
        }

        /// <summary>
        /// Escape newlines and other special characters for display in text fields
        /// Skip escaping if text contains AutoCAD MText formatting codes
        /// </summary>
        private static string EscapeForDisplay(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // If text contains AutoCAD MText formatting codes, don't escape - show as-is
            if (text.Contains("\\P") || text.Contains("\\p"))
                return text;

            return text
                .Replace("\\", "\\\\")
                .Replace("\r\n", "\\n")  // Convert \r\n to \n first
                .Replace("\n", "\\n")   // Then convert remaining \n
                .Replace("\r", "\\n")   // Convert standalone \r to \n
                .Replace("\t", "\\t");
        }

        /// <summary>
        /// Unescape display characters back to actual characters
        /// Skip unescaping if text contains AutoCAD MText formatting codes
        /// </summary>
        private static string UnescapeFromDisplay(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // If text contains AutoCAD MText formatting codes, don't unescape - use as-is
            if (text.Contains("\\P") || text.Contains("\\p"))
                return text;

            // Process in specific order to handle escaped sequences properly
            return text
                .Replace("\\n", "\n")   // Convert \n back to actual newlines
                .Replace("\\t", "\t")   // Convert \t back to actual tabs
                .Replace("\\\\", "\\"); // Convert \\ back to single backslash (must be last)
        }
    }
}