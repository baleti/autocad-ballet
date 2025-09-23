using System;
using System.Collections.Generic;
using System.Linq;
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

        #region Exposed properties

        public string PatternText => _txtPattern.Text;
        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string MathOperationText => _txtMath.Text;

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
            _rtbBefore.Clear();
            foreach (var value in _originalValues)
            {
                _rtbBefore.AppendText(value + Environment.NewLine);
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

                    _rtbAfter.AppendText(transformedValue + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                _rtbAfter.Text = "Error in preview: " + ex.Message;
            }
        }

        /// <summary>
        /// Transform a value using pattern/find/replace/math operations.
        /// This is a simplified version of the DataRenamerHelper.TransformValue logic.
        /// </summary>
        private string TransformValue(string originalValue, Dictionary<string, object> dataRow)
        {
            string result = originalValue;

            // 1. Pattern transformation (highest priority)
            if (!string.IsNullOrEmpty(_txtPattern.Text))
            {
                result = _txtPattern.Text;

                // Replace {} with current value
                result = result.Replace("{}", originalValue);

                // Replace column references if dataRow is available
                if (dataRow != null)
                {
                    foreach (var kvp in dataRow)
                    {
                        string columnValue = kvp.Value?.ToString() ?? "";
                        // Replace both quoted and unquoted column references
                        result = result.Replace($"$\"{kvp.Key}\"", columnValue);
                        result = result.Replace($"${kvp.Key}", columnValue);
                    }
                }
            }
            // 2. Find/Replace transformation
            else if (!string.IsNullOrEmpty(_txtFind.Text))
            {
                result = result.Replace(_txtFind.Text, _txtReplace.Text ?? "");
            }
            // 3. Math transformation (if result is numeric)
            else if (!string.IsNullOrEmpty(_txtMath.Text))
            {
                if (double.TryParse(result, out double numericValue))
                {
                    try
                    {
                        string mathExpression = _txtMath.Text.Replace("x", numericValue.ToString());
                        // Simple math evaluation - this is a basic implementation
                        // For full functionality, you'd want to use a proper expression evaluator
                        var dataTable = new System.Data.DataTable();
                        var computedValue = dataTable.Compute(mathExpression, null);
                        if (computedValue != DBNull.Value)
                        {
                            result = computedValue.ToString();
                        }
                    }
                    catch
                    {
                        // If math evaluation fails, keep original value
                    }
                }
            }

            return result;
        }
    }
}