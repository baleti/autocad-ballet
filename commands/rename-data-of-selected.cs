using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.IO;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Alias WinForms and Drawing to avoid type ambiguities
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

// Register the command class
[assembly: CommandClass(typeof(RenameDataOfSelectedElements))]

namespace AutoCADBallet
{
    // =============================================================================================
    // 1. Shared helper - math, reflection and RenameData worker
    // =============================================================================================
    public static class DataRenamerHelper
    {
        #region DTOs ------------------------------------------------------------------------------

        public class DataInfo
        {
            public string Name;
            public string DataType; // "Property", "XData", "BlockAttribute", "CustomData", "DictExtension"
            public string Group;
            public bool IsReadOnly;
        }

        public class DataValueInfo
        {
            public Entity Element;        // owning entity
            public string DataName;
            public string CurrentValue;
            public string DataType;       // Property, XData, etc.
            public object DataSource;     // Parameter, XData key, AttributeDefinition, etc.
            public bool IsReadOnly;
        }

        #endregion

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
        private static string ParsePatternWithDataReferences(string pattern, string currentValue, Entity element)
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

                // Try getting data value from the entity
                string dataValue = GetEntityDataValue(element, dataName);
                return !string.IsNullOrEmpty(dataValue) ? dataValue : match.Value;
            });

            return result;
        }

        #endregion

        #region Entity data access helpers -------------------------------------------------------

        /// <summary>Gets data value as string from various AutoCAD data sources</summary>
        private static string GetEntityDataValue(Entity entity, string dataName)
        {
            if (entity == null || string.IsNullOrEmpty(dataName))
                return string.Empty;

            try
            {
                // Pseudo data properties
                switch (dataName.ToLowerInvariant())
                {
                    case "layer": return entity.Layer ?? string.Empty;
                    case "color": return entity.Color.ToString();
                    case "linetype": return entity.Linetype ?? string.Empty;
                    case "handle": return entity.Handle.ToString();
                    case "objecttype": return entity.GetType().Name;
                }

                // Try to get XData
                var xdata = entity.XData;
                if (xdata != null)
                {
                    var values = xdata.AsArray();
                    for (int i = 0; i < values.Length - 1; i += 2)
                    {
                        if (values[i].TypeCode == (short)DxfCode.ExtendedDataRegAppName &&
                            values[i].Value.ToString() == dataName)
                        {
                            return values[i + 1].Value?.ToString() ?? string.Empty;
                        }
                    }
                }

                // Try block attributes for BlockReference
                if (entity is BlockReference br && br.AttributeCollection.Count > 0)
                {
                    using (var tr = br.Database.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (att != null && att.Tag.Equals(dataName, StringComparison.OrdinalIgnoreCase))
                            {
                                return att.TextString ?? string.Empty;
                            }
                        }
                        tr.Commit();
                    }
                }

                // Try extension dictionary
                if (entity.ExtensionDictionary != ObjectId.Null)
                {
                    using (var tr = entity.Database.TransactionManager.StartTransaction())
                    {
                        var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                        if (extDict.Contains(dataName))
                        {
                            var dictEntry = tr.GetObject(extDict.GetAt(dataName), OpenMode.ForRead);
                            // This would need specific handling based on what type of object is stored
                            return dictEntry?.ToString() ?? string.Empty;
                        }
                        tr.Commit();
                    }
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        #region Value transformation --------------------------------------------------------------

        private static string TransformValue(string original, RenameDataForm form, Entity element = null)
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
                if (element != null)
                {
                    value = ParsePatternWithDataReferences(form.PatternText, value, element);
                }
                else
                {
                    // Fallback for preview mode (no element available)
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

        #region Data setter -----------------------------------------------------------------------

        private static bool SetEntityDataValue(Entity entity, string dataName, string newValue, string dataType, object dataSource)
        {
            if (entity == null || string.IsNullOrEmpty(dataName))
                return false;

            try
            {
                using (var tr = entity.Database.TransactionManager.StartTransaction())
                {
                    var entForWrite = tr.GetObject(entity.ObjectId, OpenMode.ForWrite) as Entity;

                    switch (dataType)
                    {
                        case "Property":
                            // Handle basic entity properties
                            switch (dataName.ToLowerInvariant())
                            {
                                case "layer":
                                    entForWrite.Layer = newValue;
                                    break;
                                case "linetype":
                                    entForWrite.Linetype = newValue;
                                    break;
                                // Color and other complex properties would need more sophisticated handling
                            }
                            break;

                        case "XData":
                            // Set XData value (simplified - would need proper XData handling)
                            SetXDataValue(entForWrite, dataName, newValue);
                            break;

                        case "BlockAttribute":
                            if (entForWrite is BlockReference br && dataSource is AttributeReference)
                            {
                                var att = tr.GetObject(((AttributeReference)dataSource).ObjectId, OpenMode.ForWrite) as AttributeReference;
                                if (att != null)
                                    att.TextString = newValue;
                            }
                            break;

                        case "DictExtension":
                            // Handle extension dictionary entries
                            SetExtensionDictionaryValue(entForWrite, dataName, newValue, tr);
                            break;
                    }

                    tr.Commit();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void SetXDataValue(Entity entity, string appName, string value)
        {
            // Simplified XData setting - real implementation would be more complex
            var rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, value)
            );
            entity.XData = rb;
            rb.Dispose();
        }

        private static void SetExtensionDictionaryValue(Entity entity, string key, string value, Transaction tr)
        {
            // Create extension dictionary if it doesn't exist
            if (entity.ExtensionDictionary == ObjectId.Null)
            {
                entity.CreateExtensionDictionary();
            }

            var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
            // This would need specific implementation based on what type of data we want to store
        }

        #endregion

        #region Main worker - RenameData ---------------------------------------------------------

        /// <summary>
        /// Lets the user pick data fields and renames them for all selected elements.
        /// - Properties: Layer, Color, LineType, etc.
        /// - XData: Extended data attached to entities
        /// - Block Attributes: For BlockReference entities
        /// - Extension Dictionary entries
        /// Two-pass scheme prevents temporary duplicates for string data.
        /// </summary>
        public static void RenameData(Document doc, List<Entity> entities)
        {
            try
            {
                // 1. Build list of editable data present on at least one entity
                var dataInfos = GatherEntityDataInfo(entities);

                if (dataInfos.Count == 0)
                {
                    WinForms.MessageBox.Show("No editable data found.");
                    return;
                }

                // 2. Data pick UI (small grid)
                var rows = new List<Dictionary<string, object>>();
                foreach (var di in dataInfos)
                {
                    rows.Add(new Dictionary<string, object>
                    {
                        { "Group", di.Group },
                        { "Name", di.Name },
                        { "Type", di.DataType }
                    });
                }
                var headers = new List<string> { "Group", "Name", "Type" };
                var picked = CustomGUIs.DataGrid(rows, headers, /*multiSelect=*/false);

                if (picked == null || picked.Count == 0)
                    return;

                // 3. Flatten to DataValueInfo list
                var dataVals = new List<DataValueInfo>();
                foreach (var entity in entities)
                {
                    foreach (var d in picked)
                    {
                        string dataName = d["Name"].ToString();
                        string dataType = d["Type"].ToString();

                        var dataValue = GetEntityDataValueInfo(entity, dataName, dataType);
                        if (dataValue != null)
                        {
                            dataVals.Add(dataValue);
                        }
                    }
                }

                if (dataVals.Count == 0)
                {
                    WinForms.MessageBox.Show("No valid data in selection.");
                    return;
                }

                // 4. Rename UI + write-back
                using (var form = new RenameDataForm(dataVals))
                {
                    if (form.ShowDialog() != WinForms.DialogResult.OK)
                        return;

                    // build update list BEFORE opening the transaction
                    var updates = new List<(DataValueInfo Dv, string NewVal)>();
                    foreach (var dv in dataVals)
                    {
                        string newVal = TransformValue(dv.CurrentValue, form, dv.Element);
                        if (newVal != dv.CurrentValue && !dv.IsReadOnly)
                            updates.Add((dv, newVal));
                    }
                    if (updates.Count == 0) return;

                    using (var tx = doc.TransactionManager.StartTransaction())
                    {

                        // === PASS 1 – temporary unique placeholders (string data only) =========
                        foreach (var u in updates)
                        {
                            if (u.Dv.DataType.Contains("String") || u.Dv.DataType == "BlockAttribute")
                            {
                                SetEntityDataValue(u.Dv.Element, u.Dv.DataName,
                                                 $"TMP_{Guid.NewGuid():N}",
                                                 u.Dv.DataType, u.Dv.DataSource);
                            }
                        }

                        // === PASS 2 – final values =============================================
                        foreach (var u in updates)
                        {
                            try
                            {
                                SetEntityDataValue(u.Dv.Element, u.Dv.DataName, u.NewVal,
                                                 u.Dv.DataType, u.Dv.DataSource);
                            }
                            catch (System.Exception ex)
                            {
                                WinForms.MessageBox.Show($"Error setting {u.Dv.DataName}: {ex.Message}", "Error");
                            }
                        }

                        tx.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                WinForms.MessageBox.Show(ex.ToString(), "Critical Error");
            }
        }

        private static List<DataInfo> GatherEntityDataInfo(List<Entity> entities)
        {
            var dataInfos = new List<DataInfo>();

            foreach (var entity in entities)
            {
                // Add basic properties
                dataInfos.Add(new DataInfo { Name = "Layer", DataType = "Property", Group = "Basic", IsReadOnly = false });
                dataInfos.Add(new DataInfo { Name = "LineType", DataType = "Property", Group = "Basic", IsReadOnly = false });
                dataInfos.Add(new DataInfo { Name = "Color", DataType = "Property", Group = "Basic", IsReadOnly = true });
                dataInfos.Add(new DataInfo { Name = "Handle", DataType = "Property", Group = "Basic", IsReadOnly = true });

                // Add XData if present
                var xdata = entity.XData;
                if (xdata != null)
                {
                    var values = xdata.AsArray();
                    for (int i = 0; i < values.Length - 1; i += 2)
                    {
                        if (values[i].TypeCode == (short)DxfCode.ExtendedDataRegAppName)
                        {
                            string appName = values[i].Value.ToString();
                            if (!dataInfos.Any(d => d.Name == appName && d.DataType == "XData"))
                            {
                                dataInfos.Add(new DataInfo
                                {
                                    Name = appName,
                                    DataType = "XData",
                                    Group = "Extended",
                                    IsReadOnly = false
                                });
                            }
                        }
                    }
                }

                // Add block attributes if BlockReference
                if (entity is BlockReference br && br.AttributeCollection.Count > 0)
                {
                    try
                    {
                        using (var tr = br.Database.TransactionManager.StartTransaction())
                        {
                            foreach (ObjectId attId in br.AttributeCollection)
                            {
                                var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (att != null && !dataInfos.Any(d => d.Name == att.Tag && d.DataType == "BlockAttribute"))
                                {
                                    dataInfos.Add(new DataInfo
                                    {
                                        Name = att.Tag,
                                        DataType = "BlockAttribute",
                                        Group = "Attributes",
                                        IsReadOnly = att.IsConstant
                                    });
                                }
                            }
                            tr.Commit();
                        }
                    }
                    catch { /* Skip if can't read attributes */ }
                }
            }

            return dataInfos.GroupBy(d => d.Name + d.DataType)
                            .Select(g => g.First())
                            .OrderBy(d => d.Group)
                            .ThenBy(d => d.Name)
                            .ToList();
        }

        private static DataValueInfo GetEntityDataValueInfo(Entity entity, string dataName, string dataType)
        {
            try
            {
                string currentValue = GetEntityDataValue(entity, dataName);
                object dataSource = null;
                bool isReadOnly = false;

                // Get data source and readonly status based on type
                switch (dataType)
                {
                    case "Property":
                        isReadOnly = dataName.ToLowerInvariant() == "handle" ||
                                   dataName.ToLowerInvariant() == "color";
                        break;

                    case "BlockAttribute":
                        if (entity is BlockReference br)
                        {
                            using (var tr = br.Database.TransactionManager.StartTransaction())
                            {
                                foreach (ObjectId attId in br.AttributeCollection)
                                {
                                    var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                    if (att != null && att.Tag.Equals(dataName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        dataSource = att;
                                        isReadOnly = att.IsConstant;
                                        break;
                                    }
                                }
                                tr.Commit();
                            }
                        }
                        break;
                }

                return new DataValueInfo
                {
                    Element = entity,
                    DataName = dataName,
                    CurrentValue = currentValue,
                    DataType = dataType,
                    DataSource = dataSource,
                    IsReadOnly = isReadOnly
                };
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }

    // =============================================================================================
    // 2. WinForms preview dialog
    // =============================================================================================
    public class RenameDataForm : WinForms.Form
    {
        private readonly List<DataRenamerHelper.DataValueInfo> _dataVals;

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

        public RenameDataForm(List<DataRenamerHelper.DataValueInfo> dataVals)
        {
            _dataVals = dataVals;
            BuildUI();
            LoadCurrentValues();
        }

        private void BuildUI()
        {
            Text = "Modify Entity Data";
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
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 0: Find
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 36)); // Row 1: Replace (increased for spacing)
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 2: Pattern
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 22)); // Row 3: Pattern hint
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 32)); // Row 4: Math
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Absolute, 22)); // Row 5: Math hint
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));
            grid.RowStyles.Add(new WinForms.RowStyle(WinForms.SizeType.Percent, 45));

            // Find / Replace
            grid.Controls.Add(MakeLabel("Find:"), 0, 0);
            _txtFind = MakeTextBox();
            grid.Controls.Add(_txtFind, 1, 0);

            grid.Controls.Add(MakeLabel("Replace:"), 0, 1);
            _txtReplace = MakeTextBox();
            grid.Controls.Add(_txtReplace, 1, 1);

            // Pattern
            grid.Controls.Add(MakeLabel("Pattern:"), 0, 2);
            _txtPattern = MakeTextBox("{}");   // default
            grid.Controls.Add(_txtPattern, 1, 2);

            grid.Controls.Add(MakeHint("Use {} for current value. Use $\"Data Name\" or $DataName to reference other data."), 1, 3);

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
            foreach (var dv in _dataVals)
                _rtbBefore.AppendText(dv.CurrentValue + Environment.NewLine);

            RefreshPreview();
        }

        private void RefreshPreview()
        {
            _rtbAfter.Clear();
            foreach (var dv in _dataVals)
            {
                string v = dv.CurrentValue;

                // Apply same transformations as helper
                if (!string.IsNullOrEmpty(FindText))
                    v = v.Replace(FindText, ReplaceText);
                else if (!string.IsNullOrEmpty(ReplaceText))
                    v = ReplaceText;

                if (!string.IsNullOrEmpty(PatternText))
                {
                    // For preview, we can't access other data, so show placeholder
                    v = PatternText.Replace("{}", v);
                    // Show data references as-is in preview
                    var regex = new Regex(@"\$""([^""]+)""|(?<!\$)\$(\w+)");
                    v = regex.Replace(v, match => "[" + match.Value + "]");
                }

                if (!string.IsNullOrEmpty(MathOperationText))
                {
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                        v = DataRenamerHelper.ApplyMathOperation(n, MathOperationText).ToString(CultureInfo.InvariantCulture);
                    else
                        v = DataRenamerHelper.ApplyMathToNumbersInString(v, MathOperationText);
                }

                _rtbAfter.AppendText(v + Environment.NewLine);
            }
        }
    }

    // =============================================================================================
    // 3. External command – rename data of selection
    // =============================================================================================
    public class RenameDataOfSelectedElements
    {
        [CommandMethod("rename-data-of-selected", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void RenameDataOfSelectedCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                var entities = new List<Entity>();
                var currentScope = SelectionScopeManager.CurrentScope;

                // Handle selection based on current scope
                if (currentScope == SelectionScopeManager.SelectionScope.view)
                {
                    // Get pickfirst set (pre-selected objects)
                    var selResult = ed.SelectImplied();

                    // If there is no pickfirst set, request user to select objects
                    if (selResult.Status == PromptStatus.Error)
                    {
                        var selectionOpts = new PromptSelectionOptions();
                        selectionOpts.MessageForAdding = "\nSelect objects to rename data: ";
                        selResult = ed.GetSelection(selectionOpts);
                    }
                    else if (selResult.Status == PromptStatus.OK)
                    {
                        // Clear the pickfirst set since we're consuming it
                        ed.SetImpliedSelection(new ObjectId[0]);
                    }

                    if (selResult.Status == PromptStatus.OK && selResult.Value.Count > 0)
                    {
                        using (var tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            foreach (var objectId in selResult.Value.GetObjectIds())
                            {
                                try
                                {
                                    var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                                    if (entity != null)
                                    {
                                        entities.Add(entity);
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
                    }
                    else
                    {
                        ed.WriteMessage("\nNo selection found. Please select entities first when in 'view' scope.\n");
                        return;
                    }
                }
                else
                {
                    // Get entities from stored selection
                    var storedSelection = SelectionStorage.LoadSelection();
                    if (storedSelection == null || storedSelection.Count == 0)
                    {
                        ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-category' first.\n");
                        return;
                    }

                    // Process stored selection items (current document only)
                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        foreach (var item in storedSelection)
                        {
                            try
                            {
                                // Check if this is from the current document
                                if (Path.GetFullPath(item.DocumentPath) == Path.GetFullPath(doc.Name))
                                {
                                    var handle = Convert.ToInt64(item.Handle, 16);
                                    var objectId = doc.Database.GetObjectId(false, new Handle(handle), 0);

                                    if (objectId != ObjectId.Null)
                                    {
                                        var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                                        if (entity != null)
                                        {
                                            entities.Add(entity);
                                        }
                                    }
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
                }

                if (entities.Count == 0)
                {
                    ed.WriteMessage("\nNo entities found to process.\n");
                    return;
                }

                // Execute the rename operation
                DataRenamerHelper.RenameData(doc, entities);
                ed.WriteMessage($"\nProcessed {entities.Count} entities.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n");
            }
        }
    }
}