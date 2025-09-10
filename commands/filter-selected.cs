using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

[assembly: CommandClass(typeof(AutoCADBallet.FilterSelectedCommand))]

namespace AutoCADBallet
{
    public static class EntityDataHelper
    {
        public static List<Dictionary<string, object>> GetEntityData(Document doc, ObjectId[] objectIds, bool includeXData = false)
        {
            var entityData = new List<Dictionary<string, object>>();

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                foreach (var id in objectIds)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity != null)
                        {
                            var data = GetEntityDataDictionary(entity, doc, id, includeXData);
                            entityData.Add(data);
                        }
                    }
                    catch { /* Skip problematic entities */ }
                }
                tr.Commit();
            }

            return entityData;
        }

        private static Dictionary<string, object> GetEntityDataDictionary(Entity entity, Document doc, ObjectId id, bool includeXData)
        {
            var data = new Dictionary<string, object>
            {
                ["Handle"] = entity.Handle.ToString(),
                ["EntityType"] = entity.GetType().Name,
                ["Layer"] = entity.Layer,
                ["Color"] = GetColorString(entity),
                ["Linetype"] = entity.Linetype,
                ["LineWeight"] = entity.LineWeight.ToString(),
                ["ObjectId"] = id,
                ["DocumentPath"] = doc.Name,
                ["Space"] = GetSpaceName(entity, doc)
            };

            // Add block name if it's a block reference
            if (entity is BlockReference blockRef)
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var blockDef = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    data["BlockName"] = blockDef?.Name ?? "";
                    data["Position"] = $"{blockRef.Position.X:F2}, {blockRef.Position.Y:F2}, {blockRef.Position.Z:F2}";
                    data["Scale"] = $"{blockRef.ScaleFactors.X:F2}, {blockRef.ScaleFactors.Y:F2}, {blockRef.ScaleFactors.Z:F2}";
                    data["Rotation"] = $"{blockRef.Rotation * 180 / Math.PI:F2}°";

                    // Get attribute values
                    var attValues = new List<string>();
                    foreach (ObjectId attId in blockRef.AttributeCollection)
                    {
                        var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                        if (att != null && !string.IsNullOrEmpty(att.TextString))
                        {
                            attValues.Add($"{att.Tag}={att.TextString}");
                        }
                    }
                    if (attValues.Any())
                        data["Attributes"] = string.Join(", ", attValues);

                    tr.Commit();
                }
            }
            else
            {
                data["BlockName"] = "";
            }

            // Add specific properties based on entity type
            switch (entity)
            {
                case Line line:
                    data["Length"] = line.Length.ToString("F2");
                    data["StartPoint"] = $"{line.StartPoint.X:F2}, {line.StartPoint.Y:F2}";
                    data["EndPoint"] = $"{line.EndPoint.X:F2}, {line.EndPoint.Y:F2}";
                    break;

                case Circle circle:
                    data["Radius"] = circle.Radius.ToString("F2");
                    data["Center"] = $"{circle.Center.X:F2}, {circle.Center.Y:F2}";
                    data["Area"] = (Math.PI * circle.Radius * circle.Radius).ToString("F2");
                    break;

                case Arc arc:
                    data["Radius"] = arc.Radius.ToString("F2");
                    data["Center"] = $"{arc.Center.X:F2}, {arc.Center.Y:F2}";
                    data["StartAngle"] = $"{arc.StartAngle * 180 / Math.PI:F2}°";
                    data["EndAngle"] = $"{arc.EndAngle * 180 / Math.PI:F2}°";
                    data["Length"] = arc.Length.ToString("F2");
                    break;

                case Polyline pline:
                    data["Closed"] = pline.Closed.ToString();
                    data["Vertices"] = pline.NumberOfVertices.ToString();
                    data["Length"] = pline.Length.ToString("F2");
                    if (pline.Closed)
                        data["Area"] = Math.Abs(pline.Area).ToString("F2");
                    break;

                case DBText text:
                    data["TextString"] = text.TextString;
                    data["Height"] = text.Height.ToString("F2");
                    data["Position"] = $"{text.Position.X:F2}, {text.Position.Y:F2}";
                    data["Rotation"] = $"{text.Rotation * 180 / Math.PI:F2}°";
                    break;

                case MText mtext:
                    data["TextString"] = StripMTextFormatting(mtext.Contents);
                    data["Height"] = mtext.TextHeight.ToString("F2");
                    data["Width"] = mtext.Width.ToString("F2");
                    data["Location"] = $"{mtext.Location.X:F2}, {mtext.Location.Y:F2}";
                    break;

                case Dimension dim:
                    data["DimensionText"] = dim.DimensionText;
                    data["Measurement"] = dim.Measurement.ToString("F2");
                    break;

                case Hatch hatch:
                    data["PatternName"] = hatch.PatternName;
                    data["Area"] = hatch.Area.ToString("F2");
                    data["PatternScale"] = hatch.PatternScale.ToString("F2");
                    break;

                case Solid3d solid:
                    data["Volume"] = GetSolid3dVolume(solid);
                    break;
            }

            // Include XData if requested
            if (includeXData)
            {
                var xdata = entity.XData;
                if (xdata != null)
                {
                    var xdataStr = new StringBuilder();
                    foreach (TypedValue tv in xdata)
                    {
                        if (tv.TypeCode == 1001) // App name
                            xdataStr.Append($"[{tv.Value}]");
                        else if (tv.TypeCode == 1000) // String data
                            xdataStr.Append($" {tv.Value}");
                    }
                    if (xdataStr.Length > 0)
                        data["XData"] = xdataStr.ToString();
                }
            }

            // Create display name
            data["DisplayName"] = CreateDisplayName(data);

            return data;
        }

        private static string StripMTextFormatting(string mtext)
        {
            // Remove common MText formatting codes
            var result = mtext;
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\\[LO]", "");
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\\[Ff][^;]+;", "");
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\\[Hh][^;]+;", "");
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\\[Cc]\d+;", "");
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\\[Pp]", "");
            result = System.Text.RegularExpressions.Regex.Replace(result, @"\{|\}", "");
            return result;
        }

        private static string GetSolid3dVolume(Solid3d solid)
        {
            try
            {
                var massProps = solid.MassProperties;
                return massProps.Volume.ToString("F2");
            }
            catch
            {
                return "N/A";
            }
        }

        private static string CreateDisplayName(Dictionary<string, object> data)
        {
            var entityType = data["EntityType"].ToString();
            var layer = data["Layer"].ToString();

            // Special handling for blocks
            if (!string.IsNullOrEmpty(data["BlockName"].ToString()) && data["BlockName"].ToString() != "")
                return $"{data["BlockName"]} [{entityType}] on {layer}";

            // Special handling for text entities
            if (data.ContainsKey("TextString") && !string.IsNullOrEmpty(data["TextString"].ToString()))
            {
                var text = data["TextString"].ToString();
                if (text.Length > 30)
                    text = text.Substring(0, 27) + "...";
                return $"\"{text}\" [{entityType}]";
            }

            // Special handling for dimensions
            if (data.ContainsKey("DimensionText") && !string.IsNullOrEmpty(data["DimensionText"].ToString()))
            {
                return $"Dim: {data["DimensionText"]} [{entityType}]";
            }

            return $"{entityType} on {layer}";
        }

        private static string GetColorString(Entity entity)
        {
            if (entity.Color.IsByLayer)
                return "ByLayer";
            else if (entity.Color.IsByBlock)
                return "ByBlock";
            else if (entity.Color.IsByAci)
                return $"Index {entity.Color.ColorIndex}";
            else if (entity.Color.ColorMethod == Autodesk.AutoCAD.Colors.ColorMethod.ByColor)
                return $"RGB({entity.Color.Red},{entity.Color.Green},{entity.Color.Blue})";
            else
                return entity.Color.ToString();
        }

        private static string GetSpaceName(Entity entity, Document doc)
        {
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var owner = tr.GetObject(entity.OwnerId, OpenMode.ForRead) as BlockTableRecord;
                if (owner != null)
                {
                    if (owner.IsLayout)
                    {
                        if (owner.Name == "*Model_Space")
                            return "Model";
                        else
                            return "Paper";
                    }
                    else
                    {
                        return owner.Name; // Block definition name
                    }
                }
                tr.Commit();
            }
            return "Unknown";
        }
    }

    // Windows Forms DataGridView window for filtering
    public class FilterWindow : Form
    {
        private DataGridView dataGrid;
        private List<Dictionary<string, object>> allData;
        private List<Dictionary<string, object>> filteredData;
        private TextBox filterBox;
        private Label statusLabel;

        public List<Dictionary<string, object>> SelectedRows { get; private set; }

        public FilterWindow(List<Dictionary<string, object>> data, List<string> columnNames)
        {
            allData = data;
            filteredData = new List<Dictionary<string, object>>(data);
            SelectedRows = new List<Dictionary<string, object>>();

            Text = "Filter AutoCAD Selection";
            Size = new Size(1200, 700);
            StartPosition = FormStartPosition.CenterScreen;

            // Create main panel
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4,
                ColumnCount = 1
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Filter panel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // DataGrid
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));  // Status bar
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));  // Button panel

            // Filter panel
            var filterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                Padding = new Padding(5)
            };

            filterPanel.Controls.Add(new Label { Text = "Filter:", AutoSize = true, Anchor = AnchorStyles.Left });

            filterBox = new TextBox
            {
                Width = 300,
                Anchor = AnchorStyles.Left
            };
            filterBox.TextChanged += OnFilterTextChanged;
            filterPanel.Controls.Add(filterBox);

            var clearButton = new Button
            {
                Text = "Clear Filter",
                AutoSize = true
            };
            clearButton.Click += (s, e) => { filterBox.Text = ""; };
            filterPanel.Controls.Add(clearButton);

            var filterHelpLabel = new Label
            {
                Text = "(Type to filter across all columns)",
                ForeColor = Color.Gray,
                AutoSize = true
            };
            filterPanel.Controls.Add(filterHelpLabel);

            mainPanel.Controls.Add(filterPanel, 0, 0);

            // DataGrid
            dataGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                ReadOnly = true,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.AliceBlue },
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                RowHeadersVisible = false
            };

            // Create columns (skip internal columns)
            foreach (var columnName in columnNames)
            {
                if (columnName == "ObjectId" || columnName == "DocumentPath")
                    continue; // Skip internal columns

                var column = new DataGridViewTextBoxColumn
                {
                    Name = columnName,
                    HeaderText = columnName,
                    DataPropertyName = columnName
                };

                // Set column widths
                if (columnName == "DisplayName")
                    column.Width = 300;
                else if (columnName == "Handle")
                    column.Width = 100;
                else if (columnName == "EntityType")
                    column.Width = 120;
                else if (columnName == "Layer")
                    column.Width = 150;
                else if (columnName == "BlockName")
                    column.Width = 150;
                else if (columnName == "Color")
                    column.Width = 100;
                else if (columnName == "TextString")
                    column.Width = 200;
                else
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

                dataGrid.Columns.Add(column);
            }

            // Bind data
            UpdateDataGrid();

            mainPanel.Controls.Add(dataGrid, 0, 1);

            // Status bar
            statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = $"Total: {allData.Count} entities | Filtered: {filteredData.Count} entities | Selected: 0 entities",
                BackColor = Color.LightGray,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(5, 0, 0, 0)
            };
            mainPanel.Controls.Add(statusLabel, 0, 2);

            // Button panel
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                Padding = new Padding(5)
            };

            var cancelButton = new Button
            {
                Text = "Cancel",
                Size = new Size(100, 30)
            };
            cancelButton.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            buttonPanel.Controls.Add(cancelButton);

            var okButton = new Button
            {
                Text = "OK - Apply Filter",
                Size = new Size(120, 30),
                Font = new System.Drawing.Font(DefaultFont, FontStyle.Bold)
            };
            okButton.Click += OnOkClick;
            buttonPanel.Controls.Add(okButton);

            var spacing = new Label { Width = 20 };
            buttonPanel.Controls.Add(spacing);

            var selectNoneButton = new Button
            {
                Text = "Select None",
                Size = new Size(100, 30)
            };
            selectNoneButton.Click += (s, e) => { dataGrid.ClearSelection(); UpdateStatusBar(); };
            buttonPanel.Controls.Add(selectNoneButton);

            var selectAllButton = new Button
            {
                Text = "Select All",
                Size = new Size(100, 30)
            };
            selectAllButton.Click += (s, e) => { dataGrid.SelectAll(); UpdateStatusBar(); };
            buttonPanel.Controls.Add(selectAllButton);

            mainPanel.Controls.Add(buttonPanel, 0, 3);

            Controls.Add(mainPanel);

            // Handle selection changed
            dataGrid.SelectionChanged += (s, e) => UpdateStatusBar();

            // Select all items initially
            Load += (s, e) => { dataGrid.SelectAll(); UpdateStatusBar(); };
        }

        private void UpdateDataGrid()
        {
            // Convert dictionary list to DataTable for DataGridView binding
            var dataTable = new System.Data.DataTable();

            if (filteredData.Any())
            {
                // Add columns
                foreach (var key in filteredData[0].Keys)
                {
                    dataTable.Columns.Add(key, typeof(string));
                }

                // Add rows
                foreach (var dict in filteredData)
                {
                    var row = dataTable.NewRow();
                    foreach (var kvp in dict)
                    {
                        row[kvp.Key] = kvp.Value?.ToString() ?? "";
                    }
                    dataTable.Rows.Add(row);
                }
            }

            dataGrid.DataSource = dataTable;
        }

        private void UpdateStatusBar()
        {
            statusLabel.Text = $"Total: {allData.Count} entities | Filtered: {filteredData.Count} entities | Selected: {dataGrid.SelectedRows.Count} entities";
        }

        private void OnFilterTextChanged(object sender, EventArgs e)
        {
            var filterText = filterBox.Text.ToLower();

            if (string.IsNullOrWhiteSpace(filterText))
            {
                filteredData = new List<Dictionary<string, object>>(allData);
            }
            else
            {
                filteredData = allData.Where(row =>
                {
                    return row.Any(kvp =>
                        kvp.Value != null &&
                        kvp.Value.ToString().ToLower().Contains(filterText)
                    );
                }).ToList();
            }

            UpdateDataGrid();
            UpdateStatusBar();
        }

        private void OnOkClick(object sender, EventArgs e)
        {
            SelectedRows = new List<Dictionary<string, object>>();

            foreach (DataGridViewRow row in dataGrid.SelectedRows)
            {
                if (row.Index < filteredData.Count)
                {
                    SelectedRows.Add(filteredData[row.Index]);
                }
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }

    public class FilterSelectedCommand
    {
        [CommandMethod("filter-selected")]
        public static void FilterSelected()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var editor = doc.Editor;

            try
            {
                // Load selection from file
                var savedSelection = SelectionStorage.LoadSelection();
                if (!savedSelection.Any())
                {
                    editor.WriteMessage("\nNo saved selection found. Run select-all-in-opened-drawings first.");
                    return;
                }

                editor.WriteMessage($"\nLoading {savedSelection.Count} entities from saved selection...\n");

                // Group by document
                var selectionByDoc = savedSelection.GroupBy(s => s.DocumentPath);
                var allEntityData = new List<Dictionary<string, object>>();

                foreach (var docGroup in selectionByDoc)
                {
                    // Find the document
                    Document targetDoc = null;
                    foreach (Document d in Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager)
                    {
                        if (d.Name == docGroup.Key)
                        {
                            targetDoc = d;
                            break;
                        }
                    }

                    if (targetDoc == null)
                    {
                        editor.WriteMessage($"  Warning: Document '{Path.GetFileName(docGroup.Key)}' is not open - skipping {docGroup.Count()} entities\n");
                        continue;
                    }

                    // Get entities by handle
                    var objectIds = new List<ObjectId>();
                    using (targetDoc.LockDocument())
                    {
                        using (var tr = targetDoc.Database.TransactionManager.StartTransaction())
                        {
                            foreach (var item in docGroup)
                            {
                                try
                                {
                                    var handle = new Handle(Convert.ToInt64(item.Handle, 16));
                                    if (targetDoc.Database.TryGetObjectId(handle, out ObjectId id))
                                    {
                                        objectIds.Add(id);
                                    }
                                }
                                catch
                                {
                                    // Skip invalid handles
                                }
                            }
                            tr.Commit();
                        }
                    }

                    // Get entity data
                    if (objectIds.Any())
                    {
                        var entityData = EntityDataHelper.GetEntityData(targetDoc, objectIds.ToArray(), true);
                        allEntityData.AddRange(entityData);
                        editor.WriteMessage($"  Loaded {entityData.Count} entities from {Path.GetFileName(targetDoc.Name)}\n");
                    }
                }

                if (!allEntityData.Any())
                {
                    editor.WriteMessage("\nNo valid entities found in saved selection.");
                    return;
                }

                editor.WriteMessage($"\nTotal entities loaded: {allEntityData.Count}\n");

                // Get column names
                var columnNames = new List<string> { "DisplayName", "EntityType", "Layer", "BlockName",
                                                      "Handle", "Color", "Linetype", "LineWeight", "Space" };

                // Add any additional property columns found in data
                var additionalColumns = allEntityData
                    .SelectMany(d => d.Keys)
                    .Distinct()
                    .Where(k => !columnNames.Contains(k) && k != "ObjectId" && k != "DocumentPath")
                    .OrderBy(k => k);
                columnNames.AddRange(additionalColumns);

                // Show filter window
                editor.WriteMessage("\nOpening filter window...\n");
                var filterWindow = new FilterWindow(allEntityData, columnNames);
                var result = filterWindow.ShowDialog();

                if (result == DialogResult.OK && filterWindow.SelectedRows.Any())
                {
                    // Convert selected rows back to selection items
                    var newSelection = new List<SelectionItem>();
                    foreach (var row in filterWindow.SelectedRows)
                    {
                        if (row.ContainsKey("Handle") && row.ContainsKey("DocumentPath"))
                        {
                            newSelection.Add(new SelectionItem
                            {
                                DocumentPath = row["DocumentPath"].ToString(),
                                Handle = row["Handle"].ToString()
                            });
                        }
                    }

                    // Save new selection
                    SelectionStorage.SaveSelection(newSelection);

                    // Select entities in current document
                    var currentDocSelection = newSelection.Where(s => s.DocumentPath == doc.Name).ToList();
                    if (currentDocSelection.Any())
                    {
                        using (var tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            var selectedIds = new List<ObjectId>();

                            foreach (var item in currentDocSelection)
                            {
                                try
                                {
                                    var handle = new Handle(Convert.ToInt64(item.Handle, 16));
                                    if (doc.Database.TryGetObjectId(handle, out ObjectId id))
                                    {
                                        selectedIds.Add(id);
                                    }
                                }
                                catch
                                {
                                    // Skip invalid handles
                                }
                            }

                            // Set selection
                            if (selectedIds.Count > 0)
                            {
                                editor.SetImpliedSelection(selectedIds.ToArray());
                                editor.WriteMessage($"\nSelected {selectedIds.Count} entities in current drawing.");
                            }

                            tr.Commit();
                        }
                    }
                    else
                    {
                        editor.WriteMessage($"\nNo filtered entities found in current drawing.");
                    }

                    editor.WriteMessage($"\nFiltered selection saved: {newSelection.Count} total entities across all drawings.");
                }
                else
                {
                    editor.WriteMessage("\nFilter operation cancelled.");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError: {ex.Message}\n");
                editor.WriteMessage($"Stack trace: {ex.StackTrace}\n");
            }
        }
    }
}
