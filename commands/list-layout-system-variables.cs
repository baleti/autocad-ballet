using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public class LayoutSystemVariableInfo
    {
        public string Name { get; set; }
        public object Value { get; set; }
        public string Changed { get; set; }
        public string DocumentName { get; set; }
        public string LayoutName { get; set; }

        public LayoutSystemVariableInfo()
        {
            Changed = "";
            LayoutName = "";
            DocumentName = "";
        }
    }

    public static class ListLayoutSystemVariables
    {
        // Layout-specific system variables (saved per layout)
        // These require activating each layout to read/write
        private static readonly string[] LayoutSpecificSystemVariables = new[]
        {
            "PSLTSCALE",   // Paper space linetype scaling
            "UCSFOLLOW",   // UCS follows view changes
            "UCSVP",       // UCS per viewport
        };

        // Known default values for layout-specific system variables
        private static readonly Dictionary<string, string> KnownDefaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "PSLTSCALE", "1" },
            { "UCSFOLLOW", "0" },
            { "UCSVP", "1" },
        };

        public static void ExecuteDocumentScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;

            // Show warning dialog
            if (!ShowWarningDialog("current document"))
            {
                ed.WriteMessage("\nCommand cancelled by user.\n");
                return;
            }

            try
            {
                var layoutVariables = GetLayoutSystemVariablesForDocument(doc);

                if (layoutVariables.Count == 0)
                {
                    ed.WriteMessage("\nNo layout-specific system variables found.\n");
                    return;
                }

                // Convert to Dictionary format for DataGrid
                var dataList = layoutVariables.Select(lv => new Dictionary<string, object>
                {
                    { "Name", lv.Name },
                    { "Value", FormatSystemVariableValue(lv.Value) },
                    { "Changed", lv.Changed },
                    { "Layout", lv.LayoutName },
                    { "_OriginalValue", lv.Value }  // Store original for type conversion
                }).ToList();

                var propertyNames = new List<string> { "Name", "Value", "Changed", "Layout" };

                // Show DataGrid with layout-specific variables
                CustomGUIs.ResetEditsAppliedFlag();
                var selectedRows = CustomGUIs.DataGrid(
                    dataList,
                    propertyNames,
                    spanAllScreens: false,
                    initialSelectionIndices: null,
                    onDeleteEntries: null,
                    allowCreateFromSearch: false,
                    commandName: "list-layout-system-variables-in-document"
                );

                // DataGrid modifies dataList in-place during edit mode
                // Check if any values were actually changed by comparing to original values
                var modifiedItems = dataList.Where(item =>
                {
                    var currentValue = item["Value"].ToString();
                    var originalValue = FormatSystemVariableValue(item["_OriginalValue"]);
                    return currentValue != originalValue;
                }).ToList();

                if (modifiedItems.Count > 0)
                {
                    ed.WriteMessage($"\n{modifiedItems.Count} variable(s) modified, applying changes...");
                    // Apply only the modified values
                    ApplyLayoutSystemVariableChanges(doc, modifiedItems);
                    ed.WriteMessage($"\nLayout-specific system variables updated.\n");
                }
                else
                {
                    ed.WriteMessage($"\nNo changes made.\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n");
            }
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            var docs = AcadApp.DocumentManager;

            // Show warning dialog
            if (!ShowWarningDialog("all open documents"))
            {
                ed.WriteMessage("\nCommand cancelled by user.\n");
                return;
            }

            try
            {
                var allLayoutVariables = new List<LayoutSystemVariableInfo>();

                // Collect layout-specific variables from all open documents
                foreach (Document doc in docs)
                {
                    var docVars = GetLayoutSystemVariablesForDocument(doc);
                    allLayoutVariables.AddRange(docVars);
                }

                if (allLayoutVariables.Count == 0)
                {
                    ed.WriteMessage("\nNo layout-specific system variables found in any open document.\n");
                    return;
                }

                // Convert to Dictionary format for DataGrid
                var dataList = allLayoutVariables.Select(lv => new Dictionary<string, object>
                {
                    { "Name", lv.Name },
                    { "Value", FormatSystemVariableValue(lv.Value) },
                    { "Changed", lv.Changed },
                    { "Layout", lv.LayoutName },
                    { "DocumentName", lv.DocumentName },
                    { "_OriginalValue", lv.Value }  // Store original for type conversion
                }).ToList();

                var propertyNames = new List<string> { "Name", "Value", "Changed", "Layout", "DocumentName" };

                // Show DataGrid with layout-specific variables from all documents
                CustomGUIs.ResetEditsAppliedFlag();
                var selectedRows = CustomGUIs.DataGrid(
                    dataList,
                    propertyNames,
                    spanAllScreens: false,
                    initialSelectionIndices: null,
                    onDeleteEntries: null,
                    allowCreateFromSearch: false,
                    commandName: "list-layout-system-variables-in-session"
                );

                // DataGrid modifies dataList in-place during edit mode
                // Check if any values were actually changed by comparing to original values
                var modifiedItems = dataList.Where(item =>
                {
                    var currentValue = item["Value"].ToString();
                    var originalValue = FormatSystemVariableValue(item["_OriginalValue"]);
                    return currentValue != originalValue;
                }).ToList();

                if (modifiedItems.Count > 0)
                {
                    ed.WriteMessage($"\n{modifiedItems.Count} variable(s) modified, applying changes...");
                    // Group by document and apply only the modified values
                    var grouped = modifiedItems.GroupBy(row => row["DocumentName"].ToString());

                    foreach (var group in grouped)
                    {
                        var targetDoc = FindDocumentByName(group.Key);
                        if (targetDoc != null)
                        {
                            ApplyLayoutSystemVariableChanges(targetDoc, group.ToList());
                        }
                    }

                    ed.WriteMessage($"\nLayout-specific system variables updated across documents.\n");
                }
                else
                {
                    ed.WriteMessage($"\nNo changes made.\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n");
            }
        }

        private static bool ShowWarningDialog(string scope)
        {
            var result = MessageBox.Show(
                $"This command will activate every layout in {scope} to read layout-specific system variables.\n\n" +
                "This process can be slow and may cause AutoCAD to become unresponsive or crash if you have many layouts or large documents.\n\n" +
                "Do you want to continue?",
                "Layout System Variables Information",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button2  // Default to Cancel
            );

            return result == DialogResult.OK;
        }

        private static List<LayoutSystemVariableInfo> GetLayoutSystemVariablesForDocument(Document doc)
        {
            var layoutVariables = new List<LayoutSystemVariableInfo>();
            var db = doc.Database;

            try
            {
                using (var docLock = doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        var layouts = new List<Layout>();

                        foreach (DictionaryEntry entry in layoutDict)
                        {
                            var layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                            if (layout != null)
                            {
                                layouts.Add(layout);
                            }
                        }

                        // Sort layouts by tab order
                        layouts = layouts.OrderBy(l => l.TabOrder).ToList();

                        // For each layout, get the layout-specific variables
                        foreach (var layout in layouts)
                        {
                            string layoutName = layout.LayoutName;

                            // Switch to this layout to read its system variables
                            string originalLayout = LayoutManager.Current.CurrentLayout;
                            try
                            {
                                LayoutManager.Current.CurrentLayout = layoutName;

                                // Read layout-specific system variables
                                foreach (var varName in LayoutSpecificSystemVariables)
                                {
                                    try
                                    {
                                        var value = AcadApp.GetSystemVariable(varName);
                                        if (value != null)
                                        {
                                            var valueStr = FormatSystemVariableValue(value);
                                            var changed = "";

                                            // Check if value differs from known default
                                            if (KnownDefaults.ContainsKey(varName))
                                            {
                                                var defaultValue = KnownDefaults[varName];
                                                if (!string.Equals(valueStr, defaultValue, StringComparison.OrdinalIgnoreCase))
                                                {
                                                    changed = "✓";  // Mark as changed
                                                }
                                            }

                                            layoutVariables.Add(new LayoutSystemVariableInfo
                                            {
                                                Name = varName,
                                                Value = value,
                                                Changed = changed,
                                                DocumentName = System.IO.Path.GetFileName(doc.Name),
                                                LayoutName = layoutName
                                            });
                                        }
                                    }
                                    catch
                                    {
                                        // Skip system variables that can't be read
                                    }
                                }
                            }
                            finally
                            {
                                // Restore original layout
                                try
                                {
                                    LayoutManager.Current.CurrentLayout = originalLayout;
                                }
                                catch
                                {
                                    // Ignore errors restoring layout
                                }
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                // If we can't read layout-specific variables, just skip them
                System.Diagnostics.Debug.WriteLine($"Error reading layout-specific variables: {ex.Message}");
            }

            return layoutVariables;
        }

        private static string FormatSystemVariableValue(object value)
        {
            if (value == null)
                return "";

            // Handle different value types
            if (value is string str)
                return str;

            if (value is double dbl)
                return dbl.ToString("G");

            if (value is short s)
                return s.ToString();

            if (value is int i)
                return i.ToString();

            return value.ToString();
        }

        private static void ApplyLayoutSystemVariableChanges(Document doc, List<Dictionary<string, object>> items)
        {
            var ed = doc.Editor;
            var db = doc.Database;

            // Group items by layout
            var groupedByLayout = items.GroupBy(item => item.ContainsKey("Layout") ? item["Layout"].ToString() : "");

            foreach (var layoutGroup in groupedByLayout)
            {
                string layoutName = layoutGroup.Key;

                if (string.IsNullOrEmpty(layoutName))
                    continue;

                // Switch to the layout
                string originalLayout = null;
                try
                {
                    using (var docLock = doc.LockDocument())
                    {
                        originalLayout = LayoutManager.Current.CurrentLayout;
                        LayoutManager.Current.CurrentLayout = layoutName;
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nFailed to switch to layout {layoutName}: {ex.Message}");
                    continue;
                }

                // Apply variables for this layout
                foreach (var item in layoutGroup)
                {
                    try
                    {
                        var varName = item["Name"].ToString();
                        var newValueStr = item["Value"].ToString();
                        var originalValue = item["_OriginalValue"];

                        // Get the current value from AutoCAD
                        var currentValue = AcadApp.GetSystemVariable(varName);

                        // Convert the new value to the correct type
                        var newValue = ConvertToSystemVariableType(newValueStr, originalValue);

                        // Only set if value actually changed
                        if (newValue != null && !ValuesAreEqual(currentValue, newValue))
                        {
                            AcadApp.SetSystemVariable(varName, newValue);
                            ed.WriteMessage($"\n{varName} set to {FormatSystemVariableValue(newValue)} for layout '{layoutName}'");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        var varName = item.ContainsKey("Name") ? item["Name"].ToString() : "unknown";
                        ed.WriteMessage($"\nFailed to set {varName}: {ex.Message}");
                    }
                }

                // Restore original layout
                if (originalLayout != null)
                {
                    try
                    {
                        using (var docLock = doc.LockDocument())
                        {
                            LayoutManager.Current.CurrentLayout = originalLayout;
                        }
                    }
                    catch
                    {
                        // Ignore errors restoring layout
                    }
                }
            }
        }

        private static object ConvertToSystemVariableType(object newValue, object originalValue)
        {
            if (originalValue == null || newValue == null)
                return newValue;

            // If newValue is already the correct type, return it
            if (newValue.GetType() == originalValue.GetType())
                return newValue;

            // Convert string to the original type
            string newValueStr = newValue.ToString();

            try
            {
                if (originalValue is short)
                {
                    return short.Parse(newValueStr);
                }
                else if (originalValue is int)
                {
                    return int.Parse(newValueStr);
                }
                else if (originalValue is double)
                {
                    return double.Parse(newValueStr);
                }
                else if (originalValue is string)
                {
                    return newValueStr;
                }
            }
            catch
            {
                // Return original value if conversion fails
                return originalValue;
            }

            return newValue;
        }

        private static bool ValuesAreEqual(object val1, object val2)
        {
            if (val1 == null && val2 == null)
                return true;

            if (val1 == null || val2 == null)
                return false;

            // For numeric types, compare as strings to handle precision
            if ((val1 is double || val1 is float) && (val2 is double || val2 is float))
            {
                return Math.Abs(Convert.ToDouble(val1) - Convert.ToDouble(val2)) < 0.0001;
            }

            return val1.Equals(val2);
        }

        private static Document FindDocumentByName(string documentName)
        {
            var docs = AcadApp.DocumentManager;
            foreach (Document doc in docs)
            {
                if (string.Equals(System.IO.Path.GetFileName(doc.Name), documentName, StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }
            }
            return null;
        }
    }
}
