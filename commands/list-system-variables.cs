using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public class SystemVariableInfo
    {
        public string Name { get; set; }
        public object Value { get; set; }
        public string Changed { get; set; }
        public string DocumentName { get; set; }  // For session scope
        public string LayoutName { get; set; }     // Empty if document-wide, layout name if layout-specific

        public SystemVariableInfo()
        {
            Changed = "";  // Empty by default (unchanged from default)
            LayoutName = "";  // Empty by default (document-wide)
        }
    }

    public static class ListSystemVariables
    {
        // Common AutoCAD system variables to display (document-wide only)
        // This is a subset of commonly used system variables - AutoCAD has hundreds
        // Note: Layout-specific variables like PSLTSCALE, UCSFOLLOW, UCSVP are excluded
        // because they cannot be read without activating each layout, which is slow.
        private static readonly string[] CommonSystemVariables = new[]
        {
            "ACADVER", "DWGNAME", "DWGPREFIX", "DWGTITLED",
            "OSMODE", "ORTHOMODE", "POLARMODE", "SNAPMODE", "GRIDMODE",
            "CMDECHO", "FILEDIA", "ATTREQ", "ATTDIA",
            "AUNITS", "AUPREC", "LUNITS", "LUPREC",
            "DIMSCALE", "LTSCALE", "TEXTSIZE", "TEXTSTYLE",
            "CLAYER", "CELTYPE", "CECOLOR", "CELWEIGHT",
            "PICKSTYLE", "PICKFIRST", "PICKADD", "PICKAUTO",
            "MIRRTEXT", "EXPLMODE", "DELOBJ",
            "MEASUREMENT", "MEASUREINIT",
            "BACKGROUNDPLOT", "SAVETIME", "ISAVEBAK", "ISAVEPERCENT",
            "MAXACTVP", "TILEMODE", "PDMODE", "PDSIZE",
            "HIGHLIGHT", "SELECTIONAREA", "SELECTIONPREVIEW",
            "HPNAME", "HPANG", "HPSCALE", "HPASSOC",
            "LWDISPLAY", "LWDEFAULT", "LWEIGHT",
            "ELEVATION", "THICKNESS", "SURFTAB1", "SURFTAB2",
            "FILLMODE", "QTEXTMODE", "SPLFRAME", "DISPSILH",
            "PELLIPSE", "PLINEGEN", "PLINETYPE",
            "ANNOALLVISIBLE", "ANNOAUTOSCALE", "CANNOSCALE",
            "XLOADCTL", "XLOADPATH", "PROJECTNAME",
            "VIEWRES", "FACETRES", "ISOLINES",
            "PROXYGRAPHICS", "PROXYNOTICE", "PROXYSHOW",
            "MTEXTED", "DCTCUST", "DCTMAIN",
            "STARTUP", "STARTMODE", "SDI",
            "SHORTCUTMENU", "SNAPANG", "SNAPBASE", "SNAPSTYL", "SNAPUNIT",
            "GRIDUNIT", "GRIDDISPLAY", "GRIDMAJOR",
            "UCSAXISANG", "UCSBASE",
            "APERTURE", "CURSORSIZE", "GRIPBLOCK", "GRIPOBJLIMIT",
            "GRIPSIZE", "GRIPS", "GRIPCOLOR", "GRIPHOT", "GRIPHOVER",
            "TOOLTIPS", "TOOLTIPMERGE",
            "MBUTTONPAN", "ZOOMFACTOR", "ZOOMWHEEL",
            "DYNMODE", "DYNPROMPT", "DYNTOOLTIPS",
            "INPUTHISTORYMODE", "ROLLOVERTIPS",
            "TRACKPATH", "AUTOSNAP", "POLARANG", "POLARDIST",
            "VSBACKGROUNDS", "VSCURRENT", "VSEDGES", "VSEDGESMOOTH",
        };

        // Known default values for some system variables
        // This is not exhaustive - many defaults depend on template or AutoCAD version
        private static readonly Dictionary<string, string> KnownDefaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "OSMODE", "4133" },
            { "ORTHOMODE", "0" },
            { "POLARMODE", "0" },
            { "SNAPMODE", "0" },
            { "GRIDMODE", "0" },
            { "CMDECHO", "1" },
            { "FILEDIA", "1" },
            { "ATTREQ", "1" },
            { "ATTDIA", "0" },
            { "AUNITS", "0" },
            { "AUPREC", "0" },
            { "LUNITS", "2" },
            { "LUPREC", "4" },
            { "DIMSCALE", "1.0" },
            { "LTSCALE", "1.0" },
            { "PICKFIRST", "1" },
            { "PICKADD", "1" },
            { "PICKAUTO", "1" },
            { "PICKSTYLE", "1" },
            { "MIRRTEXT", "0" },
            { "EXPLMODE", "1" },
            { "DELOBJ", "1" },
            { "MEASUREMENT", "0" },
            { "BACKGROUNDPLOT", "2" },
            { "SAVETIME", "10" },
            { "ISAVEBAK", "1" },
            { "ISAVEPERCENT", "50" },
            { "MAXACTVP", "64" },
            { "TILEMODE", "1" },
            { "PDMODE", "0" },
            { "PDSIZE", "0.0" },
            { "HIGHLIGHT", "1" },
            { "LWDISPLAY", "0" },
            { "ELEVATION", "0.0" },
            { "THICKNESS", "0.0" },
            { "FILLMODE", "1" },
            { "QTEXTMODE", "0" },
            { "SPLFRAME", "0" },
            { "DISPSILH", "0" },
            { "PELLIPSE", "0" },
            { "PLINEGEN", "0" },
            { "PLINETYPE", "2" },
            { "XLOADCTL", "2" },
            { "PROXYGRAPHICS", "1" },
            { "PROXYNOTICE", "1" },
            { "PROXYSHOW", "1" },
            { "SDI", "0" },
            { "SHORTCUTMENU", "11" },
            { "SNAPANG", "0" },
            { "UCSFOLLOW", "0" },
            { "APERTURE", "10" },
            { "CURSORSIZE", "5" },
            { "GRIPBLOCK", "0" },
            { "GRIPSIZE", "5" },
            { "GRIPS", "1" },
            { "TOOLTIPS", "1" },
            { "MBUTTONPAN", "1" },
            { "ZOOMFACTOR", "60" },
            { "DYNMODE", "3" },
            { "DYNPROMPT", "1" },
            { "DYNTOOLTIPS", "1" },
            { "INPUTHISTORYMODE", "15" },
            { "ROLLOVERTIPS", "1" },
            { "TRACKPATH", "0" },
            { "AUTOSNAP", "63" },
            { "POLARANG", "90" },
            { "POLARDIST", "0.0" },
        };

        public static void ExecuteDocumentScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;

            try
            {
                var systemVariables = GetSystemVariablesForDocument(doc);

                if (systemVariables.Count == 0)
                {
                    ed.WriteMessage("\nNo system variables found.\n");
                    return;
                }

                // Convert to Dictionary format for DataGrid
                var dataList = systemVariables.Select(sysVar => new Dictionary<string, object>
                {
                    { "Name", sysVar.Name },
                    { "Value", FormatSystemVariableValue(sysVar.Value) },
                    { "Changed", sysVar.Changed },
                    { "_OriginalValue", sysVar.Value }  // Store original for type conversion
                }).ToList();

                var propertyNames = new List<string> { "Name", "Value", "Changed" };

                // Show DataGrid with system variables
                CustomGUIs.ResetEditsAppliedFlag();
                var selectedRows = CustomGUIs.DataGrid(
                    dataList,
                    propertyNames,
                    spanAllScreens: false,
                    initialSelectionIndices: null,
                    onDeleteEntries: null,
                    allowCreateFromSearch: false,
                    commandName: "list-system-variables-in-document"
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
                    ApplySystemVariableChanges(doc, modifiedItems);
                    ed.WriteMessage($"\nSystem variables updated.\n");
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

            try
            {
                var allSystemVariables = new List<SystemVariableInfo>();

                // Collect system variables from all open documents
                foreach (Document doc in docs)
                {
                    var docVars = GetSystemVariablesForDocument(doc);
                    allSystemVariables.AddRange(docVars);
                }

                if (allSystemVariables.Count == 0)
                {
                    ed.WriteMessage("\nNo system variables found in any open document.\n");
                    return;
                }

                // Convert to Dictionary format for DataGrid
                var dataList = allSystemVariables.Select(sysVar => new Dictionary<string, object>
                {
                    { "Name", sysVar.Name },
                    { "Value", FormatSystemVariableValue(sysVar.Value) },
                    { "Changed", sysVar.Changed },
                    { "DocumentName", sysVar.DocumentName },
                    { "_OriginalValue", sysVar.Value }  // Store original for type conversion
                }).ToList();

                var propertyNames = new List<string> { "Name", "Value", "Changed", "DocumentName" };

                // Show DataGrid with system variables from all documents
                CustomGUIs.ResetEditsAppliedFlag();
                var selectedRows = CustomGUIs.DataGrid(
                    dataList,
                    propertyNames,
                    spanAllScreens: false,
                    initialSelectionIndices: null,
                    onDeleteEntries: null,
                    allowCreateFromSearch: false,
                    commandName: "list-system-variables-in-session"
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
                            ApplySystemVariableChanges(targetDoc, group.ToList());
                        }
                    }

                    ed.WriteMessage($"\nSystem variables updated across documents.\n");
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

        private static List<SystemVariableInfo> GetSystemVariablesForDocument(Document doc)
        {
            var systemVariables = new List<SystemVariableInfo>();

            // Read document-wide system variables (no layout activation required)
            foreach (var varName in CommonSystemVariables)
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

                        systemVariables.Add(new SystemVariableInfo
                        {
                            Name = varName,
                            Value = value,
                            Changed = changed,
                            DocumentName = System.IO.Path.GetFileName(doc.Name),
                            LayoutName = ""  // Not used (document-wide only)
                        });
                    }
                }
                catch
                {
                    // Skip system variables that can't be read
                }
            }

            return systemVariables;
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

        private static void ApplySystemVariableChanges(Document doc, List<Dictionary<string, object>> items)
        {
            var ed = doc.Editor;

            foreach (var item in items)
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
                        ed.WriteMessage($"\n{varName} set to {FormatSystemVariableValue(newValue)}");
                    }
                }
                catch (System.Exception ex)
                {
                    var varName = item.ContainsKey("Name") ? item["Name"].ToString() : "unknown";
                    ed.WriteMessage($"\nFailed to set {varName}: {ex.Message}");
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
