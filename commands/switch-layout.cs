using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SwitchLayoutCommand))]

namespace AutoCADBallet
{
    public class SwitchLayoutCommand
    {
        [CommandMethod("switch-layout")]
        public void SwitchLayout()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;

            string projectName = Path.GetFileNameWithoutExtension(doc.Name) ?? "UnknownProject";
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logFilePath = Path.Combine(appDataPath, "autocad-ballet", "LogLayoutChanges", projectName);

            var layoutNames = new List<string>();
            
            if (File.Exists(logFilePath))
            {
                try
                {
                    var layoutEntries = File.ReadAllLines(logFilePath)
                                          .Reverse()
                                          .Select(l => l.Trim())
                                          .Where(l => l.Length > 0)
                                          .Distinct()
                                          .ToList();

                    foreach (string entry in layoutEntries)
                    {
                        var parts = entry.Split(new[] { ' ' }, 2);
                        if (parts.Length == 2)
                            layoutNames.Add(parts[1]);
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Could not read log file: {ex.Message}");
                }
            }
            
            if (layoutNames.Count == 0)
            {
                ed.WriteMessage("\nNo layout log found. Showing all layouts in current drawing.");
            }

            var availableLayouts = new List<Dictionary<string, object>>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                
                foreach (DictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                    if (layout != null && (layoutNames.Count == 0 || layoutNames.Contains(layout.LayoutName)))
                    {
                        availableLayouts.Add(new Dictionary<string, object>
                        {
                            ["LayoutName"] = layout.LayoutName,
                            ["TabOrder"] = layout.TabOrder,
                            ["LayoutObject"] = layout
                        });
                    }
                }
                tr.Commit();
            }

            if (availableLayouts.Count == 0)
            {
                ed.WriteMessage("\nNo matching layouts found in this drawing.");
                return;
            }

            availableLayouts = availableLayouts.OrderBy(l => 
            {
                if (l["TabOrder"] == null) return int.MaxValue;
                if (int.TryParse(l["TabOrder"].ToString(), out int tabOrder)) return tabOrder;
                return int.MaxValue;
            }).ToList();

            int selectedIndex = -1;
            string currentLayoutName = LayoutManager.Current.CurrentLayout;
            selectedIndex = availableLayouts.FindIndex(l => l["LayoutName"].ToString() == currentLayoutName);

            var propertyNames = new List<string> { "TabOrder", "LayoutName" };
            var initialSelectionIndices = selectedIndex >= 0 
                                            ? new List<int> { selectedIndex }
                                            : new List<int>();

            ed.WriteMessage($"\nDebug: Found {availableLayouts.Count} layouts, selectedIndex={selectedIndex}");
            if (selectedIndex >= 0)
                ed.WriteMessage($"\nDebug: Will pre-select: {availableLayouts[selectedIndex]["LayoutName"]}");

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(availableLayouts, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    string chosenLayoutName = chosen.First()["LayoutName"].ToString();
                    LayoutManager.Current.CurrentLayout = chosenLayoutName;
                    ed.WriteMessage($"\nSwitched to layout: {chosenLayoutName}");
                    return;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError showing layout picker: {ex.Message}");
            }

            ed.WriteMessage("\nAvailable layouts:");
            for (int i = 0; i < availableLayouts.Count; i++)
            {
                string marker = (i == selectedIndex) ? " [CURRENT]" : "";
                ed.WriteMessage($"\n{i + 1}: {availableLayouts[i]["LayoutName"]}{marker}");
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nSelect layout number: ");
            pio.AllowNegative = false;
            pio.AllowZero = false;
            pio.LowerLimit = 1;
            pio.UpperLimit = availableLayouts.Count;
            PromptIntegerResult pir = ed.GetInteger(pio);

            if (pir.Status == PromptStatus.OK)
            {
                string selectedLayoutName = availableLayouts[pir.Value - 1]["LayoutName"].ToString();
                LayoutManager.Current.CurrentLayout = selectedLayoutName;
                ed.WriteMessage($"\nSwitched to layout: {selectedLayoutName}");
            }
        }
    }
}