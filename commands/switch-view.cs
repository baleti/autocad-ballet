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

[assembly: CommandClass(typeof(AutoCADBallet.SwitchViewCommand))]

namespace AutoCADBallet
{
    public class SwitchViewCommand
    {
        [CommandMethod("switch-view")]
        public void SwitchView()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "LogLayoutChanges");

            var layoutNames = new HashSet<string>();
            
            // Read layout logs from all documents if they exist
            if (Directory.Exists(logDirPath))
            {
                foreach (string logFile in Directory.GetFiles(logDirPath))
                {
                    try
                    {
                        var layoutEntries = File.ReadAllLines(logFile)
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
                        ed.WriteMessage($"\nWarning: Could not read log file {logFile}: {ex.Message}");
                    }
                }
            }
            
            if (layoutNames.Count == 0)
            {
                ed.WriteMessage("\nNo layout logs found. Showing all layouts from all open documents.");
            }

            var availableViews = new List<Dictionary<string, object>>();
            string currentLayoutName = LayoutManager.Current.CurrentLayout;
            string currentDocName = Path.GetFileNameWithoutExtension(activeDoc.Name);

            // Iterate through all open documents
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileNameWithoutExtension(doc.Name);
                
                try
                {
                    // Lock the document for read access
                    using (doc.LockDocument())
                    {
                        Database db = doc.Database;
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                            
                            foreach (DictionaryEntry entry in layoutDict)
                            {
                                var layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                                if (layout != null && (layoutNames.Count == 0 || layoutNames.Contains(layout.LayoutName)))
                                {
                                    bool isCurrent = (doc == activeDoc && layout.LayoutName == currentLayoutName);
                                    
                                    availableViews.Add(new Dictionary<string, object>
                                    {
                                        ["LayoutName"] = layout.LayoutName,
                                        ["DocumentName"] = docName,
                                        ["TabOrder"] = layout.TabOrder,
                                        ["Document"] = doc,
                                        ["IsCurrent"] = isCurrent,
                                        ["DisplayName"] = $"{layout.LayoutName} ({docName})"
                                    });
                                }
                            }
                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Could not access layouts in document {docName}: {ex.Message}");
                }
            }

            if (availableViews.Count == 0)
            {
                ed.WriteMessage("\nNo matching layouts found in any open document.");
                return;
            }

            // Sort by document name first, then by tab order
            availableViews = availableViews.OrderBy(v => v["DocumentName"].ToString())
                                         .ThenBy(v => 
                                         {
                                             if (v["TabOrder"] == null) return int.MaxValue;
                                             if (int.TryParse(v["TabOrder"].ToString(), out int tabOrder)) return tabOrder;
                                             return int.MaxValue;
                                         }).ToList();

            int selectedIndex = -1;
            selectedIndex = availableViews.FindIndex(v => (bool)v["IsCurrent"]);

            var propertyNames = new List<string> { "LayoutName", "DocumentName" };
            var initialSelectionIndices = selectedIndex >= 0 
                                            ? new List<int> { selectedIndex }
                                            : new List<int>();

            ed.WriteMessage($"\nDebug: Found {availableViews.Count} layouts across {docs.Count} documents, selectedIndex={selectedIndex}");
            if (selectedIndex >= 0)
                ed.WriteMessage($"\nDebug: Will pre-select: {availableViews[selectedIndex]["DisplayName"]}");

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(availableViews, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    var selectedView = chosen.First();
                    string chosenLayoutName = selectedView["LayoutName"].ToString();
                    Document targetDoc = selectedView["Document"] as Document;
                    string targetDocName = selectedView["DocumentName"].ToString();

                    // Switch to the target document first
                    docs.MdiActiveDocument = targetDoc;
                    
                    // Then switch to the target layout
                    LayoutManager.Current.CurrentLayout = chosenLayoutName;
                    
                    ed.WriteMessage($"\nSwitched to layout '{chosenLayoutName}' in document '{targetDocName}'");
                    return;
                }
                return;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nDataGrid failed, using text prompt: {ex.Message}");
            }

            // Text-based fallback
            ed.WriteMessage("\nAvailable layouts:");
            for (int i = 0; i < availableViews.Count; i++)
            {
                string marker = (i == selectedIndex) ? " [CURRENT]" : "";
                ed.WriteMessage($"\n{i + 1}: {availableViews[i]["DisplayName"]}{marker}");
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nSelect view number: ");
            pio.AllowNegative = false;
            pio.AllowZero = false;
            pio.LowerLimit = 1;
            pio.UpperLimit = availableViews.Count;
            PromptIntegerResult pir = ed.GetInteger(pio);

            if (pir.Status == PromptStatus.OK)
            {
                var selectedView = availableViews[pir.Value - 1];
                string selectedLayoutName = selectedView["LayoutName"].ToString();
                Document targetDoc = selectedView["Document"] as Document;
                string targetDocName = selectedView["DocumentName"].ToString();

                // Switch to the target document first
                docs.MdiActiveDocument = targetDoc;
                
                // Then switch to the target layout
                LayoutManager.Current.CurrentLayout = selectedLayoutName;
                
                ed.WriteMessage($"\nSwitched to layout '{selectedLayoutName}' in document '{targetDocName}'");
            }
        }
    }
}