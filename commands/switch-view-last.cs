using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SwitchViewLastCommand))]

namespace AutoCADBallet
{
    public class SwitchViewLastCommand
    {
        [CommandMethod("switch-view-last")]
        public void SwitchViewLast()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

            // Build a chronological list of all layout entries across all documents
            var allHistoryEntries = new List<(DateTime timestamp, string layoutName, string docName)>();

            // Read layout logs from all documents if they exist
            if (Directory.Exists(logDirPath))
            {
                foreach (string logFile in Directory.GetFiles(logDirPath))
                {
                    try
                    {
                        string docName = Path.GetFileName(logFile);
                        var lines = File.ReadAllLines(logFile)
                                        .Select(l => l.Trim())
                                        .Where(l => l.Length > 0);

                        foreach (string line in lines)
                        {
                            var parts = line.Split(new[] { ' ' }, 2);
                            if (parts.Length == 2)
                            {
                                if (DateTime.TryParseExact(parts[0], "yyyy-MM-dd_HH:mm:ss",
                                                         System.Globalization.CultureInfo.InvariantCulture,
                                                         System.Globalization.DateTimeStyles.None, out DateTime timestamp))
                                {
                                    string layoutName = parts[1];
                                    allHistoryEntries.Add((timestamp, layoutName, docName));
                                }
                            }
                        }
                    }
                    catch (System.Exception)
                    {
                        // Silently skip files that can't be read
                    }
                }
            }

            if (allHistoryEntries.Count < 2)
            {
                return;
            }

            // Sort chronologically (most recent first) and create view entries
            var availableViews = new List<Dictionary<string, object>>();
            string currentLayoutName = LayoutManager.Current.CurrentLayout;
            string currentDocName = Path.GetFileNameWithoutExtension(activeDoc.Name);

            // Sort all entries by timestamp (most recent first) and deduplicate
            var sortedEntries = allHistoryEntries
                .OrderByDescending(e => e.timestamp)
                .GroupBy(e => new { e.layoutName, e.docName })
                .Select(g => g.First()) // Take the most recent occurrence of each layout/document combo
                .ToList();

            // Create view entries for each unique layout/document combination
            foreach (var entry in sortedEntries)
            {
                // Check if this document is currently open and if the layout exists
                Document targetDoc = null;
                bool layoutExists = false;

                foreach (Document doc in docs)
                {
                    if (Path.GetFileNameWithoutExtension(doc.Name) == entry.docName)
                    {
                        targetDoc = doc;

                        // Check if the layout actually exists in the document
                        try
                        {
                            using (doc.LockDocument())
                            {
                                Database db = doc.Database;
                                using (var tr = db.TransactionManager.StartTransaction())
                                {
                                    // Model space is always available, doesn't exist in LayoutDictionary
                                    if (entry.layoutName.Equals("Model", StringComparison.OrdinalIgnoreCase))
                                    {
                                        layoutExists = true;
                                    }
                                    else
                                    {
                                        var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                        layoutExists = layoutDict.Contains(entry.layoutName);
                                    }
                                    tr.Commit();
                                }
                            }
                        }
                        catch
                        {
                            layoutExists = false;
                        }
                        break;
                    }
                }

                // Only add entries where the document is open and the layout exists
                if (targetDoc != null && layoutExists)
                {
                    availableViews.Add(new Dictionary<string, object>
                    {
                        ["LayoutName"] = entry.layoutName,
                        ["DocumentName"] = entry.docName,
                        ["Document"] = targetDoc,
                        ["Timestamp"] = entry.timestamp
                    });
                }
            }

            if (availableViews.Count < 2)
            {
                return;
            }

            // Select the second most recent layout (index 1)
            var selectedView = availableViews[1];
            string chosenLayoutName = selectedView["LayoutName"].ToString();
            Document chosenDoc = selectedView["Document"] as Document;
            string targetDocName = selectedView["DocumentName"].ToString();

            // Use the proper cross-document switching pattern from switch-view-recent.cs
            if (docs.MdiActiveDocument != chosenDoc)
            {
                // Set up event handler for when document activation completes
                DocumentCollectionEventHandler handler = null;
                handler = (sender, e) =>
                {
                    if (e.Document == chosenDoc)
                    {
                        // Unsubscribe from event to avoid memory leaks
                        docs.DocumentActivated -= handler;

                        try
                        {
                            // Switch to the target layout with document lock
                            using (DocumentLock docLock = chosenDoc.LockDocument())
                            {
                                LayoutManager.Current.CurrentLayout = chosenLayoutName;
                            }

                            LayoutChangeTracker.LogLayoutChange(targetDocName, chosenLayoutName, true);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception)
                        {
                        }
                    }
                };

                docs.DocumentActivated += handler;
                docs.MdiActiveDocument = chosenDoc;
            }
            else
            {
                // Already in the target document, just switch layout
                try
                {
                    // Direct layout switch with document lock
                    using (DocumentLock docLock = chosenDoc.LockDocument())
                    {
                        LayoutManager.Current.CurrentLayout = chosenLayoutName;
                    }

                    LayoutChangeTracker.LogLayoutChange(targetDocName, chosenLayoutName, true);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception)
                {
                }
            }
        }

    }
}