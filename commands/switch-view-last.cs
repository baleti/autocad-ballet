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
        // Simple class to replace tuple for .NET Framework 4.6/4.7 compatibility
        private class LayoutHistoryEntry
        {
            public DateTime Timestamp { get; set; }
            public string LayoutName { get; set; }
            public string DocName { get; set; }
            public string DocumentPath { get; set; }  // Full absolute path
        }

        [CommandMethod("switch-view-last", CommandFlags.Session)]
        public void SwitchViewLast()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

            // Build a chronological list of all layout entries across all documents
            var allHistoryEntries = new List<LayoutHistoryEntry>();

            // Read layout logs from all documents if they exist
            // Now supports multiple document sections per log file
            if (Directory.Exists(logDirPath))
            {
                foreach (string logFile in Directory.GetFiles(logDirPath))
                {
                    try
                    {
                        var lines = File.ReadAllLines(logFile);

                        // Parse document sections
                        string sectionDocumentPath = null;
                        string sectionDocName = null;

                        foreach (var line in lines)
                        {
                            if (string.IsNullOrWhiteSpace(line))
                                continue;

                            if (line.StartsWith("DOCUMENT_PATH:"))
                            {
                                sectionDocumentPath = line.Substring("DOCUMENT_PATH:".Length).Trim();
                                sectionDocName = Path.GetFileNameWithoutExtension(sectionDocumentPath);
                            }
                            else if (!line.StartsWith("LAST_SESSION_PID:") && !line.StartsWith("DOCUMENT_OPENED:"))
                            {
                                // Layout entry
                                var parts = line.Split(new[] { ' ' }, 2);
                                if (parts.Length == 2)
                                {
                                    if (DateTime.TryParseExact(parts[0], "yyyy-MM-dd_HH:mm:ss",
                                                             System.Globalization.CultureInfo.InvariantCulture,
                                                             System.Globalization.DateTimeStyles.None, out DateTime timestamp))
                                    {
                                        string layoutName = parts[1];
                                        allHistoryEntries.Add(new LayoutHistoryEntry
                                        {
                                            Timestamp = timestamp,
                                            LayoutName = layoutName,
                                            DocName = sectionDocName ?? Path.GetFileName(logFile),
                                            DocumentPath = sectionDocumentPath
                                        });
                                    }
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
                .OrderByDescending(e => e.Timestamp)
                .GroupBy(e => new { e.LayoutName, e.DocumentPath })
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
                    // Match by full path if available, otherwise fall back to filename
                    bool isMatch = false;
                    if (!string.IsNullOrEmpty(entry.DocumentPath))
                    {
                        try
                        {
                            string docFullPath = Path.GetFullPath(doc.Name);
                            string entryFullPath = Path.GetFullPath(entry.DocumentPath);
                            isMatch = string.Equals(docFullPath, entryFullPath, StringComparison.OrdinalIgnoreCase);
                        }
                        catch
                        {
                            isMatch = false;
                        }
                    }
                    else
                    {
                        // Fallback for old log entries without document path
                        isMatch = Path.GetFileNameWithoutExtension(doc.Name) == entry.DocName;
                    }

                    if (isMatch)
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
                                    if (entry.LayoutName.Equals("Model", StringComparison.OrdinalIgnoreCase))
                                    {
                                        layoutExists = true;
                                    }
                                    else
                                    {
                                        var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                        layoutExists = layoutDict.Contains(entry.LayoutName);
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
                        ["LayoutName"] = entry.LayoutName,
                        ["DocumentName"] = entry.DocName,
                        ["Document"] = targetDoc,
                        ["Timestamp"] = entry.Timestamp
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

                            SwitchViewLogging.LogLayoutChange(targetDocName, chosenDoc.Name, chosenLayoutName, true);
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

                    SwitchViewLogging.LogLayoutChange(targetDocName, chosenDoc.Name, chosenLayoutName, true);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception)
                {
                }
            }
        }

    }
}