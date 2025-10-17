using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.OpenDocumentsRecentReadWriteCommand))]

namespace AutoCADBallet
{
    public class OpenDocumentsRecentReadWriteCommand
    {
        [CommandMethod("open-documents-recent-read-write", CommandFlags.Session)]
        public void OpenDocumentsRecentReadWrite()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

            // Build dictionaries mapping document names to their info from log files
            var documentPaths = new Dictionary<string, string>();
            var documentSessionPids = new Dictionary<string, string>();

            // Build a list of all document access history
            var documentAccessTimes = new Dictionary<string, DateTime>();

            // Read layout logs from all documents to determine last access times and absolute paths
            if (Directory.Exists(logDirPath))
            {
                foreach (string logFile in Directory.GetFiles(logDirPath))
                {
                    try
                    {
                        string docName = Path.GetFileName(logFile);
                        var lines = File.ReadAllLines(logFile)
                                        .Select(l => l.Trim())
                                        .Where(l => l.Length > 0)
                                        .ToList();

                        if (lines.Count == 0) continue;

                        // Parse header lines
                        string absolutePath = null;
                        string sessionPid = null;
                        int startIndex = 0;

                        // First line should contain the absolute document path
                        if (lines.Count > 0 && lines[0].StartsWith("DOCUMENT_PATH:"))
                        {
                            absolutePath = lines[0].Substring("DOCUMENT_PATH:".Length);
                            documentPaths[docName] = absolutePath;
                            startIndex = 1;
                        }

                        // Second line should contain the session PID
                        if (lines.Count > 1 && lines[1].StartsWith("LAST_SESSION_PID:"))
                        {
                            sessionPid = lines[1].Substring("LAST_SESSION_PID:".Length);
                            documentSessionPids[docName] = sessionPid;
                            startIndex = 2; // Skip both header lines when processing timestamps
                        }

                        DateTime lastAccess = DateTime.MinValue;
                        for (int i = startIndex; i < lines.Count; i++)
                        {
                            string line = lines[i];
                            var parts = line.Split(new[] { ' ' }, 2);
                            if (parts.Length == 2)
                            {
                                if (DateTime.TryParseExact(parts[0], "yyyy-MM-dd_HH:mm:ss",
                                                         System.Globalization.CultureInfo.InvariantCulture,
                                                         System.Globalization.DateTimeStyles.None, out DateTime timestamp))
                                {
                                    if (timestamp > lastAccess)
                                    {
                                        lastAccess = timestamp;
                                    }
                                }
                            }
                        }

                        if (lastAccess != DateTime.MinValue)
                        {
                            documentAccessTimes[docName] = lastAccess;
                        }
                    }
                    catch (System.Exception)
                    {
                        // Silently skip files that can't be read
                    }
                }
            }

            // Get all currently closed documents with access history
            var availableDocuments = new List<Dictionary<string, object>>();
            string currentDocName = Path.GetFileNameWithoutExtension(activeDoc.Name);

            // Get list of currently open documents
            var openDocuments = new HashSet<string>();
            foreach (Document doc in docs)
            {
                openDocuments.Add(Path.GetFileNameWithoutExtension(doc.Name));
            }

            // Create entries for documents with access history that are not currently open
            foreach (var docAccess in documentAccessTimes)
            {
                string docName = docAccess.Key;
                DateTime lastAccess = docAccess.Value;

                // Skip if document is currently open
                if (openDocuments.Contains(docName))
                {
                    continue;
                }

                // Get the document info from log file
                string documentPath = documentPaths.ContainsKey(docName) ? documentPaths[docName] : null;
                string sessionPid = documentSessionPids.ContainsKey(docName) ? documentSessionPids[docName] : "Unknown";

                availableDocuments.Add(new Dictionary<string, object>
                {
                    ["document name"] = docName,
                    ["last opened"] = lastAccess.ToString("yyyy-MM-dd HH:mm:ss"),
                    ["session"] = sessionPid,
                    ["absolute path"] = documentPath ?? "Unknown",
                    ["DocumentName"] = docName,
                    ["DocumentPath"] = documentPath,
                    ["LastAccessed"] = lastAccess,
                    ["SessionPid"] = sessionPid
                });
            }

            if (availableDocuments.Count == 0)
            {
                ed.WriteMessage("\nNo recently accessed documents found that are not currently open.\n");
                return;
            }

            // Sort by last opened time descending
            availableDocuments = availableDocuments.OrderByDescending(doc => doc["LastAccessed"])
                .ToList();

            var propertyNames = new List<string> { "document name", "last opened", "session", "absolute path" };
            var initialSelectionIndices = new List<int>(); // No initial selection

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(
                    availableDocuments,
                    propertyNames,
                    false,
                    initialSelectionIndices,
                    onDeleteEntries: (entriesToDelete) =>
                    {
                        // Delete log files from runtime folder
                        foreach (var entry in entriesToDelete)
                        {
                            if (entry.ContainsKey("DocumentName"))
                            {
                                string docName = entry["DocumentName"]?.ToString();
                                if (!string.IsNullOrEmpty(docName))
                                {
                                    string logFilePath = Path.Combine(logDirPath, docName);
                                    if (File.Exists(logFilePath))
                                    {
                                        try
                                        {
                                            File.Delete(logFilePath);
                                        }
                                        catch (System.Exception)
                                        {
                                            // Silently continue if file deletion fails
                                        }
                                    }
                                }
                            }
                        }
                        return true; // Always return true to remove from grid
                    });

                if (chosen != null && chosen.Count > 0)
                {
                    Document lastOpenedDoc = null;

                    foreach (var selectedDoc in chosen)
                    {
                        string docName = selectedDoc["DocumentName"].ToString();
                        string docPath = selectedDoc["DocumentPath"] as string;

                        try
                        {
                            if (string.IsNullOrEmpty(docPath))
                            {
                                ed.WriteMessage($"\nCould not open document {docName}: No file path available\n");
                                continue;
                            }

                            // Verify the file exists before attempting to open it
                            if (!File.Exists(docPath))
                            {
                                ed.WriteMessage($"\nCould not open document {docName}: File not found at {docPath}\n");
                                continue;
                            }

                            // Try to open the document
                            Document openedDoc = docs.Open(docPath, false);

                            if (openedDoc != null)
                            {
                                ed.WriteMessage($"\nOpened document: {docName}\n");
                                lastOpenedDoc = openedDoc;
                            }
                            else
                            {
                                ed.WriteMessage($"\nFailed to open document: {docName}\n");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError opening document {docName}: {ex.Message}\n");
                        }
                    }

                    // Set the last successfully opened document as active using event-driven approach to minimize flicker
                    if (lastOpenedDoc != null)
                    {
                        DocumentCollectionEventHandler handler = null;
                        handler = (sender, e) => {
                            if (e.Document == lastOpenedDoc)
                            {
                                docs.DocumentActivated -= handler;
                                // Document is now properly activated
                            }
                        };
                        docs.DocumentActivated += handler;
                        docs.MdiActiveDocument = lastOpenedDoc;
                    }
                }
            }
            catch (System.Exception)
            {
                // DataGrid failed, command ends silently
            }
        }
    }
}
