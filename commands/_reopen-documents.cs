using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class ReopenDocumentsUtility
    {
        private static Dictionary<string, DateTime> _documentOpenTimes = new Dictionary<string, DateTime>();
        private static readonly object _lockObject = new object();

        static ReopenDocumentsUtility()
        {
            try
            {
                // Register for document events to track open times
                AcadApp.DocumentManager.DocumentCreated += OnDocumentCreated;
                AcadApp.DocumentManager.DocumentActivated += OnDocumentActivated;

                // Load existing open times from file
                LoadDocumentOpenTimes();
            }
            catch (System.Exception)
            {
                // Ignore initialization errors
            }
        }

        private static void OnDocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            TrackDocumentOpenTime(e.Document);
        }

        private static void OnDocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            // Track activation time for existing documents that might not have been tracked yet
            TrackDocumentOpenTime(e.Document);
        }

        public static void TrackDocumentOpenTime(Document doc)
        {
            if (doc == null) return;

            try
            {
                string fullPath = doc.Name;
                if (string.IsNullOrEmpty(fullPath) || fullPath.StartsWith("Drawing")) return;

                lock (_lockObject)
                {
                    // Only track if we haven't seen this document before
                    if (!_documentOpenTimes.ContainsKey(fullPath))
                    {
                        _documentOpenTimes[fullPath] = DateTime.Now;
                        SaveDocumentOpenTimes();
                    }
                }
            }
            catch (System.Exception)
            {
                // Ignore tracking errors
            }
        }

        private static void LoadDocumentOpenTimes()
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

                if (!Directory.Exists(logDirPath))
                    return;

                lock (_lockObject)
                {
                    // Read all log files from switch-view-logs directory
                    foreach (string logFile in Directory.GetFiles(logDirPath))
                    {
                        try
                        {
                            var lines = File.ReadAllLines(logFile);
                            if (lines.Length < 3)
                                continue;

                            string documentPath = null;
                            DateTime openTime = DateTime.MinValue;

                            // Parse headers
                            foreach (string line in lines)
                            {
                                if (line.StartsWith("DOCUMENT_PATH:"))
                                {
                                    documentPath = line.Substring("DOCUMENT_PATH:".Length);
                                }
                                else if (line.StartsWith("DOCUMENT_OPENED:"))
                                {
                                    string timeStr = line.Substring("DOCUMENT_OPENED:".Length);
                                    DateTime.TryParse(timeStr, out openTime);
                                }
                                else if (line.StartsWith("LAST_SESSION_PID:"))
                                {
                                    // Skip this header
                                    continue;
                                }
                                else
                                {
                                    // Reached layout entries, stop parsing headers
                                    break;
                                }
                            }

                            if (!string.IsNullOrEmpty(documentPath) && openTime != DateTime.MinValue)
                            {
                                _documentOpenTimes[documentPath] = openTime;
                            }
                        }
                        catch (System.Exception)
                        {
                            // Skip files that can't be read
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                // Ignore load errors
            }
        }

        private static void SaveDocumentOpenTimes()
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");
                Directory.CreateDirectory(logDirPath);

                lock (_lockObject)
                {
                    foreach (var kvp in _documentOpenTimes)
                    {
                        string documentPath = kvp.Key;
                        DateTime openTime = kvp.Value;
                        string projectName = Path.GetFileNameWithoutExtension(documentPath);
                        string logFilePath = Path.Combine(logDirPath, projectName);

                        try
                        {
                            // Read existing log file if it exists
                            var lines = File.Exists(logFilePath) ? File.ReadAllLines(logFilePath).ToList() : new List<string>();
                            bool foundOpenedHeader = false;
                            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;

                            // Update or add DOCUMENT_OPENED header
                            for (int i = 0; i < lines.Count; i++)
                            {
                                if (lines[i].StartsWith("DOCUMENT_OPENED:"))
                                {
                                    lines[i] = $"DOCUMENT_OPENED:{openTime:yyyy-MM-dd HH:mm:ss}";
                                    foundOpenedHeader = true;
                                    break;
                                }
                            }

                            if (!foundOpenedHeader)
                            {
                                // Need to add the header - find the right position
                                var headerLines = new List<string>();
                                int insertIndex = 0;

                                // Ensure DOCUMENT_PATH exists
                                if (lines.Count == 0 || !lines[0].StartsWith("DOCUMENT_PATH:"))
                                {
                                    headerLines.Add($"DOCUMENT_PATH:{Path.GetFullPath(documentPath)}");
                                }
                                else
                                {
                                    headerLines.Add(lines[0]);
                                    insertIndex = 1;
                                }

                                // Ensure LAST_SESSION_PID exists
                                if (lines.Count < 2 || !lines[1].StartsWith("LAST_SESSION_PID:"))
                                {
                                    headerLines.Add($"LAST_SESSION_PID:{processId}");
                                }
                                else
                                {
                                    headerLines.Add($"LAST_SESSION_PID:{processId}");
                                    insertIndex = 2;
                                }

                                // Add DOCUMENT_OPENED
                                headerLines.Add($"DOCUMENT_OPENED:{openTime:yyyy-MM-dd HH:mm:ss}");

                                // Add remaining lines
                                headerLines.AddRange(lines.Skip(insertIndex));
                                lines = headerLines;
                            }

                            File.WriteAllLines(logFilePath, lines);
                        }
                        catch (System.Exception)
                        {
                            // Skip files that can't be written
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                // Ignore save errors
            }
        }

        public static void ReopenDocuments(bool forceReadOnly)
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            // Helper function to get a valid editor
            Func<Editor> GetValidEditor = () =>
            {
                try
                {
                    var currentActive = docs.MdiActiveDocument;
                    return currentActive?.Editor;
                }
                catch
                {
                    return null;
                }
            };

            // Helper function to safely write messages
            Action<string> SafeWriteMessage = (message) =>
            {
                try
                {
                    var currentEd = GetValidEditor();
                    if (currentEd != null)
                    {
                        currentEd.WriteMessage(message);
                    }
                }
                catch
                {
                    // Ignore editor write errors
                }
            };

            // Generate current session identifier
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            string currentSessionId = $"{processId}_{sessionId}";

            // Ensure all currently open documents are tracked
            foreach (Document doc in docs)
            {
                TrackDocumentOpenTime(doc);
            }

            // Collect information about all opened documents
            var documentEntries = new List<Dictionary<string, object>>();

            foreach (Document doc in docs)
            {
                try
                {
                    string docName = Path.GetFileNameWithoutExtension(doc.Name);
                    string fullPath = doc.Name;
                    bool isReadOnly = false;
                    DateTime lastModified = DateTime.MinValue;
                    bool modifiedSinceOpened = false;

                    // Check if document corresponds to an actual file
                    bool isRealFile = !string.IsNullOrEmpty(fullPath) &&
                                     !fullPath.StartsWith("Drawing") &&
                                     fullPath.Contains("\\") || fullPath.Contains("/");

                    if (isRealFile && File.Exists(fullPath))
                    {
                        try
                        {
                            // Get file information
                            FileInfo fileInfo = new FileInfo(fullPath);
                            lastModified = fileInfo.LastWriteTime;

                            // Check if document is opened as read-only (not just file system read-only)
                            bool fileSystemReadOnly = fileInfo.IsReadOnly;
                            isReadOnly = doc.IsReadOnly;

                            // For read-only documents, check if file has been modified since opened
                            if (isReadOnly)
                            {
                                lock (_lockObject)
                                {
                                    if (_documentOpenTimes.ContainsKey(fullPath))
                                    {
                                        DateTime openTime = _documentOpenTimes[fullPath];
                                        modifiedSinceOpened = lastModified > openTime;
                                    }
                                    else
                                    {
                                        // If we don't have open time, assume not modified
                                        modifiedSinceOpened = false;
                                    }
                                }
                            }
                        }
                        catch (System.Exception)
                        {
                            // If we can't access file info, assume defaults
                            isReadOnly = false;
                            modifiedSinceOpened = false;
                        }
                    }

                    documentEntries.Add(new Dictionary<string, object>
                    {
                        ["name"] = docName,
                        ["read only"] = isReadOnly ? "Yes" : "No",
                        ["last modified"] = isRealFile && lastModified != DateTime.MinValue
                            ? lastModified.ToString("yyyy-MM-dd HH:mm:ss")
                            : "N/A",
                        ["updated"] = isReadOnly && modifiedSinceOpened ? "Yes" : "No",
                        ["session"] = currentSessionId,
                        // Internal properties for processing
                        ["Document"] = doc,
                        ["FullPath"] = fullPath,
                        ["IsRealFile"] = isRealFile,
                        ["IsReadOnly"] = isReadOnly,
                        ["ModifiedSinceOpened"] = modifiedSinceOpened
                    });
                }
                catch (System.Exception)
                {
                    // Skip documents that can't be processed
                    continue;
                }
            }

            if (documentEntries.Count == 0)
            {
                SafeWriteMessage("\nNo documents are currently open.\n");
                return;
            }

            // Sort by document name
            documentEntries = documentEntries.OrderBy(entry => entry["name"].ToString()).ToList();

            // Find current document for initial selection
            int selectedIndex = documentEntries.FindIndex(entry =>
            {
                Document entryDoc = entry["Document"] as Document;
                return entryDoc == activeDoc;
            });

            var propertyNames = new List<string> { "name", "read only", "last modified", "updated", "session" };
            var initialSelectionIndices = selectedIndex >= 0
                                            ? new List<int> { selectedIndex }
                                            : new List<int>();

            try
            {
                string commandMode = forceReadOnly ? "read-only" : "read-write";
                SafeWriteMessage($"\n=== Reopen Documents as {commandMode.ToUpper()} ===\n");

                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(
                    documentEntries,
                    propertyNames,
                    false,
                    initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    SafeWriteMessage($"\nProcessing {chosen.Count} selected document(s) for reopening as {commandMode}...\n");

                    int successCount = 0;
                    int failureCount = 0;

                    // Process each selected document
                    for (int i = 0; i < chosen.Count; i++)
                    {
                        var selectedDocument = chosen[i];
                        Document targetDoc = selectedDocument["Document"] as Document;
                        string fullPath = selectedDocument["FullPath"] as string;
                        bool isRealFile = (bool)selectedDocument["IsRealFile"];

                        SafeWriteMessage($"\n=== Processing document {i + 1} of {chosen.Count}: {selectedDocument["name"]} ===\n");

                        if (targetDoc != null && isRealFile && File.Exists(fullPath))
                        {
                            try
                            {
                                // Close the document first
                                SafeWriteMessage($"\nClosing document: {selectedDocument["name"]}\n");
                                targetDoc.CloseAndDiscard();

                                // Get current document count for diagnostics
                                int docCountBefore = docs.Count;
                                string activeDocBefore = docs.MdiActiveDocument?.Name ?? "None";
                                SafeWriteMessage($"\nBefore close: {docCountBefore} documents, active: {activeDocBefore}\n");

                                // Wait for the document to be fully closed
                                int maxWaitTime = 5000; // 5 seconds max
                                int waitInterval = 100;  // Check every 100ms
                                int totalWaitTime = 0;
                                bool documentClosed = false;

                                while (totalWaitTime < maxWaitTime)
                                {
                                    // Force processing of any pending events
                                    System.Windows.Forms.Application.DoEvents();

                                    // Check if document is still in the collection
                                    documentClosed = true;
                                    foreach (Document doc in docs)
                                    {
                                        if (doc == targetDoc)
                                        {
                                            documentClosed = false;
                                            break;
                                        }
                                    }

                                    if (documentClosed)
                                    {
                                        int docCountAfter = docs.Count;
                                        string activeDocAfter = docs.MdiActiveDocument?.Name ?? "None";
                                        SafeWriteMessage($"\nDocument closed after {totalWaitTime}ms. {docCountAfter} documents remain, active: {activeDocAfter}\n");
                                        SafeWriteMessage($"\nReopening as {commandMode}: {selectedDocument["name"]}\n");
                                        break;
                                    }

                                    System.Threading.Thread.Sleep(waitInterval);
                                    totalWaitTime += waitInterval;
                                }

                                if (!documentClosed)
                                {
                                    SafeWriteMessage($"\nWarning: Document may not be fully closed. Attempting to reopen anyway...\n");
                                }

                                // Now try to reopen the document
                                SafeWriteMessage($"\nAttempting to reopen as {commandMode}: {selectedDocument["name"]}\n");
                                SafeWriteMessage($"\nFile path: {fullPath}\n");

                                try
                                {
                                    // Use forceReadOnly parameter to determine open mode
                                    Document newDoc = AcadApp.DocumentManager.Open(fullPath, forceReadOnly);

                                    if (newDoc != null)
                                    {
                                        SafeWriteMessage($"\nDocument opened successfully. Making it active...\n");

                                        // Make it the active document
                                        AcadApp.DocumentManager.MdiActiveDocument = newDoc;

                                        // Track the new open time
                                        TrackDocumentOpenTime(newDoc);

                                        // Use safe message for success
                                        SafeWriteMessage($"\nSuccessfully reopened document as {commandMode}: {selectedDocument["name"]}\n");
                                        successCount++;
                                    }
                                    else
                                    {
                                        SafeWriteMessage($"\nFailed to reopen document - Open() returned null: {selectedDocument["name"]}\n");
                                        failureCount++;
                                    }
                                }
                                catch (System.Exception reopenEx)
                                {
                                    SafeWriteMessage($"\nException during reopen: {reopenEx.Message}\n");
                                    SafeWriteMessage($"\nException type: {reopenEx.GetType().Name}\n");
                                    if (reopenEx.InnerException != null)
                                    {
                                        SafeWriteMessage($"\nInner exception: {reopenEx.InnerException.Message}\n");
                                    }
                                    failureCount++;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                SafeWriteMessage($"\nFailed to reopen document: {ex.Message}\n");
                                failureCount++;
                            }
                        }
                        else if (targetDoc != null && !isRealFile)
                        {
                            SafeWriteMessage($"\nCannot reopen '{selectedDocument["name"]}' - it's not saved to a file yet.\n");
                            failureCount++;
                        }
                        else if (targetDoc != null && !File.Exists(fullPath))
                        {
                            SafeWriteMessage($"\nCannot reopen '{selectedDocument["name"]}' - file not found: {fullPath}\n");
                            failureCount++;
                        }
                        else
                        {
                            SafeWriteMessage($"\nSkipping '{selectedDocument["name"]}' - invalid document state.\n");
                            failureCount++;
                        }
                    }

                    // Summary
                    SafeWriteMessage($"\n=== Reopen Summary ===\n");
                    SafeWriteMessage($"Successfully reopened: {successCount} document(s) as {commandMode}\n");
                    if (failureCount > 0)
                    {
                        SafeWriteMessage($"Failed to reopen: {failureCount} document(s)\n");
                    }
                    SafeWriteMessage($"Total processed: {chosen.Count} document(s)\n");
                }
            }
            catch (System.Exception)
            {
                // DataGrid failed, command ends silently
            }
        }
    }
}