using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.OpenDocumentsRecentReadOnlyCommand))]

namespace AutoCADBallet
{
    public class OpenDocumentsRecentReadOnlyCommand
    {
        // Safe message writing that handles null editor references
        private void SafeWriteMessage(string message)
        {
            try
            {
                var currentDoc = AcadApp.DocumentManager.MdiActiveDocument;
                currentDoc?.Editor?.WriteMessage(message);
            }
            catch
            {
                // Silently fail if we can't write to editor
            }
        }

        // Diagnostic logging to track crashes
        private void LogDiagnostic(string message)
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string diagPath = Path.Combine(appDataPath, "autocad-ballet", "diagnostics");
                Directory.CreateDirectory(diagPath);

                string logFile = Path.Combine(diagPath, $"open-docs-diagnostics_{DateTime.Now:yyyy-MM-dd}.log");
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                File.AppendAllText(logFile, $"[{timestamp}] {message}\n");
            }
            catch
            {
                // Silently fail if we can't write diagnostics
            }
        }

        [CommandMethod("open-documents-recent-read-only", CommandFlags.Session)]
        public void OpenDocumentsRecentReadOnly()
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

                        // Third line might contain the document opened time
                        if (lines.Count > 2 && lines[2].StartsWith("DOCUMENT_OPENED:"))
                        {
                            startIndex = 3; // Skip all three header lines when processing timestamps
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
                    LogDiagnostic($"=== START: Opening {chosen.Count} documents ===");
                    LogDiagnostic($"Current active document: {activeDoc.Name}");
                    LogDiagnostic($"Current document count: {docs.Count}");

                    Document lastOpenedDoc = null;
                    int successCount = 0;
                    int failCount = 0;

                    for (int i = 0; i < chosen.Count; i++)
                    {
                        var selectedDoc = chosen[i];
                        string docName = selectedDoc["DocumentName"].ToString();
                        string docPath = selectedDoc["DocumentPath"] as string;

                        LogDiagnostic($"--- Document {i+1}/{chosen.Count}: {docName} ---");
                        LogDiagnostic($"Path: {docPath}");

                        try
                        {
                            if (string.IsNullOrEmpty(docPath))
                            {
                                LogDiagnostic($"SKIP: Empty path");
                                SafeWriteMessage($"\nCould not open document {docName}: No file path available\n");
                                failCount++;
                                continue;
                            }

                            // Verify the file exists before attempting to open it
                            if (!File.Exists(docPath))
                            {
                                LogDiagnostic($"SKIP: File not found");
                                SafeWriteMessage($"\nCould not open document {docName}: File not found at {docPath}\n");
                                failCount++;
                                continue;
                            }

                            LogDiagnostic($"File exists, checking if already open...");

                            // Check if document is already open (race condition protection)
                            bool alreadyOpen = false;
                            int openDocCount = 0;
                            foreach (Document existingDoc in docs)
                            {
                                openDocCount++;
                                try
                                {
                                    if (string.Equals(Path.GetFullPath(existingDoc.Name), Path.GetFullPath(docPath), StringComparison.OrdinalIgnoreCase))
                                    {
                                        alreadyOpen = true;
                                        LogDiagnostic($"Already open (found in position {openDocCount})");
                                        SafeWriteMessage($"\nDocument already open: {docName}\n");
                                        break;
                                    }
                                }
                                catch (System.Exception checkEx)
                                {
                                    LogDiagnostic($"Error checking document {openDocCount}: {checkEx.Message}");
                                }
                            }

                            LogDiagnostic($"Total open documents checked: {openDocCount}");

                            if (alreadyOpen)
                            {
                                continue;
                            }

                            LogDiagnostic($"BEFORE docs.Open() - Active doc: {AcadApp.DocumentManager.MdiActiveDocument?.Name ?? "NULL"}");
                            LogDiagnostic($"BEFORE docs.Open() - Doc count: {docs.Count}");

                            // Try to open the document in read-only mode with additional safety
                            Document openedDoc = null;
                            try
                            {
                                // docs.Open is synchronous and blocks until document is loaded
                                openedDoc = docs.Open(docPath, true); // true = read-only

                                LogDiagnostic($"AFTER docs.Open() - Returned: {(openedDoc != null ? openedDoc.Name : "NULL")}");
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                            {
                                LogDiagnostic($"AutoCAD exception during open: {acEx.ErrorStatus} - {acEx.Message}");
                                SafeWriteMessage($"\nAutoCAD error opening {docName}: {acEx.ErrorStatus}\n");
                                failCount++;
                                continue;
                            }

                            if (openedDoc != null)
                            {
                                LogDiagnostic($"SUCCESS: Opened {openedDoc.Name}");
                                LogDiagnostic($"AFTER open - Active doc: {AcadApp.DocumentManager.MdiActiveDocument?.Name ?? "NULL"}");
                                LogDiagnostic($"AFTER open - Doc count: {docs.Count}");

                                SafeWriteMessage($"\nOpened document in read-only mode: {docName}\n");
                                lastOpenedDoc = openedDoc;
                                successCount++;

                                // Add delay and cleanup between document opens
                                if (i < chosen.Count - 1) // Don't delay after last document
                                {
                                    LogDiagnostic($"Forcing garbage collection...");
                                    // Force garbage collection to free memory
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    GC.Collect();
                                    LogDiagnostic($"GC complete");

                                    // Allow Windows message pump to process - critical for UI stability
                                    LogDiagnostic($"Processing Windows messages...");
                                    System.Windows.Forms.Application.DoEvents();
                                    LogDiagnostic($"DoEvents complete");

                                    // Delay to let AutoCAD stabilize
                                    int delayMs = 500; // Fixed 500ms delay
                                    LogDiagnostic($"Sleeping {delayMs}ms before next document...");
                                    System.Threading.Thread.Sleep(delayMs);
                                    LogDiagnostic($"Sleep complete");
                                }
                            }
                            else
                            {
                                LogDiagnostic($"FAIL: docs.Open returned null");
                                SafeWriteMessage($"\nFailed to open document: {docName}\n");
                                failCount++;
                            }
                        }
                        catch (System.Runtime.InteropServices.SEHException sehEx)
                        {
                            LogDiagnostic($"SEH EXCEPTION: {sehEx.Message}");
                            SafeWriteMessage($"\nCritical error opening {docName}. Stopping to prevent crash.\n");
                            break;
                        }
                        catch (System.AccessViolationException avEx)
                        {
                            LogDiagnostic($"ACCESS VIOLATION: {avEx.Message}");
                            SafeWriteMessage($"\nMemory access error opening {docName}. Stopping to prevent crash.\n");
                            break;
                        }
                        catch (System.Exception ex)
                        {
                            LogDiagnostic($"EXCEPTION: {ex.GetType().Name} - {ex.Message}");
                            SafeWriteMessage($"\nError opening document {docName}: {ex.Message}\n");
                            failCount++;

                            // If we get multiple failures in a row, stop to prevent cascade failures
                            if (failCount >= 3 && successCount == 0)
                            {
                                LogDiagnostic($"STOPPING: 3 consecutive failures");
                                SafeWriteMessage($"\nStopping after {failCount} consecutive failures to prevent further issues.\n");
                                break;
                            }
                        }
                    }

                    LogDiagnostic($"=== END: Success={successCount}, Fail={failCount} ===");

                    // Set the last successfully opened document as active
                    if (lastOpenedDoc != null)
                    {
                        try
                        {
                            docs.MdiActiveDocument = lastOpenedDoc;
                            SafeWriteMessage($"\n{successCount} document(s) opened successfully");
                            if (failCount > 0)
                            {
                                SafeWriteMessage($", {failCount} failed");
                            }
                            SafeWriteMessage($".\n");
                        }
                        catch (System.Exception ex)
                        {
                            SafeWriteMessage($"\nWarning: Could not activate last document: {ex.Message}\n");
                        }
                    }
                    else if (failCount > 0)
                    {
                        SafeWriteMessage($"\nFailed to open {failCount} document(s).\n");
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