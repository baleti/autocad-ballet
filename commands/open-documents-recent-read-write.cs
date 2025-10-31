using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.OpenDocumentsRecentReadWriteCommand))]

namespace AutoCADBallet
{
    public class OpenDocumentsRecentReadWriteCommand
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

        [CommandMethod("open-documents-recent-read-write", CommandFlags.Session)]
        public void OpenDocumentsRecentReadWrite()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

            // Build a list of all document access history (keyed by absolute path)
            var documentAccessTimes = new Dictionary<string, DateTime>();
            var documentPaths = new Dictionary<string, string>();
            var documentSessionPids = new Dictionary<string, string>();

            // Read layout logs from all documents to determine last access times and absolute paths
            // Now supports multiple document sections per log file
            if (Directory.Exists(logDirPath))
            {
                foreach (string logFile in Directory.GetFiles(logDirPath))
                {
                    try
                    {
                        var lines = File.ReadAllLines(logFile);

                        // Parse document sections (log file may contain multiple documents)
                        var sections = new List<Dictionary<string, object>>();
                        Dictionary<string, object> currentSection = null;

                        foreach (var line in lines)
                        {
                            if (string.IsNullOrWhiteSpace(line))
                                continue;

                            if (line.StartsWith("DOCUMENT_PATH:"))
                            {
                                // Start new section
                                if (currentSection != null)
                                {
                                    sections.Add(currentSection);
                                }
                                currentSection = new Dictionary<string, object>
                                {
                                    ["DocumentPath"] = line.Substring("DOCUMENT_PATH:".Length).Trim(),
                                    ["LayoutEntries"] = new List<string>()
                                };
                            }
                            else if (currentSection != null)
                            {
                                if (line.StartsWith("LAST_SESSION_PID:"))
                                {
                                    currentSection["SessionPid"] = line.Substring("LAST_SESSION_PID:".Length).Trim();
                                }
                                else if (line.StartsWith("DOCUMENT_OPENED:"))
                                {
                                    string openedTimeStr = line.Substring("DOCUMENT_OPENED:".Length).Trim();
                                    if (DateTime.TryParseExact(openedTimeStr, "yyyy-MM-dd HH:mm:ss",
                                                             System.Globalization.CultureInfo.InvariantCulture,
                                                             System.Globalization.DateTimeStyles.None, out DateTime openedTime))
                                    {
                                        currentSection["DocumentOpened"] = openedTime;
                                    }
                                }
                                else
                                {
                                    // Layout entry
                                    ((List<string>)currentSection["LayoutEntries"]).Add(line);
                                }
                            }
                        }

                        // Add last section
                        if (currentSection != null)
                        {
                            sections.Add(currentSection);
                        }

                        // Process each section to find last access time
                        foreach (var section in sections)
                        {
                            string absolutePath = section.ContainsKey("DocumentPath") ? section["DocumentPath"].ToString() : null;
                            string sessionPid = section.ContainsKey("SessionPid") ? section["SessionPid"].ToString() : "Unknown";

                            if (string.IsNullOrEmpty(absolutePath))
                                continue;

                            // Find most recent layout access time in this section
                            DateTime lastAccess = DateTime.MinValue;
                            var layoutEntries = (List<string>)section["LayoutEntries"];

                            foreach (var entry in layoutEntries)
                            {
                                var parts = entry.Split(new[] { ' ' }, 2);
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

                            // Use DOCUMENT_OPENED as fallback if no layout entries
                            if (lastAccess == DateTime.MinValue && section.ContainsKey("DocumentOpened"))
                            {
                                lastAccess = (DateTime)section["DocumentOpened"];
                            }

                            if (lastAccess != DateTime.MinValue)
                            {
                                // Use absolute path as unique key
                                string uniqueKey = absolutePath;
                                documentAccessTimes[uniqueKey] = lastAccess;
                                documentPaths[uniqueKey] = absolutePath;
                                documentSessionPids[uniqueKey] = sessionPid;
                            }
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

            // Get list of currently open documents (by absolute path for accurate comparison)
            var openDocumentPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Document doc in docs)
            {
                try
                {
                    openDocumentPaths.Add(Path.GetFullPath(doc.Name));
                }
                catch
                {
                    // Skip documents where we can't get the full path
                }
            }

            // Create entries for documents with access history that are not currently open
            foreach (var docAccess in documentAccessTimes)
            {
                string absolutePath = docAccess.Key; // Now keyed by absolute path
                DateTime lastAccess = docAccess.Value;

                // Skip if document is currently open (compare by full path)
                if (openDocumentPaths.Contains(absolutePath))
                {
                    continue;
                }

                // Get document info
                string documentPath = documentPaths.ContainsKey(absolutePath) ? documentPaths[absolutePath] : absolutePath;
                string sessionPid = documentSessionPids.ContainsKey(absolutePath) ? documentSessionPids[absolutePath] : "Unknown";
                string docName = Path.GetFileNameWithoutExtension(absolutePath);

                availableDocuments.Add(new Dictionary<string, object>
                {
                    ["document name"] = docName,
                    ["last opened"] = lastAccess.ToString("yyyy-MM-dd HH:mm:ss"),
                    ["session"] = sessionPid,
                    ["absolute path"] = documentPath,
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
                        // Confirm deletion
                        if (entriesToDelete.Count == 0)
                            return false;

                        string message = entriesToDelete.Count == 1
                            ? $"Remove '{entriesToDelete[0]["document name"]}' from recent documents?"
                            : $"Remove {entriesToDelete.Count} documents from recent list?";

                        var result = MessageBox.Show(
                            message,
                            "Confirm Removal",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button1); // OK is default (focused)

                        if (result == DialogResult.OK)
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
                            return true; // Remove from grid
                        }

                        return false; // Don't remove from grid if cancelled
                    });

                if (chosen != null && chosen.Count > 0)
                {
                    Document lastOpenedDoc = null;
                    int successCount = 0;
                    int failCount = 0;

                    for (int i = 0; i < chosen.Count; i++)
                    {
                        var selectedDoc = chosen[i];
                        string docName = selectedDoc["DocumentName"].ToString();
                        string docPath = selectedDoc["DocumentPath"] as string;

                        try
                        {
                            if (string.IsNullOrEmpty(docPath))
                            {
                                SafeWriteMessage($"\nCould not open document {docName}: No file path available\n");
                                failCount++;
                                continue;
                            }

                            // Verify the file exists before attempting to open it
                            if (!File.Exists(docPath))
                            {
                                SafeWriteMessage($"\nCould not open document {docName}: File not found at {docPath}\n");
                                failCount++;
                                continue;
                            }

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
                                        SafeWriteMessage($"\nDocument already open: {docName}\n");
                                        break;
                                    }
                                }
                                catch (System.Exception checkEx)
                                {
                                    // Error checking document
                                }
                            }

                            if (alreadyOpen)
                            {
                                continue;
                            }

                            // Try to open the document in read-write mode with additional safety
                            Document openedDoc = null;
                            try
                            {
                                // docs.Open is synchronous and blocks until document is loaded
                                openedDoc = docs.Open(docPath, false); // false = read-write
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                            {
                                SafeWriteMessage($"\nAutoCAD error opening {docName}: {acEx.ErrorStatus}\n");
                                failCount++;
                                continue;
                            }

                            if (openedDoc != null)
                            {
                                SafeWriteMessage($"\nOpened document: {docName}\n");
                                lastOpenedDoc = openedDoc;
                                successCount++;

                                // Add delay and cleanup between document opens
                                if (i < chosen.Count - 1) // Don't delay after last document
                                {
                                    // Force garbage collection to free memory
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    GC.Collect();

                                    // Allow Windows message pump to process - critical for UI stability
                                    System.Windows.Forms.Application.DoEvents();

                                    // Delay to let AutoCAD stabilize
                                    int delayMs = 500; // Fixed 500ms delay
                                    System.Threading.Thread.Sleep(delayMs);
                                }
                            }
                            else
                            {
                                SafeWriteMessage($"\nFailed to open document: {docName}\n");
                                failCount++;
                            }
                        }
                        catch (System.Runtime.InteropServices.SEHException sehEx)
                        {
                            SafeWriteMessage($"\nCritical error opening {docName}. Stopping to prevent crash.\n");
                            break;
                        }
                        catch (System.AccessViolationException avEx)
                        {
                            SafeWriteMessage($"\nMemory access error opening {docName}. Stopping to prevent crash.\n");
                            break;
                        }
                        catch (System.Exception ex)
                        {
                            SafeWriteMessage($"\nError opening document {docName}: {ex.Message}\n");
                            failCount++;

                            // If we get multiple failures in a row, stop to prevent cascade failures
                            if (failCount >= 3 && successCount == 0)
                            {
                                SafeWriteMessage($"\nStopping after {failCount} consecutive failures to prevent further issues.\n");
                                break;
                            }
                        }
                    }

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
