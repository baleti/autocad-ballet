using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using LayoutEventArgs = Autodesk.AutoCAD.DatabaseServices.LayoutEventArgs;

[assembly: ExtensionApplication(typeof(AutoCADBallet.SwitchViewLogging))]

namespace AutoCADBallet
{
    // Helper class to represent a document section in the log file
    internal class DocumentSection
    {
        public string DocumentPath { get; set; }
        public string SessionPid { get; set; }
        public string DocumentOpened { get; set; }
        public List<string> LayoutEntries { get; set; } = new List<string>();

        public List<string> ToLines()
        {
            var lines = new List<string>();
            lines.Add($"DOCUMENT_PATH:{DocumentPath}");
            lines.Add($"LAST_SESSION_PID:{SessionPid}");
            lines.Add($"DOCUMENT_OPENED:{DocumentOpened}");
            lines.AddRange(LayoutEntries);
            return lines;
        }
    }

    public class SwitchViewLogging : IExtensionApplication
    {
        private static string _sessionId;
        private static string _lastActiveDocumentLayout = null; // Tracks last active document+layout globally
        private static string _previousActiveDocumentLayout = null; // Tracks the document+layout before the last one
        private static DateTime _lastLogTime = DateTime.MinValue; // Tracks when last log occurred
        private static Dictionary<Document, System.EventHandler> _documentHandlers =
            new Dictionary<Document, System.EventHandler>();
        private static HashSet<Document> _newlyCreatedDocuments = new HashSet<Document>(); // Track newly opened documents

        public void Initialize()
        {
            // Generate unique session identifier combining process ID and session ID
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            _sessionId = $"{processId}_{sessionId}";

            try
            {
                // Register for layout switch events
                LayoutManager.Current.LayoutSwitched += OnLayoutSwitched;

                // Register for document events
                AcadApp.DocumentManager.DocumentActivated += OnDocumentActivated;
                AcadApp.DocumentManager.DocumentCreated += OnDocumentCreated;
                AcadApp.DocumentManager.DocumentBecameCurrent += OnDocumentBecameCurrent;
                AcadApp.DocumentManager.DocumentToBeDestroyed += OnDocumentToBeDestroyed;

                // Hook up selection logging for existing documents
                foreach (Document doc in AcadApp.DocumentManager)
                {
                    RegisterDocumentSelectionEvents(doc);
                }
            }
            catch (System.Exception ex)
            {
                // Silently handle initialization errors to not interfere with AutoCAD startup
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging initialization error: {ex.Message}");
            }
        }

        public void Terminate()
        {
            try
            {
                // Unregister events to prevent memory leaks
                LayoutManager.Current.LayoutSwitched -= OnLayoutSwitched;
                AcadApp.DocumentManager.DocumentActivated -= OnDocumentActivated;
                AcadApp.DocumentManager.DocumentCreated -= OnDocumentCreated;
                AcadApp.DocumentManager.DocumentBecameCurrent -= OnDocumentBecameCurrent;
                AcadApp.DocumentManager.DocumentToBeDestroyed -= OnDocumentToBeDestroyed;

                // Unregister document-specific selection events
                foreach (var kvp in _documentHandlers.ToList())
                {
                    UnregisterDocumentSelectionEvents(kvp.Key);
                }
                _documentHandlers.Clear();
            }
            catch (System.Exception ex)
            {
                // Silently handle termination errors
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging termination error: {ex.Message}");
            }
        }

        private static void OnLayoutSwitched(object sender, LayoutEventArgs e)
        {
            try
            {
                Document activeDoc = AcadApp.DocumentManager.MdiActiveDocument;
                if (activeDoc != null)
                {
                    string projectName = Path.GetFileNameWithoutExtension(activeDoc.Name) ?? "UnknownProject";
                    LogLayoutChange(projectName, activeDoc.Name, e.Name);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnLayoutSwitched error: {ex.Message}");
            }
        }

        private static void OnDocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            try
            {
                if (e.Document != null)
                {
                    string projectName = Path.GetFileNameWithoutExtension(e.Document.Name) ?? "UnknownProject";

                    // Check if we need to initialize the log file for this document
                    InitializeLogFileIfNeeded(projectName, e.Document);

                    // Check if this is a newly created/opened document
                    bool isNewlyOpened = _newlyCreatedDocuments.Contains(e.Document);
                    if (isNewlyOpened)
                    {
                        _newlyCreatedDocuments.Remove(e.Document);
                    }

                    // Get the layout from the activated document's database to ensure accuracy
                    string currentLayout = null;
                    try
                    {
                        using (var docLock = e.Document.LockDocument(DocumentLockMode.Read, null, null, false))
                        {
                            currentLayout = LayoutManager.Current.CurrentLayout;
                        }
                    }
                    catch
                    {
                        // Fallback if locking fails
                        currentLayout = LayoutManager.Current.CurrentLayout;
                    }

                    if (!string.IsNullOrEmpty(currentLayout))
                    {
                        // Force update for newly opened documents to ensure they appear with current timestamp
                        LogLayoutChange(projectName, e.Document.Name, currentLayout, isNewlyOpened);
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnDocumentActivated error: {ex.Message}");
            }
        }

        private static void OnDocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            try
            {
                if (e.Document != null)
                {
                    string projectName = Path.GetFileNameWithoutExtension(e.Document.Name) ?? "UnknownProject";
                    InitializeLogFileIfNeeded(projectName, e.Document, clearExisting: false);
                    RegisterDocumentSelectionEvents(e.Document);

                    // Mark as newly created so first activation logs with forceUpdate
                    _newlyCreatedDocuments.Add(e.Document);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnDocumentCreated error: {ex.Message}");
            }
        }

        private static void OnDocumentToBeDestroyed(object sender, DocumentCollectionEventArgs e)
        {
            try
            {
                if (e.Document != null)
                {
                    UnregisterDocumentSelectionEvents(e.Document);
                    // Clean up newly created documents tracking
                    _newlyCreatedDocuments.Remove(e.Document);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnDocumentToBeDestroyed error: {ex.Message}");
            }
        }

        private static void OnDocumentBecameCurrent(object sender, DocumentCollectionEventArgs e)
        {
            try
            {
                if (e.Document != null)
                {
                    string projectName = Path.GetFileNameWithoutExtension(e.Document.Name) ?? "UnknownProject";

                    // Check if this is a newly created/opened document
                    bool isNewlyOpened = _newlyCreatedDocuments.Contains(e.Document);
                    if (isNewlyOpened)
                    {
                        _newlyCreatedDocuments.Remove(e.Document);
                    }

                    string currentLayout = LayoutManager.Current.CurrentLayout;

                    if (!string.IsNullOrEmpty(currentLayout))
                    {
                        // Force update for newly opened documents to ensure they appear with current timestamp
                        LogLayoutChange(projectName, e.Document.Name, currentLayout, isNewlyOpened);
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnDocumentBecameCurrent error: {ex.Message}");
            }
        }

        // Parse log file into document sections
        private static List<DocumentSection> ParseDocumentSections(string logFilePath)
        {
            var sections = new List<DocumentSection>();

            if (!File.Exists(logFilePath))
                return sections;

            var lines = File.ReadAllLines(logFilePath);
            DocumentSection currentSection = null;

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                if (line.StartsWith("DOCUMENT_PATH:"))
                {
                    // Start a new section
                    if (currentSection != null)
                    {
                        sections.Add(currentSection);
                    }
                    currentSection = new DocumentSection
                    {
                        DocumentPath = line.Substring("DOCUMENT_PATH:".Length).Trim()
                    };
                }
                else if (currentSection != null)
                {
                    if (line.StartsWith("LAST_SESSION_PID:"))
                    {
                        currentSection.SessionPid = line.Substring("LAST_SESSION_PID:".Length).Trim();
                    }
                    else if (line.StartsWith("DOCUMENT_OPENED:"))
                    {
                        currentSection.DocumentOpened = line.Substring("DOCUMENT_OPENED:".Length).Trim();
                    }
                    else
                    {
                        // Layout entry
                        currentSection.LayoutEntries.Add(line);
                    }
                }
            }

            // Add the last section
            if (currentSection != null)
            {
                sections.Add(currentSection);
            }

            return sections;
        }

        // Find section by document path (case-insensitive)
        private static DocumentSection FindSectionByPath(List<DocumentSection> sections, string documentPath)
        {
            string normalizedPath = Path.GetFullPath(documentPath).ToLowerInvariant();
            return sections.FirstOrDefault(s =>
                !string.IsNullOrEmpty(s.DocumentPath) &&
                Path.GetFullPath(s.DocumentPath).ToLowerInvariant() == normalizedPath);
        }

        // Write all sections back to log file
        private static void WriteSections(string logFilePath, List<DocumentSection> sections)
        {
            var allLines = new List<string>();

            for (int i = 0; i < sections.Count; i++)
            {
                allLines.AddRange(sections[i].ToLines());

                // Add blank line between sections (except after last one)
                if (i < sections.Count - 1)
                {
                    allLines.Add("");
                }
            }

            File.WriteAllLines(logFilePath, allLines);
        }

        private static void InitializeLogFileIfNeeded(string projectName, Document document, bool clearExisting = false)
        {
            try
            {
                if (document == null)
                    return;

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

                if (!Directory.Exists(logDirPath))
                    Directory.CreateDirectory(logDirPath);

                string logFilePath = Path.Combine(logDirPath, projectName);
                string absolutePath = Path.GetFullPath(document.Name);

                if (clearExisting && File.Exists(logFilePath))
                {
                    // Clear the contents of the log file
                    File.WriteAllText(logFilePath, string.Empty);
                }

                // Parse existing sections or create empty list
                var sections = ParseDocumentSections(logFilePath);

                // Find or create section for this document
                var section = FindSectionByPath(sections, absolutePath);

                if (section == null)
                {
                    // Create new section for this document
                    section = new DocumentSection
                    {
                        DocumentPath = absolutePath,
                        SessionPid = System.Diagnostics.Process.GetCurrentProcess().Id.ToString(),
                        DocumentOpened = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    };
                    sections.Add(section);

                    // Write updated sections
                    WriteSections(logFilePath, sections);
                }
                else
                {
                    // Update session PID for existing section (document was reopened)
                    section.SessionPid = System.Diagnostics.Process.GetCurrentProcess().Id.ToString();
                    WriteSections(logFilePath, sections);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.InitializeLogFileIfNeeded error: {ex.Message}");
            }
        }

        public static void LogLayoutChange(string projectName, string documentPath, string layoutName)
        {
            LogLayoutChange(projectName, documentPath, layoutName, false);
        }

        public static void LogLayoutChange(string projectName, string documentPath, string layoutName, bool forceUpdate)
        {
            try
            {
                string absolutePath = Path.GetFullPath(documentPath);
                string currentDocLayoutKey = $"{absolutePath}|{layoutName}";

                // Check if this is the same as the last active document+layout combination globally
                // This allows logging when switching between documents even if they're on the same layout
                // Skip duplicate check if forceUpdate is true (for manual switch commands)
                if (!forceUpdate && _lastActiveDocumentLayout != null &&
                    _lastActiveDocumentLayout == currentDocLayoutKey)
                {
                    return; // Don't log duplicate consecutive document+layout changes
                }

                // Detect "bounce-back" pattern: Doc A → Doc B → Doc A (when Session command returns to original context)
                // This happens when a Session command opens a document, then AutoCAD returns control to the original document
                // Suppress bounce-back if it happens within 3 seconds and we're not forcing an update
                if (!forceUpdate && _previousActiveDocumentLayout != null &&
                    _previousActiveDocumentLayout == currentDocLayoutKey)
                {
                    TimeSpan timeSinceLastLog = DateTime.Now - _lastLogTime;
                    if (timeSinceLastLog.TotalSeconds < 3.0)
                    {
                        // This is a bounce-back to the previous document within 3 seconds - suppress it
                        return;
                    }
                }

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

                if (!Directory.Exists(logDirPath))
                    Directory.CreateDirectory(logDirPath);

                string logFilePath = Path.Combine(logDirPath, projectName);

                // Parse existing sections
                var sections = ParseDocumentSections(logFilePath);

                // Find section for this document
                var section = FindSectionByPath(sections, absolutePath);

                if (section == null)
                {
                    // Create new section if document isn't tracked yet
                    section = new DocumentSection
                    {
                        DocumentPath = absolutePath,
                        SessionPid = System.Diagnostics.Process.GetCurrentProcess().Id.ToString(),
                        DocumentOpened = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    };
                    sections.Add(section);
                }

                // Update DOCUMENT_OPENED to current time
                section.DocumentOpened = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Update or add layout entry
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss");
                string logEntry = $"{timestamp} {layoutName}";

                bool layoutExists = false;
                for (int i = 0; i < section.LayoutEntries.Count; i++)
                {
                    var parts = section.LayoutEntries[i].Split(new[] { ' ' }, 2);
                    if (parts.Length == 2 && parts[1] == layoutName)
                    {
                        // Update timestamp for existing layout
                        section.LayoutEntries[i] = logEntry;
                        layoutExists = true;
                        break;
                    }
                }

                if (!layoutExists)
                {
                    section.LayoutEntries.Add(logEntry);
                }

                // Write all sections back to file
                WriteSections(logFilePath, sections);

                // Update our global tracking: shift current to previous, then set new current
                _previousActiveDocumentLayout = _lastActiveDocumentLayout;
                _lastActiveDocumentLayout = currentDocLayoutKey;
                _lastLogTime = DateTime.Now;
            }
            catch (System.Exception ex)
            {
                // Silently ignore logging errors to not interfere with layout switching
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.LogLayoutChange error: {ex.Message}");
            }
        }

        // Selection logging methods

        private static void RegisterDocumentSelectionEvents(Document doc)
        {
            try
            {
                // Only register if not already registered
                if (_documentHandlers.ContainsKey(doc))
                    return;

                // Create event handler for this document
                System.EventHandler handler = (sender, e) =>
                {
                    OnSelectionChanged(doc, sender, e);
                };

                // Register the handler
                doc.ImpliedSelectionChanged += handler;
                _documentHandlers[doc] = handler;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.RegisterDocumentSelectionEvents error: {ex.Message}");
            }
        }

        private static void UnregisterDocumentSelectionEvents(Document doc)
        {
            try
            {
                if (_documentHandlers.ContainsKey(doc))
                {
                    doc.ImpliedSelectionChanged -= _documentHandlers[doc];
                    _documentHandlers.Remove(doc);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.UnregisterDocumentSelectionEvents error: {ex.Message}");
            }
        }

        private static void OnSelectionChanged(Document doc, object sender, System.EventArgs e)
        {
            try
            {
                if (doc == null)
                    return;

                // Get the current implied selection
                var ed = doc.Editor;
                var selResult = ed.SelectImplied();

                if (selResult.Status != PromptStatus.OK || selResult.Value == null || selResult.Value.Count == 0)
                    return;

                // Get document name for log file
                string projectName = Path.GetFileNameWithoutExtension(doc.Name) ?? "UnknownProject";
                string documentPath = Path.GetFullPath(doc.Name);

                // Get the selection handles
                var handles = new List<string>();
                foreach (SelectedObject selObj in selResult.Value)
                {
                    if (selObj != null && !selObj.ObjectId.IsNull && selObj.ObjectId.IsValid)
                    {
                        handles.Add(selObj.ObjectId.Handle.ToString());
                    }
                }

                if (handles.Count > 0)
                {
                    LogSelection(projectName, documentPath, handles);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnSelectionChanged error: {ex.Message}");
            }
        }

        private static void LogSelection(string projectName, string documentPath, List<string> handles)
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "selection-logs");

                if (!Directory.Exists(logDirPath))
                    Directory.CreateDirectory(logDirPath);

                string logFilePath = Path.Combine(logDirPath, projectName);
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss.fff");

                // Format: timestamp|documentPath|handle1,handle2,handle3,...
                string logEntry = $"{timestamp}|{documentPath}|{string.Join(",", handles)}";

                // Append to log file
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
            catch (System.Exception ex)
            {
                // Silently ignore logging errors to not interfere with selection operations
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.LogSelection error: {ex.Message}");
            }
        }
    }
}