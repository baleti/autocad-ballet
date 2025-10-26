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
    public class SwitchViewLogging : IExtensionApplication
    {
        private static string _sessionId;
        private static Dictionary<string, string> _lastKnownLayouts = new Dictionary<string, string>();
        private static Dictionary<Document, System.EventHandler> _documentHandlers =
            new Dictionary<Document, System.EventHandler>();

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
                    LogLayoutChange(projectName, e.Name);
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
                    InitializeLogFileIfNeeded(projectName);

                    // Log the current layout of the activated document
                    string currentLayout = LayoutManager.Current.CurrentLayout;
                    if (!string.IsNullOrEmpty(currentLayout))
                    {
                        LogLayoutChange(projectName, currentLayout);
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
                    InitializeLogFileIfNeeded(projectName, clearExisting: false);
                    RegisterDocumentSelectionEvents(e.Document);
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
                    string currentLayout = LayoutManager.Current.CurrentLayout;

                    if (!string.IsNullOrEmpty(currentLayout))
                    {
                        LogLayoutChange(projectName, currentLayout);
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnDocumentBecameCurrent error: {ex.Message}");
            }
        }

        private static void InitializeLogFileIfNeeded(string projectName, bool clearExisting = false)
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

                if (!Directory.Exists(logDirPath))
                    Directory.CreateDirectory(logDirPath);

                string logFilePath = Path.Combine(logDirPath, projectName);

                if (clearExisting && File.Exists(logFilePath))
                {
                    // Clear the contents of the log file
                    File.WriteAllText(logFilePath, string.Empty);
                    _lastKnownLayouts.Remove(projectName);
                }

                // Ensure the log file starts with the absolute document path and session info
                if (!File.Exists(logFilePath) || clearExisting)
                {
                    Document activeDoc = AcadApp.DocumentManager.MdiActiveDocument;
                    if (activeDoc != null)
                    {
                        string absolutePath = Path.GetFullPath(activeDoc.Name);
                        int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
                        string openTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        File.WriteAllText(logFilePath, $"DOCUMENT_PATH:{absolutePath}\nLAST_SESSION_PID:{processId}\nDOCUMENT_OPENED:{openTime}\n");
                    }
                }
                else
                {
                    // Check if the header lines already exist
                    var lines = File.ReadAllLines(logFilePath).ToList();
                    bool needsUpdate = false;
                    var headerLines = new List<string>();
                    int headerCount = 0;

                    Document activeDoc = AcadApp.DocumentManager.MdiActiveDocument;
                    if (activeDoc != null)
                    {
                        string absolutePath = Path.GetFullPath(activeDoc.Name);
                        int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
                        string openTime = null;

                        // Check for document path (first header)
                        if (lines.Count == 0 || !lines[0].StartsWith("DOCUMENT_PATH:"))
                        {
                            headerLines.Add($"DOCUMENT_PATH:{absolutePath}");
                            needsUpdate = true;
                        }
                        else
                        {
                            headerLines.Add(lines[0]);
                            headerCount = 1;
                        }

                        // Check for session PID (second header)
                        if (lines.Count < 2 || !lines[1].StartsWith("LAST_SESSION_PID:"))
                        {
                            headerLines.Add($"LAST_SESSION_PID:{processId}");
                            needsUpdate = true;
                        }
                        else
                        {
                            headerLines.Add($"LAST_SESSION_PID:{processId}");
                            headerCount = 2;
                            needsUpdate = true; // Always update PID
                        }

                        // Check for document opened time (third header)
                        if (lines.Count < 3 || !lines[2].StartsWith("DOCUMENT_OPENED:"))
                        {
                            // Use current time as default
                            openTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            headerLines.Add($"DOCUMENT_OPENED:{openTime}");
                            needsUpdate = true;
                        }
                        else
                        {
                            headerLines.Add(lines[2]); // Keep existing open time
                            headerCount = 3;
                        }

                        // Add remaining content after headers
                        headerLines.AddRange(lines.Skip(headerCount));

                        if (needsUpdate)
                        {
                            File.WriteAllLines(logFilePath, headerLines);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.InitializeLogFileIfNeeded error: {ex.Message}");
            }
        }

        public static void LogLayoutChange(string projectName, string layoutName)
        {
            LogLayoutChange(projectName, layoutName, false);
        }

        public static void LogLayoutChange(string projectName, string layoutName, bool forceUpdate)
        {
            try
            {
                // Check if this is the same as the last known layout for this project
                // Skip duplicate check if forceUpdate is true (for manual switch commands)
                if (!forceUpdate && _lastKnownLayouts.ContainsKey(projectName) &&
                    _lastKnownLayouts[projectName] == layoutName)
                {
                    return; // Don't log duplicate consecutive layout changes
                }

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "switch-view-logs");

                if (!Directory.Exists(logDirPath))
                    Directory.CreateDirectory(logDirPath);

                string logFilePath = Path.Combine(logDirPath, projectName);
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss");
                string logEntry = $"{timestamp} {layoutName}";

                // Read existing entries and separate headers from layout entries
                var allLines = File.Exists(logFilePath) ?
                    File.ReadAllLines(logFilePath).ToList() : new List<string>();

                var headerLines = new List<string>();
                var layoutEntries = new List<string>();

                // Separate headers (lines starting with known header prefixes) from layout entries
                foreach (var line in allLines)
                {
                    if (line.StartsWith("DOCUMENT_PATH:") ||
                        line.StartsWith("LAST_SESSION_PID:") ||
                        line.StartsWith("DOCUMENT_OPENED:"))
                    {
                        headerLines.Add(line);
                    }
                    else if (!string.IsNullOrWhiteSpace(line))
                    {
                        layoutEntries.Add(line);
                    }
                }

                // Check if this layout already exists and update its timestamp, or add new entry
                bool layoutExists = false;
                for (int i = 0; i < layoutEntries.Count; i++)
                {
                    var parts = layoutEntries[i].Split(new[] { ' ' }, 2);
                    if (parts.Length == 2 && parts[1] == layoutName)
                    {
                        // Update timestamp for existing layout
                        layoutEntries[i] = logEntry;
                        layoutExists = true;
                        break;
                    }
                }

                // If layout doesn't exist, add it as a new entry
                if (!layoutExists)
                {
                    layoutEntries.Add(logEntry);
                }

                // Write headers first, then layout entries
                var finalLines = new List<string>();
                finalLines.AddRange(headerLines);
                finalLines.AddRange(layoutEntries);
                File.WriteAllLines(logFilePath, finalLines);

                // Update our cache
                _lastKnownLayouts[projectName] = layoutName;
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