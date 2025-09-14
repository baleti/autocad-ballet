using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
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
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.OnDocumentCreated error: {ex.Message}");
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

                // Read existing entries to check for duplicates
                List<string> logEntries = File.Exists(logFilePath) ?
                    File.ReadAllLines(logFilePath).ToList() : new List<string>();

                // Check if this layout already exists and update its timestamp, or add new entry
                bool layoutExists = false;
                for (int i = 0; i < logEntries.Count; i++)
                {
                    var parts = logEntries[i].Split(new[] { ' ' }, 2);
                    if (parts.Length == 2 && parts[1] == layoutName)
                    {
                        // Update timestamp for existing layout
                        logEntries[i] = logEntry;
                        layoutExists = true;
                        break;
                    }
                }

                // If layout doesn't exist, add it as a new entry
                if (!layoutExists)
                {
                    logEntries.Add(logEntry);
                }

                // Write the updated entries back to the file
                File.WriteAllLines(logFilePath, logEntries);

                // Update our cache
                _lastKnownLayouts[projectName] = layoutName;
            }
            catch (System.Exception ex)
            {
                // Silently ignore logging errors to not interfere with layout switching
                System.Diagnostics.Debug.WriteLine($"SwitchViewLogging.LogLayoutChange error: {ex.Message}");
            }
        }
    }
}