using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.PurgeDocuments))]

namespace AutoCADBallet
{
    public static class PurgeDiagnostics
    {
        private static string _logFilePath = null;

        private static void InitializeLogFile()
        {
            if (_logFilePath == null)
            {
                try
                {
                    string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string logDirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "diagnostics");

                    if (!Directory.Exists(logDirPath))
                        Directory.CreateDirectory(logDirPath);

                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
                    _logFilePath = Path.Combine(logDirPath, $"purge-documents_{timestamp}.log");

                    File.WriteAllText(_logFilePath, $"=== Purge Documents Diagnostics Log - {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ==={Environment.NewLine}");
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"PurgeDiagnostics initialization error: {ex.Message}");
                }
            }
        }

        public static void Log(string message)
        {
            try
            {
                InitializeLogFile();
                if (_logFilePath != null)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    File.AppendAllText(_logFilePath, $"[{timestamp}] {message}{Environment.NewLine}");
                }
            }
            catch
            {
                // Silently ignore logging errors
            }
        }

        public static void LogException(string context, System.Exception ex)
        {
            Log($"EXCEPTION in {context}: {ex.GetType().Name} - {ex.Message}");
            if (ex.StackTrace != null)
            {
                Log($"  Stack trace: {ex.StackTrace}");
            }
        }
    }

    public class PurgeDocuments
    {
        [CommandMethod("purge-documents", CommandFlags.Session)]
        public void PurgeDocumentsCommand()
        {
            PurgeDiagnostics.Log("=== COMMAND START ===");

            var activeDoc = AcadApp.DocumentManager.MdiActiveDocument;
            if (activeDoc == null)
            {
                PurgeDiagnostics.Log("ERROR: No active document found");
                System.Windows.Forms.MessageBox.Show("No active document found.", "Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }

            PurgeDiagnostics.Log($"Active document: {activeDoc.Name}");
            var ed = activeDoc.Editor;

            try
            {
                ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                PurgeDiagnostics.LogException("PurgeDocumentsCommand", ex);
                ed.WriteMessage($"\nError in purge-documents: {ex.Message}\n");
            }
            finally
            {
                PurgeDiagnostics.Log("=== COMMAND END ===");
            }
        }

        private static void ExecuteApplicationScope(Editor ed)
        {
            PurgeDiagnostics.Log("=== ExecuteApplicationScope START ===");

            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null)
            {
                PurgeDiagnostics.Log("ERROR: No active document in ExecuteApplicationScope");
                return;
            }

            PurgeDiagnostics.Log($"Active doc in scope: {activeDoc.Name}");

            // Generate session identifier
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            string currentSessionId = $"{processId}_{sessionId}";
            PurgeDiagnostics.Log($"Session ID: {currentSessionId}");

            var allDocuments = new List<Dictionary<string, object>>();
            int currentDocIndex = -1;
            int docIndex = 0;

            PurgeDiagnostics.Log("Iterating through open documents...");
            // Iterate through all open documents
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileName(doc.Name);
                string docFullPath = doc.Name;
                bool isActiveDoc = (doc == activeDoc);

                PurgeDiagnostics.Log($"Processing doc: {docName} (active={isActiveDoc})");

                if (isActiveDoc)
                {
                    currentDocIndex = docIndex;
                }

                allDocuments.Add(new Dictionary<string, object>
                {
                    ["document"] = Path.GetFileNameWithoutExtension(docName),
                    ["autocad session"] = currentSessionId,
                    ["IsActive"] = isActiveDoc,
                    ["DocumentObject"] = doc,
                    ["DocumentPath"] = docFullPath
                });

                docIndex++;
            }

            PurgeDiagnostics.Log($"Total documents found: {allDocuments.Count}");

            if (allDocuments.Count == 0)
            {
                PurgeDiagnostics.Log("No documents found, exiting");
                ed.WriteMessage("\nNo documents found in session.");
                return;
            }

            // Sort documents by name
            PurgeDiagnostics.Log("Sorting documents by name...");
            allDocuments = allDocuments.OrderBy(d => d["document"].ToString()).ToList();

            // Update currentDocIndex after sorting
            currentDocIndex = allDocuments.FindIndex(d => (bool)d["IsActive"]);
            PurgeDiagnostics.Log($"Current doc index after sorting: {currentDocIndex}");

            var propertyNames = new List<string> { "document", "autocad session" };
            var initialSelectionIndices = new List<int>();
            if (currentDocIndex >= 0)
            {
                initialSelectionIndices.Add(currentDocIndex);
            }

            try
            {
                PurgeDiagnostics.Log($"Showing DataGrid with {allDocuments.Count} documents");
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allDocuments, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    PurgeDiagnostics.Log($"User selected {chosen.Count} documents");
                    foreach (var doc in chosen)
                    {
                        PurgeDiagnostics.Log($"  Selected: {doc["document"]}");
                    }

                    int totalProcessed = 0;
                    int totalPurged = 0;

                    // Process each selected document
                    PurgeDiagnostics.Log($"Starting to process {chosen.Count} documents");
                    foreach (var docInfo in chosen)
                    {
                        Document targetDoc = null;
                        string docName = "";

                        try
                        {
                            targetDoc = docInfo["DocumentObject"] as Document;
                            docName = Path.GetFileNameWithoutExtension(targetDoc.Name);

                            PurgeDiagnostics.Log($"\n>>> Processing document: {docName}");
                            PurgeDiagnostics.Log($"    Full path: {targetDoc.Name}");

                            // Use current active document's editor for messages
                            Document currentActiveDoc = docs.MdiActiveDocument;
                            if (currentActiveDoc != null)
                            {
                                currentActiveDoc.Editor.WriteMessage($"\n\nProcessing document: {docName}...");
                            }

                            // Activate the target document if needed
                            if (docs.MdiActiveDocument != targetDoc)
                            {
                                PurgeDiagnostics.Log($"    Need to switch to document '{docName}'");
                                docs.MdiActiveDocument = targetDoc;

                                // Wait for document switch to complete
                                int switchWaitCount = 0;
                                while (docs.MdiActiveDocument != targetDoc && switchWaitCount < 50)
                                {
                                    System.Windows.Forms.Application.DoEvents();
                                    System.Threading.Thread.Sleep(100);
                                    switchWaitCount++;
                                }

                                if (docs.MdiActiveDocument != targetDoc)
                                {
                                    PurgeDiagnostics.Log($"    ERROR: Document switch FAILED after {switchWaitCount * 100}ms");
                                    if (docs.MdiActiveDocument != null)
                                    {
                                        docs.MdiActiveDocument.Editor.WriteMessage($"\nERROR: Failed to switch to document '{docName}', skipping.");
                                    }
                                    continue;
                                }
                                PurgeDiagnostics.Log($"    Document switch successful after {switchWaitCount * 100}ms");
                            }
                            else
                            {
                                PurgeDiagnostics.Log($"    Document '{docName}' is already active");
                            }

                            // Purge the document
                            PurgeDiagnostics.Log($"    Calling PurgeDocument...");
                            int purgedCount = PurgeDocument(targetDoc);
                            PurgeDiagnostics.Log($"    PurgeDocument returned: {purgedCount} objects purged");

                            if (purgedCount > 0)
                            {
                                targetDoc.Editor.WriteMessage($"\nPurged {purgedCount} unreferenced object(s).");
                                totalPurged += purgedCount;
                            }
                            else
                            {
                                targetDoc.Editor.WriteMessage($"\nNo objects to purge.");
                            }

                            totalProcessed++;
                            PurgeDiagnostics.Log($"    Document processed successfully. Total so far: {totalProcessed}");
                        }
                        catch (System.Exception ex)
                        {
                            PurgeDiagnostics.LogException($"Processing document '{docName}'", ex);
                            // Use current active document's editor for error message
                            try
                            {
                                Document activeDocForError = docs.MdiActiveDocument;
                                if (activeDocForError != null)
                                {
                                    activeDocForError.Editor.WriteMessage($"\nERROR processing document '{docName}': {ex.Message}");
                                }
                            }
                            catch
                            {
                                // If even writing the error message fails, just continue
                            }
                        }
                    }

                    // Final summary
                    PurgeDiagnostics.Log($"=== Final Summary ===");
                    PurgeDiagnostics.Log($"Total documents processed: {totalProcessed}");
                    PurgeDiagnostics.Log($"Total objects purged: {totalPurged}");

                    Document currentDoc = docs.MdiActiveDocument;
                    if (currentDoc != null)
                    {
                        currentDoc.Editor.WriteMessage($"\n\n=== Summary ===");
                        currentDoc.Editor.WriteMessage($"\nTotal documents processed: {totalProcessed}");
                        currentDoc.Editor.WriteMessage($"\nTotal objects purged: {totalPurged}");
                    }
                }
                else
                {
                    PurgeDiagnostics.Log("User cancelled or no documents selected");
                }
            }
            catch (System.Exception ex)
            {
                PurgeDiagnostics.LogException("ExecuteApplicationScope", ex);
                Document currentDoc = docs.MdiActiveDocument;
                if (currentDoc != null)
                {
                    currentDoc.Editor.WriteMessage($"\nError in purge-documents command: {ex.Message}");
                }
            }
            finally
            {
                PurgeDiagnostics.Log("=== ExecuteApplicationScope END ===");
            }
        }

        private static int PurgeDocument(Document doc)
        {
            PurgeDiagnostics.Log($"      >> PurgeDocument START for {doc.Name}");
            int totalPurged = 0;
            Database db = doc.Database;
            int maxIterations = 10; // Purge up to 10 times to catch nested references
            int iteration = 0;

            PurgeDiagnostics.Log($"      Starting purge iterations (max {maxIterations})");

            while (iteration < maxIterations)
            {
                iteration++;
                PurgeDiagnostics.Log($"      === PASS {iteration} START ===");
                int purgedThisPass = 0;

                // Lock document for each pass to avoid long-running lock issues
                PurgeDiagnostics.Log($"      Locking document for pass {iteration}...");
                using (DocumentLock docLock = doc.LockDocument())
                {
                    PurgeDiagnostics.Log($"      Document locked");

                    // Build a complete collection of all objects to purge
                    ObjectIdCollection allObjectsToPurge = new ObjectIdCollection();

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        PurgeDiagnostics.Log($"      Transaction started for pass {iteration}");
                        PurgeDiagnostics.Log($"      Collecting all objects to purge...");

                        // Collect all symbol table objects
                        CollectSymbolTableObjects(db.BlockTableId, tr, allObjectsToPurge);
                        CollectSymbolTableObjects(db.LayerTableId, tr, allObjectsToPurge);
                        CollectSymbolTableObjects(db.TextStyleTableId, tr, allObjectsToPurge);
                        CollectSymbolTableObjects(db.DimStyleTableId, tr, allObjectsToPurge);
                        CollectSymbolTableObjects(db.LinetypeTableId, tr, allObjectsToPurge);
                        CollectSymbolTableObjects(db.RegAppTableId, tr, allObjectsToPurge);

                        // Collect dictionary objects
                        CollectDictionaryObjects(db, "ACAD_MATERIAL", tr, allObjectsToPurge);
                        CollectDictionaryObjects(db, "ACAD_TABLESTYLE", tr, allObjectsToPurge);
                        CollectDictionaryObjects(db, "ACAD_MLINESTYLE", tr, allObjectsToPurge);
                        CollectDictionaryObjects(db, "ACAD_MLEADERSTYLE", tr, allObjectsToPurge);
                        CollectDictionaryObjects(db, "ACAD_PLOTSTYLENAME", tr, allObjectsToPurge);
                        CollectDictionaryObjects(db, "ACAD_VISUALSTYLE", tr, allObjectsToPurge);
                        CollectDictionaryObjects(db, "ACAD_DETAILVIEWSTYLE", tr, allObjectsToPurge);
                        CollectDictionaryObjects(db, "ACAD_SECTIONVIEWSTYLE", tr, allObjectsToPurge);

                        PurgeDiagnostics.Log($"      Collected {allObjectsToPurge.Count} total objects");

                        // Call db.Purge once with all objects
                        PurgeDiagnostics.Log($"      Calling db.Purge on all collected objects...");
                        db.Purge(allObjectsToPurge);
                        PurgeDiagnostics.Log($"      db.Purge returned {allObjectsToPurge.Count} purgeable objects");

                        // Erase all purgeable objects
                        if (allObjectsToPurge.Count > 0)
                        {
                            PurgeDiagnostics.Log($"      Erasing {allObjectsToPurge.Count} purgeable objects...");
                            foreach (ObjectId objId in allObjectsToPurge)
                            {
                                try
                                {
                                    DBObject obj = tr.GetObject(objId, OpenMode.ForWrite);
                                    obj.Erase();
                                    purgedThisPass++;
                                }
                                catch (System.Exception ex)
                                {
                                    PurgeDiagnostics.Log($"        Failed to erase object {objId}: {ex.Message}");
                                }
                            }
                            PurgeDiagnostics.Log($"      Successfully erased {purgedThisPass} objects");
                        }

                        PurgeDiagnostics.Log($"      Committing transaction for pass {iteration}...");
                        tr.Commit();
                        PurgeDiagnostics.Log($"      Transaction committed");
                    }
                    PurgeDiagnostics.Log($"      Releasing document lock for pass {iteration}");
                }
                // Document lock released here

                totalPurged += purgedThisPass;
                PurgeDiagnostics.Log($"      === PASS {iteration} END: Purged {purgedThisPass} objects ===");
                doc.Editor.WriteMessage($"\n  Pass {iteration}: Purged {purgedThisPass} object(s)");

                // If nothing was purged, we're done
                if (purgedThisPass == 0)
                {
                    PurgeDiagnostics.Log($"      No objects purged in pass {iteration}, stopping iterations");
                    break;
                }
            }

            PurgeDiagnostics.Log($"      Completed {iteration} purge passes, total purged: {totalPurged}");
            PurgeDiagnostics.Log($"      << PurgeDocument END, returning {totalPurged}");
            return totalPurged;
        }

        private static void CollectSymbolTableObjects(ObjectId tableId, Transaction tr, ObjectIdCollection collection)
        {
            try
            {
                SymbolTable table = tr.GetObject(tableId, OpenMode.ForRead) as SymbolTable;
                if (table != null)
                {
                    foreach (ObjectId objId in table)
                    {
                        collection.Add(objId);
                    }
                }
            }
            catch (System.Exception ex)
            {
                PurgeDiagnostics.Log($"      Error collecting from symbol table: {ex.Message}");
            }
        }

        private static void CollectDictionaryObjects(Database db, string dictionaryKey, Transaction tr, ObjectIdCollection collection)
        {
            try
            {
                DBDictionary namedObjDict = tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (namedObjDict != null && namedObjDict.Contains(dictionaryKey))
                {
                    ObjectId dictId = namedObjDict.GetAt(dictionaryKey);
                    DBDictionary dict = tr.GetObject(dictId, OpenMode.ForRead) as DBDictionary;
                    if (dict != null)
                    {
                        foreach (DBDictionaryEntry entry in dict)
                        {
                            collection.Add(entry.Value);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                PurgeDiagnostics.Log($"      Error collecting from dictionary {dictionaryKey}: {ex.Message}");
            }
        }

    }
}
