using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.BindXrefsAndPurgeDocumentsInSession))]

namespace AutoCADBallet
{
    public class BindAndPurgeProgressDialog : Form
    {
        private Label titleLabel;
        private Label currentDocLabel;
        private Label statusLabel;
        private Label detailLabel;
        private ProgressBar progressBar;
        private Label progressTextLabel;

        public BindAndPurgeProgressDialog(int totalDocuments)
        {
            InitializeComponents(totalDocuments);
        }

        private void InitializeComponents(int totalDocuments)
        {
            this.Text = "Binding Xrefs and Purging";
            this.Size = new Size(550, 240);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true;

            // Title
            titleLabel = new Label
            {
                Text = "Binding Xrefs and Purging Documents",
                Font = new System.Drawing.Font("Arial", 10, FontStyle.Bold),
                Location = new Point(20, 15),
                AutoSize = true
            };
            this.Controls.Add(titleLabel);

            // Progress text (e.g., "Processing 2 of 5")
            progressTextLabel = new Label
            {
                Text = "Preparing...",
                Location = new Point(20, 45),
                AutoSize = true,
                Font = new System.Drawing.Font("Arial", 9, FontStyle.Bold)
            };
            this.Controls.Add(progressTextLabel);

            // Progress bar
            progressBar = new ProgressBar
            {
                Location = new Point(20, 70),
                Size = new Size(500, 25),
                Minimum = 0,
                Maximum = totalDocuments,
                Value = 0,
                Style = ProgressBarStyle.Continuous
            };
            this.Controls.Add(progressBar);

            // Current document
            currentDocLabel = new Label
            {
                Text = "",
                Location = new Point(20, 105),
                Size = new Size(500, 20),
                AutoEllipsis = true,
                Font = new System.Drawing.Font("Arial", 8, FontStyle.Regular)
            };
            this.Controls.Add(currentDocLabel);

            // Status
            statusLabel = new Label
            {
                Text = "",
                Location = new Point(20, 130),
                Size = new Size(500, 20),
                AutoEllipsis = true,
                Font = new System.Drawing.Font("Arial", 8, FontStyle.Regular)
            };
            this.Controls.Add(statusLabel);

            // Detailed status info
            detailLabel = new Label
            {
                Text = "",
                Location = new Point(20, 155),
                Size = new Size(500, 40),
                AutoEllipsis = false,
                Font = new System.Drawing.Font("Arial", 8, FontStyle.Italic),
                ForeColor = System.Drawing.Color.DarkGray
            };
            this.Controls.Add(detailLabel);
        }

        public void UpdateProgress(int current, int total, string documentName, string status, string detail = "")
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(current, total, documentName, status, detail)));
                return;
            }

            progressBar.Value = Math.Min(current, progressBar.Maximum);
            progressTextLabel.Text = $"Processing document {current} of {total}";
            currentDocLabel.Text = $"Document: {documentName}";
            statusLabel.Text = status;
            detailLabel.Text = detail;
            this.Refresh();
            System.Windows.Forms.Application.DoEvents();
        }

        public void Complete()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => Complete()));
                return;
            }

            statusLabel.Text = "Complete!";
            detailLabel.Text = "";
            this.Refresh();
        }
    }

    public static class BindAndPurgeDiagnostics
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
                    _logFilePath = Path.Combine(logDirPath, $"bind-and-purge_{timestamp}.log");

                    File.WriteAllText(_logFilePath, $"=== Bind and Purge Diagnostics Log - {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ==={Environment.NewLine}");
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"BindAndPurgeDiagnostics initialization error: {ex.Message}");
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

        public static void LogPath()
        {
            InitializeLogFile();
            if (_logFilePath != null)
            {
                System.Diagnostics.Debug.WriteLine($"Diagnostics log: {_logFilePath}");
            }
        }
    }

    public class BindXrefsAndPurgeDocumentsInSession
    {
        [CommandMethod("bind-xrefs-and-purge-documents-in-session", CommandFlags.Session)]
        public void BindXrefsAndPurgeDocumentsInSessionCommand()
        {
            BindAndPurgeDiagnostics.LogPath();
            BindAndPurgeDiagnostics.Log("=== COMMAND START ===");

            var activeDoc = AcadApp.DocumentManager.MdiActiveDocument;
            if (activeDoc == null)
            {
                BindAndPurgeDiagnostics.Log("ERROR: No active document found");
                System.Windows.Forms.MessageBox.Show("No active document found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            BindAndPurgeDiagnostics.Log($"Active document: {activeDoc.Name}");
            var ed = activeDoc.Editor;

            try
            {
                ExecuteApplicationScope(ed);
            }
            catch (System.Exception ex)
            {
                BindAndPurgeDiagnostics.LogException("BindXrefsAndPurgeDocumentsInSessionCommand", ex);
                ed.WriteMessage($"\nError in bind-xrefs-and-purge-documents-in-session: {ex.Message}\n");
            }
            finally
            {
                BindAndPurgeDiagnostics.Log("=== COMMAND END ===");
            }
        }

        private static void ExecuteApplicationScope(Editor ed)
        {
            BindAndPurgeDiagnostics.Log("=== ExecuteApplicationScope START ===");

            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null)
            {
                BindAndPurgeDiagnostics.Log("ERROR: No active document in ExecuteApplicationScope");
                return;
            }

            BindAndPurgeDiagnostics.Log($"Active doc in scope: {activeDoc.Name}");

            // Generate session identifier for this AutoCAD process
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            string currentSessionId = $"{processId}_{sessionId}";
            BindAndPurgeDiagnostics.Log($"Session ID: {currentSessionId}");

            var allDocuments = new List<Dictionary<string, object>>();
            int currentDocIndex = -1;
            int docIndex = 0;

            BindAndPurgeDiagnostics.Log("Iterating through open documents...");
            // Iterate through all open documents
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileName(doc.Name);
                string docFullPath = doc.Name;
                bool isActiveDoc = (doc == activeDoc);

                BindAndPurgeDiagnostics.Log($"Processing doc: {docName} (active={isActiveDoc})");

                if (isActiveDoc)
                {
                    currentDocIndex = docIndex;
                }

                // Count xrefs in the document
                int xrefCount = 0;
                try
                {
                    BindAndPurgeDiagnostics.Log($"  Counting xrefs in {docName}...");
                    Database db = doc.Database;
                    using (DocumentLock docLock = doc.LockDocument())
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                            foreach (ObjectId blockId in blockTable)
                            {
                                BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                                if (btr.IsFromExternalReference)
                                {
                                    xrefCount++;
                                }
                            }
                            tr.Commit();
                        }
                    }
                    BindAndPurgeDiagnostics.Log($"  Found {xrefCount} xrefs in {docName}");
                }
                catch (System.Exception ex)
                {
                    BindAndPurgeDiagnostics.Log($"  ERROR counting xrefs in {docName}: {ex.Message}");
                    // Silently skip documents that can't be read
                }

                allDocuments.Add(new Dictionary<string, object>
                {
                    ["document"] = Path.GetFileNameWithoutExtension(docName),
                    ["xref count"] = xrefCount,
                    ["autocad session"] = currentSessionId,
                    ["IsActive"] = isActiveDoc,
                    ["DocumentObject"] = doc,
                    ["DocumentPath"] = docFullPath
                });

                docIndex++;
            }

            BindAndPurgeDiagnostics.Log($"Total documents found: {allDocuments.Count}");
            if (allDocuments.Count == 0)
            {
                BindAndPurgeDiagnostics.Log("No documents found, exiting");
                ed.WriteMessage("\nNo documents found in session.");
                return;
            }

            // Sort documents by name
            BindAndPurgeDiagnostics.Log("Sorting documents by name...");
            allDocuments = allDocuments.OrderBy(d => d["document"].ToString()).ToList();

            // Update currentDocIndex after sorting
            currentDocIndex = allDocuments.FindIndex(d => (bool)d["IsActive"]);
            BindAndPurgeDiagnostics.Log($"Current doc index after sorting: {currentDocIndex}");

            var propertyNames = new List<string> { "document", "xref count", "autocad session" };
            var initialSelectionIndices = new List<int>();
            if (currentDocIndex >= 0)
            {
                initialSelectionIndices.Add(currentDocIndex);
            }

            try
            {
                BindAndPurgeDiagnostics.Log($"Showing DataGrid with {allDocuments.Count} documents");
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allDocuments, propertyNames, false, initialSelectionIndices);

                BindAndPurgeProgressDialog progressDialog = null;
                System.Threading.Thread dialogThread = null;

                if (chosen != null && chosen.Count > 0)
                {
                    BindAndPurgeDiagnostics.Log($"User selected {chosen.Count} documents");
                    foreach (var doc in chosen)
                    {
                        BindAndPurgeDiagnostics.Log($"  Selected: {doc["document"]} ({doc["xref count"]} xrefs)");
                    }

                    // Prompt for output folder
                    BindAndPurgeDiagnostics.Log("Prompting for output folder...");
                    string outputFolder = PromptForFolder(ed);
                    if (string.IsNullOrEmpty(outputFolder))
                    {
                        BindAndPurgeDiagnostics.Log("No folder selected, operation cancelled");
                        ed.WriteMessage("\nOperation cancelled - no folder selected.");
                        return;
                    }
                    BindAndPurgeDiagnostics.Log($"Output folder selected: {outputFolder}");

                    int totalProcessed = 0;
                    int totalBound = 0;
                    int totalPurged = 0;

                    try
                    {
                        // Create and show progress dialog
                        dialogThread = new System.Threading.Thread(() =>
                        {
                            progressDialog = new BindAndPurgeProgressDialog(chosen.Count);
                            progressDialog.Show();
                            System.Windows.Forms.Application.Run(progressDialog);
                        });
                        dialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
                        dialogThread.Start();

                        // Wait for dialog to be created
                        while (progressDialog == null)
                        {
                            System.Threading.Thread.Sleep(50);
                        }

                        // Process each selected document
                        BindAndPurgeDiagnostics.Log($"Starting to process {chosen.Count} documents");
                        int docCounter = 0;
                        foreach (var docInfo in chosen)
                        {
                            Document targetDoc = null;
                            string docName = "";
                            string docFullPath = "";
                            string outputPath = "";

                            try
                            {
                                docCounter++;
                                targetDoc = docInfo["DocumentObject"] as Document;
                                docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                                docFullPath = targetDoc.Name;
                                string docExtension = Path.GetExtension(targetDoc.Name);
                                outputPath = Path.Combine(outputFolder, Path.GetFileName(docFullPath));

                                BindAndPurgeDiagnostics.Log($"\n>>> Processing document: {docName}");
                                BindAndPurgeDiagnostics.Log($"    Full path: {docFullPath}");
                                BindAndPurgeDiagnostics.Log($"    Output path: {outputPath}");

                                // Update progress dialog
                                progressDialog.UpdateProgress(docCounter, chosen.Count, docName, "Preparing...",
                                    $"Switching to document and initializing...");

                                // Use current active document's editor for messages
                                Document currentActiveDoc = docs.MdiActiveDocument;
                                if (currentActiveDoc != null)
                                {
                                    currentActiveDoc.Editor.WriteMessage($"\n\nProcessing {docCounter}/{chosen.Count}: {docName}...");
                                }

                                // Activate the target document if it's not already active
                                if (docs.MdiActiveDocument != targetDoc)
                                {
                                    BindAndPurgeDiagnostics.Log($"    Need to switch to document '{docName}'");
                                    if (docs.MdiActiveDocument != null)
                                    {
                                        docs.MdiActiveDocument.Editor.WriteMessage($"\nSwitching to document '{docName}'...");
                                    }
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
                                        BindAndPurgeDiagnostics.Log($"    ERROR: Document switch FAILED after {switchWaitCount * 100}ms");
                                        if (docs.MdiActiveDocument != null)
                                        {
                                            docs.MdiActiveDocument.Editor.WriteMessage($"\nERROR: Failed to switch to document '{docName}', skipping.");
                                        }
                                        continue;
                                    }
                                    BindAndPurgeDiagnostics.Log($"    Document switch successful after {switchWaitCount * 100}ms");
                                }
                                else
                                {
                                    BindAndPurgeDiagnostics.Log($"    Document '{docName}' is already active");
                                }

                                // STEP 1: Bind xrefs
                                BindAndPurgeDiagnostics.Log($"    === STEP 1: Binding xrefs ===");
                                progressDialog.UpdateProgress(docCounter, chosen.Count, docName, "Step 1/3: Counting xrefs...", "");
                                targetDoc.Editor.WriteMessage($"\n  Step 1/3: Binding xrefs...");

                                int xrefCountBefore = CountXrefs(targetDoc);
                                BindAndPurgeDiagnostics.Log($"    Found {xrefCountBefore} xrefs before binding");

                                if (xrefCountBefore > 0)
                                {
                                    progressDialog.UpdateProgress(docCounter, chosen.Count, docName, "Step 1/3: Binding xrefs...",
                                        $"Found {xrefCountBefore} xrefs to bind. This may take a while...");
                                    targetDoc.Editor.WriteMessage($"\n    Binding {xrefCountBefore} xref(s)...");

                                    int boundCount = BindXrefsInDocument(targetDoc, targetDoc.Editor, progressDialog, docCounter, chosen.Count, docName);
                                    int xrefCountAfter = CountXrefs(targetDoc);

                                    BindAndPurgeDiagnostics.Log($"    BindXrefsInDocument returned: {boundCount} xrefs processed");
                                    BindAndPurgeDiagnostics.Log($"    Xref count before: {xrefCountBefore}, after: {xrefCountAfter}");

                                    if (xrefCountAfter < xrefCountBefore)
                                    {
                                        int actualBound = xrefCountBefore - xrefCountAfter;
                                        totalBound += actualBound;
                                        BindAndPurgeDiagnostics.Log($"    Successfully bound {actualBound} xrefs");
                                        targetDoc.Editor.WriteMessage($"\n    Successfully bound {actualBound} xref(s)!");
                                    }
                                    else if (xrefCountAfter == xrefCountBefore)
                                    {
                                        BindAndPurgeDiagnostics.Log($"    WARNING: Xref count unchanged - binding may have failed!");
                                        targetDoc.Editor.WriteMessage($"\n    WARNING: Xref count unchanged - binding may have failed.");
                                    }
                                }
                                else
                                {
                                    BindAndPurgeDiagnostics.Log($"    No xrefs found to bind");
                                    targetDoc.Editor.WriteMessage($"\n    No xrefs found to bind.");
                                }

                                // STEP 2: Purge
                                BindAndPurgeDiagnostics.Log($"    === STEP 2: Purging document ===");
                                progressDialog.UpdateProgress(docCounter, chosen.Count, docName, "Step 2/3: Purging...",
                                    "Removing unreferenced objects...");
                                targetDoc.Editor.WriteMessage($"\n  Step 2/3: Purging document...");

                                int purgedCount = PurgeDocument(targetDoc, progressDialog, docCounter, chosen.Count, docName);
                                BindAndPurgeDiagnostics.Log($"    PurgeDocument returned: {purgedCount} objects purged");

                                if (purgedCount > 0)
                                {
                                    targetDoc.Editor.WriteMessage($"\n    Purged {purgedCount} unreferenced object(s).");
                                    totalPurged += purgedCount;
                                }
                                else
                                {
                                    targetDoc.Editor.WriteMessage($"\n    No objects to purge.");
                                }

                                // STEP 3: Save
                                BindAndPurgeDiagnostics.Log($"    === STEP 3: Saving document ===");
                                BindAndPurgeDiagnostics.Log($"    Saving document to '{outputPath}'...");
                                progressDialog.UpdateProgress(docCounter, chosen.Count, docName, "Step 3/3: Saving document...",
                                    $"Saving to: {Path.GetFileName(outputPath)}");
                                targetDoc.Editor.WriteMessage($"\n  Step 3/3: Saving document to '{outputPath}'...");

                                using (DocumentLock docLock = targetDoc.LockDocument())
                                {
                                    targetDoc.Database.SaveAs(outputPath, DwgVersion.Current);
                                }

                                BindAndPurgeDiagnostics.Log($"    Document saved successfully");
                                progressDialog.UpdateProgress(docCounter, chosen.Count, docName, $"Complete!",
                                    $"Saved to {Path.GetFileName(outputPath)}");
                                targetDoc.Editor.WriteMessage($"\n    Document saved successfully.");

                                totalProcessed++;
                                BindAndPurgeDiagnostics.Log($"    Document processed successfully. Total so far: {totalProcessed}");

                                // Add a small delay before processing next document
                                BindAndPurgeDiagnostics.Log($"    Sleeping 500ms before next document...");
                                System.Threading.Thread.Sleep(500);
                            }
                            catch (System.Exception ex)
                            {
                                BindAndPurgeDiagnostics.LogException($"Processing document '{docName}'", ex);
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
                                // Continue with next document
                            }
                        }

                        // Final summary
                        BindAndPurgeDiagnostics.Log($"=== Final Summary ===");
                        BindAndPurgeDiagnostics.Log($"Total documents processed: {totalProcessed}");
                        BindAndPurgeDiagnostics.Log($"Total xrefs bound: {totalBound}");
                        BindAndPurgeDiagnostics.Log($"Total objects purged: {totalPurged}");
                        BindAndPurgeDiagnostics.Log($"Output folder: {outputFolder}");

                        Document currentDoc = docs.MdiActiveDocument;
                        if (currentDoc != null)
                        {
                            currentDoc.Editor.WriteMessage($"\n\n=== Summary ===");
                            currentDoc.Editor.WriteMessage($"\nTotal documents processed: {totalProcessed}");
                            currentDoc.Editor.WriteMessage($"\nTotal xrefs bound: {totalBound}");
                            currentDoc.Editor.WriteMessage($"\nTotal objects purged: {totalPurged}");
                            currentDoc.Editor.WriteMessage($"\nOutput folder: {outputFolder}");
                        }
                    }
                    finally
                    {
                        // Close progress dialog
                        if (progressDialog != null && dialogThread != null)
                        {
                            try
                            {
                                progressDialog.Complete();
                                System.Threading.Thread.Sleep(1000); // Show "Complete!" for 1 second
                                progressDialog.Invoke(new Action(() => progressDialog.Close()));
                                dialogThread.Join(1000);
                            }
                            catch
                            {
                                // Silently handle dialog cleanup errors
                            }
                        }
                    }
                }
                else
                {
                    BindAndPurgeDiagnostics.Log("User cancelled or no documents selected");
                }
            }
            catch (System.Exception ex)
            {
                BindAndPurgeDiagnostics.LogException("ExecuteApplicationScope", ex);
                // Use current active document's editor for error message
                Document currentDoc = docs.MdiActiveDocument;
                if (currentDoc != null)
                {
                    currentDoc.Editor.WriteMessage($"\nError in bind-xrefs-and-purge-documents-in-session command: {ex.Message}");
                }
            }
            finally
            {
                BindAndPurgeDiagnostics.Log("=== ExecuteApplicationScope END ===");
            }
        }

        private static int CountXrefs(Document doc)
        {
            int count = 0;
            try
            {
                Database db = doc.Database;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        foreach (ObjectId blockId in blockTable)
                        {
                            BlockTableRecord btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                            if (btr.IsFromExternalReference)
                            {
                                count++;
                            }
                        }
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                BindAndPurgeDiagnostics.Log($"ERROR counting xrefs: {ex.Message}");
            }
            return count;
        }

        private static int BindXrefsInDocument(Document doc, Editor ed, BindAndPurgeProgressDialog progressDialog, int docNum, int totalDocs, string docName)
        {
            BindAndPurgeDiagnostics.Log($"      >> BindXrefsInDocument START for {doc.Name}");
            Database db = doc.Database;
            int boundCount = 0;
            int attemptedCount = 0;
            int skippedCount = 0;

            try
            {
                using (DocumentLock docLock = doc.LockDocument())
                {
                    BindAndPurgeDiagnostics.Log($"      Building xref graph...");
                    using (var xrefGraph = db.GetHostDwgXrefGraph(true))
                    {
                        BindAndPurgeDiagnostics.Log($"      Xref graph has {xrefGraph.NumNodes} nodes");

                        // First, enumerate all xrefs and log their status
                        BindAndPurgeDiagnostics.Log($"      === Enumerating all xrefs and their status ===");
                        for (int i = 0; i < xrefGraph.NumNodes; i++)
                        {
                            XrefGraphNode node = xrefGraph.GetXrefNode(i);

                            // Skip the root node (the current database itself)
                            if (node.Database != null && node.Database.Equals(db) && !node.IsNested)
                            {
                                BindAndPurgeDiagnostics.Log($"      Node {i}: [ROOT DATABASE] - skipping");
                                continue;
                            }

                            string xrefName = node.Name;
                            BindAndPurgeDiagnostics.Log($"      Node {i}: Name='{xrefName}', IsNested={node.IsNested}");

                            // Get the BlockTableRecord for detailed info
                            try
                            {
                                using (Transaction tr = db.TransactionManager.StartTransaction())
                                {
                                    BlockTableRecord btr = tr.GetObject(node.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                                    if (btr != null)
                                    {
                                        bool isXref = btr.IsFromExternalReference;
                                        bool isResolved = btr.IsResolved;
                                        bool isUnloaded = btr.IsUnloaded;
                                        XrefStatus xrefStatus = btr.XrefStatus;
                                        string pathName = btr.PathName;

                                        BindAndPurgeDiagnostics.Log($"        IsFromExternalReference={isXref}");
                                        BindAndPurgeDiagnostics.Log($"        IsResolved={isResolved}");
                                        BindAndPurgeDiagnostics.Log($"        IsUnloaded={isUnloaded}");
                                        BindAndPurgeDiagnostics.Log($"        XrefStatus={xrefStatus}");
                                        BindAndPurgeDiagnostics.Log($"        PathName='{pathName}'");

                                        // Check if file exists
                                        if (!string.IsNullOrEmpty(pathName))
                                        {
                                            bool fileExists = File.Exists(pathName);
                                            BindAndPurgeDiagnostics.Log($"        File exists: {fileExists}");
                                            if (!fileExists)
                                            {
                                                BindAndPurgeDiagnostics.Log($"        WARNING: Xref file not found at path!");
                                            }
                                        }
                                    }
                                    tr.Commit();
                                }
                            }
                            catch (System.Exception ex)
                            {
                                BindAndPurgeDiagnostics.Log($"        ERROR getting xref details: {ex.Message}");
                            }
                        }

                        // Now iterate through all nodes to bind them
                        BindAndPurgeDiagnostics.Log($"      === Starting binding process ===");
                        for (int i = 0; i < xrefGraph.NumNodes; i++)
                        {
                            XrefGraphNode node = xrefGraph.GetXrefNode(i);

                            // Skip the root node (the current database itself)
                            if (node.Database != null && node.Database.Equals(db) && !node.IsNested)
                            {
                                continue;
                            }

                            string xrefName = node.Name;
                            BindAndPurgeDiagnostics.Log($"      Processing node {i} for binding: {xrefName}");

                            if (xrefName != null && xrefName.Length > 0)
                            {
                                // Get xref status before attempting to bind
                                bool shouldAttemptBind = false;
                                string skipReason = "";

                                try
                                {
                                    using (Transaction tr = db.TransactionManager.StartTransaction())
                                    {
                                        BlockTableRecord btr = tr.GetObject(node.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                                        if (btr != null)
                                        {
                                            if (!btr.IsFromExternalReference)
                                            {
                                                skipReason = "Not an external reference";
                                            }
                                            else if (btr.IsUnloaded)
                                            {
                                                skipReason = "Xref is unloaded";
                                            }
                                            else if (!btr.IsResolved)
                                            {
                                                skipReason = "Xref is not resolved (file may be missing or inaccessible)";
                                            }
                                            else if (btr.XrefStatus != XrefStatus.Resolved)
                                            {
                                                skipReason = $"Xref status is {btr.XrefStatus} (expected Resolved)";
                                            }
                                            else
                                            {
                                                shouldAttemptBind = true;
                                            }
                                        }
                                        tr.Commit();
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    BindAndPurgeDiagnostics.Log($"      ERROR checking xref status for '{xrefName}': {ex.Message}");
                                    skipReason = $"Error checking status: {ex.Message}";
                                }

                                if (!shouldAttemptBind)
                                {
                                    BindAndPurgeDiagnostics.Log($"      SKIPPING xref '{xrefName}': {skipReason}");
                                    ed.WriteMessage($"\n      Skipping xref '{xrefName}': {skipReason}");
                                    skippedCount++;
                                    continue;
                                }

                                // Attempt to bind
                                try
                                {
                                    attemptedCount++;
                                    BindAndPurgeDiagnostics.Log($"      ATTEMPTING to bind xref: {xrefName}");
                                    ed.WriteMessage($"\n      Binding xref: {xrefName}");

                                    progressDialog.UpdateProgress(docNum, totalDocs, docName, "Step 1/3: Binding xrefs...",
                                        $"Binding {xrefName}... ({attemptedCount})");

                                    // Bind the xref (insert type) - measure time
                                    var startTime = DateTime.Now;
                                    db.BindXrefs(new ObjectIdCollection(new[] { node.BlockTableRecordId }), false);
                                    var elapsed = (DateTime.Now - startTime).TotalSeconds;

                                    boundCount++;
                                    BindAndPurgeDiagnostics.Log($"      SUCCESS binding xref: {xrefName} in {elapsed:F2}s");

                                    // Log warning if binding took unusually long
                                    if (elapsed > 5.0)
                                    {
                                        BindAndPurgeDiagnostics.Log($"      WARNING: Xref '{xrefName}' took {elapsed:F2}s to bind (unusually long)");
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    BindAndPurgeDiagnostics.LogException($"Binding xref '{xrefName}'", ex);
                                    ed.WriteMessage($"\n      ERROR: Could not bind xref '{xrefName}': {ex.Message}");
                                }
                            }
                        }
                    }
                }

                BindAndPurgeDiagnostics.Log($"      Binding summary: Attempted={attemptedCount}, Succeeded={boundCount}, Skipped={skippedCount}");
                ed.WriteMessage($"\n    Binding complete: {boundCount} bound, {skippedCount} skipped");
            }
            catch (System.Exception ex)
            {
                BindAndPurgeDiagnostics.LogException("BindXrefsInDocument", ex);
                ed.WriteMessage($"\nError binding xrefs: {ex.Message}");
            }

            BindAndPurgeDiagnostics.Log($"      << BindXrefsInDocument END, returning {boundCount}");
            return boundCount;
        }

        private static int PurgeDocument(Document doc, BindAndPurgeProgressDialog progressDialog, int docNum, int totalDocs, string docName)
        {
            BindAndPurgeDiagnostics.Log($"      >> PurgeDocument START for {doc.Name}");
            int totalPurged = 0;
            Database db = doc.Database;
            int maxIterations = 10; // Purge up to 10 times to catch nested references
            int iteration = 0;

            BindAndPurgeDiagnostics.Log($"      Starting purge iterations (max {maxIterations})");

            while (iteration < maxIterations)
            {
                iteration++;
                BindAndPurgeDiagnostics.Log($"      === PURGE PASS {iteration} START ===");
                int purgedThisPass = 0;

                BindAndPurgeDiagnostics.Log($"      Locking document for pass {iteration}...");
                using (DocumentLock docLock = doc.LockDocument())
                {
                    // Build a complete collection of all objects to purge
                    ObjectIdCollection allObjectsToPurge = new ObjectIdCollection();

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BindAndPurgeDiagnostics.Log($"      Transaction started for pass {iteration}");
                        progressDialog.UpdateProgress(docNum, totalDocs, docName, "Step 2/3: Purging...",
                            $"Purge pass {iteration}/{maxIterations} - collecting objects...");

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

                        BindAndPurgeDiagnostics.Log($"      Collected {allObjectsToPurge.Count} total objects");

                        progressDialog.UpdateProgress(docNum, totalDocs, docName, "Step 2/3: Purging...",
                            $"Purge pass {iteration}/{maxIterations} - purging {allObjectsToPurge.Count} objects...");

                        // Call db.Purge once with all objects
                        BindAndPurgeDiagnostics.Log($"      Calling db.Purge on all collected objects...");
                        db.Purge(allObjectsToPurge);
                        BindAndPurgeDiagnostics.Log($"      db.Purge returned {allObjectsToPurge.Count} purgeable objects");

                        // Erase all purgeable objects
                        if (allObjectsToPurge.Count > 0)
                        {
                            BindAndPurgeDiagnostics.Log($"      Erasing {allObjectsToPurge.Count} purgeable objects...");
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
                                    BindAndPurgeDiagnostics.Log($"        Failed to erase object {objId}: {ex.Message}");
                                }
                            }
                            BindAndPurgeDiagnostics.Log($"      Successfully erased {purgedThisPass} objects");
                        }

                        tr.Commit();
                    }
                }

                totalPurged += purgedThisPass;
                BindAndPurgeDiagnostics.Log($"      === PURGE PASS {iteration} END: Purged {purgedThisPass} objects ===");
                doc.Editor.WriteMessage($"\n      Purge pass {iteration}: {purgedThisPass} object(s)");

                // If nothing was purged, we're done
                if (purgedThisPass == 0)
                {
                    BindAndPurgeDiagnostics.Log($"      No objects purged in pass {iteration}, stopping iterations");
                    break;
                }
            }

            BindAndPurgeDiagnostics.Log($"      Completed {iteration} purge passes, total purged: {totalPurged}");
            BindAndPurgeDiagnostics.Log($"      << PurgeDocument END, returning {totalPurged}");
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
                BindAndPurgeDiagnostics.Log($"      Error collecting from symbol table: {ex.Message}");
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
                BindAndPurgeDiagnostics.Log($"      Error collecting from dictionary {dictionaryKey}: {ex.Message}");
            }
        }

        private static string PromptForFolder(Editor ed)
        {
            string selectedFolder = null;

            // Use STA thread for folder browser dialog
            var thread = new System.Threading.Thread(() =>
            {
                using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.ShowNewFolderButton = true;

                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        selectedFolder = folderDialog.SelectedPath;
                    }
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();

            return selectedFolder;
        }
    }
}
