using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.CloseDocumentsCommand))]

namespace AutoCADBallet
{
    public class CloseDocumentsCommand
    {
        [CommandMethod("close-documents", CommandFlags.Session)]
        public void CloseDocuments()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

            // Use application-level messaging in case active document gets closed
            void WriteAppMessage(string message)
            {
                AcadApp.DocumentManager.MdiActiveDocument?.Editor?.WriteMessage(message);
            }

            var availableDocuments = new List<Dictionary<string, object>>();
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileNameWithoutExtension(doc.Name);

                // Check for unsaved changes using COM Saved property (doesn't require switching documents)
                bool hasUnsavedChanges = false;
                try
                {
                    dynamic acadDoc = doc.GetAcadDocument();
                    // Saved property: true = no unsaved changes, false = has unsaved changes
                    hasUnsavedChanges = !acadDoc.Saved;
                }
                catch (System.Exception)
                {
                    // If we can't check Saved property, assume no unsaved changes
                    hasUnsavedChanges = false;
                }

                availableDocuments.Add(new Dictionary<string, object>
                {
                    ["document"] = docName,
                    ["ReadOnly"] = doc.IsReadOnly ? "read only" : "",
                    ["UnsavedChanges"] = hasUnsavedChanges ? "unsaved" : "",
                    ["IsActive"] = doc == activeDoc,
                    ["Document"] = doc
                });
            }

            if (availableDocuments.Count == 0)
                return;

            availableDocuments = availableDocuments.OrderBy(d => d["document"].ToString()).ToList();

            int selectedIndex = -1;
            selectedIndex = availableDocuments.FindIndex(d => (bool)d["IsActive"]);

            var propertyNames = new List<string> { "document", "ReadOnly", "UnsavedChanges" };
            var initialSelectionIndices = selectedIndex >= 0
                                            ? new List<int> { selectedIndex }
                                            : new List<int>();

            // Enable multi-select to allow closing multiple documents (DataGrid supports multi-select by default)
            List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(
                availableDocuments,
                propertyNames,
                spanAllScreens: false,
                initialSelectionIndices: initialSelectionIndices
            );

            if (chosen == null || chosen.Count == 0)
            {
                return;
            }

            // Collect documents to close
            var documentsToClose = new List<Document>();
            foreach (var choice in chosen)
            {
                Document doc = choice["Document"] as Document;
                if (doc != null)
                {
                    documentsToClose.Add(doc);
                }
            }

            if (documentsToClose.Count == 0)
            {
                return;
            }

            WriteAppMessage($"\nClosing {documentsToClose.Count} document(s) without saving...\n");

            int successCount = 0;
            int failureCount = 0;

            // Close selected documents without prompting to save
            foreach (Document doc in documentsToClose)
            {
                try
                {
                    string docName = Path.GetFileNameWithoutExtension(doc.Name);
                    WriteAppMessage($"Closing: {docName}\n");

                    // Close and discard without prompting to save
                    doc.CloseAndDiscard();
                    successCount++;
                }
                catch (System.Exception ex)
                {
                    WriteAppMessage($"Failed to close {Path.GetFileNameWithoutExtension(doc.Name)}: {ex.Message}\n");
                    failureCount++;
                }
            }

            // Final summary - only show if any documents remain
            try
            {
                var remainingActive = docs.MdiActiveDocument;
                if (remainingActive?.Editor != null)
                {
                    remainingActive.Editor.WriteMessage($"\n=== Close Summary ===\n");
                    remainingActive.Editor.WriteMessage($"Successfully closed: {successCount} document(s)\n");
                    if (failureCount > 0)
                    {
                        remainingActive.Editor.WriteMessage($"Failed to close: {failureCount} document(s)\n");
                    }
                }
            }
            catch
            {
                // If no documents remain, we can't show summary - that's expected and OK
            }
        }
    }
}
