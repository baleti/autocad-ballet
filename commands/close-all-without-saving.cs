using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.CloseAllWithoutSavingCommand))]

namespace AutoCADBallet
{
    public class CloseAllWithoutSavingCommand
    {
        [CommandMethod("close-all-without-saving", CommandFlags.Session)]
        public void CloseAllWithoutSaving()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            var ed = activeDoc.Editor;

            // Collect all documents to close (need to do this first since collection will change)
            var documentsToClose = new List<Document>();
            foreach (Document doc in docs)
            {
                documentsToClose.Add(doc);
            }

            if (documentsToClose.Count == 0)
            {
                ed.WriteMessage("\nNo documents are currently open.\n");
                return;
            }

            ed.WriteMessage($"\nClosing {documentsToClose.Count} document(s) without saving...\n");

            int successCount = 0;
            int failureCount = 0;

            // Sort documents: close non-active documents first, then the active one
            var nonActiveDocuments = documentsToClose.Where(doc => doc != activeDoc).ToList();
            var activeDocuments = documentsToClose.Where(doc => doc == activeDoc).ToList();

            // Process non-active documents first
            foreach (Document doc in nonActiveDocuments)
            {
                try
                {
                    string docName = System.IO.Path.GetFileNameWithoutExtension(doc.Name);
                    ed.WriteMessage($"Closing: {docName}\n");

                    // Close and discard without prompting to save
                    doc.CloseAndDiscard();
                    successCount++;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"Failed to close document: {ex.Message}\n");
                    failureCount++;
                }
            }

            // Process active document last, if any remain
            foreach (Document doc in activeDocuments)
            {
                try
                {
                    string docName = System.IO.Path.GetFileNameWithoutExtension(doc.Name);
                    ed.WriteMessage($"Closing active document: {docName}\n");

                    // For the active document, we need to be more careful
                    // Try to switch to another document first if possible, or just close it
                    if (docs.Count > 1)
                    {
                        // Find another document to make active
                        Document anotherDoc = null;
                        foreach (Document otherDoc in docs)
                        {
                            if (otherDoc != doc)
                            {
                                anotherDoc = otherDoc;
                                break;
                            }
                        }

                        if (anotherDoc != null)
                        {
                            docs.MdiActiveDocument = anotherDoc;
                            // Allow some time for the document switch
                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Now close the previously active document
                        doc.CloseAndDiscard();
                        successCount++;
                    }
                    else
                    {
                        // This is the last document - context becomes invalid during close
                        ed.WriteMessage("This is the last document. Run the command again to close it.\n");
                        failureCount++;
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"Failed to close active document: {ex.Message}\n");
                    failureCount++;
                }
            }

            // Summary (only show if there's still an active document to write to)
            try
            {
                var currentActive = docs.MdiActiveDocument;
                if (currentActive?.Editor != null)
                {
                    currentActive.Editor.WriteMessage($"\n=== Close Summary ===\n");
                    currentActive.Editor.WriteMessage($"Successfully closed: {successCount} document(s)\n");
                    if (failureCount > 0)
                    {
                        currentActive.Editor.WriteMessage($"Failed to close: {failureCount} document(s)\n");
                    }
                }
            }
            catch
            {
                // If no documents remain, we can't write the summary - that's expected
            }
        }
    }
}