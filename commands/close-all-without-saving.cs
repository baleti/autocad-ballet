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

            // Use application-level messaging instead of document-bound Editor
            void WriteAppMessage(string message)
            {
                AcadApp.DocumentManager.MdiActiveDocument?.Editor?.WriteMessage(message);
            }

            // Collect all documents to close (snapshot since collection will change)
            var documentsToClose = new List<Document>();
            foreach (Document doc in docs)
            {
                documentsToClose.Add(doc);
            }

            if (documentsToClose.Count == 0)
            {
                WriteAppMessage("\nNo documents are currently open.\n");
                return;
            }

            WriteAppMessage($"\nClosing {documentsToClose.Count} document(s) without saving...\n");

            int successCount = 0;
            int failureCount = 0;

            // Close all documents - no special handling needed since we're using application context
            foreach (Document doc in documentsToClose)
            {
                try
                {
                    string docName = System.IO.Path.GetFileNameWithoutExtension(doc.Name);
                    WriteAppMessage($"Closing: {docName}\n");

                    // Close and discard without prompting to save
                    doc.CloseAndDiscard();
                    successCount++;
                }
                catch (System.Exception ex)
                {
                    WriteAppMessage($"Failed to close {System.IO.Path.GetFileNameWithoutExtension(doc.Name)}: {ex.Message}\n");
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