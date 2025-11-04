using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SaveAllForceCommand))]

namespace AutoCADBallet
{
    public class SaveAllForceCommand
    {
        [CommandMethod("save-all-force", CommandFlags.Session)]
        public void SaveAllForce()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;

            int savedCount = 0;
            int skippedCount = 0;

            foreach (Document doc in docs)
            {
                // Skip read-only documents
                if (doc.IsReadOnly)
                {
                    skippedCount++;
                    continue;
                }

                // Skip documents that haven't been saved yet (no file path)
                if (string.IsNullOrEmpty(doc.Name) || doc.Name.StartsWith("Drawing"))
                {
                    skippedCount++;
                    continue;
                }

                // Check if document is in reference edit mode (xref editing)
                bool isInRefEdit = false;
                try
                {
                    dynamic acadDoc = doc.GetAcadDocument();
                    string refEditName = acadDoc.GetVariable("REFEDITNAME");
                    isInRefEdit = !string.IsNullOrEmpty(refEditName);
                }
                catch (System.Exception)
                {
                    // If we can't check REFEDITNAME, assume not in reference edit mode
                    isInRefEdit = false;
                }

                if (isInRefEdit)
                {
                    skippedCount++;
                    continue;
                }

                // Try to save the document using COM interface (same as LISP vla-save)
                try
                {
                    dynamic acadDoc = doc.GetAcadDocument();
                    acadDoc.Save();
                    savedCount++;
                }
                catch (System.Exception)
                {
                    // Silently skip documents that fail to save
                    skippedCount++;
                }
            }

            // Write message to active document's editor if available
            if (activeDoc != null)
            {
                var ed = activeDoc.Editor;
                ed.WriteMessage($"\nSaved {savedCount} document(s), skipped {skippedCount} document(s).\n");
            }
        }
    }
}
