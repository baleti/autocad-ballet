using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SwitchDocumentCommand))]

namespace AutoCADBallet
{
    public class SwitchDocumentCommand
    {
        [CommandMethod("switch-document", CommandFlags.Session)]
        public void SwitchDocument()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

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

            List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(availableDocuments, propertyNames, false, initialSelectionIndices);

            if (chosen != null && chosen.Count > 0)
            {
                Document chosenDoc = chosen.First()["Document"] as Document;
                if (chosenDoc != null && chosenDoc != activeDoc)
                {
                    // Set up event handler to log layout change after document activation
                    DocumentCollectionEventHandler handler = null;
                    handler = (sender, e) =>
                    {
                        if (e.Document == chosenDoc)
                        {
                            // Unsubscribe from event to avoid memory leaks
                            docs.DocumentActivated -= handler;

                            try
                            {
                                // Log the current layout of the activated document
                                string projectName = Path.GetFileNameWithoutExtension(chosenDoc.Name) ?? "UnknownProject";
                                string currentLayout = LayoutManager.Current.CurrentLayout;
                                if (!string.IsNullOrEmpty(currentLayout))
                                {
                                    SwitchViewLogging.LogLayoutChange(projectName, chosenDoc.Name, currentLayout, true);
                                }
                            }
                            catch (System.Exception)
                            {
                                // Silently handle logging errors
                            }
                        }
                    };

                    docs.DocumentActivated += handler;
                    docs.MdiActiveDocument = chosenDoc;
                }
            }
        }
    }
}