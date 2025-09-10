using Autodesk.AutoCAD.ApplicationServices;
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
        [CommandMethod("switch-document")]
        public void SwitchDocument()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logFilePath = Path.Combine(appDataPath, "autocad-ballet", "LogDocumentChanges", "session");

            var documentNames = new List<string>();
            
            if (File.Exists(logFilePath))
            {
                try
                {
                    var documentEntries = File.ReadAllLines(logFilePath)
                                            .Reverse()
                                            .Select(l => l.Trim())
                                            .Where(l => l.Length > 0)
                                            .Distinct()
                                            .ToList();

                    foreach (string entry in documentEntries)
                    {
                        var parts = entry.Split(new[] { ' ' }, 2);
                        if (parts.Length == 2)
                            documentNames.Add(parts[1]);
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Could not read log file: {ex.Message}");
                }
            }
            
            if (documentNames.Count == 0)
            {
                ed.WriteMessage("\nNo document log found. Showing all open documents.");
            }

            var availableDocuments = new List<Dictionary<string, object>>();
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileName(doc.Name);
                if (documentNames.Count == 0 || documentNames.Contains(docName) || documentNames.Contains(doc.Name))
                {
                    availableDocuments.Add(new Dictionary<string, object>
                    {
                        ["DocumentName"] = docName,
                        ["FullPath"] = doc.Name,
                        ["IsActive"] = doc == activeDoc,
                        ["Document"] = doc
                    });
                }
            }

            if (availableDocuments.Count == 0)
            {
                ed.WriteMessage("\nNo matching documents found in this session.");
                return;
            }

            availableDocuments = availableDocuments.OrderBy(d => d["DocumentName"].ToString()).ToList();

            int selectedIndex = -1;
            selectedIndex = availableDocuments.FindIndex(d => (bool)d["IsActive"]);

            var propertyNames = new List<string> { "DocumentName", "FullPath" };
            var initialSelectionIndices = selectedIndex >= 0 
                                            ? new List<int> { selectedIndex }
                                            : new List<int>();

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(availableDocuments, propertyNames, false);

                if (chosen != null && chosen.Count > 0)
                {
                    Document chosenDoc = chosen.First()["Document"] as Document;
                    if (chosenDoc != null)
                    {
                        docs.MdiActiveDocument = chosenDoc;
                        ed.WriteMessage($"\nSwitched to document: {chosenDoc.Name}");
                        return;
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError showing document picker: {ex.Message}");
            }

            ed.WriteMessage("\nAvailable documents:");
            for (int i = 0; i < availableDocuments.Count; i++)
            {
                string marker = (i == selectedIndex) ? " [CURRENT]" : "";
                ed.WriteMessage($"\n{i + 1}: {availableDocuments[i]["DocumentName"]}{marker}");
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nSelect document number: ");
            pio.AllowNegative = false;
            pio.AllowZero = false;
            pio.LowerLimit = 1;
            pio.UpperLimit = availableDocuments.Count;
            PromptIntegerResult pir = ed.GetInteger(pio);

            if (pir.Status == PromptStatus.OK)
            {
                Document selectedDoc = availableDocuments[pir.Value - 1]["Document"] as Document;
                if (selectedDoc != null)
                {
                    docs.MdiActiveDocument = selectedDoc;
                    ed.WriteMessage($"\nSwitched to document: {selectedDoc.Name}");
                }
            }
        }
    }
}