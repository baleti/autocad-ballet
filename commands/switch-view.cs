using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SwitchViewCommand))]

namespace AutoCADBallet
{
    public class SwitchViewCommand
    {
        [CommandMethod("switch-view")]
        public void SwitchView()
        {
            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null) return;

            Editor ed = activeDoc.Editor;
            string currentLayoutName = LayoutManager.Current.CurrentLayout;

            var allViews = new List<Dictionary<string, object>>();
            int currentViewIndex = -1;
            int viewIndex = 0;

            // Iterate through all open documents
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileName(doc.Name);
                string docFullPath = doc.Name;
                bool isActiveDoc = (doc == activeDoc);

                // Get layouts from this document
                // Note: We need to be careful accessing non-active document databases
                Database db = doc.Database;

                try
                {
                    // For non-active documents, we might need to lock the document
                    using (DocumentLock docLock = doc.LockDocument())
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                            var layoutsInDoc = new List<Dictionary<string, object>>();

                            foreach (DictionaryEntry entry in layoutDict)
                            {
                                Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                                if (layout != null)
                                {
                                    layoutsInDoc.Add(new Dictionary<string, object>
                                    {
                                        ["LayoutName"] = layout.LayoutName,
                                        ["TabOrder"] = layout.TabOrder,
                                        ["LayoutObject"] = layout
                                    });
                                }
                            }

                            // Sort layouts by tab order
                            layoutsInDoc = layoutsInDoc.OrderBy(l =>
                            {
                                if (l["TabOrder"] == null) return int.MaxValue;
                                if (int.TryParse(l["TabOrder"].ToString(), out int tabOrder)) return tabOrder;
                                return int.MaxValue;
                            }).ToList();

                            // Add each layout to the combined list
                            foreach (var layoutInfo in layoutsInDoc)
                            {
                                string layoutName = layoutInfo["LayoutName"].ToString();
                                bool isCurrentView = isActiveDoc && layoutName == currentLayoutName;

                                if (isCurrentView)
                                {
                                    currentViewIndex = viewIndex;
                                }

                                allViews.Add(new Dictionary<string, object>
                                {
                                    ["ViewName"] = $"{layoutName} ({docName})",
                                    ["DocumentName"] = docName,
                                    ["LayoutName"] = layoutName,
                                    ["FullPath"] = docFullPath,
                                    ["TabOrder"] = layoutInfo["TabOrder"],
                                    ["IsActive"] = isCurrentView,
                                    ["Document"] = doc
                                });

                                viewIndex++;
                            }

                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Could not read layouts from {docName}: {ex.Message}");
                }
            }

            if (allViews.Count == 0)
            {
                ed.WriteMessage("\nNo layouts found in any open documents.");
                return;
            }

            // Sort views by document name first, then by tab order
            allViews = allViews.OrderBy(v => v["DocumentName"].ToString())
                              .ThenBy(v =>
                              {
                                  if (v["TabOrder"] == null) return int.MaxValue;
                                  if (int.TryParse(v["TabOrder"].ToString(), out int tabOrder)) return tabOrder;
                                  return int.MaxValue;
                              })
                              .ToList();

            // Update currentViewIndex after sorting
            currentViewIndex = allViews.FindIndex(v => (bool)v["IsActive"]);

            var propertyNames = new List<string> { "ViewName", "FullPath" };
            var initialSelectionIndices = currentViewIndex >= 0
                                            ? new List<int> { currentViewIndex }
                                            : new List<int>();

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allViews, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    var selected = chosen.First();
                    Document targetDoc = selected["Document"] as Document;
                    string targetLayout = selected["LayoutName"].ToString();

                    if (targetDoc != null)
                    {
                        // First switch to the document
                        if (docs.MdiActiveDocument != targetDoc)
                        {
                            // Set up event handler for when document activation completes
                            DocumentCollectionEventHandler handler = null;
                            handler = (sender, e) =>
                            {
                                if (e.Document == targetDoc)
                                {
                                    // Unsubscribe from event to avoid memory leaks
                                    docs.DocumentActivated -= handler;

                                    try
                                    {
                                        // Use the activated document's editor, not the original one
                                        Editor targetEd = targetDoc.Editor;

                                        // Now safely switch to the layout
                                        LayoutManager.Current.CurrentLayout = targetLayout;

                                        // Log the layout change
                                        string projectName = Path.GetFileNameWithoutExtension(targetDoc.Name) ?? "UnknownProject";
                                        LayoutLogger.LogLayoutChange(projectName, targetLayout);

                                        targetEd.WriteMessage($"\nSwitched to layout: {targetLayout}");
                                    }
                                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                    {
                                        // Use the activated document's editor for error messages too
                                        targetDoc.Editor.WriteMessage($"\nFailed to switch to layout '{targetLayout}': {ex.Message}");
                                    }
                                }
                            };

                            docs.DocumentActivated += handler;
                            docs.MdiActiveDocument = targetDoc;
                            ed.WriteMessage($"\nSwitching to document: {targetDoc.Name}");
                        }
                        else
                        {
                            // Already in the target document, just switch layout
                            try
                            {
                                LayoutManager.Current.CurrentLayout = targetLayout;

                                string projectName = Path.GetFileNameWithoutExtension(targetDoc.Name) ?? "UnknownProject";
                                LayoutLogger.LogLayoutChange(projectName, targetLayout);

                                ed.WriteMessage($"\nSwitched to layout: {targetLayout}");
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to switch to layout '{targetLayout}': {ex.Message}");
                            }
                        }
                        return;
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError showing view picker: {ex.Message}");
            }

            // Fallback to text-based selection
            ed.WriteMessage("\nAvailable views (Layout - Document):");
            for (int i = 0; i < allViews.Count; i++)
            {
                string marker = (i == currentViewIndex) ? " [CURRENT]" : "";
                ed.WriteMessage($"\n{i + 1}: {allViews[i]["ViewName"]}{marker}");
            }

            PromptIntegerOptions pio = new PromptIntegerOptions("\nSelect view number: ");
            pio.AllowNegative = false;
            pio.AllowZero = false;
            pio.LowerLimit = 1;
            pio.UpperLimit = allViews.Count;
            PromptIntegerResult pir = ed.GetInteger(pio);

            if (pir.Status == PromptStatus.OK)
            {
                var selected = allViews[pir.Value - 1];
                Document targetDoc = selected["Document"] as Document;
                string targetLayout = selected["LayoutName"].ToString();

                if (targetDoc != null)
                {
                    // First switch to the document
                    if (docs.MdiActiveDocument != targetDoc)
                    {
                        // Set up event handler for when document activation completes
                        DocumentCollectionEventHandler handler = null;
                        handler = (sender, e) =>
                        {
                            if (e.Document == targetDoc)
                            {
                                // Unsubscribe from event to avoid memory leaks
                                docs.DocumentActivated -= handler;

                                try
                                {
                                    // Use the activated document's editor
                                    Editor targetEd = targetDoc.Editor;

                                    // Verify layout exists before trying to switch
                                    // Now safely switch to the layout
                                    LayoutManager.Current.CurrentLayout = targetLayout;

                                    // Log the layout change
                                    string projectName = Path.GetFileNameWithoutExtension(targetDoc.Name) ?? "UnknownProject";
                                    LayoutLogger.LogLayoutChange(projectName, targetLayout);

                                    targetEd.WriteMessage($"\nSwitched to: {targetLayout} in {Path.GetFileName(targetDoc.Name)}");
                                }
                                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                {
                                    targetDoc.Editor.WriteMessage($"\nFailed to switch to layout '{targetLayout}': {ex.Message}");
                                }
                            }
                        };

                        docs.DocumentActivated += handler;
                        docs.MdiActiveDocument = targetDoc;
                    }
                    else
                    {
                        // Already in the target document, just switch layout
                        try
                        {
                            LayoutManager.Current.CurrentLayout = targetLayout;

                            string projectName = Path.GetFileNameWithoutExtension(targetDoc.Name) ?? "UnknownProject";
                            LayoutLogger.LogLayoutChange(projectName, targetLayout);

                            ed.WriteMessage($"\nSwitched to: {targetLayout} in {Path.GetFileName(targetDoc.Name)}");
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            ed.WriteMessage($"\nFailed to switch to layout '{targetLayout}': {ex.Message}");
                        }
                    }
                }
            }
        }
    }
}
