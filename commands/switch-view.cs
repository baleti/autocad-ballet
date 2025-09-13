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

            // Generate session identifier for this AutoCAD process
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            string currentSessionId = $"{processId}_{sessionId}";

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
                                    ["layout"] = layoutName,
                                    ["document"] = Path.GetFileNameWithoutExtension(docName),
                                    ["autocad session"] = currentSessionId,
                                    ["LayoutName"] = layoutName,
                                    ["FullPath"] = docFullPath,
                                    ["TabOrder"] = layoutInfo["TabOrder"],
                                    ["IsActive"] = isCurrentView,
                                    ["DocumentObject"] = doc
                                });

                                viewIndex++;
                            }

                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    // Silently skip documents that can't be read
                }
            }

            if (allViews.Count == 0)
            {
                return;
            }

            // Sort views by document name first, then by tab order (Model space typically has TabOrder=0)
            allViews = allViews.OrderBy(v => v["document"].ToString())
                              .ThenBy(v =>
                              {
                                  if (v["TabOrder"] == null) return int.MaxValue;
                                  if (int.TryParse(v["TabOrder"].ToString(), out int tabOrder))
                                  {
                                      // Model space layout typically has TabOrder=0 and should come first
                                      return tabOrder;
                                  }
                                  return int.MaxValue;
                              })
                              .ToList();

            // Update currentViewIndex after sorting
            currentViewIndex = allViews.FindIndex(v => (bool)v["IsActive"]);


            var propertyNames = new List<string> { "layout", "document", "autocad session" };
            var initialSelectionIndices = currentViewIndex >= 0
                                            ? new List<int> { currentViewIndex }
                                            : new List<int>();


            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allViews, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    var selected = chosen.First();
                    Document targetDoc = selected["DocumentObject"] as Document;
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
                                        // Switch to the target layout with document lock
                                        using (DocumentLock docLock = targetDoc.LockDocument())
                                        {
                                            LayoutManager.Current.CurrentLayout = targetLayout;
                                        }

                                        // Log the layout change
                                        string projectName = Path.GetFileNameWithoutExtension(targetDoc.Name) ?? "UnknownProject";
                                        LayoutChangeTracker.LogLayoutChange(projectName, targetLayout, true);

                                    }
                                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                    {
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
                                // Direct layout switch with document lock
                                using (DocumentLock docLock = targetDoc.LockDocument())
                                {
                                    LayoutManager.Current.CurrentLayout = targetLayout;
                                }

                                string projectName = Path.GetFileNameWithoutExtension(targetDoc.Name) ?? "UnknownProject";
                                LayoutChangeTracker.LogLayoutChange(projectName, targetLayout, true);

                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                            }
                        }
                        return;
                    }
                }
            }
            catch (System.Exception ex)
            {
                // DataGrid failed, command ends silently
            }
        }
    }
}
