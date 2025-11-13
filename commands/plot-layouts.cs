using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class PlotDiagnostics
    {
        // Logging disabled - no diagnostic files will be created
        public static void Log(string message)
        {
            // No-op: logging disabled
        }

        public static void LogException(string context, Exception ex)
        {
            // No-op: logging disabled
        }
    }

    public static class PlotLayouts
    {
        public static void ExecuteDocumentScope(Editor ed)
        {
            Document activeDoc = AcadApp.DocumentManager.MdiActiveDocument;
            if (activeDoc == null) return;

            // Disable background plotting to avoid delays
            object oldBackgroundPlot = AcadApp.GetSystemVariable("BACKGROUNDPLOT");
            AcadApp.SetSystemVariable("BACKGROUNDPLOT", 0);

            Database db = activeDoc.Database;
            string currentLayoutName = LayoutManager.Current.CurrentLayout;

            var allLayouts = new List<Dictionary<string, object>>();
            int currentViewIndex = -1;

            try
            {
                using (DocumentLock docLock = activeDoc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                        int viewIndex = 0;
                        foreach (DictionaryEntry entry in layoutDict)
                        {
                            Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                            if (layout != null && !layout.ModelType)
                            {
                                bool isCurrentView = layout.LayoutName == currentLayoutName;

                                if (isCurrentView)
                                {
                                    currentViewIndex = viewIndex;
                                }

                                allLayouts.Add(new Dictionary<string, object>
                                {
                                    ["layout"] = layout.LayoutName,
                                    ["LayoutName"] = layout.LayoutName,
                                    ["TabOrder"] = layout.TabOrder,
                                    ["IsActive"] = isCurrentView,
                                    ["ObjectId"] = (ObjectId)entry.Value,
                                    ["Handle"] = layout.Handle.ToString(),
                                    ["DocumentObject"] = activeDoc,
                                    ["DocumentPath"] = activeDoc.Name
                                });

                                viewIndex++;
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError reading layouts: {ex.Message}");
                return;
            }

            if (allLayouts.Count == 0)
            {
                ed.WriteMessage("\nNo paper space layouts found to plot.");
                return;
            }

            // Sort by tab order
            allLayouts = allLayouts.OrderBy(v =>
            {
                if (v["TabOrder"] == null) return int.MaxValue;
                if (int.TryParse(v["TabOrder"].ToString(), out int tabOrder))
                {
                    return tabOrder;
                }
                return int.MaxValue;
            }).ToList();

            // Update currentViewIndex after sorting
            currentViewIndex = allLayouts.FindIndex(v => (bool)v["IsActive"]);

            var propertyNames = new List<string> { "layout" };
            var initialSelectionIndices = new List<int>();
            if (currentViewIndex >= 0)
            {
                initialSelectionIndices.Add(currentViewIndex);
            }

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allLayouts, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    // Prompt for output folder
                    string outputFolder = PromptForFolder(ed);
                    if (string.IsNullOrEmpty(outputFolder))
                    {
                        ed.WriteMessage("\nPlot operation cancelled - no folder selected.");
                        return;
                    }

                    int totalPlotted = PlotLayoutsToFolder(chosen, outputFolder, activeDoc, db, ed);

                    ed.WriteMessage($"\n\nTotal: {totalPlotted} layout(s) plotted successfully to '{outputFolder}'.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in plot-layouts-in-document command: {ex.Message}");
            }
            finally
            {
                // Restore original BACKGROUNDPLOT setting
                AcadApp.SetSystemVariable("BACKGROUNDPLOT", oldBackgroundPlot);
            }
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            PlotDiagnostics.Log("=== ExecuteApplicationScope START ===");

            DocumentCollection docs = AcadApp.DocumentManager;
            Document activeDoc = docs.MdiActiveDocument;
            if (activeDoc == null)
            {
                PlotDiagnostics.Log("ERROR: No active document");
                return;
            }

            PlotDiagnostics.Log($"Active document: {activeDoc.Name}");

            // Disable background plotting to avoid 35+ second delays between plots
            object oldBackgroundPlot = AcadApp.GetSystemVariable("BACKGROUNDPLOT");
            PlotDiagnostics.Log($"Original BACKGROUNDPLOT value: {oldBackgroundPlot}");
            AcadApp.SetSystemVariable("BACKGROUNDPLOT", 0);
            PlotDiagnostics.Log("BACKGROUNDPLOT set to 0 (foreground plotting)");

            string currentLayoutName = LayoutManager.Current.CurrentLayout;
            PlotDiagnostics.Log($"Current layout: {currentLayoutName}");

            // Generate session identifier for this AutoCAD process
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            string currentSessionId = $"{processId}_{sessionId}";

            var allLayouts = new List<Dictionary<string, object>>();
            int currentViewIndex = -1;
            int viewIndex = 0;

            // Iterate through all open documents
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileName(doc.Name);
                string docFullPath = doc.Name;
                bool isActiveDoc = (doc == activeDoc);

                Database db = doc.Database;

                try
                {
                    using (DocumentLock docLock = doc.LockDocument())
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                            var layoutsInDoc = new List<Dictionary<string, object>>();

                            foreach (DictionaryEntry entry in layoutDict)
                            {
                                Layout layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                                if (layout != null && !layout.ModelType)
                                {
                                    layoutsInDoc.Add(new Dictionary<string, object>
                                    {
                                        ["LayoutName"] = layout.LayoutName,
                                        ["TabOrder"] = layout.TabOrder,
                                        ["ObjectId"] = (ObjectId)entry.Value,
                                        ["Handle"] = layout.Handle.ToString()
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

                                allLayouts.Add(new Dictionary<string, object>
                                {
                                    ["layout"] = layoutName,
                                    ["document"] = Path.GetFileNameWithoutExtension(docName),
                                    ["autocad session"] = currentSessionId,
                                    ["LayoutName"] = layoutName,
                                    ["FullPath"] = docFullPath,
                                    ["TabOrder"] = layoutInfo["TabOrder"],
                                    ["IsActive"] = isCurrentView,
                                    ["DocumentObject"] = doc,
                                    ["ObjectId"] = layoutInfo["ObjectId"],
                                    ["DocumentPath"] = docFullPath,
                                    ["Handle"] = layoutInfo["Handle"]
                                });

                                viewIndex++;
                            }

                            tr.Commit();
                        }
                    }
                }
                catch (System.Exception)
                {
                    // Silently skip documents that can't be read
                }
            }

            if (allLayouts.Count == 0)
            {
                ed.WriteMessage("\nNo paper space layouts found to plot.");
                return;
            }

            // Sort views by document name first, then by tab order
            allLayouts = allLayouts.OrderBy(v => v["document"].ToString())
                                  .ThenBy(v =>
                                  {
                                      if (v["TabOrder"] == null) return int.MaxValue;
                                      if (int.TryParse(v["TabOrder"].ToString(), out int tabOrder))
                                      {
                                          return tabOrder;
                                      }
                                      return int.MaxValue;
                                  })
                                  .ToList();

            // Update currentViewIndex after sorting
            currentViewIndex = allLayouts.FindIndex(v => (bool)v["IsActive"]);

            var propertyNames = new List<string> { "layout", "document", "autocad session" };
            var initialSelectionIndices = new List<int>();
            if (currentViewIndex >= 0)
            {
                initialSelectionIndices.Add(currentViewIndex);
            }

            try
            {
                PlotDiagnostics.Log($"Showing DataGrid with {allLayouts.Count} layouts");
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allLayouts, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    PlotDiagnostics.Log($"User selected {chosen.Count} layouts");
                    foreach (var layout in chosen)
                    {
                        PlotDiagnostics.Log($"  - {layout["document"]} :: {layout["layout"]}");
                    }

                    // Prompt for output folder
                    PlotDiagnostics.Log("Prompting for output folder");
                    string outputFolder = PromptForFolder(ed);
                    if (string.IsNullOrEmpty(outputFolder))
                    {
                        PlotDiagnostics.Log("No folder selected, operation cancelled");
                        ed.WriteMessage("\nPlot operation cancelled - no folder selected.");
                        return;
                    }
                    PlotDiagnostics.Log($"Output folder selected: {outputFolder}");

                    // Group selected layouts by document
                    var layoutsByDocument = chosen.GroupBy(l => l["DocumentObject"] as Document).ToList();
                    PlotDiagnostics.Log($"Layouts grouped into {layoutsByDocument.Count} documents");

                    int totalPlotted = 0;

                    // Process all documents sequentially
                    foreach (var docGroup in layoutsByDocument)
                    {
                        Document targetDoc = docGroup.Key;
                        Database targetDb = targetDoc.Database;
                        string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);

                        PlotDiagnostics.Log($"Processing document: {docName} ({docGroup.Count()} layouts)");

                        // Activate the target document if it's not already active
                        if (docs.MdiActiveDocument != targetDoc)
                        {
                            PlotDiagnostics.Log($"  Need to switch to document '{docName}'");
                            // Use current active document's editor (not the original ed which may be stale)
                            docs.MdiActiveDocument.Editor.WriteMessage($"\nSwitching to document '{docName}'...");

                            try
                            {
                                docs.MdiActiveDocument = targetDoc;
                                PlotDiagnostics.Log($"  Document switch initiated");

                                // Wait for document switch to complete (verify it actually switched)
                                int switchWaitCount = 0;
                                while (docs.MdiActiveDocument != targetDoc && switchWaitCount < 50) // 5 second timeout
                                {
                                    System.Windows.Forms.Application.DoEvents();
                                    System.Threading.Thread.Sleep(100);
                                    switchWaitCount++;
                                }

                                if (docs.MdiActiveDocument != targetDoc)
                                {
                                    PlotDiagnostics.Log($"  ERROR: Document switch FAILED after {switchWaitCount * 100}ms");
                                    PlotDiagnostics.Log($"  Active doc is still: {docs.MdiActiveDocument.Name}");
                                    docs.MdiActiveDocument.Editor.WriteMessage($"\nERROR: Failed to switch to document '{docName}', skipping layouts from this document.");
                                    continue;
                                }

                                PlotDiagnostics.Log($"  Document switch successful after {switchWaitCount * 100}ms");
                                PlotDiagnostics.Log($"  Active doc is now: {docs.MdiActiveDocument.Name}");
                            }
                            catch (Exception ex)
                            {
                                PlotDiagnostics.LogException("Document switch", ex);
                                targetDoc.Editor.WriteMessage($"\nError switching to document '{docName}': {ex.Message}");
                                continue;
                            }
                        }
                        else
                        {
                            PlotDiagnostics.Log($"  Document '{docName}' is already active");
                        }

                        // Plot layouts from this document
                        PlotDiagnostics.Log($"  Calling PlotLayoutsToFolder for {docGroup.Count()} layouts");
                        int plottedInDoc = PlotLayoutsToFolder(docGroup.ToList(), outputFolder, targetDoc, targetDb, targetDoc.Editor);
                        totalPlotted += plottedInDoc;
                        PlotDiagnostics.Log($"  PlotLayoutsToFolder returned: {plottedInDoc} plotted");

                        if (plottedInDoc > 0)
                        {
                            targetDoc.Editor.WriteMessage($"\nSuccessfully plotted {plottedInDoc} layout(s) from '{docName}'.");
                        }
                    }

                    PlotDiagnostics.Log($"=== TOTAL PLOTTED: {totalPlotted} ===");

                    // Use the current active document's editor for final message
                    Document currentDoc = docs.MdiActiveDocument;
                    if (currentDoc != null)
                    {
                        currentDoc.Editor.WriteMessage($"\n\nTotal: {totalPlotted} layout(s) plotted successfully to '{outputFolder}'.");
                    }
                }
            }
            catch (System.Exception ex)
            {
                PlotDiagnostics.LogException("ExecuteApplicationScope", ex);

                // Use current active document's editor for error message
                Document currentDoc = docs.MdiActiveDocument;
                if (currentDoc != null)
                {
                    currentDoc.Editor.WriteMessage($"\nError in plot-layouts-in-session command: {ex.Message}");
                }
            }
            finally
            {
                // Restore original BACKGROUNDPLOT setting
                AcadApp.SetSystemVariable("BACKGROUNDPLOT", oldBackgroundPlot);
                PlotDiagnostics.Log($"BACKGROUNDPLOT restored to: {oldBackgroundPlot}");
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

        private static int PlotLayoutsToFolder(List<Dictionary<string, object>> selectedLayouts,
            string outputFolder, Document targetDoc, Database targetDb, Editor ed)
        {
            PlotDiagnostics.Log($"    PlotLayoutsToFolder START for {selectedLayouts.Count} layouts");
            int plottedCount = 0;

            try
            {
                PlotDiagnostics.Log($"    Locking document: {targetDoc.Name}");
                using (DocumentLock docLock = targetDoc.LockDocument())
                {
                    PlotDiagnostics.Log($"    Document locked, starting transaction");
                    using (Transaction tr = targetDb.TransactionManager.StartTransaction())
                    {
                        DBDictionary layoutDict = tr.GetObject(targetDb.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                        foreach (var selectedLayout in selectedLayouts)
                        {
                            try
                            {
                                string layoutName = selectedLayout["LayoutName"].ToString();
                                ObjectId layoutId = (ObjectId)selectedLayout["ObjectId"];

                                PlotDiagnostics.Log($"    Processing layout: {layoutName}");

                                Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                                if (layout != null)
                                {
                                    // Create output file name: LayoutName.pdf
                                    string fileName = $"{layoutName}.pdf";
                                    string outputPath = Path.Combine(outputFolder, fileName);
                                    PlotDiagnostics.Log($"      Output file: {fileName}");

                                    // Switch to the layout before plotting (required to avoid eLayoutNotCurrent error)
                                    PlotDiagnostics.Log($"      Switching to layout: {layoutName}");
                                    PlotDiagnostics.Log($"      Current layout before: {LayoutManager.Current.CurrentLayout}");
                                    LayoutManager.Current.CurrentLayout = layoutName;
                                    PlotDiagnostics.Log($"      Current layout after: {LayoutManager.Current.CurrentLayout}");

                                    // Wait for plot engine to be ready (with longer timeout for background plotting)
                                    PlotDiagnostics.Log($"      Plot engine state: {PlotFactory.ProcessPlotState}");
                                    int waitCount = 0;
                                    while (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting && waitCount < 600) // 60 seconds max
                                    {
                                        if (waitCount == 0)
                                            PlotDiagnostics.Log($"      Waiting for plot engine to be ready...");
                                        System.Windows.Forms.Application.DoEvents();
                                        System.Threading.Thread.Sleep(100);
                                        waitCount++;

                                        // Log progress every 5 seconds
                                        if (waitCount % 50 == 0)
                                            PlotDiagnostics.Log($"      Still waiting... {waitCount * 100}ms elapsed, state: {PlotFactory.ProcessPlotState}");
                                    }
                                    if (waitCount > 0)
                                        PlotDiagnostics.Log($"      Waited {waitCount * 100}ms for plot engine");
                                    PlotDiagnostics.Log($"      Plot engine state after wait: {PlotFactory.ProcessPlotState}");

                                    // Plot the layout to PDF using its existing settings
                                    PlotDiagnostics.Log($"      Calling PlotLayoutToPdf");
                                    bool success = PlotLayoutToPdf(layout, outputPath, targetDoc, targetDb, ed);
                                    PlotDiagnostics.Log($"      PlotLayoutToPdf returned: {success}");

                                    if (success)
                                    {
                                        plottedCount++;
                                        ed.WriteMessage($"\nPlotted '{layoutName}' to '{fileName}'");

                                        // Wait a bit after successful plot to ensure plot engine is ready for next plot
                                        System.Threading.Thread.Sleep(500);
                                        PlotDiagnostics.Log($"      Plot successful, plotted count: {plottedCount}");
                                    }
                                    else
                                    {
                                        ed.WriteMessage($"\nFailed to plot '{layoutName}'");
                                        PlotDiagnostics.Log($"      Plot FAILED for {layoutName}");
                                    }
                                }
                                else
                                {
                                    PlotDiagnostics.Log($"      Layout object is null for {layoutName}");
                                }
                            }
                            catch (System.Exception ex)
                            {
                                PlotDiagnostics.LogException($"Plotting layout {selectedLayout["LayoutName"]}", ex);
                                ed.WriteMessage($"\nError plotting layout '{selectedLayout["LayoutName"]}': {ex.Message}");
                            }
                        }

                        PlotDiagnostics.Log($"    Committing transaction");
                        tr.Commit();
                        PlotDiagnostics.Log($"    Transaction committed");
                    }
                }
            }
            catch (System.Exception ex)
            {
                string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                PlotDiagnostics.LogException($"PlotLayoutsToFolder for {docName}", ex);
                ed.WriteMessage($"\nError processing document '{docName}': {ex.Message}");
            }

            PlotDiagnostics.Log($"    PlotLayoutsToFolder END, returning {plottedCount}");
            return plottedCount;
        }

        private static bool PlotLayoutToPdf(Layout layout, string outputPath, Document doc, Database db, Editor ed)
        {
            PlotDiagnostics.Log($"        PlotLayoutToPdf START for {layout.LayoutName}");
            try
            {
                // Get the PlotSettings from the layout
                PlotDiagnostics.Log($"        Creating PlotSettings, ModelType={layout.ModelType}");
                using (PlotSettings plotSettings = new PlotSettings(layout.ModelType))
                {
                    plotSettings.CopyFrom(layout);
                    PlotDiagnostics.Log($"        PlotSettings copied from layout");

                    // Set the output device to PDF
                    PlotSettingsValidator plotValidator = PlotSettingsValidator.Current;

                    // Try to find a PDF plotter - common names include "DWG To PDF.pc3", "AutoCAD PDF (General Documentation).pc3"
                    string pdfPlotter = FindPdfPlotter();
                    PlotDiagnostics.Log($"        PDF Plotter: {pdfPlotter ?? "null"}");
                    if (!string.IsNullOrEmpty(pdfPlotter))
                    {
                        plotValidator.SetPlotConfigurationName(plotSettings, pdfPlotter, null);
                        PlotDiagnostics.Log($"        Plot configuration set to {pdfPlotter}");
                    }
                    else
                    {
                        // If no PDF plotter found, use the layout's configured plotter
                        ed.WriteMessage($"\nWarning: No PDF plotter found, using layout's configured plotter for '{layout.LayoutName}'");
                        PlotDiagnostics.Log($"        Using layout's configured plotter");
                    }

                    // Set the plot type to layout
                    plotValidator.SetPlotType(plotSettings, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
                    PlotDiagnostics.Log($"        Plot type set to Layout");

                    // Use the layout's configured plot settings (paper size, plot area, etc.)
                    // These are already set in the plotSettings copied from layout

                    // Validate the plot settings
                    plotValidator.RefreshLists(plotSettings);
                    PlotDiagnostics.Log($"        Plot settings validated");

                    // Create PlotInfo
                    PlotInfo plotInfo = new PlotInfo();
                    plotInfo.Layout = layout.ObjectId;
                    PlotDiagnostics.Log($"        PlotInfo created");

                    PlotInfoValidator plotInfoValidator = new PlotInfoValidator();
                    plotInfoValidator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                    plotInfoValidator.Validate(plotInfo);
                    PlotDiagnostics.Log($"        PlotInfo validated");

                    // Check if PlotEngine is already busy
                    PlotDiagnostics.Log($"        PlotEngine state: {PlotFactory.ProcessPlotState}");
                    if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                    {
                        PlotDiagnostics.Log($"        Creating PlotEngine");
                        using (PlotEngine plotEngine = PlotFactory.CreatePublishEngine())
                        {
                            // Create PlotProgressDialog
                            PlotProgressDialog plotProgress = new PlotProgressDialog(false, 1, true);

                            using (plotProgress)
                            {
                                plotProgress.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Plotting to PDF");
                                plotProgress.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                                plotProgress.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                                plotProgress.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                                plotProgress.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");

                                plotProgress.LowerPlotProgressRange = 0;
                                plotProgress.UpperPlotProgressRange = 100;
                                plotProgress.PlotProgressPos = 0;

                                // Start the plot
                                plotProgress.OnBeginPlot();
                                plotProgress.IsVisible = true;

                                plotEngine.BeginPlot(plotProgress, null);

                                // Start document
                                plotEngine.BeginDocument(plotInfo, doc.Name, null, 1, true, outputPath);

                                // Plot page info
                                PlotPageInfo pageInfo = new PlotPageInfo();
                                plotEngine.BeginPage(pageInfo, plotInfo, true, null);

                                plotProgress.OnBeginSheet();
                                plotProgress.LowerSheetProgressRange = 0;
                                plotProgress.UpperSheetProgressRange = 100;
                                plotProgress.SheetProgressPos = 0;

                                // Actually plot the page
                                PlotInfo plotInfoToPlot = new PlotInfo();
                                plotInfoToPlot.Layout = layout.ObjectId;
                                plotInfoToPlot.OverrideSettings = plotSettings;
                                plotInfoValidator.Validate(plotInfoToPlot);

                                plotEngine.BeginGenerateGraphics(null);
                                plotEngine.EndGenerateGraphics(null);

                                // End the page
                                plotProgress.SheetProgressPos = 100;
                                plotProgress.OnEndSheet();

                                plotEngine.EndPage(null);

                                // End the document
                                plotEngine.EndDocument(null);

                                // End the plot
                                plotProgress.PlotProgressPos = 100;
                                plotProgress.OnEndPlot();

                                plotEngine.EndPlot(null);
                            }
                        }

                        PlotDiagnostics.Log($"        Plot completed successfully");
                        return true;
                    }
                    else
                    {
                        PlotDiagnostics.Log($"        ERROR: Plot engine is busy (state: {PlotFactory.ProcessPlotState})");
                        ed.WriteMessage($"\nPlot engine is busy, cannot plot '{layout.LayoutName}'");
                        return false;
                    }
                }
            }
            catch (System.Exception ex)
            {
                PlotDiagnostics.LogException($"PlotLayoutToPdf for {layout.LayoutName}", ex);
                ed.WriteMessage($"\nError plotting layout '{layout.LayoutName}': {ex.Message}");
                return false;
            }
        }

        private static string FindPdfPlotter()
        {
            try
            {
                // Return the most common PDF plotter name
                // The PlotSettingsValidator will handle validation if it doesn't exist
                return "DWG To PDF.pc3";
            }
            catch
            {
                return null;
            }
        }
    }
}
