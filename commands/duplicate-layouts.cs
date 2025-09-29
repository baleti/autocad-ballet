using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.DuplicateLayoutsCommand))]

namespace AutoCADBallet
{
    public class DuplicateLayoutsCommand
    {
        [CommandMethod("duplicate-layouts", CommandFlags.Modal)]
        public void DuplicateLayouts()
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

            var allLayouts = new List<Dictionary<string, object>>();
            int currentViewIndex = -1;
            int viewIndex = 0;

            // Iterate through all open documents
            foreach (Document doc in docs)
            {
                string docName = Path.GetFileName(doc.Name);
                string docFullPath = doc.Name;
                bool isActiveDoc = (doc == activeDoc);

                // Get layouts from this document
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
                                    // Exclude model space layouts (Model tab)
                                    if (!layout.ModelType)
                                    {
                                        layoutsInDoc.Add(new Dictionary<string, object>
                                        {
                                            ["LayoutName"] = layout.LayoutName,
                                            ["TabOrder"] = layout.TabOrder,
                                            ["LayoutObject"] = layout,
                                            ["ObjectId"] = (ObjectId)entry.Value,
                                            ["Handle"] = layout.Handle.ToString()
                                        });
                                    }
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
                ed.WriteMessage("\nNo paper space layouts found to duplicate.");
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
            var initialSelectionIndices = currentViewIndex >= 0
                                            ? new List<int> { currentViewIndex }
                                            : new List<int>();

            try
            {
                List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(allLayouts, propertyNames, false, initialSelectionIndices);

                if (chosen != null && chosen.Count > 0)
                {
                    // Group selected layouts by document
                    var layoutsByDocument = chosen.GroupBy(l => l["DocumentObject"] as Document);

                    int totalDuplicated = 0;

                    // Process all documents without switching
                    foreach (var docGroup in layoutsByDocument)
                    {
                        Document targetDoc = docGroup.Key;
                        int duplicatedInDoc = ProcessLayoutsWithCloning(docGroup, targetDoc, ed);
                        totalDuplicated += duplicatedInDoc;
                        
                        string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                        if (duplicatedInDoc > 0)
                        {
                            ed.WriteMessage($"\nSuccessfully duplicated {duplicatedInDoc} layout(s) in '{docName}'.");
                        }
                    }

                    ed.WriteMessage($"\n\nTotal: {totalDuplicated} layout(s) duplicated successfully.");
                    
                    // Request regen for the active document only if layouts were duplicated there
                    if (layoutsByDocument.Any(g => g.Key == activeDoc))
                    {
                        activeDoc.SendStringToExecute("_.REGENALL ", true, false, false);
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in duplicate-layouts command: {ex.Message}");
            }
        }

        private int ProcessLayoutsWithCloning(IGrouping<Document, Dictionary<string, object>> docGroup, 
            Document targetDoc, Editor ed)
        {
            int duplicatedCount = 0;
            Database targetDb = targetDoc.Database;

            try
            {
                using (DocumentLock docLock = targetDoc.LockDocument())
                {
                    using (Transaction tr = targetDb.TransactionManager.StartTransaction())
                    {
                        DBDictionary layoutDict = tr.GetObject(targetDb.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        BlockTable bt = tr.GetObject(targetDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                        // Process selected layouts in reverse tab order to maintain proper positioning
                        var selectedLayouts = docGroup.OrderByDescending(l =>
                        {
                            if (l["TabOrder"] == null) return int.MinValue;
                            if (int.TryParse(l["TabOrder"].ToString(), out int tabOrder)) return tabOrder;
                            return int.MinValue;
                        }).ToList();

                        foreach (var selectedLayout in selectedLayouts)
                        {
                            try
                            {
                                string originalName = selectedLayout["LayoutName"].ToString();
                                ObjectId sourceLayoutId = (ObjectId)selectedLayout["ObjectId"];
                                Document sourceDoc = selectedLayout["DocumentObject"] as Document;

                                // Generate unique name
                                string newName = GenerateUniqueName(originalName + " - Copy", layoutDict);

                                if (sourceDoc == targetDoc)
                                {
                                    // Same document - clone within same transaction
                                    ObjectId newLayoutId = CloneLayoutInSameDocument(
                                        sourceLayoutId, newName, layoutDict, bt, tr);

                                    if (newLayoutId != ObjectId.Null)
                                    {
                                        duplicatedCount++;
                                        ed.WriteMessage($"\nDuplicated layout '{originalName}' as '{newName}'");
                                    }
                                }
                                else
                                {
                                    // Cross-document - need to handle with separate transaction
                                    ObjectId newLayoutId = CloneLayoutCrossDocument(
                                        sourceDoc, sourceLayoutId, targetDoc, newName, tr);

                                    if (newLayoutId != ObjectId.Null)
                                    {
                                        duplicatedCount++;
                                        ed.WriteMessage($"\nDuplicated layout '{originalName}' as '{newName}' from '{Path.GetFileNameWithoutExtension(sourceDoc.Name)}'");
                                    }
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nError duplicating layout '{selectedLayout["LayoutName"]}': {ex.Message}");
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                string docName = Path.GetFileNameWithoutExtension(targetDoc.Name);
                ed.WriteMessage($"\nError processing document '{docName}': {ex.Message}");
            }

            return duplicatedCount;
        }

        private ObjectId CloneLayoutInSameDocument(ObjectId sourceLayoutId, string newName, 
            DBDictionary layoutDict, BlockTable bt, Transaction tr)
        {
            try
            {
                Layout sourceLayout = tr.GetObject(sourceLayoutId, OpenMode.ForRead) as Layout;
                if (sourceLayout == null) return ObjectId.Null;

                BlockTableRecord sourceBtr = tr.GetObject(sourceLayout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                // Create new block table record for the new layout's paper space
                BlockTableRecord newBtr = new BlockTableRecord();
                newBtr.Name = "*Paper_Space" + GetNextPaperSpaceIndex(bt, tr);
                
                ObjectId newBtrId = bt.Add(newBtr);
                tr.AddNewlyCreatedDBObject(newBtr, true);

                // Clone all entities from source BTR to new BTR
                ObjectIdCollection entityIds = new ObjectIdCollection();
                foreach (ObjectId id in sourceBtr)
                {
                    entityIds.Add(id);
                }

                if (entityIds.Count > 0)
                {
                    IdMapping idMap = new IdMapping();
                    sourceLayout.Database.DeepCloneObjects(entityIds, newBtrId, idMap, false);
                }

                // Create new layout
                Layout newLayout = new Layout();
                newLayout.LayoutName = newName;
                newLayout.BlockTableRecordId = newBtrId;

                // Copy properties from source layout using CopyFrom method
                newLayout.CopyFrom(sourceLayout);
                
                // Ensure the layout name and BTR link survive CopyFrom
                newLayout.LayoutName = newName;
                newLayout.BlockTableRecordId = newBtrId;

                // Add to layout dictionary
                ObjectId newLayoutId = layoutDict.SetAt(newName, newLayout);
                tr.AddNewlyCreatedDBObject(newLayout, true);

                // Link the new block table record back to this layout
                newBtr.UpgradeOpen();
                newBtr.LayoutId = newLayoutId;

                // Adjust tab orders
                AdjustTabOrders(layoutDict, sourceLayout.TabOrder, newLayout, tr);

                // Copy viewports separately (they need special handling)
                CopyViewports(sourceBtr, newBtr, tr);

                return newLayoutId;
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Failed to clone layout: {ex.Message}", ex);
            }
        }

        private ObjectId CloneLayoutCrossDocument(Document sourceDoc, ObjectId sourceLayoutId, 
            Document targetDoc, string newName, Transaction targetTr)
        {
            try
            {
                Database sourceDb = sourceDoc.Database;
                Database targetDb = targetDoc.Database;

                using (DocumentLock sourceLock = sourceDoc.LockDocument())
                {
                    using (Transaction sourceTr = sourceDb.TransactionManager.StartTransaction())
                    {
                        Layout sourceLayout = sourceTr.GetObject(sourceLayoutId, OpenMode.ForRead) as Layout;
                        if (sourceLayout == null) return ObjectId.Null;

                        BlockTableRecord sourceBtr = sourceTr.GetObject(sourceLayout.BlockTableRecordId, 
                            OpenMode.ForRead) as BlockTableRecord;

                        // Get target database objects
                        DBDictionary layoutDict = targetTr.GetObject(targetDb.LayoutDictionaryId, 
                            OpenMode.ForWrite) as DBDictionary;
                        BlockTable bt = targetTr.GetObject(targetDb.BlockTableId, OpenMode.ForWrite) as BlockTable;

                        // Create new block table record for the new layout's paper space
                        BlockTableRecord newBtr = new BlockTableRecord();
                        newBtr.Name = "*Paper_Space" + GetNextPaperSpaceIndex(bt, targetTr);
                        
                        ObjectId newBtrId = bt.Add(newBtr);
                        targetTr.AddNewlyCreatedDBObject(newBtr, true);

                        // Clone all entities from source BTR to new BTR (cross-database)
                        ObjectIdCollection entityIds = new ObjectIdCollection();
                        foreach (ObjectId id in sourceBtr)
                        {
                            entityIds.Add(id);
                        }

                        if (entityIds.Count > 0)
                        {
                            IdMapping idMap = new IdMapping();
                            sourceDb.WblockCloneObjects(entityIds, newBtrId, idMap, 
                                DuplicateRecordCloning.Replace, false);
                        }

                        // Create new layout in target database
                        Layout newLayout = new Layout();
                        newLayout.LayoutName = newName;
                        newLayout.BlockTableRecordId = newBtrId;

                        // Copy plot settings using PlotSettingsValidator (cross-database safe)
                        CopyPlotSettings(sourceLayout, newLayout, targetDb);

                        // Add to layout dictionary
                        ObjectId newLayoutId = layoutDict.SetAt(newName, newLayout);
                        targetTr.AddNewlyCreatedDBObject(newLayout, true);

                        // Link the new block table record back to this layout
                        newBtr.UpgradeOpen();
                        newBtr.LayoutId = newLayoutId;

                        // Adjust tab orders
                        AdjustTabOrders(layoutDict, sourceLayout.TabOrder, newLayout, targetTr);

                        sourceTr.Commit();
                        return newLayoutId;
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Failed to clone layout cross-document: {ex.Message}", ex);
            }
        }

        private void CopyPlotSettings(Layout source, Layout target, Database targetDb)
        {
            try
            {
                // Use PlotSettingsValidator to properly copy plot settings
                using (PlotSettingsValidator psv = PlotSettingsValidator.Current)
                {
                    // Copy basic settings that can be directly assigned
                    target.PlotSettingsName = source.PlotSettingsName;
                    target.PrintLineweights = source.PrintLineweights;
                    target.ShowPlotStyles = source.ShowPlotStyles;
                    target.PlotTransparency = source.PlotTransparency;
                    target.PlotPlotStyles = source.PlotPlotStyles;
                    target.DrawViewportsFirst = source.DrawViewportsFirst;
                    target.PlotHidden = source.PlotHidden;
                    target.PlotViewportBorders = source.PlotViewportBorders;
                    target.ScaleLineweights = source.ScaleLineweights;

                    // Try to set the plot configuration name using PlotSettingsValidator
                    try
                    {
                        if (!string.IsNullOrEmpty(source.PlotConfigurationName))
                        {
                            psv.SetPlotConfigurationName(target, source.PlotConfigurationName, 
                                source.CanonicalMediaName);
                        }
                    }
                    catch { /* If device not available, skip */ }

                    // Set plot type and related settings
                    try
                    {
                        psv.SetPlotType(target, source.PlotType);
                        
                        if (source.PlotType == Autodesk.AutoCAD.DatabaseServices.PlotType.Window)
                        {
                            psv.SetPlotWindowArea(target, source.PlotWindowArea);
                        }
                    }
                    catch { /* Use defaults if this fails */ }

                    // Set rotation
                    try
                    {
                        psv.SetPlotRotation(target, source.PlotRotation);
                    }
                    catch { /* Use default if this fails */ }

                    // Set plot origin
                    try
                    {
                        psv.SetPlotOrigin(target, source.PlotOrigin);
                    }
                    catch { /* Use default if this fails */ }

                    // Set plot centering
                    try
                    {
                        psv.SetPlotCentered(target, source.PlotCentered);
                    }
                    catch { /* Use default if this fails */ }

                    // Set scale
                    try
                    {
                        if (source.UseStandardScale)
                        {
                            psv.SetStdScaleType(target, source.StdScaleType);
                        }
                        else
                        {
                            psv.SetCustomPrintScale(target, source.CustomPrintScale);
                        }
                    }
                    catch { /* Use defaults if this fails */ }

                    // Set the current style sheet
                    try
                    {
                        if (!string.IsNullOrEmpty(source.CurrentStyleSheet))
                        {
                            psv.SetCurrentStyleSheet(target, source.CurrentStyleSheet);
                        }
                    }
                    catch { /* Use default if this fails */ }
                }
            }
            catch
            {
                // If PlotSettingsValidator fails, at least the basic properties are copied
            }
        }

        private void CopyViewports(BlockTableRecord sourceBtr, BlockTableRecord targetBtr, Transaction tr)
        {
            try
            {
                // Find viewports in source BTR
                List<Viewport> sourceViewports = new List<Viewport>();
                foreach (ObjectId id in sourceBtr)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent is Viewport vp && vp.Number != 1) // Skip paper space viewport
                    {
                        sourceViewports.Add(vp);
                    }
                }

                // Find corresponding viewports in target BTR
                List<Viewport> targetViewports = new List<Viewport>();
                foreach (ObjectId id in targetBtr)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent is Viewport vp && vp.Number != 1)
                    {
                        targetViewports.Add(vp);
                    }
                }

                // Match and copy viewport properties
                for (int i = 0; i < Math.Min(sourceViewports.Count, targetViewports.Count); i++)
                {
                    Viewport sourceVp = sourceViewports[i];
                    Viewport targetVp = targetViewports[i];
                    
                    targetVp.UpgradeOpen();
                    
                    // Copy viewport properties
                    targetVp.ViewCenter = sourceVp.ViewCenter;
                    targetVp.ViewHeight = sourceVp.ViewHeight;
                    targetVp.ViewTarget = sourceVp.ViewTarget;
                    targetVp.ViewDirection = sourceVp.ViewDirection;
                    targetVp.TwistAngle = sourceVp.TwistAngle;
                    targetVp.Locked = sourceVp.Locked;
                    targetVp.On = sourceVp.On;
                    
                    if (!sourceVp.NonRectClipOn)
                    {
                        targetVp.Width = sourceVp.Width;
                        targetVp.Height = sourceVp.Height;
                        targetVp.CenterPoint = sourceVp.CenterPoint;
                    }
                    
                    // Copy frozen layers in viewport
                    try
                    {
                        ObjectIdCollection frozenLayers = sourceVp.GetFrozenLayers();
                        if (frozenLayers != null && frozenLayers.Count > 0)
                        {
                            foreach (ObjectId layerId in frozenLayers)
                            {
                                targetVp.FreezeLayersInViewport(frozenLayers.Cast<ObjectId>().GetEnumerator());
                                break; // Only need to do this once
                            }
                        }
                    }
                    catch { /* Skip layer freezing if it fails */ }
                }
            }
            catch (System.Exception)
            {
                // If viewport copying fails, continue without them
            }
        }

        private string GenerateUniqueName(string baseName, DBDictionary layoutDict)
        {
            string newName = baseName;
            int counter = 1;

            while (layoutDict.Contains(newName))
            {
                newName = $"{baseName} ({counter})";
                counter++;
            }

            return newName;
        }

        private string GetNextPaperSpaceIndex(BlockTable bt, Transaction tr)
        {
            int index = 0;
            string baseName = "*Paper_Space";
            
            // Find the next available index
            while (bt.Has(baseName + index))
            {
                index++;
            }
            
            return index.ToString();
        }

        private void AdjustTabOrders(DBDictionary layoutDict, int originalTabOrder, 
            Layout newLayout, Transaction tr)
        {
            int targetTabOrder = originalTabOrder + 1;

            // Shift existing tab orders
            foreach (DBDictionaryEntry entry in layoutDict)
            {
                if (entry.Value != newLayout.ObjectId)
                {
                    Layout layout = tr.GetObject(entry.Value, OpenMode.ForWrite) as Layout;
                    if (layout != null && layout.TabOrder >= targetTabOrder)
                    {
                        layout.TabOrder++;
                    }
                }
            }

            newLayout.TabOrder = targetTabOrder;
        }
    }
}
