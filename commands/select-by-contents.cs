using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace AutoCADBallet
{
    public static class SelectByContents
    {
        /// <summary>
        /// View scope: Search current view/layout for entities containing text
        /// </summary>
        public static void ExecuteViewScope(Editor ed, Database db)
        {
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();
            string modeMsg = searchInBlocks ? " (including blocks/xrefs)" : "";

            // Prompt for search text
            string searchText = PromptForSearchText("Search Text in Current View");
            if (string.IsNullOrEmpty(searchText))
            {
                ed.WriteMessage("\nSearch cancelled.\n");
                return;
            }

            ed.WriteMessage($"\nSearching for '{searchText}' in current view{modeMsg}...\n");

            var matchingIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                foreach (ObjectId id in currentSpace)
                {
                    try
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity == null) continue;

                        // Check if entity contains the search text
                        if (EntityContainsText(entity, searchText, tr))
                        {
                            matchingIds.Add(id);
                        }

                        // If search-in-blocks mode is enabled and this is a block reference, search within
                        if (searchInBlocks && entity is BlockReference blockRef)
                        {
                            var blockEntities = BlockEntityUtilities.GetEntitiesInBlock(blockRef.BlockTableRecord, tr);
                            foreach (var blockEntityId in blockEntities)
                            {
                                try
                                {
                                    var blockEntity = tr.GetObject(blockEntityId, OpenMode.ForRead) as Entity;
                                    if (blockEntity != null && EntityContainsText(blockEntity, searchText, tr))
                                    {
                                        matchingIds.Add(blockEntityId);
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                tr.Commit();
            }

            if (matchingIds.Count > 0)
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;

                // Build selection references for ALL entities (including those in blocks)
                var selectionItems = new List<SelectionItem>();
                var selectableIds = new List<ObjectId>();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var id in matchingIds)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity != null)
                            {
                                // Add to selection storage (all entities, including block internals)
                                selectionItems.Add(new SelectionItem
                                {
                                    DocumentPath = doc.Name,
                                    Handle = entity.Handle.ToString(),
                                    SessionId = null
                                });

                                // Only add entities in current space to selectable list
                                if (entity.BlockId == db.CurrentSpaceId)
                                {
                                    selectableIds.Add(id);
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    tr.Commit();
                }

                // Save selection references (including block entities) for filter commands
                try
                {
                    var docName = Path.GetFileName(doc.Name);
                    SelectionStorage.SaveSelection(selectionItems, docName);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nWarning: Could not save selection references: {ex.Message}\n");
                }

                // Set implied selection for selectable entities only
                if (selectableIds.Count > 0)
                {
                    ed.SetImpliedSelection(selectableIds.ToArray());
                    ed.WriteMessage($"\nSelected {selectableIds.Count} entities containing '{searchText}'.\n");

                    // Report if some entities were in blocks and stored as references
                    int blockedCount = selectionItems.Count - selectableIds.Count;
                    if (blockedCount > 0)
                    {
                        ed.WriteMessage($"  ({blockedCount} entities within blocks stored as references for filter commands)\n");
                    }
                }
                else
                {
                    ed.WriteMessage($"\nNo directly selectable entities. All {selectionItems.Count} entities within blocks stored as references.\n");
                }
            }
            else
            {
                ed.WriteMessage($"\nNo entities found containing '{searchText}'.\n");
            }
        }

        /// <summary>
        /// Document scope: Search entire document for entities containing text
        /// </summary>
        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();
            string modeMsg = searchInBlocks ? " (including blocks/xrefs)" : "";

            // Prompt for search text
            string searchText = PromptForSearchText("Search Text in Document");
            if (string.IsNullOrEmpty(searchText))
            {
                ed.WriteMessage("\nSearch cancelled.\n");
                return;
            }

            ed.WriteMessage($"\nSearching for '{searchText}' in entire document{modeMsg}...\n");

            var matchingReferences = new List<ContentEntityReference>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Get all layouts
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity == null) continue;

                            if (EntityContainsText(entity, searchText, tr))
                            {
                                matchingReferences.Add(new ContentEntityReference
                                {
                                    DocumentPath = doc.Name,
                                    DocumentName = Path.GetFileName(doc.Name),
                                    Handle = entity.Handle.ToString(),
                                    SpaceName = spaceName,
                                    TextContent = GetEntityTextContent(entity, tr),
                                    IsInBlock = false
                                });
                            }

                            // If search-in-blocks mode is enabled and this is a block reference, search within
                            if (searchInBlocks && entity is BlockReference blockRef)
                            {
                                var blockName = BlockEntityUtilities.GetBlockName(blockRef, tr);
                                var blockEntities = BlockEntityUtilities.GetEntitiesInBlock(blockRef.BlockTableRecord, tr, blockName, blockRef.Handle.ToString());

                                foreach (var blockEntityId in blockEntities)
                                {
                                    try
                                    {
                                        var blockEntity = tr.GetObject(blockEntityId, OpenMode.ForRead) as Entity;
                                        if (blockEntity != null && EntityContainsText(blockEntity, searchText, tr))
                                        {
                                            matchingReferences.Add(new ContentEntityReference
                                            {
                                                DocumentPath = doc.Name,
                                                DocumentName = Path.GetFileName(doc.Name),
                                                Handle = blockEntity.Handle.ToString(),
                                                SpaceName = spaceName,
                                                TextContent = GetEntityTextContent(blockEntity, tr),
                                                IsInBlock = true
                                            });
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                tr.Commit();
            }

            if (matchingReferences.Count == 0)
            {
                ed.WriteMessage($"\nNo entities found containing '{searchText}'.\n");
                return;
            }

            var selectionItems = matchingReferences.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            // Save selection (replacing previous document selection)
            try
            {
                var docName = Path.GetFileName(doc.Name);
                SelectionStorage.SaveSelection(selectionItems, docName);
                ed.WriteMessage($"\nDocument selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {matchingReferences.Count}\n");
                ed.WriteMessage($"  Search text: '{searchText}'\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection to subset of matching entities in current space
            var currentViewIds = new List<ObjectId>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var entityRef in matchingReferences)
                {
                    try
                    {
                        var handle = Convert.ToInt64(entityRef.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (!objectId.IsNull && objectId.IsValid)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.BlockId == db.CurrentSpaceId)
                            {
                                currentViewIds.Add(objectId);
                            }
                        }
                    }
                    catch
                    {
                        // Skip invalid entities
                    }
                }
                tr.Commit();
            }

            if (currentViewIds.Count > 0)
            {
                try
                {
                    ed.SetImpliedSelection(currentViewIds.ToArray());
                    ed.WriteMessage($"  Selected {currentViewIds.Count} entities in current view.\n");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError setting current view selection: {ex.Message}\n");
                }
            }
        }

        /// <summary>
        /// Session scope: Search all open documents for entities containing text
        /// </summary>
        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Prompt for search text
            string searchText = PromptForSearchText("Search Text in All Open Documents");
            if (string.IsNullOrEmpty(searchText))
            {
                ed.WriteMessage("\nSearch cancelled.\n");
                return;
            }

            ed.WriteMessage($"\nSearching for '{searchText}' in all open documents...\n");

            var docManager = AcadApp.DocumentManager;
            var allReferences = new List<ContentEntityReference>();

            foreach (Autodesk.AutoCAD.ApplicationServices.Document openDoc in docManager)
            {
                string docPath = openDoc.Name;
                string docName = Path.GetFileName(docPath);

                ed.WriteMessage($"\nScanning: {docName}...");

                try
                {
                    var refs = GatherMatchingEntitiesFromDocument(openDoc.Database, docPath, docName, searchText);
                    allReferences.AddRange(refs);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
                }
            }

            if (allReferences.Count == 0)
            {
                ed.WriteMessage($"\n\nNo entities found containing '{searchText}' in any open documents.\n");
                return;
            }

            var selectionItems = allReferences.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            // Clear all existing selections and save (session scope behavior)
            ClearAllStoredSelections();

            try
            {
                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage($"\n\nApplication selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {allReferences.Count}\n");
                ed.WriteMessage($"  Search text: '{searchText}'\n");

                // Group by document
                var docGroups = allReferences.GroupBy(e => e.DocumentName)
                                             .OrderBy(g => g.Key)
                                             .ToList();
                ed.WriteMessage($"  Documents: {docGroups.Count}\n");
                foreach (var docGroup in docGroups)
                {
                    ed.WriteMessage($"    {docGroup.Key}: {docGroup.Count()} entities\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection to subset of matching entities in current document's current space
            var currentDocPath = Path.GetFullPath(doc.Name);
            var currentViewIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var entityRef in allReferences)
                {
                    try
                    {
                        // Only process entities from current document
                        var entityDocPath = Path.GetFullPath(entityRef.DocumentPath);
                        if (!string.Equals(entityDocPath, currentDocPath, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var handle = Convert.ToInt64(entityRef.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (!objectId.IsNull && objectId.IsValid)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.BlockId == db.CurrentSpaceId)
                            {
                                currentViewIds.Add(objectId);
                            }
                        }
                    }
                    catch
                    {
                        // Skip invalid entities
                    }
                }
                tr.Commit();
            }

            if (currentViewIds.Count > 0)
            {
                try
                {
                    ed.SetImpliedSelection(currentViewIds.ToArray());
                    ed.WriteMessage($"  Selected {currentViewIds.Count} entities in current view.\n");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError setting current view selection: {ex.Message}\n");
                }
            }
        }

        #region Helper Methods

        /// <summary>
        /// Prompt user for search text using a simple Windows Forms dialog
        /// </summary>
        private static string PromptForSearchText(string title)
        {
            string result = null;

            using (var form = new WinForms.Form())
            {
                form.Text = title;
                form.Width = 500;
                form.Height = 150;
                form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                form.StartPosition = WinForms.FormStartPosition.CenterScreen;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.Font = new Drawing.Font("Segoe UI", 9);

                var label = new WinForms.Label
                {
                    Text = "Search for text (case insensitive):",
                    Location = new Drawing.Point(10, 15),
                    AutoSize = true
                };

                var textBox = new WinForms.TextBox
                {
                    Location = new Drawing.Point(10, 40),
                    Width = 460,
                    Font = new Drawing.Font("Segoe UI", 10)
                };

                var okButton = new WinForms.Button
                {
                    Text = "OK",
                    DialogResult = WinForms.DialogResult.OK,
                    Location = new Drawing.Point(300, 75),
                    Width = 80
                };

                var cancelButton = new WinForms.Button
                {
                    Text = "Cancel",
                    DialogResult = WinForms.DialogResult.Cancel,
                    Location = new Drawing.Point(390, 75),
                    Width = 80
                };

                form.Controls.Add(label);
                form.Controls.Add(textBox);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);

                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                // Set focus to textbox when shown
                form.Shown += (s, e) => textBox.Focus();

                if (form.ShowDialog() == WinForms.DialogResult.OK)
                {
                    result = textBox.Text?.Trim();
                }
            }

            return result;
        }

        /// <summary>
        /// Check if an entity contains the specified text (case insensitive)
        /// </summary>
        private static bool EntityContainsText(Entity entity, string searchText, Transaction tr)
        {
            if (entity == null || string.IsNullOrEmpty(searchText))
                return false;

            string entityText = GetEntityTextContent(entity, tr);
            if (string.IsNullOrEmpty(entityText))
                return false;

            return entityText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Extract text content from various entity types
        /// </summary>
        private static string GetEntityTextContent(Entity entity, Transaction tr)
        {
            if (entity == null)
                return null;

            try
            {
                if (entity is DBText dbText)
                {
                    return dbText.TextString;
                }
                else if (entity is MText mText)
                {
                    return mText.Contents;
                }
                else if (entity is AttributeReference attrRef)
                {
                    return attrRef.TextString;
                }
                else if (entity is Dimension dimension)
                {
                    return dimension.DimensionText;
                }
                else if (entity is MLeader mLeader)
                {
                    // MLeaders can have multiple text contents
                    var textContents = new List<string>();

                    // Try to get MText content if it exists
                    if (mLeader.ContentType == ContentType.MTextContent)
                    {
                        var mtext = mLeader.MText;
                        if (mtext != null)
                        {
                            textContents.Add(mtext.Contents);
                        }
                    }

                    return textContents.Count > 0 ? string.Join(" ", textContents) : null;
                }
                else if (entity is BlockReference blockRef)
                {
                    // Get attribute text from block references
                    var attrTexts = new List<string>();

                    foreach (ObjectId attrId in blockRef.AttributeCollection)
                    {
                        var attr = tr.GetObject(attrId, OpenMode.ForRead) as AttributeReference;
                        if (attr != null && !string.IsNullOrEmpty(attr.TextString))
                        {
                            attrTexts.Add(attr.TextString);
                        }
                    }

                    return attrTexts.Count > 0 ? string.Join(" ", attrTexts) : null;
                }
            }
            catch
            {
                return null;
            }

            return null;
        }

        /// <summary>
        /// Gather matching entities from a specific document
        /// </summary>
        private static List<ContentEntityReference> GatherMatchingEntitiesFromDocument(Database db, string docPath, string docName, string searchText)
        {
            var references = new List<ContentEntityReference>();
            bool searchInBlocks = SelectInBlocksMode.IsEnabled();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity == null) continue;

                            if (EntityContainsText(entity, searchText, tr))
                            {
                                references.Add(new ContentEntityReference
                                {
                                    DocumentPath = docPath,
                                    DocumentName = docName,
                                    Handle = entity.Handle.ToString(),
                                    SpaceName = spaceName,
                                    TextContent = GetEntityTextContent(entity, tr),
                                    IsInBlock = false
                                });
                            }

                            // If search-in-blocks mode is enabled and this is a block reference, search within
                            if (searchInBlocks && entity is BlockReference blockRef)
                            {
                                var blockName = BlockEntityUtilities.GetBlockName(blockRef, tr);
                                var blockEntities = BlockEntityUtilities.GetEntitiesInBlock(blockRef.BlockTableRecord, tr, blockName, blockRef.Handle.ToString());

                                foreach (var blockEntityId in blockEntities)
                                {
                                    try
                                    {
                                        var blockEntity = tr.GetObject(blockEntityId, OpenMode.ForRead) as Entity;
                                        if (blockEntity != null && EntityContainsText(blockEntity, searchText, tr))
                                        {
                                            references.Add(new ContentEntityReference
                                            {
                                                DocumentPath = docPath,
                                                DocumentName = docName,
                                                Handle = blockEntity.Handle.ToString(),
                                                SpaceName = spaceName,
                                                TextContent = GetEntityTextContent(blockEntity, tr),
                                                IsInBlock = true
                                            });
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                tr.Commit();
            }

            return references;
        }

        /// <summary>
        /// Clear all stored selection files
        /// </summary>
        private static void ClearAllStoredSelections()
        {
            try
            {
                var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var selectionDir = Path.Combine(appDataPath, "autocad-ballet", "runtime", "selection");

                if (Directory.Exists(selectionDir))
                {
                    foreach (var file in Directory.GetFiles(selectionDir))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch
                        {
                            // Skip files that can't be deleted
                        }
                    }
                }
            }
            catch
            {
                // If clearing fails, continue anyway
            }
        }

        #endregion
    }

    /// <summary>
    /// Simple class to store entity references with text content
    /// </summary>
    public class ContentEntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string SpaceName { get; set; }
        public string TextContent { get; set; }
        public bool IsInBlock { get; set; }
    }
}
