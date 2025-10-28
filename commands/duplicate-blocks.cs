using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

/// <summary>
/// Command to duplicate selected blocks by creating new block definitions with " 2" appended to their names
/// </summary>
public class DuplicateBlocks
{
    public static void ExecuteViewScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        try
        {
            var selectedBlockRefs = new List<BlockReference>();
            var blockDefsToProcess = new HashSet<string>();

            // Get pickfirst set (pre-selected objects)
            var selResult = ed.SelectImplied();

            // If there is no pickfirst set, request user to select objects
            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect block references to duplicate: ";

                // Filter to only allow block references
                var filter = new SelectionFilter(new TypedValue[] {
                    new TypedValue((int)DxfCode.Start, "INSERT")
                });

                selResult = ed.GetSelection(selectionOpts, filter);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                // Clear the pickfirst set since we're consuming it
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo block references selected.\n");
                return;
            }

            // Process selected block references
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId objId in selResult.Value.GetObjectIds())
                {
                    var blockRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                    if (blockRef != null)
                    {
                        var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null && !btr.IsFromExternalReference && !btr.IsAnonymous)
                        {
                            selectedBlockRefs.Add(blockRef);
                            blockDefsToProcess.Add(btr.Name);
                        }
                    }
                }
                tr.Commit();
            }

            DuplicateAndReplaceBlocks(doc, selectedBlockRefs, blockDefsToProcess, ed, null);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError duplicating blocks: {ex.Message}\n");
        }
    }

    public static void ExecuteDocumentScope(Editor ed, Database db)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;

        try
        {
            var docName = Path.GetFileName(doc.Name);
            var storedSelection = SelectionStorage.LoadSelection(docName);
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage($"\nNo stored selection found for current document '{docName}'. Use commands like 'select-by-categories-in-document' first.\n");
                return;
            }

            var selectedBlockRefs = new List<BlockReference>();
            var blockDefsToProcess = new HashSet<string>();

            // Process stored selections - only handle current document blocks
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var item in storedSelection)
                {
                    try
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (objectId != ObjectId.Null)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead);
                            if (entity is BlockReference blockRef)
                            {
                                var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null && !btr.IsFromExternalReference && !btr.IsAnonymous)
                                {
                                    selectedBlockRefs.Add(blockRef);
                                    blockDefsToProcess.Add(btr.Name);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Skip problematic entities
                        continue;
                    }
                }
                tr.Commit();
            }

            DuplicateAndReplaceBlocks(doc, selectedBlockRefs, blockDefsToProcess, ed, docName);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError duplicating blocks: {ex.Message}\n");
        }
    }

    public static void ExecuteApplicationScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        try
        {
            var storedSelection = SelectionStorage.LoadSelectionFromOpenDocuments();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-categories-in-session' first.\n");
                return;
            }

            var selectedBlockRefs = new List<BlockReference>();
            var blockDefsToProcess = new HashSet<string>();

            // Process stored selections - only handle current document blocks
            var currentDocSelections = storedSelection.Where(item =>
            {
                try
                {
                    return string.Equals(Path.GetFullPath(item.DocumentPath), Path.GetFullPath(doc.Name), StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return string.Equals(item.DocumentPath, doc.Name, StringComparison.OrdinalIgnoreCase);
                }
            }).ToList();

            if (currentDocSelections.Count == 0)
            {
                ed.WriteMessage("\nNo block references found in current document from stored selection. Only blocks in the current document can be duplicated.\n");
                return;
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var item in currentDocSelections)
                {
                    try
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (objectId != ObjectId.Null)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead);
                            if (entity is BlockReference blockRef)
                            {
                                var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null && !btr.IsFromExternalReference && !btr.IsAnonymous)
                                {
                                    selectedBlockRefs.Add(blockRef);
                                    blockDefsToProcess.Add(btr.Name);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Skip problematic entities
                        continue;
                    }
                }
                tr.Commit();
            }

            // Report external entities (from other documents) that can't be duplicated
            var externalSelections = storedSelection.Where(item =>
            {
                try
                {
                    return !string.Equals(Path.GetFullPath(item.DocumentPath), Path.GetFullPath(doc.Name), StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return !string.Equals(item.DocumentPath, doc.Name, StringComparison.OrdinalIgnoreCase);
                }
            }).ToList();

            if (externalSelections.Count > 0)
            {
                ed.WriteMessage($"\nNote: {externalSelections.Count} external entities found but cannot be duplicated (blocks can only be duplicated within their source document).\n");
            }

            DuplicateAndReplaceBlocks(doc, selectedBlockRefs, blockDefsToProcess, ed, null);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError duplicating blocks: {ex.Message}\n");
        }
    }

    private static void DuplicateAndReplaceBlocks(Document doc, List<BlockReference> selectedBlockRefs, HashSet<string> blockDefsToProcess, Editor ed, string docNameForStorage)
    {
        var db = doc.Database;

        if (selectedBlockRefs.Count == 0)
        {
            ed.WriteMessage("\nNo valid block references found. Only regular (non-xref, non-anonymous) blocks can be duplicated.\n");
            return;
        }

        // Dictionary to map original block names to new duplicated block names
        var blockNameMapping = new Dictionary<string, string>();
        var duplicatedCount = 0;

        // Second pass: duplicate each unique block definition
        using (var docLock = doc.LockDocument())
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (var originalBlockName in blockDefsToProcess)
                {
                    var newBlockName = GenerateUniqueBlockName(blockTable, originalBlockName);

                    if (DuplicateBlockDefinition(tr, blockTable, originalBlockName, newBlockName))
                    {
                        blockNameMapping[originalBlockName] = newBlockName;
                        duplicatedCount++;
                    }
                }

                tr.Commit();
            }
        }

        if (duplicatedCount == 0)
        {
            ed.WriteMessage("\nNo block definitions could be duplicated.\n");
            return;
        }

        // Third pass: replace the selected block references with references to the new block definitions
        var replacedCount = 0;
        using (var docLock = doc.LockDocument())
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (var blockRef in selectedBlockRefs)
                {
                    var currentBlockRef = tr.GetObject(blockRef.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (currentBlockRef != null)
                    {
                        var originalBtr = tr.GetObject(currentBlockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (originalBtr != null && blockNameMapping.ContainsKey(originalBtr.Name))
                        {
                            var newBlockName = blockNameMapping[originalBtr.Name];
                            var newBtrId = blockTable[newBlockName];

                            // Create a new block reference with the same properties
                            var newBlockRef = new BlockReference(currentBlockRef.Position, newBtrId)
                            {
                                Rotation = currentBlockRef.Rotation,
                                ScaleFactors = currentBlockRef.ScaleFactors,
                                Layer = currentBlockRef.Layer,
                                Color = currentBlockRef.Color,
                                Linetype = currentBlockRef.Linetype,
                                LinetypeScale = currentBlockRef.LinetypeScale,
                                LineWeight = currentBlockRef.LineWeight,
                                Transparency = currentBlockRef.Transparency,
                                Visible = currentBlockRef.Visible
                            };

                            // Copy attributes if they exist
                            CopyBlockAttributes(tr, currentBlockRef, newBlockRef);

                            // Add the new block reference to the same container
                            var container = tr.GetObject(currentBlockRef.BlockId, OpenMode.ForWrite) as BlockTableRecord;
                            container.AppendEntity(newBlockRef);
                            tr.AddNewlyCreatedDBObject(newBlockRef, true);

                            // Remove the original block reference
                            currentBlockRef.UpgradeOpen();
                            currentBlockRef.Erase();

                            replacedCount++;
                        }
                    }
                }

                tr.Commit();
            }
        }

        ed.WriteMessage($"\nDuplicated {duplicatedCount} block definition(s) and replaced {replacedCount} block reference(s).\n");

        // List the new block names created
        if (blockNameMapping.Count > 0)
        {
            ed.WriteMessage("New block definitions created:\n");
            foreach (var mapping in blockNameMapping)
            {
                ed.WriteMessage($"  {mapping.Key} -> {mapping.Value}\n");
            }
        }

        // Update stored selection if docNameForStorage is provided
        if (docNameForStorage != null && replacedCount > 0)
        {
            // Collect the new block references for selection storage
            var newSelectionItems = new List<SelectionItem>();
            using (var docLock = doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var container = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId id in container)
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead);
                        if (entity is BlockReference blockRef)
                        {
                            var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            if (btr != null && blockNameMapping.ContainsValue(btr.Name))
                            {
                                newSelectionItems.Add(new SelectionItem
                                {
                                    DocumentPath = doc.Name,
                                    Handle = entity.Handle.ToString(),
                                    SessionId = null // Will be auto-generated
                                });
                            }
                        }
                    }
                    tr.Commit();
                }
            }

            if (newSelectionItems.Count > 0)
            {
                SelectionStorage.SaveSelection(newSelectionItems, docNameForStorage);
                ed.WriteMessage("Updated stored selection with new block references.\n");
            }
        }
    }

    private static string GenerateUniqueBlockName(BlockTable blockTable, string originalName)
    {
        var baseName = originalName + " 2";
        var uniqueName = baseName;
        var counter = 3;

        // Keep incrementing until we find a unique name
        while (blockTable.Has(uniqueName))
        {
            uniqueName = originalName + " " + counter.ToString();
            counter++;
        }

        return uniqueName;
    }

    private static bool DuplicateBlockDefinition(Transaction tr, BlockTable blockTable, string originalName, string newName)
    {
        try
        {
            if (!blockTable.Has(originalName))
                return false;

            var originalBtrId = blockTable[originalName];
            var originalBtr = tr.GetObject(originalBtrId, OpenMode.ForRead) as BlockTableRecord;

            if (originalBtr == null || originalBtr.IsFromExternalReference || originalBtr.IsAnonymous)
                return false;

            // Create new block table record
            var newBtr = new BlockTableRecord
            {
                Name = newName,
                Origin = originalBtr.Origin,
                Comments = originalBtr.Comments,
                BlockScaling = originalBtr.BlockScaling,
                Explodable = originalBtr.Explodable
            };

            // Add to block table
            blockTable.UpgradeOpen();
            var newBtrId = blockTable.Add(newBtr);
            tr.AddNewlyCreatedDBObject(newBtr, true);

            // Copy all entities from original block to new block
            foreach (ObjectId entId in originalBtr)
            {
                var originalEnt = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                if (originalEnt != null)
                {
                    var clonedEnt = originalEnt.Clone() as Entity;
                    newBtr.AppendEntity(clonedEnt);
                    tr.AddNewlyCreatedDBObject(clonedEnt, true);
                }
            }

            return true;
        }
        catch (System.Exception ex)
        {
            var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage($"\nError duplicating block '{originalName}': {ex.Message}\n");
            return false;
        }
    }

    private static void CopyBlockAttributes(Transaction tr, BlockReference sourceBlock, BlockReference targetBlock)
    {
        try
        {
            // Get the block table record for the target block to check for attribute definitions
            var targetBtr = tr.GetObject(targetBlock.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            if (targetBtr == null)
                return;

            var hasAttributeDefs = false;
            var attributeDefs = new Dictionary<string, AttributeDefinition>();

            // Collect attribute definitions from the target block
            foreach (ObjectId id in targetBtr)
            {
                var attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                if (attDef != null)
                {
                    hasAttributeDefs = true;
                    attributeDefs[attDef.Tag.ToUpper()] = attDef;
                }
            }

            if (!hasAttributeDefs)
                return;

            // Copy attribute values from source block
            var sourceAttributes = new Dictionary<string, string>();
            foreach (ObjectId attId in sourceBlock.AttributeCollection)
            {
                var sourceAttRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (sourceAttRef != null)
                {
                    sourceAttributes[sourceAttRef.Tag.ToUpper()] = sourceAttRef.TextString;
                }
            }

            // Create attribute references for the target block
            foreach (var kvp in attributeDefs)
            {
                var attDef = kvp.Value;
                var attRef = new AttributeReference();

                attRef.SetAttributeFromBlock(attDef, targetBlock.BlockTransform);

                // Set the value from source if it exists, otherwise use default
                if (sourceAttributes.ContainsKey(attDef.Tag.ToUpper()))
                {
                    attRef.TextString = sourceAttributes[attDef.Tag.ToUpper()];
                }
                else
                {
                    attRef.TextString = attDef.TextString;
                }

                targetBlock.AttributeCollection.AppendAttribute(attRef);
                tr.AddNewlyCreatedDBObject(attRef, true);
            }
        }
        catch
        {
            // Skip attribute copying if there are any issues
        }
    }
}