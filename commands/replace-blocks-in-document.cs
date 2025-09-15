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

[assembly: CommandClass(typeof(ReplaceBlocksInDocument))]

/// <summary>
/// Command to replace selected blocks with a user-chosen block type from current document
/// </summary>
public class ReplaceBlocksInDocument
{
    [CommandMethod("replace-blocks-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void ReplaceBlocksInDocumentCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            var currentScope = SelectionScopeManager.CurrentScope;
            var selectedBlockRefs = new List<BlockReference>();

            // Handle selection based on current scope
            if (currentScope == SelectionScopeManager.SelectionScope.view)
            {
                // Get pickfirst set (pre-selected objects)
                var selResult = ed.SelectImplied();

                // If there is no pickfirst set, request user to select objects
                if (selResult.Status == PromptStatus.Error)
                {
                    var selectionOpts = new PromptSelectionOptions();
                    selectionOpts.MessageForAdding = "\nSelect block references to replace: ";

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
                            if (btr != null && !btr.IsFromExternalReference)
                            {
                                selectedBlockRefs.Add(blockRef);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            else
            {
                // For document, process, desktop, network scopes - use stored selection
                var storedSelection = SelectionStorage.LoadSelection();
                if (storedSelection == null || storedSelection.Count == 0)
                {
                    ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-category' first or switch to 'view' scope.\n");
                    return;
                }

                // Filter to current document if in document scope
                if (currentScope == SelectionScopeManager.SelectionScope.document)
                {
                    var currentDocPath = Path.GetFullPath(doc.Name);
                    storedSelection = storedSelection.Where(item =>
                    {
                        try
                        {
                            var itemPath = Path.GetFullPath(item.DocumentPath);
                            return string.Equals(itemPath, currentDocPath, StringComparison.OrdinalIgnoreCase);
                        }
                        catch
                        {
                            return string.Equals(item.DocumentPath, doc.Name, StringComparison.OrdinalIgnoreCase);
                        }
                    }).ToList();
                }

                if (storedSelection.Count == 0)
                {
                    ed.WriteMessage($"\nNo stored selection found for current scope '{currentScope}'. Use commands like 'select-by-category' first.\n");
                    return;
                }

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
                    ed.WriteMessage("\nNo block references found in current document from stored selection. Only blocks in the current document can be replaced.\n");
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
                                    if (btr != null && !btr.IsFromExternalReference)
                                    {
                                        selectedBlockRefs.Add(blockRef);
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

                // Report external entities (from other documents) that can't be replaced
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
                    ed.WriteMessage($"\nNote: {externalSelections.Count} external entities found but cannot be replaced (blocks can only be replaced within their source document).\n");
                }
            }

            if (selectedBlockRefs.Count == 0)
            {
                ed.WriteMessage("\nNo valid block references found. Only regular (non-xref) blocks can be replaced.\n");
                return;
            }

            // Get all available block definitions in current document
            var availableBlocks = GetAvailableBlockDefinitions(db);
            if (availableBlocks.Count == 0)
            {
                ed.WriteMessage("\nNo block definitions found in current document.\n");
                return;
            }

            // Get unique block names from selected block references for initial selection
            var selectedBlockNames = new HashSet<string>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var blockRef in selectedBlockRefs)
                {
                    var currentBlockRef = tr.GetObject(blockRef.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (currentBlockRef != null)
                    {
                        var btr = tr.GetObject(currentBlockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null)
                        {
                            selectedBlockNames.Add(btr.Name);
                        }
                    }
                }
                tr.Commit();
            }

            // Show datagrid for block selection
            var blockData = availableBlocks.Select(name => new Dictionary<string, object>
            {
                ["Block Name"] = name
            }).ToList();

            // Create initial selection indices for blocks that match selected instances
            var initialSelectionIndices = new List<int>();
            for (int i = 0; i < blockData.Count; i++)
            {
                var blockName = blockData[i]["Block Name"].ToString();
                if (selectedBlockNames.Contains(blockName))
                {
                    initialSelectionIndices.Add(i);
                }
            }

            var chosenBlocks = CustomGUIs.DataGrid(blockData, new List<string> { "Block Name" }, false, initialSelectionIndices);
            if (chosenBlocks.Count != 1)
            {
                ed.WriteMessage("\nPlease select exactly one block type to replace with.\n");
                return;
            }

            var targetBlockName = chosenBlocks[0]["Block Name"].ToString();

            // Replace the selected block references with the chosen block type
            var replacedCount = 0;
            using (var docLock = doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var targetBtrId = blockTable[targetBlockName];

                    foreach (var blockRef in selectedBlockRefs)
                    {
                        var currentBlockRef = tr.GetObject(blockRef.ObjectId, OpenMode.ForRead) as BlockReference;
                        if (currentBlockRef != null)
                        {
                            // Create a new block reference with the same properties but different block
                            var newBlockRef = new BlockReference(currentBlockRef.Position, targetBtrId)
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

                            // Copy compatible attributes
                            CopyCompatibleAttributes(tr, currentBlockRef, newBlockRef);

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

                    tr.Commit();
                }
            }

            ed.WriteMessage($"\nReplaced {replacedCount} block reference(s) with '{targetBlockName}'.\n");

            // Update stored selection if not in view scope
            if (currentScope != SelectionScopeManager.SelectionScope.view && replacedCount > 0)
            {
                // Collect the new block references for selection storage
                var newSelectionItems = new List<SelectionItem>();
                using (var docLock = doc.LockDocument())
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var targetBtrId = blockTable[targetBlockName];

                        // Search all layouts for new block references
                        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                            foreach (ObjectId id in btr)
                            {
                                var entity = tr.GetObject(id, OpenMode.ForRead);
                                if (entity is BlockReference blockRef && blockRef.BlockTableRecord == targetBtrId)
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
                    SelectionStorage.SaveSelection(newSelectionItems);
                    ed.WriteMessage("Updated stored selection with replaced block references.\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError replacing blocks: {ex.Message}\n");
        }
    }

    private List<string> GetAvailableBlockDefinitions(Database db)
    {
        var blockNames = new List<string>();

        using (var tr = db.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            foreach (ObjectId btrId in blockTable)
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr != null && !btr.IsLayout && !btr.IsFromExternalReference &&
                    !btr.IsAnonymous && !btr.Name.StartsWith("*"))
                {
                    blockNames.Add(btr.Name);
                }
            }

            tr.Commit();
        }

        return blockNames.OrderBy(name => name).ToList();
    }

    private void CopyCompatibleAttributes(Transaction tr, BlockReference sourceBlock, BlockReference targetBlock)
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

            // Get attribute values from source block
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

                // Set the value from source if compatible tag exists, otherwise use default
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