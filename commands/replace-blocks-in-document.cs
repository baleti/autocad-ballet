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
            var selectedBlockRefs = new List<BlockReference>();

            // This command always uses document scope - load stored selection from current document
            var docName = Path.GetFileName(doc.Name);
            var storedSelection = SelectionStorage.LoadSelection(docName);
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage($"\nNo stored selection found for current document '{docName}'. Use commands like 'select-by-categories-in-document' first.\n");
                return;
            }

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
                    var targetBtr = tr.GetObject(targetBtrId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (var blockRef in selectedBlockRefs)
                    {
                        var currentBlockRef = tr.GetObject(blockRef.ObjectId, OpenMode.ForWrite) as BlockReference;
                        if (currentBlockRef != null)
                        {
                            // Save old block attributes before swapping
                            var oldAttributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            foreach (ObjectId attId in currentBlockRef.AttributeCollection)
                            {
                                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (attRef != null)
                                {
                                    oldAttributes[attRef.Tag] = attRef.TextString;
                                }
                            }

                            // Collect old attribute IDs before changing the block
                            var oldAttributeIds = new List<ObjectId>();
                            foreach (ObjectId attId in currentBlockRef.AttributeCollection)
                            {
                                oldAttributeIds.Add(attId);
                            }

                            // Change the block reference to point to the new block definition
                            currentBlockRef.BlockTableRecord = targetBtrId;

                            // Remove old attributes after changing the block reference
                            foreach (ObjectId attId in oldAttributeIds)
                            {
                                if (!attId.IsErased)
                                {
                                    var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                                    if (attRef != null)
                                    {
                                        attRef.Erase();
                                    }
                                }
                            }

                            // Add new attributes from the new block definition
                            int copiedCount = 0;
                            int newCount = 0;
                            foreach (ObjectId defId in targetBtr)
                            {
                                var attDef = tr.GetObject(defId, OpenMode.ForRead) as AttributeDefinition;
                                if (attDef != null && !attDef.Constant)
                                {
                                    var attRef = new AttributeReference();

                                    // CRITICAL: Set database defaults to associate AttributeReference with correct database
                                    // This prevents eWrongDatabase errors when appending to the block reference
                                    attRef.SetDatabaseDefaults(db);

                                    attRef.SetAttributeFromBlock(attDef, currentBlockRef.BlockTransform);

                                    // Try to copy value from old attributes if tag matches
                                    if (oldAttributes.TryGetValue(attDef.Tag, out string oldValue))
                                    {
                                        attRef.TextString = oldValue;
                                        copiedCount++;
                                    }
                                    else
                                    {
                                        attRef.TextString = attDef.TextString; // Use default value
                                        newCount++;
                                    }

                                    // Add the attribute to the block reference
                                    currentBlockRef.AttributeCollection.AppendAttribute(attRef);
                                    tr.AddNewlyCreatedDBObject(attRef, true);
                                }
                            }

                            replacedCount++;
                            ed.WriteMessage($"\nReplaced block ({copiedCount} attrs copied, {newCount} new)");
                        }
                    }

                    tr.Commit();
                }
            }

            ed.WriteMessage($"\nReplaced {replacedCount} block reference(s) with '{targetBlockName}'.\n");

            // Update stored selection for document scope
            if (replacedCount > 0)
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
                    SelectionStorage.SaveSelection(newSelectionItems, docName);
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
}