using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCADBallet
{
    /// <summary>
    /// Extended entity reference that includes block path information.
    /// This allows tracking entities that exist within block definitions, which cannot be directly selected in AutoCAD.
    /// </summary>
    public class BlockEntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string SpaceName { get; set; }

        /// <summary>
        /// Path showing the nesting of blocks containing this entity.
        /// Format: "BlockA > BlockB > BlockC" for nested blocks.
        /// Empty string for entities in top-level spaces (Model/Paper).
        /// </summary>
        public string BlockPath { get; set; }

        /// <summary>
        /// Indicates whether this entity is within a block definition (true) or in top-level space (false).
        /// </summary>
        public bool IsInBlock { get; set; }

        /// <summary>
        /// Block reference handle that contains this entity (if IsInBlock is true).
        /// Null or empty for top-level entities.
        /// </summary>
        public string ContainingBlockRefHandle { get; set; }
    }

    /// <summary>
    /// Utilities for gathering entities from within block definitions, xrefs, and dynamic blocks.
    /// </summary>
    public static class BlockEntityUtilities
    {
        /// <summary>
        /// Recursively gathers all entities from a block definition, including nested blocks.
        /// </summary>
        /// <param name="blockId">The BlockTableRecord ObjectId to search</param>
        /// <param name="tr">The active transaction</param>
        /// <param name="currentPath">Current block nesting path (for recursion)</param>
        /// <param name="blockRefHandle">Handle of the block reference that contains these entities</param>
        /// <param name="processedBlocks">Set of already-processed block IDs to prevent infinite recursion</param>
        /// <returns>List of entity ObjectIds found within the block</returns>
        public static List<ObjectId> GetEntitiesInBlock(
            ObjectId blockId,
            Transaction tr,
            string currentPath = "",
            string blockRefHandle = "",
            HashSet<ObjectId> processedBlocks = null)
        {
            var entities = new List<ObjectId>();

            if (processedBlocks == null)
                processedBlocks = new HashSet<ObjectId>();

            // Prevent infinite recursion from circular block references
            if (processedBlocks.Contains(blockId))
                return entities;

            processedBlocks.Add(blockId);

            try
            {
                var btr = tr.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null)
                    return entities;

                // Skip layout blocks (Model, Paper spaces)
                if (btr.IsLayout)
                    return entities;

                foreach (ObjectId entityId in btr)
                {
                    try
                    {
                        var entity = tr.GetObject(entityId, OpenMode.ForRead);

                        // Add the entity itself
                        entities.Add(entityId);

                        // If this is a block reference, recursively search it
                        if (entity is BlockReference blockRef)
                        {
                            var nestedBlockId = blockRef.BlockTableRecord;
                            var nestedBtr = tr.GetObject(nestedBlockId, OpenMode.ForRead) as BlockTableRecord;

                            string nestedPath = string.IsNullOrEmpty(currentPath)
                                ? nestedBtr.Name
                                : $"{currentPath} > {nestedBtr.Name}";

                            var nestedEntities = GetEntitiesInBlock(
                                nestedBlockId,
                                tr,
                                nestedPath,
                                blockRef.Handle.ToString(),
                                processedBlocks);

                            entities.AddRange(nestedEntities);
                        }
                    }
                    catch
                    {
                        // Skip entities that can't be accessed
                        continue;
                    }
                }
            }
            catch
            {
                // Skip blocks that can't be accessed
            }

            return entities;
        }

        /// <summary>
        /// Gets all block references in a space, including references to xrefs and dynamic blocks.
        /// </summary>
        /// <param name="spaceId">The space (layout) BlockTableRecord ObjectId</param>
        /// <param name="tr">The active transaction</param>
        /// <returns>List of BlockReference objects</returns>
        public static List<BlockReference> GetBlockReferencesInSpace(ObjectId spaceId, Transaction tr)
        {
            var blockRefs = new List<BlockReference>();

            try
            {
                var btr = tr.GetObject(spaceId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null)
                    return blockRefs;

                foreach (ObjectId entityId in btr)
                {
                    try
                    {
                        var entity = tr.GetObject(entityId, OpenMode.ForRead);
                        if (entity is BlockReference blockRef)
                        {
                            blockRefs.Add(blockRef);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch
            {
                // Skip spaces that can't be accessed
            }

            return blockRefs;
        }

        /// <summary>
        /// Creates a BlockEntityReference from an entity, including block path information.
        /// </summary>
        public static BlockEntityReference CreateBlockEntityReference(
            DBObject entity,
            string docPath,
            string docName,
            string spaceName,
            string blockPath,
            bool isInBlock,
            string containingBlockRefHandle)
        {
            return new BlockEntityReference
            {
                DocumentPath = docPath,
                DocumentName = docName,
                Handle = entity.Handle.ToString(),
                SpaceName = spaceName,
                BlockPath = blockPath,
                IsInBlock = isInBlock,
                ContainingBlockRefHandle = containingBlockRefHandle
            };
        }

        /// <summary>
        /// Gets block type description for a block reference.
        /// </summary>
        public static string GetBlockType(BlockReference blockRef, Transaction tr)
        {
            try
            {
                var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null)
                    return "Block";

                if (btr.IsFromExternalReference)
                    return "XRef";
                else if (btr.IsDynamicBlock || btr.IsAnonymous)
                    return "Dynamic Block";
                else
                    return "Block Reference";
            }
            catch
            {
                return "Block";
            }
        }

        /// <summary>
        /// Gets the block name for a block reference.
        /// </summary>
        public static string GetBlockName(BlockReference blockRef, Transaction tr)
        {
            try
            {
                var btr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null)
                    return blockRef.Name;

                // For dynamic blocks, try to get the original block name
                if (btr.IsDynamicBlock || btr.IsAnonymous)
                {
                    // Get the dynamic block table record from the block reference
                    var dynamicBlockId = blockRef.DynamicBlockTableRecord;
                    if (dynamicBlockId != ObjectId.Null)
                    {
                        var dynamicBtr = tr.GetObject(dynamicBlockId, OpenMode.ForRead) as BlockTableRecord;
                        if (dynamicBtr != null)
                            return dynamicBtr.Name;
                    }
                }

                return btr.Name;
            }
            catch
            {
                return blockRef.Name;
            }
        }
    }
}
