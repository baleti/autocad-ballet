using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class CopyTagsFromSelection
    {
        /// <summary>
        /// Copies tags from selected entities to other entities in the current view
        /// </summary>
        public static void ExecuteViewScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Get source selection (entities with tags to copy from)
            var sourceSelResult = ed.SelectImplied();

            if (sourceSelResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect source objects with tags to copy from: ";
                sourceSelResult = ed.GetSelection(selectionOpts);
            }
            else if (sourceSelResult.Status == PromptStatus.OK)
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (sourceSelResult.Status != PromptStatus.OK || sourceSelResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo source objects selected.\n");
                return;
            }

            var sourceIds = sourceSelResult.Value.GetObjectIds();

            // Collect all unique tags from source entities
            var allTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sourceId in sourceIds)
            {
                try
                {
                    var tags = sourceId.GetTags(db);
                    foreach (var tag in tags)
                    {
                        allTags.Add(tag);
                    }
                }
                catch { }
            }

            if (allTags.Count == 0)
            {
                ed.WriteMessage("\nNo tags found on source objects.\n");
                return;
            }

            ed.WriteMessage($"\nFound {allTags.Count} unique tag(s) on source objects: {string.Join(", ", allTags)}\n");

            // Prompt for target selection (entities to copy tags to)
            var targetSelOpts = new PromptSelectionOptions();
            targetSelOpts.MessageForAdding = "\nSelect target objects to copy tags to: ";
            var targetSelResult = ed.GetSelection(targetSelOpts);

            if (targetSelResult.Status != PromptStatus.OK || targetSelResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo target objects selected.\n");
                return;
            }

            var targetIds = targetSelResult.Value.GetObjectIds();

            // Copy tags to target entities
            int successCount = 0;
            foreach (var targetId in targetIds)
            {
                try
                {
                    targetId.AddTags(db, allTags.ToArray());
                    successCount++;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError copying tags to entity: {ex.Message}\n");
                }
            }

            ed.WriteMessage($"\nSuccessfully copied {allTags.Count} tag(s) to {successCount} object(s).\n");
        }

    }
}
