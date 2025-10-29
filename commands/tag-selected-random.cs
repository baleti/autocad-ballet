using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class TagSelectedRandom
    {
        private static readonly Random Random = new Random();

        /// <summary>
        /// Generates a random alphanumeric tag (e.g., "tag-abc123")
        /// </summary>
        private static string GenerateRandomTag()
        {
            const string chars = "abcdefghijklmnopqrstuvwxyz0123456789";
            const int length = 6;
            var randomPart = new char[length];

            for (int i = 0; i < length; i++)
            {
                randomPart[i] = chars[Random.Next(chars.Length)];
            }

            return "tag-" + new string(randomPart);
        }

        /// <summary>
        /// Tags entities in the current view with a randomly generated tag
        /// </summary>
        public static void ExecuteViewScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            // Get pickfirst set or prompt for selection
            var selResult = ed.SelectImplied();

            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect objects to tag: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo objects selected.\n");
                return;
            }

            var selectedIds = selResult.Value.GetObjectIds();

            // Generate random tag
            var randomTag = GenerateRandomTag();
            var tagNames = new List<string> { randomTag };

            // Tag the selected entities
            TagEntities(selectedIds, tagNames, db, doc.Name, ed);

            ed.WriteMessage($"\nSuccessfully tagged {selectedIds.Length} object(s) with tag: {randomTag}\n");
        }

        /// <summary>
        /// Helper method to tag entities with the given tag names
        /// </summary>
        private static void TagEntities(ObjectId[] objectIds, List<string> tagNames, Database db,
            string documentPath, Editor ed)
        {
            foreach (var objectId in objectIds)
            {
                try
                {
                    // Add tags to entity (stored as text attribute in extension dictionary)
                    objectId.AddTags(db, tagNames.ToArray());
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError tagging entity: {ex.Message}\n");
                }
            }
        }
    }
}
