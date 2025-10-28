using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADBallet
{
    public static class TagSelected
    {
        /// <summary>
        /// Tags entities in the current view with user-selected or newly created tags
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

            // Show DataGrid with existing tags from current database
            var existingTags = TagHelpers.GetAllTagsInDatabase(db);
            var entries = existingTags.Select(t => new Dictionary<string, object>
            {
                { "TagName", t.Name },
                { "UsageCount", t.UsageCount }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect existing tags or type a new tag name and press Enter...\n");

            // Use DataGrid with allowCreateFromSearch enabled
            var selectedTags = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: true);

            if (selectedTags == null || selectedTags.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract tag names - either from selected entries or from search text
            var tagNames = new List<string>();

            foreach (var tag in selectedTags)
            {
                if (tag.ContainsKey("__SEARCH_TEXT__"))
                {
                    // User typed a new tag name
                    tagNames.Add((string)tag["__SEARCH_TEXT__"]);
                }
                else if (tag.ContainsKey("TagName"))
                {
                    // User selected existing tag
                    tagNames.Add((string)tag["TagName"]);
                }
            }

            if (tagNames.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags.\n");
                return;
            }

            // Tag the selected entities
            TagEntities(selectedIds, tagNames, db, doc.Name, ed);

            ed.WriteMessage($"\nSuccessfully tagged {selectedIds.Length} object(s) with {tagNames.Count} tag(s).\n");
        }

        /// <summary>
        /// Tags entities from stored selection in current document
        /// </summary>
        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var docName = Path.GetFileName(doc.Name);

            var storedSelection = SelectionStorage.LoadSelection(docName);
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage($"\nNo stored selection found for document. Use 'select-by-categories-in-document' first.\n");
                return;
            }

            // Filter to current document only
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

            if (storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection in current document.\n");
                return;
            }

            // Convert to ObjectIds
            var objectIds = new List<ObjectId>();
            foreach (var item in storedSelection)
            {
                try
                {
                    var handle = Convert.ToInt64(item.Handle, 16);
                    var objectId = db.GetObjectId(false, new Handle(handle), 0);
                    if (objectId != ObjectId.Null)
                    {
                        objectIds.Add(objectId);
                    }
                }
                catch { }
            }

            if (objectIds.Count == 0)
            {
                ed.WriteMessage("\nNo valid objects found in stored selection.\n");
                return;
            }

            // Show DataGrid with existing tags from current database
            var existingTags = TagHelpers.GetAllTagsInDatabase(db);
            var entries = existingTags.Select(t => new Dictionary<string, object>
            {
                { "TagName", t.Name },
                { "UsageCount", t.UsageCount }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect existing tags or type a new tag name and press Enter...\n");

            var selectedTags = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: true);

            if (selectedTags == null || selectedTags.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract tag names
            var tagNames = new List<string>();
            foreach (var tag in selectedTags)
            {
                if (tag.ContainsKey("__SEARCH_TEXT__"))
                {
                    tagNames.Add((string)tag["__SEARCH_TEXT__"]);
                }
                else if (tag.ContainsKey("TagName"))
                {
                    tagNames.Add((string)tag["TagName"]);
                }
            }

            if (tagNames.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags.\n");
                return;
            }

            // Tag the entities
            TagEntities(objectIds.ToArray(), tagNames, db, doc.Name, ed);

            ed.WriteMessage($"\nSuccessfully tagged {objectIds.Count} object(s) with {tagNames.Count} tag(s).\n");
        }

        /// <summary>
        /// Tags entities from stored selection across all documents
        /// </summary>
        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            var storedSelection = SelectionStorage.LoadSelectionFromAllDocuments();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use 'select-by-categories-in-application' first.\n");
                return;
            }

            // Show DataGrid with existing tags from all open documents
            var docManager = AcadApp.DocumentManager;
            var databases = new List<Database>();
            foreach (Document openDoc in docManager)
            {
                databases.Add(openDoc.Database);
            }
            var existingTags = TagHelpers.GetAllTagsInDocuments(databases);
            var entries = existingTags.Select(t => new Dictionary<string, object>
            {
                { "TagName", t.Name },
                { "UsageCount", t.UsageCount }
            }).ToList();

            var propertyNames = new List<string> { "TagName", "UsageCount" };

            ed.WriteMessage($"\nSelect existing tags or type a new tag name and press Enter...\n");

            var selectedTags = CustomGUIs.DataGrid(entries, propertyNames, spanAllScreens: false,
                initialSelectionIndices: null, onDeleteEntries: null, allowCreateFromSearch: true);

            if (selectedTags == null || selectedTags.Count == 0)
            {
                ed.WriteMessage("\nNo tags selected.\n");
                return;
            }

            // Extract tag names
            var tagNames = new List<string>();
            foreach (var tag in selectedTags)
            {
                if (tag.ContainsKey("__SEARCH_TEXT__"))
                {
                    tagNames.Add((string)tag["__SEARCH_TEXT__"]);
                }
                else if (tag.ContainsKey("TagName"))
                {
                    tagNames.Add((string)tag["TagName"]);
                }
            }

            if (tagNames.Count == 0)
            {
                ed.WriteMessage("\nNo valid tags.\n");
                return;
            }

            // Separate current document from external documents
            var currentDocEntities = new List<ObjectId>();
            var externalDocuments = new Dictionary<string, List<SelectionItem>>();

            foreach (var item in storedSelection)
            {
                try
                {
                    var itemPath = Path.GetFullPath(item.DocumentPath);
                    var currentPath = Path.GetFullPath(doc.Name);

                    if (string.Equals(itemPath, currentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);
                        if (objectId != ObjectId.Null)
                        {
                            currentDocEntities.Add(objectId);
                        }
                    }
                    else
                    {
                        if (!externalDocuments.ContainsKey(item.DocumentPath))
                        {
                            externalDocuments[item.DocumentPath] = new List<SelectionItem>();
                        }
                        externalDocuments[item.DocumentPath].Add(item);
                    }
                }
                catch { }
            }

            // Tag current document entities
            if (currentDocEntities.Count > 0)
            {
                TagEntities(currentDocEntities.ToArray(), tagNames, db, doc.Name, ed);
                ed.WriteMessage($"\nTagged {currentDocEntities.Count} object(s) in current document.\n");
            }

            // Tag external document entities
            var docs = AcadApp.DocumentManager;
            foreach (var kvp in externalDocuments)
            {
                var extDocPath = kvp.Key;
                var extItems = kvp.Value;

                // Try to find already open document
                Document extDoc = null;
                foreach (Document openDoc in docs)
                {
                    try
                    {
                        if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(extDocPath),
                            StringComparison.OrdinalIgnoreCase))
                        {
                            extDoc = openDoc;
                            break;
                        }
                    }
                    catch { }
                }

                if (extDoc != null)
                {
                    // Document is already open - tag entities in it
                    var extDb = extDoc.Database;
                    var extObjectIds = new List<ObjectId>();

                    foreach (var item in extItems)
                    {
                        try
                        {
                            var handle = Convert.ToInt64(item.Handle, 16);
                            var objectId = extDb.GetObjectId(false, new Handle(handle), 0);
                            if (objectId != ObjectId.Null)
                            {
                                extObjectIds.Add(objectId);
                            }
                        }
                        catch { }
                    }

                    if (extObjectIds.Count > 0)
                    {
                        TagEntities(extObjectIds.ToArray(), tagNames, extDb, extDoc.Name, ed);
                        ed.WriteMessage($"\nTagged {extObjectIds.Count} object(s) in {Path.GetFileName(extDocPath)}.\n");
                    }
                }
                else
                {
                    ed.WriteMessage($"\nSkipping {Path.GetFileName(extDocPath)} - document not open.\n");
                }
            }

            ed.WriteMessage("\nTagging operation completed.\n");
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
