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

public class DeleteSelected
{
    public static void ExecuteViewScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        try
        {
            var selResult = ed.SelectImplied();

            if (selResult.Status == PromptStatus.Error)
            {
                var selectionOpts = new PromptSelectionOptions();
                selectionOpts.MessageForAdding = "\nSelect entities to delete: ";
                selResult = ed.GetSelection(selectionOpts);
            }
            else if (selResult.Status == PromptStatus.OK)
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
            {
                ed.WriteMessage("\nNo entities selected.\n");
                return;
            }

            var selectedIds = selResult.Value.GetObjectIds();
            DeleteEntitiesInCurrentDocument(doc, selectedIds, ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError deleting entities: {ex.Message}\n");
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

            var currentSessionId = GetCurrentSessionId();
            storedSelection = storedSelection.Where(item =>
                string.IsNullOrEmpty(item.SessionId) || item.SessionId == currentSessionId).ToList();

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
                ed.WriteMessage($"\nNo stored selection found for current document '{docName}'.\n");
                return;
            }

            var currentDocEntities = new List<ObjectId>();

            foreach (var item in storedSelection)
            {
                try
                {
                    var handle = Convert.ToInt64(item.Handle, 16);
                    var objectId = db.GetObjectId(false, new Handle(handle), 0);
                    if (objectId != ObjectId.Null)
                    {
                        currentDocEntities.Add(objectId);
                    }
                }
                catch
                {
                    continue;
                }
            }

            if (currentDocEntities.Count > 0)
            {
                int deleted = DeleteEntitiesInCurrentDocument(doc, currentDocEntities.ToArray(), ed);
                if (deleted > 0)
                {
                    SelectionStorage.SaveSelection(new List<SelectionItem>(), docName);
                    ed.WriteMessage("Cleared stored selection.\n");
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError deleting entities: {ex.Message}\n");
        }
    }

    public static void ExecuteApplicationScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        try
        {
            var storedSelection = SelectionStorage.LoadSelectionFromAllDocuments();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-categories-in-application' first.\n");
                return;
            }

            var currentSessionId = GetCurrentSessionId();
            storedSelection = storedSelection.Where(item =>
                string.IsNullOrEmpty(item.SessionId) || item.SessionId == currentSessionId).ToList();

            if (storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found after filtering.\n");
                return;
            }

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
                catch
                {
                    continue;
                }
            }

            int totalDeleted = 0;

            if (currentDocEntities.Count > 0)
            {
                int deleted = DeleteEntitiesInCurrentDocument(doc, currentDocEntities.ToArray(), ed);
                totalDeleted += deleted;
            }

            foreach (var externalDoc in externalDocuments)
            {
                int deleted = DeleteEntitiesInExternalDocument(externalDoc.Key, externalDoc.Value, ed);
                totalDeleted += deleted;
            }

            if (totalDeleted > 0)
            {
                // Clear all document-specific selection files for application scope
                ClearAllStoredSelections();
                ed.WriteMessage("Cleared stored selection.\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError deleting entities: {ex.Message}\n");
        }
    }

    /// <summary>
    /// Clear all stored selections across all documents for application scope
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
            // If clearing fails, ignore - entities were already deleted
        }
    }

    private static string GetCurrentSessionId()
    {
        int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
        int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
        return $"{processId}_{sessionId}";
    }

    private static int DeleteEntitiesInCurrentDocument(Document doc, ObjectId[] entityIds, Editor ed)
    {
        var db = doc.Database;
        int deletedCount = 0;

        using (var docLock = doc.LockDocument())
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var objectId in entityIds)
                {
                    try
                    {
                        var dbObj = tr.GetObject(objectId, OpenMode.ForWrite);
                        if (dbObj != null && !dbObj.IsErased)
                        {
                            dbObj.Erase();
                            deletedCount++;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                tr.Commit();
            }
        }

        if (deletedCount > 0)
        {
            ed.WriteMessage($"\nDeleted {deletedCount} entities from current document.\n");
        }
        else
        {
            ed.WriteMessage("\nNo entities could be deleted from current document.\n");
        }

        return deletedCount;
    }

    private static int DeleteEntitiesInExternalDocument(string documentPath, List<SelectionItem> entities, Editor ed)
    {
        var docs = AcadApp.DocumentManager;
        Document externalDoc = null;
        bool docWasAlreadyOpen = false;
        int deletedCount = 0;

        try
        {
            foreach (Document openDoc in docs)
            {
                if (string.Equals(Path.GetFullPath(openDoc.Name), Path.GetFullPath(documentPath), StringComparison.OrdinalIgnoreCase))
                {
                    externalDoc = openDoc;
                    docWasAlreadyOpen = true;
                    break;
                }
            }

            if (externalDoc == null && File.Exists(documentPath))
            {
                try
                {
                    externalDoc = docs.Open(documentPath, false);
                    docWasAlreadyOpen = false;
                }
                catch
                {
                    ed.WriteMessage($"\nCould not open document: {Path.GetFileName(documentPath)}\n");
                    return 0;
                }
            }

            if (externalDoc != null)
            {
                var objectIds = new List<ObjectId>();

                foreach (var item in entities)
                {
                    try
                    {
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = externalDoc.Database.GetObjectId(false, new Handle(handle), 0);
                        if (objectId != ObjectId.Null)
                        {
                            objectIds.Add(objectId);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (objectIds.Count > 0)
                {
                    using (var docLock = externalDoc.LockDocument())
                    {
                        using (var tr = externalDoc.Database.TransactionManager.StartTransaction())
                        {
                            foreach (var objectId in objectIds)
                            {
                                try
                                {
                                    var dbObj = tr.GetObject(objectId, OpenMode.ForWrite);
                                    if (dbObj != null && !dbObj.IsErased)
                                    {
                                        dbObj.Erase();
                                        deletedCount++;
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }

                            tr.Commit();
                        }
                    }

                    if (deletedCount > 0)
                    {
                        ed.WriteMessage($"\nDeleted {deletedCount} entities from {Path.GetFileName(documentPath)}.\n");
                    }
                }
            }
        }
        finally
        {
            if (externalDoc != null && !docWasAlreadyOpen)
            {
                try
                {
                    externalDoc.CloseAndSave(documentPath);
                }
                catch
                {
                    try
                    {
                        externalDoc.CloseAndDiscard();
                    }
                    catch
                    {
                    }
                }
            }
        }

        return deletedCount;
    }
}