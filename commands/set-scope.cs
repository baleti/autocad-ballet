using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Register the command classes
[assembly: CommandClass(typeof(ToggleSelectionScope))]
[assembly: CommandClass(typeof(SetSelectionScope))]

// Static class to manage selection scope across different contexts
public static class SelectionScopeManager
{
    private static readonly string AppDataPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "autocad-ballet", "runtime"
    );

    private static readonly string ScopeFilePath = Path.Combine(AppDataPath, "scope");
    private static readonly string StoredSelectionsPath = Path.Combine(AppDataPath, "StoredSelections");

    public enum SelectionScope
    {
        view,           // Current active space/layout only (was SpaceLayout)
        document,       // Entire current document (all layouts)
        process,        // All opened documents in current AutoCAD process
        desktop,        // All documents in all AutoCAD instances on desktop (IPC later)
        network         // Across network (to implement later)
    }

    // Structure to store selection information
    public class StoredSelection
    {
        public string DrawingPath { get; set; }
        public string LayoutName { get; set; }
        public List<string> Handles { get; set; } = new List<string>();
        public DateTime Timestamp { get; set; }
    }

    static SelectionScopeManager()
    {
        // Ensure directories exist
        Directory.CreateDirectory(AppDataPath);
        // Note: StoredSelectionsPath is legacy - we now use unified SelectionStorage
    }

    public static SelectionScope CurrentScope
    {
        get
        {
            if (File.Exists(ScopeFilePath))
            {
                string scope = File.ReadAllText(ScopeFilePath).Trim();
                if (Enum.TryParse<SelectionScope>(scope, out var result))
                    return result;
            }
            return SelectionScope.view;
        }
        set
        {
            File.WriteAllText(ScopeFilePath, value.ToString());
        }
    }

    // Get selection file path for current drawing
    private static string GetSelectionFilePath(Document doc)
    {
        var drawingName = Path.GetFileNameWithoutExtension(doc.Name);
        if (string.IsNullOrEmpty(drawingName))
            drawingName = "Unnamed";

        return Path.Combine(StoredSelectionsPath, $"{drawingName}_selection.txt");
    }

    // Save selection to file
    private static void SaveStoredSelection(Document doc, StoredSelection selection)
    {
        var filePath = GetSelectionFilePath(doc);
        var lines = new List<string>
        {
            selection.DrawingPath,
            selection.LayoutName,
            selection.Timestamp.ToString("o"),
            string.Join(",", selection.Handles)
        };
        File.WriteAllLines(filePath, lines);
    }

    // Load selection from file
    private static StoredSelection LoadStoredSelection(Document doc)
    {
        var filePath = GetSelectionFilePath(doc);
        if (!File.Exists(filePath))
            return null;

        try
        {
            var lines = File.ReadAllLines(filePath);
            if (lines.Length >= 4)
            {
                return new StoredSelection
                {
                    DrawingPath = lines[0],
                    LayoutName = lines[1],
                    Timestamp = DateTime.Parse(lines[2]),
                    Handles = lines[3].Split(',').Where(h => !string.IsNullOrEmpty(h)).ToList()
                };
            }
        }
        catch { }

        return null;
    }

    // Extension method to get selection based on current scope
    public static PromptSelectionResult GetSelectionEx(this Editor ed)
    {
        return GetSelectionEx(ed, null);
    }

    public static PromptSelectionResult GetSelectionEx(this Editor ed, SelectionFilter filter)
    {
        var scope = CurrentScope;
        var doc = ed.Document;
        var db = doc.Database;

        switch (scope)
        {
            case SelectionScope.view:
                // Default behavior - current space/layout
                return filter != null ? ed.GetSelection(filter) : ed.GetSelection();

            case SelectionScope.document:
                // Get selection from all layouts in current document
                return GetDocumentWideSelection(ed, db, filter);

            case SelectionScope.process:
                // Get selection from all open documents
                return GetProcessWideSelection(ed, filter);

            case SelectionScope.desktop:
                // TODO: Implement IPC mechanism to communicate with other AutoCAD instances
                ed.WriteMessage("\nDesktop scope not yet implemented. Using Process scope instead.\n");
                return GetProcessWideSelection(ed, filter);

            case SelectionScope.network:
                // TODO: Implement network communication
                ed.WriteMessage("\nNetwork scope not yet implemented. Using Process scope instead.\n");
                return GetProcessWideSelection(ed, filter);

            default:
                return filter != null ? ed.GetSelection(filter) : ed.GetSelection();
        }
    }

    // Get selection from entire document (all layouts)
    private static PromptSelectionResult GetDocumentWideSelection(Editor ed, Database db, SelectionFilter filter)
    {
        var allIds = new List<ObjectId>();

        using (var tr = db.TransactionManager.StartTransaction())
        {
            // Get block table
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            // Iterate through all layouts
            var layoutMgr = LayoutManager.Current;
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                // Collect all entities in this layout
                foreach (ObjectId id in btr)
                {
                    if (filter != null)
                    {
                        // Apply filter if provided
                        var entity = tr.GetObject(id, OpenMode.ForRead);
                        // Note: SelectionFilter application would need custom implementation
                        // For now, add all entities
                        allIds.Add(id);
                    }
                    else
                    {
                        allIds.Add(id);
                    }
                }
            }

            tr.Commit();
        }

        // Create selection set from collected IDs
        if (allIds.Count > 0)
        {
            // Use SelectAll and then filter to our IDs
            var result = ed.SelectAll(filter);
            if (result.Status == PromptStatus.OK)
            {
                // Filter to only our collected IDs
                var filteredIds = result.Value.GetObjectIds().Where(id => allIds.Contains(id)).ToArray();
                if (filteredIds.Length > 0)
                {
                    ed.SetImpliedSelection(filteredIds);
                    return ed.SelectImplied();
                }
            }
        }

        // Return empty selection result
        ed.SetImpliedSelection(new ObjectId[0]);
        return ed.SelectImplied();
    }

    // Get selection from all open documents in process
    private static PromptSelectionResult GetProcessWideSelection(Editor ed, SelectionFilter filter)
    {
        var allIds = new List<ObjectId>();
        var docMgr = AcadApp.DocumentManager;
        var currentDoc = docMgr.MdiActiveDocument;

        // Store current document's selections
        var currentDocIds = new List<ObjectId>();

        foreach (Document doc in docMgr)
        {
            if (doc == currentDoc)
            {
                // For current document, use regular selection
                var result = GetDocumentWideSelection(ed, doc.Database, filter);
                if (result.Status == PromptStatus.OK)
                {
                    currentDocIds.AddRange(result.Value.GetObjectIds());
                }
            }
            else
            {
                // For other documents, we need to activate them temporarily
                // Note: This is simplified - in production, you'd want to handle this more carefully
                ed.WriteMessage($"\nScanning document: {doc.Name}\n");

                // Load stored selection from unified storage if it exists
                var storedSelection = AutoCADBallet.SelectionStorage.LoadSelection();
                var docSelections = storedSelection.Where(s => Path.GetFullPath(s.DocumentPath) == Path.GetFullPath(doc.Name)).ToList();
                if (docSelections.Count > 0)
                {
                    ed.WriteMessage($"  Found {docSelections.Count} stored objects\n");
                }
            }
        }

        // For now, return selections from current document only
        // Full implementation would require document switching or stored selections
        if (currentDocIds.Count > 0)
        {
            ed.SetImpliedSelection(currentDocIds.ToArray());
            return ed.SelectImplied();
        }

        // Return empty selection
        ed.SetImpliedSelection(new ObjectId[0]);
        return ed.SelectImplied();
    }

    // Extension method to set selection based on current mode
    public static void SetImpliedSelectionEx(this Editor ed, ObjectId[] ids)
    {
        var scope = CurrentScope;
        var doc = ed.Document;

        switch (scope)
        {
            case SelectionScope.view:
                // Default behavior - use AutoCAD's built-in selection (no storage needed)
                ed.SetImpliedSelection(ids);
                break;

            case SelectionScope.document:
            case SelectionScope.process:
            case SelectionScope.desktop:
            case SelectionScope.network:
                // Store selection using unified SelectionStorage
                var selectionItems = new List<AutoCADBallet.SelectionItem>();
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    foreach (var id in ids)
                    {
                        if (id.IsValid && !id.IsErased)
                        {
                            var obj = tr.GetObject(id, OpenMode.ForRead);
                            selectionItems.Add(new AutoCADBallet.SelectionItem
                            {
                                DocumentPath = doc.Name,
                                Handle = obj.Handle.ToString(),
                                SessionId = null // Will be auto-generated by SelectionStorage
                            });
                        }
                    }
                    tr.Commit();
                }

                AutoCADBallet.SelectionStorage.SaveSelection(selectionItems);
                break;
        }
    }

    // Clear stored selections
    public static void ClearStoredSelections()
    {
        // Clear unified selection storage
        try
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var unifiedSelectionPath = Path.Combine(appDataPath, "autocad-ballet", "runtime", "selection");
            if (File.Exists(unifiedSelectionPath))
            {
                File.Delete(unifiedSelectionPath);
            }
        }
        catch { }

        // Also clear legacy stored selections for compatibility
        if (Directory.Exists(StoredSelectionsPath))
        {
            foreach (var file in Directory.GetFiles(StoredSelectionsPath, "*_selection.txt"))
            {
                try { File.Delete(file); } catch { }
            }
        }
    }

    // Get handles from stored selection
    public static List<ObjectId> GetStoredSelectionIds(Document doc)
    {
        var selection = LoadStoredSelection(doc);
        if (selection == null || selection.Handles.Count == 0)
            return new List<ObjectId>();

        var ids = new List<ObjectId>();
        var db = doc.Database;

        using (var tr = doc.TransactionManager.StartTransaction())
        {
            foreach (var handleStr in selection.Handles)
            {
                if (long.TryParse(handleStr, System.Globalization.NumberStyles.HexNumber, null, out long handleValue))
                {
                    var handle = new Handle(handleValue);
                    if (db.TryGetObjectId(handle, out ObjectId id))
                    {
                        if (id.IsValid && !id.IsErased)
                            ids.Add(id);
                    }
                }
            }
            tr.Commit();
        }

        return ids;
    }
}

// Toggle command - quickly switch between modes
public class ToggleSelectionScope
{
    [CommandMethod("toggle-selection-scope")]
    public void Toggle()
    {
        var currentScope = SelectionScopeManager.CurrentScope;
        var scopes = Enum.GetValues(typeof(SelectionScopeManager.SelectionScope));
        var currentIndex = Array.IndexOf(scopes, currentScope);
        var nextIndex = (currentIndex + 1) % scopes.Length;
        var newScope = (SelectionScopeManager.SelectionScope)scopes.GetValue(nextIndex);

        SelectionScopeManager.CurrentScope = newScope;

        var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
        ed.WriteMessage($"\nSelection scope changed to: {newScope}\n");

        // Clear stored selections when changing scope
        if (newScope != SelectionScopeManager.SelectionScope.view)
        {
            SelectionScopeManager.ClearStoredSelections();
        }
    }
}

// Command to switch selection mode with UI
public class SetSelectionScope
{
    [CommandMethod("set-scope")]
    public void Switch()
    {
        var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
        var currentScope = SelectionScopeManager.CurrentScope;

        // Create scope options for DataGrid
        var scopeOptions = new List<Dictionary<string, object>>();
        var scopes = Enum.GetValues(typeof(SelectionScopeManager.SelectionScope));

        foreach (SelectionScopeManager.SelectionScope scope in scopes)
        {
            scopeOptions.Add(new Dictionary<string, object>
            {
                ["scope"] = scope.ToString(),
                ["ScopeEnum"] = scope
            });
        }

        var propertyNames = new List<string> { "scope" };

        // Find current selection index
        int currentIndex = Array.IndexOf(scopes, currentScope);
        var initialSelectionIndices = currentIndex >= 0
            ? new List<int> { currentIndex }
            : new List<int>();

        try
        {
            var chosen = CustomGUIs.DataGrid(scopeOptions, propertyNames, false, initialSelectionIndices);

            if (chosen != null && chosen.Count > 0)
            {
                var selectedScope = (SelectionScopeManager.SelectionScope)chosen.First()["ScopeEnum"];

                SelectionScopeManager.CurrentScope = selectedScope;
                ed.WriteMessage($"\nSelection scope set to: {selectedScope}\n");

                // Clear stored selections when changing scope
                if (selectedScope != SelectionScopeManager.SelectionScope.view)
                {
                    SelectionScopeManager.ClearStoredSelections();
                }
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError displaying scope selection: {ex.Message}\n");
            // Fall back to current behavior if DataGrid fails
            ed.WriteMessage($"\nCurrent scope: {currentScope}\n");
        }
    }

}
