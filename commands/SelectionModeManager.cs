using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Static class to manage selection mode across different scopes
public static class SelectionModeManager
{
    private static readonly string AppDataPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "autocad-ballet", "runtime"
    );

    private static readonly string ModeFilePath = Path.Combine(AppDataPath, "SelectionMode");
    private static readonly string StoredSelectionsPath = Path.Combine(AppDataPath, "StoredSelections");

    public enum SelectionMode
    {
        SpaceLayout,    // Current active space/layout only
        Drawing,        // Entire current drawing (all layouts)
        Process,        // All opened drawings in current AutoCAD process
        Desktop,        // All drawings in all AutoCAD instances on desktop (IPC later)
        Network         // Across network (to implement later)
    }

    // Structure to store selection information
    public class StoredSelection
    {
        public string DrawingPath { get; set; }
        public string LayoutName { get; set; }
        public List<string> Handles { get; set; } = new List<string>();
        public DateTime Timestamp { get; set; }
    }

    static SelectionModeManager()
    {
        // Ensure directories exist
        Directory.CreateDirectory(AppDataPath);
        Directory.CreateDirectory(StoredSelectionsPath);
    }

    public static SelectionMode CurrentMode
    {
        get
        {
            if (File.Exists(ModeFilePath))
            {
                string mode = File.ReadAllText(ModeFilePath).Trim();
                if (Enum.TryParse<SelectionMode>(mode, out var result))
                    return result;
            }
            return SelectionMode.SpaceLayout;
        }
        set
        {
            File.WriteAllText(ModeFilePath, value.ToString());
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

    // Extension method to get selection based on current mode
    public static PromptSelectionResult GetSelectionEx(this Editor ed)
    {
        return GetSelectionEx(ed, null);
    }

    public static PromptSelectionResult GetSelectionEx(this Editor ed, SelectionFilter filter)
    {
        var mode = CurrentMode;
        var doc = ed.Document;
        var db = doc.Database;

        switch (mode)
        {
            case SelectionMode.SpaceLayout:
                // Default behavior - current space/layout
                return filter != null ? ed.GetSelection(filter) : ed.GetSelection();

            case SelectionMode.Drawing:
                // Get selection from all layouts in current drawing
                return GetDrawingWideSelection(ed, db, filter);

            case SelectionMode.Process:
                // Get selection from all open drawings
                return GetProcessWideSelection(ed, filter);

            case SelectionMode.Desktop:
                // TODO: Implement IPC mechanism to communicate with other AutoCAD instances
                ed.WriteMessage("\nDesktop mode not yet implemented. Using Process mode instead.\n");
                return GetProcessWideSelection(ed, filter);

            case SelectionMode.Network:
                // TODO: Implement network communication
                ed.WriteMessage("\nNetwork mode not yet implemented. Using Process mode instead.\n");
                return GetProcessWideSelection(ed, filter);

            default:
                return filter != null ? ed.GetSelection(filter) : ed.GetSelection();
        }
    }

    // Get selection from entire drawing (all layouts)
    private static PromptSelectionResult GetDrawingWideSelection(Editor ed, Database db, SelectionFilter filter)
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

    // Get selection from all open drawings in process
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
                var result = GetDrawingWideSelection(ed, doc.Database, filter);
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

                // Load stored selection for this document if it exists
                var storedSel = LoadStoredSelection(doc);
                if (storedSel != null && storedSel.Handles.Count > 0)
                {
                    ed.WriteMessage($"  Found {storedSel.Handles.Count} stored objects\n");
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
        var mode = CurrentMode;
        var doc = ed.Document;

        switch (mode)
        {
            case SelectionMode.SpaceLayout:
                // Default behavior
                ed.SetImpliedSelection(ids);
                break;

            case SelectionMode.Drawing:
            case SelectionMode.Process:
            case SelectionMode.Desktop:
            case SelectionMode.Network:
                // Store selection for retrieval
                var selection = new StoredSelection
                {
                    DrawingPath = doc.Name,
                    LayoutName = LayoutManager.Current.CurrentLayout,
                    Timestamp = DateTime.Now
                };

                // Convert ObjectIds to handles for persistence
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    foreach (var id in ids)
                    {
                        if (id.IsValid && !id.IsErased)
                        {
                            var obj = tr.GetObject(id, OpenMode.ForRead);
                            selection.Handles.Add(obj.Handle.ToString());
                        }
                    }
                    tr.Commit();
                }

                SaveStoredSelection(doc, selection);
                ed.SetImpliedSelection(ids);
                break;
        }
    }

    // Clear stored selections
    public static void ClearStoredSelections()
    {
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
public class ToggleSelectionMode
{
    [CommandMethod("toggle-select-mode")]
    public void Toggle()
    {
        var currentMode = SelectionModeManager.CurrentMode;
        var modes = Enum.GetValues(typeof(SelectionModeManager.SelectionMode));
        var currentIndex = Array.IndexOf(modes, currentMode);
        var nextIndex = (currentIndex + 1) % modes.Length;
        var newMode = (SelectionModeManager.SelectionMode)modes.GetValue(nextIndex);

        SelectionModeManager.CurrentMode = newMode;

        var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
        ed.WriteMessage($"\nSelection mode changed to: {newMode}\n");

        // Clear stored selections when changing mode
        if (newMode != SelectionModeManager.SelectionMode.SpaceLayout)
        {
            SelectionModeManager.ClearStoredSelections();
        }
    }
}

// Command to switch selection mode with UI
public class SwitchSelectionMode
{
    [CommandMethod("switch-select-mode")]
    public void Switch()
    {
        var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
        var currentMode = SelectionModeManager.CurrentMode;

        // Create keywords for selection
        var pko = new PromptKeywordOptions("\nSelect mode:");
        pko.Keywords.Add("SpaceLayout");
        pko.Keywords.Add("Drawing");
        pko.Keywords.Add("Process");
        pko.Keywords.Add("Desktop");
        pko.Keywords.Add("Network");
        pko.Keywords.Default = currentMode.ToString();
        pko.AllowNone = true;

        var result = ed.GetKeywords(pko);
        if (result.Status != PromptStatus.OK)
            return;

        if (Enum.TryParse<SelectionModeManager.SelectionMode>(result.StringResult, out var newMode))
        {
            SelectionModeManager.CurrentMode = newMode;
            ed.WriteMessage($"\nSelection mode set to: {newMode}\n");

            // Clear stored selections when changing mode
            if (newMode != SelectionModeManager.SelectionMode.SpaceLayout)
            {
                SelectionModeManager.ClearStoredSelections();
            }
        }
    }
}
