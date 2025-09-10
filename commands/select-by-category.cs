using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

public class SelectByCategory
{
    [CommandMethod("select-by-cat")]
    public void SelectByCategoryCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        // Show current selection mode
        var currentMode = SelectionModeManager.CurrentMode;
        ed.WriteMessage($"\nSelection Mode: {currentMode}\n");

        // Gather entity categories based on current mode
        var categories = GatherEntityCategories(db, currentMode);

        if (categories.Count == 0)
        {
            ed.WriteMessage("\nNo entities found in current mode.\n");
            return;
        }

        // Prepare data for DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var cat in categories.OrderBy(c => c.Key))
        {
            entries.Add(new Dictionary<string, object>
            {
                { "Category", cat.Key },
                { "DXF Name", GetDxfName(cat.Key) },
                { "Count", cat.Value.Count },
                { "Selection Mode", currentMode.ToString() }
            });
        }

        var propertyNames = new List<string> { "Category", "DXF Name", "Count", "Selection Mode" };

        // Show DataGrid for selection
        ed.WriteMessage("\nSelect categories to include in selection...\n");
        var selectedCategories = CustomGUIs.DataGrid(entries, propertyNames, false);

        if (selectedCategories == null || selectedCategories.Count == 0)
        {
            ed.WriteMessage("\nNo categories selected.\n");
            return;
        }

        // Collect all ObjectIds for selected categories
        var allSelectedIds = new List<ObjectId>();
        foreach (var selected in selectedCategories)
        {
            string categoryName = selected["Category"].ToString();
            if (categories.ContainsKey(categoryName))
            {
                allSelectedIds.AddRange(categories[categoryName]);
            }
        }

        // Set selection using extension method
        if (allSelectedIds.Count > 0)
        {
            ed.SetImpliedSelectionEx(allSelectedIds.ToArray());
            ed.WriteMessage($"\n{allSelectedIds.Count} objects selected from {selectedCategories.Count} categories.\n");

            // Report details
            foreach (var selected in selectedCategories)
            {
                string categoryName = selected["Category"].ToString();
                int count = categories[categoryName].Count;
                ed.WriteMessage($"  {categoryName}: {count} objects\n");
            }
        }
        else
        {
            ed.WriteMessage("\nNo objects found for selected categories.\n");
        }
    }

    [CommandMethod("quick-select-cat")]
    public void QuickSelectCategory()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        // Show current selection mode
        var currentMode = SelectionModeManager.CurrentMode;

        // Get all available categories
        var categories = GatherEntityCategories(db, currentMode);

        if (categories.Count == 0)
        {
            ed.WriteMessage("\nNo entities found.\n");
            return;
        }

        // Create keyword options from categories
        var pko = new PromptKeywordOptions($"\nSelect category (Mode: {currentMode}):");

        // Add top 10 most common categories as keywords
        var topCategories = categories.OrderByDescending(c => c.Value.Count)
                                     .Take(10)
                                     .ToList();

        foreach (var cat in topCategories)
        {
            string keyword = cat.Key.Replace(" ", "");
            pko.Keywords.Add(keyword);
            ed.WriteMessage($"\n  {keyword} ({cat.Value.Count} objects)");
        }

        pko.Keywords.Add("ALL");
        pko.Keywords.Add("LIST");
        pko.Keywords.Default = "LIST";

        ed.WriteMessage("\n  ALL (select all entities)");
        ed.WriteMessage("\n  LIST (show full list)\n");

        var result = ed.GetKeywords(pko);
        if (result.Status != PromptStatus.OK)
            return;

        if (result.StringResult == "LIST")
        {
            // Show full list
            ed.WriteMessage("\n=== All Categories ===\n");
            foreach (var cat in categories.OrderBy(c => c.Key))
            {
                ed.WriteMessage($"  {cat.Key}: {cat.Value.Count} objects\n");
            }

            // Ask again
            var pko2 = new PromptKeywordOptions("\nEnter category name:");
            foreach (var cat in categories.Keys)
            {
                pko2.Keywords.Add(cat.Replace(" ", ""));
            }

            var result2 = ed.GetKeywords(pko2);
            if (result2.Status != PromptStatus.OK)
                return;

            SelectCategoryByName(ed, categories, result2.StringResult);
        }
        else if (result.StringResult == "ALL")
        {
            // Select all entities
            var allIds = categories.SelectMany(c => c.Value).ToList();
            ed.SetImpliedSelectionEx(allIds.ToArray());
            ed.WriteMessage($"\n{allIds.Count} objects selected.\n");
        }
        else
        {
            // Select specific category
            SelectCategoryByName(ed, categories, result.StringResult);
        }
    }

    private void SelectCategoryByName(Editor ed, Dictionary<string, List<ObjectId>> categories, string keyword)
    {
        // Find matching category (handle spaces removed from keyword)
        var matchingCategory = categories.FirstOrDefault(c =>
            c.Key.Replace(" ", "").Equals(keyword, StringComparison.OrdinalIgnoreCase));

        if (matchingCategory.Value != null && matchingCategory.Value.Count > 0)
        {
            ed.SetImpliedSelectionEx(matchingCategory.Value.ToArray());
            ed.WriteMessage($"\n{matchingCategory.Value.Count} {matchingCategory.Key} objects selected.\n");
        }
        else
        {
            ed.WriteMessage($"\nCategory '{keyword}' not found.\n");
        }
    }

    private Dictionary<string, List<ObjectId>> GatherEntityCategories(Database db, SelectionModeManager.SelectionMode mode)
    {
        var categories = new Dictionary<string, List<ObjectId>>();

        switch (mode)
        {
            case SelectionModeManager.SelectionMode.SpaceLayout:
                GatherFromCurrentSpace(db, categories);
                break;

            case SelectionModeManager.SelectionMode.Drawing:
                GatherFromEntireDrawing(db, categories);
                break;

            case SelectionModeManager.SelectionMode.Process:
                GatherFromAllDocuments(categories);
                break;

            case SelectionModeManager.SelectionMode.Desktop:
            case SelectionModeManager.SelectionMode.Network:
                // For now, fall back to Process mode
                GatherFromAllDocuments(categories);
                break;
        }

        return categories;
    }

    private void GatherFromCurrentSpace(Database db, Dictionary<string, List<ObjectId>> categories)
    {
        using (var tr = db.TransactionManager.StartTransaction())
        {
            var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

            foreach (ObjectId id in currentSpace)
            {
                var entity = tr.GetObject(id, OpenMode.ForRead);
                string categoryName = GetEntityCategory(entity);

                if (!categories.ContainsKey(categoryName))
                    categories[categoryName] = new List<ObjectId>();

                categories[categoryName].Add(id);
            }

            tr.Commit();
        }
    }

    private void GatherFromEntireDrawing(Database db, Dictionary<string, List<ObjectId>> categories)
    {
        using (var tr = db.TransactionManager.StartTransaction())
        {
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead);
                    string categoryName = GetEntityCategory(entity);

                    if (!categories.ContainsKey(categoryName))
                        categories[categoryName] = new List<ObjectId>();

                    categories[categoryName].Add(id);
                }
            }

            tr.Commit();
        }
    }

    private void GatherFromAllDocuments(Dictionary<string, List<ObjectId>> categories)
    {
        var docMgr = AcadApp.DocumentManager;
        var currentDoc = docMgr.MdiActiveDocument;

        // For now, only gather from current document
        // Full implementation would require document switching
        GatherFromEntireDrawing(currentDoc.Database, categories);

        // Note: In a full implementation, you would iterate through all documents
        // and aggregate the results, but that requires more complex handling
    }

    private string GetEntityCategory(DBObject entity)
    {
        // Get the base type name
        string typeName = entity.GetType().Name;

        // Special handling for common types
        if (entity is BlockReference)
        {
            var br = entity as BlockReference;
            using (var tr = br.Database.TransactionManager.StartTransaction())
            {
                var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (btr.IsFromExternalReference)
                    return "XRef";
                else if (btr.IsAnonymous)
                    return "Dynamic Block";
                else
                    return "Block Reference";
            }
        }
        else if (entity is Dimension)
        {
            // More specific dimension types
            if (entity is AlignedDimension)
                return "Aligned Dimension";
            else if (entity is RotatedDimension)
                return "Linear Dimension";
            else if (entity is RadialDimension)
                return "Radial Dimension";
            else if (entity is DiametricDimension)
                return "Diametric Dimension";
            else if (entity is OrdinateDimension)
                return "Ordinate Dimension";
            else if (entity is ArcDimension)
                return "Arc Dimension";
            else
                return "Dimension";
        }
        else if (entity is MText)
            return "MText";
        else if (entity is DBText)
            return "Text";
        else if (entity is Polyline)
            return "Polyline";
        else if (entity is Polyline2d)
            return "Polyline2D";
        else if (entity is Polyline3d)
            return "Polyline3D";
        else if (entity is Line)
            return "Line";
        else if (entity is Circle)
            return "Circle";
        else if (entity is Arc)
            return "Arc";
        else if (entity is Ellipse)
            return "Ellipse";
        else if (entity is Spline)
            return "Spline";
        else if (entity is Hatch)
            return "Hatch";
        else if (entity is Autodesk.AutoCAD.DatabaseServices.Region)
            return "Region";
        else if (entity is Solid3d)
            return "3D Solid";
        else if (entity is Surface)
            return "Surface";
        else if (entity is Leader)
            return "Leader";
        else if (entity is MLeader)
            return "Multileader";
        else if (entity is Table)
            return "Table";
        else if (entity is Viewport)
            return "Viewport";
        else if (entity is AttributeDefinition)
            return "Attribute Definition";
        else if (entity is AttributeReference)
            return "Attribute Reference";
        else if (entity is RasterImage)
            return "Raster Image";
        else if (entity is Wipeout)
            return "Wipeout";
        else if (entity is DBPoint)
            return "Point";
        else if (entity is Ray)
            return "Ray";
        else if (entity is Xline)
            return "Construction Line";
        else if (entity is SubDMesh)
            return "Mesh";
        else if (entity is Light)
            return "Light";
        else if (entity is Camera)
            return "Camera";
        else
        {
            // Return the class name without "Autodesk.AutoCAD." prefix
            return typeName;
        }
    }

    private string GetDxfName(string categoryName)
    {
        // Map friendly names to DXF names for common types
        var dxfMap = new Dictionary<string, string>
        {
            { "Line", "LINE" },
            { "Circle", "CIRCLE" },
            { "Arc", "ARC" },
            { "Polyline", "LWPOLYLINE" },
            { "Polyline2D", "POLYLINE" },
            { "Polyline3D", "POLYLINE" },
            { "Text", "TEXT" },
            { "MText", "MTEXT" },
            { "Block Reference", "INSERT" },
            { "XRef", "INSERT" },
            { "Dynamic Block", "INSERT" },
            { "Hatch", "HATCH" },
            { "Dimension", "DIMENSION" },
            { "Linear Dimension", "DIMENSION" },
            { "Aligned Dimension", "DIMENSION" },
            { "Radial Dimension", "DIMENSION" },
            { "Diametric Dimension", "DIMENSION" },
            { "Ordinate Dimension", "DIMENSION" },
            { "Leader", "LEADER" },
            { "Multileader", "MULTILEADER" },
            { "Spline", "SPLINE" },
            { "Ellipse", "ELLIPSE" },
            { "Point", "POINT" },
            { "3D Solid", "3DSOLID" },
            { "Region", "REGION" },
            { "Surface", "SURFACE" },
            { "Mesh", "MESH" },
            { "SubD Mesh", "MESH" },
            { "Table", "ACAD_TABLE" },
            { "Viewport", "VIEWPORT" },
            { "Attribute Definition", "ATTDEF" },
            { "Attribute Reference", "ATTRIB" },
            { "Raster Image", "IMAGE" },
            { "Wipeout", "WIPEOUT" },
            { "Ray", "RAY" },
            { "Construction Line", "XLINE" },
            { "Light", "LIGHT" },
            { "Camera", "CAMERA" }
        };

        return dxfMap.ContainsKey(categoryName) ? dxfMap[categoryName] : categoryName.ToUpper();
    }

    [CommandMethod("cat-info")]
    public void CategoryInfo()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        var currentMode = SelectionModeManager.CurrentMode;
        ed.WriteMessage($"\n=== Entity Categories (Mode: {currentMode}) ===\n");

        var categories = GatherEntityCategories(db, currentMode);

        if (categories.Count == 0)
        {
            ed.WriteMessage("\nNo entities found.\n");
            return;
        }

        // Calculate totals
        int totalEntities = categories.Sum(c => c.Value.Count);

        // Sort by count (descending)
        var sortedCategories = categories.OrderByDescending(c => c.Value.Count);

        ed.WriteMessage($"\nTotal: {totalEntities} entities in {categories.Count} categories\n");
        ed.WriteMessage("\n");

        // Display table
        ed.WriteMessage($"{"Category",-25} {"DXF Name",-15} {"Count",10} {"Percentage",12}\n");
        ed.WriteMessage(new string('-', 62) + "\n");

        foreach (var cat in sortedCategories)
        {
            double percentage = (double)cat.Value.Count / totalEntities * 100;
            string dxfName = GetDxfName(cat.Key);
            ed.WriteMessage($"{cat.Key,-25} {dxfName,-15} {cat.Value.Count,10} {percentage,11:F1}%\n");
        }

        ed.WriteMessage("\n");
        ed.WriteMessage("Commands:\n");
        ed.WriteMessage("  select-by-cat  - Select categories using DataGrid\n");
        ed.WriteMessage("  quick-select-cat  - Quick select by category\n");
        ed.WriteMessage("  switch-select-mode - Change selection mode\n");
    }
}
