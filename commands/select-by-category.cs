using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

// Register the command class
[assembly: CommandClass(typeof(SelectByCategory))]

public class SelectByCategory
{
    // Simple class to store entity references for process scope
    public class EntityReference
    {
        public string DocumentPath { get; set; }
        public string DocumentName { get; set; }
        public string Handle { get; set; }
        public string Category { get; set; }
        public string SpaceName { get; set; }
    }

    [CommandMethod("select-by-category")]
    public void SelectByCategoryCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        var currentMode = SelectionScopeManager.CurrentScope;

        if (currentMode == SelectionScopeManager.SelectionScope.process)
        {
            HandleProcessMode(ed);
        }
        else
        {
            HandleNormalModes(db, ed, currentMode);
        }
    }

    private void HandleProcessMode(Editor ed)
    {
        ed.WriteMessage("\nProcess Mode: Gathering entities from all open documents...\n");

        // Gather from all open documents
        var docManager = AcadApp.DocumentManager;
        var allReferences = new List<EntityReference>();
        var categoryGroups = new Dictionary<string, List<EntityReference>>();

        foreach (Autodesk.AutoCAD.ApplicationServices.Document doc in docManager)
        {
            string docPath = doc.Name;
            string docName = Path.GetFileName(docPath);

            ed.WriteMessage($"\nScanning: {docName}...");

            try
            {
                // Read from document database without switching
                var refs = GatherEntityReferencesFromDocument(doc.Database, docPath, docName);
                allReferences.AddRange(refs);

                // Also gather layout references
                var layoutRefs = GatherLayoutReferencesFromDocument(doc.Database, docPath, docName);
                allReferences.AddRange(layoutRefs);

                // Group by category (including layouts)
                foreach (var entityRef in refs.Concat(layoutRefs))
                {
                    if (!categoryGroups.ContainsKey(entityRef.Category))
                        categoryGroups[entityRef.Category] = new List<EntityReference>();

                    categoryGroups[entityRef.Category].Add(entityRef);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n  Error reading {docName}: {ex.Message}");
            }
        }

        if (categoryGroups.Count == 0)
        {
            ed.WriteMessage("\nNo entities found across open documents.\n");
            return;
        }

        // Prepare summary for DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var cat in categoryGroups.OrderBy(c => c.Key))
        {
            // Count entities per document
            var docCounts = cat.Value.GroupBy(e => e.DocumentName)
                                     .Select(g => $"{g.Key}: {g.Count()}")
                                     .ToList();

            entries.Add(new Dictionary<string, object>
            {
                { "Category", cat.Key },
                { "Total Count", cat.Value.Count },
                { "Documents", string.Join(", ", docCounts) }
            });
        }

        var propertyNames = new List<string> { "Category", "Total Count", "Documents" };

        // Show DataGrid for selection
        ed.WriteMessage("\nSelect categories to include in process selection...\n");
        var selectedCategories = CustomGUIs.DataGrid(entries, propertyNames, false);

        if (selectedCategories == null || selectedCategories.Count == 0)
        {
            ed.WriteMessage("\nNo categories selected.\n");
            return;
        }

        // Filter entities to only selected categories
        var selectedCategoryNames = selectedCategories.Select(s => s["Category"].ToString()).ToList();
        var selectedEntities = new List<EntityReference>();
        var categoryCounts = new Dictionary<string, int>();

        foreach (var catName in selectedCategoryNames)
        {
            if (categoryGroups.ContainsKey(catName))
            {
                selectedEntities.AddRange(categoryGroups[catName]);
                categoryCounts[catName] = categoryGroups[catName].Count;
            }
        }

        // Convert to unified SelectionItem format and save using SelectionStorage
        var selectionItems = new List<AutoCADBallet.SelectionItem>();
        foreach (var entityRef in selectedEntities)
        {
            selectionItems.Add(new AutoCADBallet.SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null // Will be auto-generated
            });
        }

        try
        {
            // Save using unified selection storage
            AutoCADBallet.SelectionStorage.SaveSelection(selectionItems);

            ed.WriteMessage($"\nProcess selection saved to unified selection storage.\n");
            ed.WriteMessage($"\nSummary:\n");
            ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
            ed.WriteMessage($"  Categories: {selectedCategoryNames.Count}\n");
            ed.WriteMessage($"  Documents: {selectedEntities.Select(e => e.DocumentName).Distinct().Count()}\n");

            foreach (var cat in categoryCounts)
            {
                ed.WriteMessage($"    {cat.Key}: {cat.Value} entities\n");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
        }
    }

    private List<EntityReference> GatherEntityReferencesFromDocument(Database db, string docPath, string docName)
    {
        var references = new List<EntityReference>();

        // Important: Use a separate transaction for each external database
        using (var tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    string spaceName = layout.LayoutName;

                    foreach (ObjectId id in btr)
                    {
                        try
                        {
                            var entity = tr.GetObject(id, OpenMode.ForRead);

                            // Skip certain system entities
                            if (entity is Viewport && spaceName != "Model")
                            {
                                var vp = entity as Viewport;
                                if (vp.Number == 1) // Paper space viewport
                                    continue;
                            }

                            string category = GetEntityCategory(entity);

                            var entityRef = new EntityReference
                            {
                                DocumentPath = docPath,
                                DocumentName = docName,
                                Handle = entity.Handle.ToString(),
                                Category = category,
                                SpaceName = spaceName
                            };

                            references.Add(entityRef);
                        }
                        catch
                        {
                            // Skip entities that can't be read
                            continue;
                        }
                    }
                }

                tr.Commit();
            }
            catch (System.Exception)
            {
                tr.Abort();
                throw;
            }
        }

        return references;
    }

    private List<EntityReference> GatherLayoutReferencesFromDocument(Database db, string docPath, string docName)
    {
        var references = new List<EntityReference>();

        using (var tr = db.TransactionManager.StartTransaction())
        {
            try
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                    var layoutRef = new EntityReference
                    {
                        DocumentPath = docPath,
                        DocumentName = docName,
                        Handle = layout.Handle.ToString(),
                        Category = "Layout",
                        SpaceName = layout.LayoutName
                    };

                    references.Add(layoutRef);
                }

                tr.Commit();
            }
            catch (System.Exception)
            {
                tr.Abort();
                throw;
            }
        }

        return references;
    }

    private void HandleNormalModes(Database db, Editor ed, SelectionScopeManager.SelectionScope scope)
    {
        // Gather entity categories based on current scope
        var categories = GatherEntityCategories(db, scope);

        if (categories.Count == 0)
        {
            ed.WriteMessage("\nNo entities found in current scope.\n");
            return;
        }

        // Prepare data for DataGrid
        var entries = new List<Dictionary<string, object>>();
        foreach (var cat in categories.OrderBy(c => c.Key))
        {
            entries.Add(new Dictionary<string, object>
            {
                { "Category", cat.Key },
                { "Count", cat.Value.Count }
            });
        }

        var propertyNames = new List<string> { "Category", "Count" };

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

    private Dictionary<string, List<ObjectId>> GatherEntityCategories(Database db, SelectionScopeManager.SelectionScope scope)
    {
        var categories = new Dictionary<string, List<ObjectId>>();

        switch (scope)
        {
            case SelectionScopeManager.SelectionScope.view:
                GatherFromCurrentSpace(db, categories);
                break;

            case SelectionScopeManager.SelectionScope.document:
                GatherFromEntireDrawing(db, categories);
                GatherLayouts(db, categories);
                break;

            case SelectionScopeManager.SelectionScope.desktop:
            case SelectionScopeManager.SelectionScope.network:
                // For now, fall back to Document scope
                GatherFromEntireDrawing(db, categories);
                GatherLayouts(db, categories);
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
                var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;

                // Skip entities that are hidden (not visible) in current view
                if (entity != null && !entity.Visible)
                    continue;

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

    private void GatherLayouts(Database db, Dictionary<string, List<ObjectId>> categories)
    {
        using (var tr = db.TransactionManager.StartTransaction())
        {
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
            var layoutIds = new List<ObjectId>();

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                layoutIds.Add(entry.Value);
            }

            if (layoutIds.Count > 0)
            {
                categories["Layout"] = layoutIds;
            }

            tr.Commit();
        }
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
            else if (entity is RadialDimensionLarge)
                return "Jogged Dimension";
            else
                return "Dimension";
        }
        // Text entities
        else if (entity is MText)
            return "MText";
        else if (entity is DBText)
            return "Text";
        // Polyline variants
        else if (entity is Polyline)
            return "Polyline";
        else if (entity is Polyline2d)
            return "Polyline2D";
        else if (entity is Polyline3d)
            return "Polyline3D";
        // Basic geometry
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
        // Pattern fills
        else if (entity is Hatch)
            return "Hatch";
        else if (entity is Solid)
            return "2D Solid";
        else if (entity is Trace)
            return "Trace";
        // 3D entities
        else if (entity is Autodesk.AutoCAD.DatabaseServices.Region)
            return "Region";
        else if (entity is Solid3d)
            return "3D Solid";
        else if (entity is Surface)
            return "Surface";
        else if (entity is Body)
            return "Body";
        else if (entity is PolygonMesh)
            return "Polygon Mesh";
        else if (entity is SubDMesh)
            return "SubD Mesh";
        // Annotation entities
        else if (entity is Leader)
            return "Leader";
        else if (entity is MLeader)
            return "Multileader";
        else if (entity is Table)
            return "Table";
        // Layout entities
        else if (entity is Viewport)
            return "Viewport";
        // Attributes
        else if (entity is AttributeDefinition)
            return "Attribute Definition";
        else if (entity is AttributeReference)
            return "Attribute Reference";
        // Images and underlays
        else if (entity is RasterImage)
            return "Raster Image";
        else if (entity is Wipeout)
            return "Wipeout";
        else if (entity is Ole2Frame)
            return "OLE Object";
        // Construction geometry
        else if (entity is DBPoint)
            return "Point";
        else if (entity is Ray)
            return "Ray";
        else if (entity is Xline)
            return "Construction Line";
        // Visualization entities
        else if (entity is Light)
            return "Light";
        else if (entity is Camera)
            return "Camera";
        else if (entity is RenderEnvironment)
            return "Render Environment";
        else if (entity is Sun)
            return "Sun";
        // Section entities
        else if (entity is Section)
            return "Section";
        else if (entity is SectionSettings)
            return "Section Settings";
        // Proxy entities
        else if (entity is ProxyEntity)
            return "Proxy Entity";
        // Field entities
        else if (entity is Field)
            return "Field";
        // Helix
        else if (entity is Helix)
            return "Helix";
        // ACIS entities
        else if (entity is SequenceEnd)
            return "Sequence End";
        // Group
        else if (entity is Group)
            return "Group";
        // Mline
        else if (entity is Mline)
            return "Multiline";
        // Shape
        else if (entity is Shape)
            return "Shape";
        else
        {
            // Return the class name without "Autodesk.AutoCAD." prefix
            if (typeName.StartsWith("Autodesk.AutoCAD."))
            {
                typeName = typeName.Replace("Autodesk.AutoCAD.", "");
            }
            return typeName;
        }
    }

}
