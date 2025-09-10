using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;

namespace AutoCADBallet
{
    public class SelectAllInOpenedDrawingsCommand
    {
        [CommandMethod("SelectAllInOpenedDrawings")]
        public static void SelectAllInOpenedDrawings()
        {
            var docManager = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            var editor = docManager.MdiActiveDocument.Editor;
            var selectionItems = new List<SelectionItem>();

            try
            {
                editor.WriteMessage("\nCollecting entities from all open drawings...\n");

                foreach (Document doc in docManager)
                {
                    int docEntityCount = 0;

                    using (doc.LockDocument())
                    {
                        using (var tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            // Get all entities from model space
                            var modelSpace = tr.GetObject(
                                SymbolUtilityServices.GetBlockModelSpaceId(doc.Database),
                                OpenMode.ForRead) as BlockTableRecord;

                            foreach (ObjectId id in modelSpace)
                            {
                                try
                                {
                                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                    if (entity != null && entity.Visible)
                                    {
                                        selectionItems.Add(new SelectionItem
                                        {
                                            DocumentPath = doc.Name,
                                            Handle = entity.Handle.ToString()
                                        });
                                        docEntityCount++;
                                    }
                                }
                                catch
                                {
                                    // Skip entities that can't be accessed
                                }
                            }

                            // Get all entities from paper space layouts
                            var layoutDict = tr.GetObject(doc.Database.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                            foreach (DBDictionaryEntry entry in layoutDict)
                            {
                                try
                                {
                                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                    if (layout != null && layout.LayoutName != "Model")
                                    {
                                        var paperSpace = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                                        foreach (ObjectId id in paperSpace)
                                        {
                                            try
                                            {
                                                var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                                                if (entity != null && entity.Visible)
                                                {
                                                    selectionItems.Add(new SelectionItem
                                                    {
                                                        DocumentPath = doc.Name,
                                                        Handle = entity.Handle.ToString()
                                                    });
                                                    docEntityCount++;
                                                }
                                            }
                                            catch
                                            {
                                                // Skip entities that can't be accessed
                                            }
                                        }
                                    }
                                }
                                catch
                                {
                                    // Skip layouts that can't be accessed
                                }
                            }

                            tr.Commit();
                        }
                    }

                    editor.WriteMessage($"  {Path.GetFileName(doc.Name)}: {docEntityCount} entities\n");
                }

                // Save selection
                SelectionStorage.SaveSelection(selectionItems);

                editor.WriteMessage($"\n==============================================");
                editor.WriteMessage($"\nTotal: {selectionItems.Count} entities from {docManager.Count} drawings");
                editor.WriteMessage($"\nSelection saved to: %appdata%\\autocad-ballet\\selection");
                editor.WriteMessage($"\n==============================================\n");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError: {ex.Message}\n");
                editor.WriteMessage($"Stack trace: {ex.StackTrace}\n");
            }
        }
    }
}
