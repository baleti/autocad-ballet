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

// Register the command class
[assembly: CommandClass(typeof(AutoCADBallet.SelectLayoutsInSession))]

namespace AutoCADBallet
{
    public class SelectLayoutsInSession
    {
        [CommandMethod("select-layouts-in-session", CommandFlags.Modal)]
        public void SelectLayoutsInSessionCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                var selectionItems = new List<SelectionItem>();

                // Iterate through all open documents
                foreach (Document openDoc in AcadApp.DocumentManager)
                {
                    try
                    {
                        var db = openDoc.Database;
                        var docPath = openDoc.Name;

                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            // Get layout dictionary
                            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                            // Iterate through all layouts
                            foreach (DBDictionaryEntry entry in layoutDict)
                            {
                                try
                                {
                                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);

                                    // Add layout to selection
                                    selectionItems.Add(new SelectionItem
                                    {
                                        DocumentPath = docPath,
                                        Handle = layout.Handle.ToString(),
                                        SessionId = null // Will be auto-generated
                                    });
                                }
                                catch
                                {
                                    // Skip problematic layouts
                                    continue;
                                }
                            }

                            tr.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError processing document {openDoc.Name}: {ex.Message}\n");
                        continue;
                    }
                }

                if (selectionItems.Count == 0)
                {
                    ed.WriteMessage("\nNo layouts found in open documents.\n");
                    return;
                }

                // Save the selection
                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage($"\nSelected {selectionItems.Count} layouts across all open documents.\n");
                ed.WriteMessage("Use 'filter-selection-in-session' to view and edit layout names.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n");
            }
        }
    }
}
