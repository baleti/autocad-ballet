using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;
using System.IO;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADCommands.EditSelectedTextCommand))]

namespace AutoCADCommands
{
    public class EditSelectedTextCommand
    {
        [CommandMethod("edit-selected-text", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void EditSelectedText()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;
            var currentScope = SelectionScopeManager.CurrentScope;

            List<TextEntity> textEntities = new List<TextEntity>();

            if (currentScope == SelectionScopeManager.SelectionScope.view)
            {
                // Handle view scope: Use pickfirst set or prompt for selection
                var selResult = ed.SelectImplied();

                if (selResult.Status == PromptStatus.Error)
                {
                    var selectionOpts = new PromptSelectionOptions();
                    selectionOpts.MessageForAdding = "\nSelect text/mtext entities: ";
                    var typeFilter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Operator, "<OR"),
                        new TypedValue((int)DxfCode.Start, "TEXT"),
                        new TypedValue((int)DxfCode.Start, "MTEXT"),
                        new TypedValue((int)DxfCode.Operator, "OR>")
                    };
                    var filter = new SelectionFilter(typeFilter);
                    selResult = ed.GetSelection(selectionOpts, filter);
                }
                else if (selResult.Status == PromptStatus.OK)
                {
                    ed.SetImpliedSelection(new ObjectId[0]);
                }

                if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
                {
                    ed.WriteMessage("\nNo text entities selected.\n");
                    return;
                }

                // Process selection from current document
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var objId in selResult.Value.GetObjectIds())
                    {
                        var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        if (IsTextEntity(entity))
                        {
                            var textEntity = CreateTextEntity(entity, doc.Name, objId.Handle.Value);
                            if (textEntity != null)
                                textEntities.Add(textEntity);
                        }
                    }
                    tr.Commit();
                }
            }
            else
            {
                // Handle document, process, desktop, network scopes: Use stored selection
                var storedSelection = SelectionStorage.LoadSelection();
                if (storedSelection == null || storedSelection.Count == 0)
                {
                    ed.WriteMessage("\nNo stored selection found. Use 'select-by-categories' first or switch to 'view' scope.\n");
                    return;
                }

                // Filter to current document if in document scope
                if (currentScope == SelectionScopeManager.SelectionScope.document)
                {
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
                }

                if (storedSelection.Count == 0)
                {
                    ed.WriteMessage($"\nNo stored selection found for current scope '{currentScope}'.\n");
                    return;
                }

                // Process stored selection items
                ProcessStoredSelection(storedSelection, textEntities, doc);
            }

            if (textEntities.Count == 0)
            {
                ed.WriteMessage("\nNo text/mtext entities found in selection.\n");
                return;
            }

            // Show text editing dialog
            var originalTexts = textEntities.Select(t => t.Text).ToList();
            using (var editForm = new AdvancedEditDialog(originalTexts, null, "Edit Selected Text"))
            {
                if (editForm.ShowDialog() == WinForms.DialogResult.OK)
                {
                    ApplyTextEdits(textEntities, editForm, ed);
                }
                else
                {
                    ed.WriteMessage("\nText editing cancelled.\n");
                }
            }
        }

        private void ProcessStoredSelection(List<SelectionItem> storedSelection, List<TextEntity> textEntities, Document currentDoc)
        {
            var docs = AcadApp.DocumentManager;
            var currentDocPath = Path.GetFullPath(currentDoc.Name);

            foreach (var item in storedSelection)
            {
                try
                {
                    var itemDocPath = Path.GetFullPath(item.DocumentPath);

                    if (string.Equals(itemDocPath, currentDocPath, StringComparison.OrdinalIgnoreCase))
                    {
                        // Handle current document items
                        var handle = Convert.ToInt64(item.Handle, 16);
                        var objectId = currentDoc.Database.GetObjectId(false, new Handle(handle), 0);

                        using (var tr = currentDoc.Database.TransactionManager.StartTransaction())
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (IsTextEntity(entity))
                            {
                                var textEntity = CreateTextEntity(entity, currentDoc.Name, objectId.Handle.Value);
                                if (textEntity != null)
                                    textEntities.Add(textEntity);
                            }
                            tr.Commit();
                        }
                    }
                    else
                    {
                        // Handle external document items
                        var targetDoc = FindDocumentByPath(itemDocPath);
                        if (targetDoc != null)
                        {
                            var handle = Convert.ToInt64(item.Handle, 16);
                            var objectId = targetDoc.Database.GetObjectId(false, new Handle(handle), 0);

                            using (var tr = targetDoc.Database.TransactionManager.StartTransaction())
                            {
                                var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                                if (IsTextEntity(entity))
                                {
                                    var textEntity = CreateTextEntity(entity, targetDoc.Name, objectId.Handle.Value);
                                    if (textEntity != null)
                                        textEntities.Add(textEntity);
                                }
                                tr.Commit();
                            }
                        }
                    }
                }
                catch (System.Exception)
                {
                    // Skip invalid handles/documents
                    continue;
                }
            }
        }

        private Document FindDocumentByPath(string documentPath)
        {
            var docs = AcadApp.DocumentManager;
            foreach (Document doc in docs)
            {
                try
                {
                    if (string.Equals(Path.GetFullPath(doc.Name), documentPath, StringComparison.OrdinalIgnoreCase))
                        return doc;
                }
                catch
                {
                    if (string.Equals(doc.Name, documentPath, StringComparison.OrdinalIgnoreCase))
                        return doc;
                }
            }
            return null;
        }

        private bool IsTextEntity(Entity entity)
        {
            return entity is DBText || entity is MText;
        }

        private TextEntity CreateTextEntity(Entity entity, string documentPath, long handle)
        {
            if (entity is DBText dbText)
            {
                return new TextEntity
                {
                    DocumentPath = documentPath,
                    Handle = handle,
                    Text = dbText.TextString,
                    EntityType = TextEntityType.DBText
                };
            }
            else if (entity is MText mtext)
            {
                return new TextEntity
                {
                    DocumentPath = documentPath,
                    Handle = handle,
                    Text = mtext.Contents,
                    EntityType = TextEntityType.MText
                };
            }
            return null;
        }

        private void ApplyTextEdits(List<TextEntity> textEntities, AdvancedEditDialog editForm, Editor ed)
        {
            int modifiedCount = 0;
            var docs = AcadApp.DocumentManager;

            // Group by document for efficient processing
            var groupedEntities = textEntities.GroupBy(t => t.DocumentPath).ToList();

            foreach (var docGroup in groupedEntities)
            {
                var targetDoc = FindDocumentByPath(docGroup.Key);
                if (targetDoc == null) continue;

                // Apply transformations and collect results
                var transformedTexts = new List<string>();
                foreach (var textEntity in docGroup)
                {
                    var transformedText = ApplyTransformation(textEntity.Text, editForm);
                    transformedTexts.Add(transformedText);
                }

                // Apply changes to database with document lock
                using (var docLock = targetDoc.LockDocument())
                {
                    using (var tr = targetDoc.Database.TransactionManager.StartTransaction())
                    {
                        int index = 0;
                        foreach (var textEntity in docGroup)
                        {
                            var newText = transformedTexts[index++];
                            if (newText != textEntity.Text)
                            {
                                var objectId = targetDoc.Database.GetObjectId(false, new Handle(textEntity.Handle), 0);
                                var entity = tr.GetObject(objectId, OpenMode.ForWrite) as Entity;

                                if (entity is DBText dbText)
                                {
                                    dbText.TextString = newText;
                                    modifiedCount++;
                                }
                                else if (entity is MText mtext)
                                {
                                    mtext.Contents = newText;
                                    modifiedCount++;
                                }
                            }
                        }
                        tr.Commit();
                    }
                }
            }

            ed.WriteMessage($"\nText editing complete. Modified {modifiedCount} entities.\n");
        }

        private string ApplyTransformation(string originalText, AdvancedEditDialog editForm)
        {
            string result = originalText;

            // Use the built-in transformation logic from AdvancedEditDialog
            // This ensures consistency with the DataGrid functionality

            // 1. Pattern transformation (highest priority)
            if (!string.IsNullOrEmpty(editForm.PatternText))
            {
                result = editForm.PatternText;
                // Replace {} with current value
                result = result.Replace("{}", originalText);
            }
            // 2. Find/Replace transformation
            else if (!string.IsNullOrEmpty(editForm.FindText))
            {
                result = result.Replace(editForm.FindText, editForm.ReplaceText ?? "");
            }
            // 3. Math transformation (if result is numeric)
            else if (!string.IsNullOrEmpty(editForm.MathOperationText))
            {
                if (double.TryParse(result, out double numericValue))
                {
                    try
                    {
                        string mathExpression = editForm.MathOperationText.Replace("x", numericValue.ToString());
                        var dataTable = new System.Data.DataTable();
                        var computedValue = dataTable.Compute(mathExpression, null);
                        if (computedValue != DBNull.Value)
                        {
                            result = computedValue.ToString();
                        }
                    }
                    catch
                    {
                        // If math evaluation fails, keep original value
                    }
                }
            }

            return result;
        }

        // Helper classes
        public class TextEntity
        {
            public string DocumentPath { get; set; }
            public long Handle { get; set; }
            public string Text { get; set; }
            public TextEntityType EntityType { get; set; }
        }

        public enum TextEntityType
        {
            DBText,
            MText
        }
    }

}