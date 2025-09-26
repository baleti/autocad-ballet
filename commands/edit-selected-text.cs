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

            ed.WriteMessage($"\n=== DEBUG: Starting edit-selected-text command ===");
            ed.WriteMessage($"\nCurrent scope: {currentScope}");

            List<TextEntity> textEntities = new List<TextEntity>();

            if (currentScope == SelectionScopeManager.SelectionScope.view)
            {
                // Handle view scope: Use pickfirst set or prompt for selection
                var selResult = ed.SelectImplied();
                ed.WriteMessage($"\nDEBUG: SelectImplied status: {selResult.Status}");

                if (selResult.Status == PromptStatus.Error)
                {
                    ed.WriteMessage($"\nDEBUG: No implied selection, prompting user...");
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
                    ed.WriteMessage($"\nDEBUG: GetSelection status: {selResult.Status}");
                }
                else if (selResult.Status == PromptStatus.OK)
                {
                    ed.WriteMessage($"\nDEBUG: Found implied selection with {selResult.Value.Count} entities");
                    ed.SetImpliedSelection(new ObjectId[0]);
                }

                if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
                {
                    ed.WriteMessage("\nNo text entities selected.\n");
                    return;
                }

                ed.WriteMessage($"\nDEBUG: Processing {selResult.Value.Count} selected entities");

                // Process selection from current document
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var objId in selResult.Value.GetObjectIds())
                    {
                        var entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        ed.WriteMessage($"\nDEBUG: Entity type: {entity?.GetType().Name ?? "null"}");
                        if (IsTextEntity(entity))
                        {
                            ed.WriteMessage($"\nDEBUG: Found text entity, creating TextEntity...");
                            var textEntity = CreateTextEntity(entity, doc.Name, objId.Handle.Value);
                            if (textEntity != null)
                            {
                                ed.WriteMessage($"\nDEBUG: Created TextEntity with text: '{textEntity.Text}'");
                                textEntities.Add(textEntity);
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"\nDEBUG: Entity is not a text entity");
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

            ed.WriteMessage($"\nDEBUG: Found {textEntities.Count} text entities");

            if (textEntities.Count == 0)
            {
                ed.WriteMessage("\nNo text/mtext entities found in selection.\n");
                return;
            }

            // Show text editing dialog
            var originalTexts = textEntities.Select(t => t.Text).ToList();
            ed.WriteMessage($"\nDEBUG: Original texts: [{string.Join(", ", originalTexts.Select(t => $"'{t}'"))}]");

            using (var editForm = new AdvancedEditDialog(originalTexts, null, "Edit Selected Text"))
            {
                ed.WriteMessage($"\nDEBUG: Showing dialog...");
                var dialogResult = editForm.ShowDialog();
                ed.WriteMessage($"\nDEBUG: Dialog result: {dialogResult}");

                if (dialogResult == WinForms.DialogResult.OK)
                {
                    ed.WriteMessage($"\nDEBUG: Dialog OK - applying edits...");
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
                    // Compare full paths if documentPath appears to be a full path
                    if (Path.IsPathRooted(documentPath))
                    {
                        if (string.Equals(Path.GetFullPath(doc.Name), Path.GetFullPath(documentPath), StringComparison.OrdinalIgnoreCase))
                            return doc;
                    }
                    else
                    {
                        // Compare just the filename if documentPath is just a filename
                        if (string.Equals(Path.GetFileName(doc.Name), documentPath, StringComparison.OrdinalIgnoreCase))
                            return doc;
                    }

                    // Fallback: also try direct string comparison
                    if (string.Equals(doc.Name, documentPath, StringComparison.OrdinalIgnoreCase))
                        return doc;
                }
                catch
                {
                    // If path operations fail, try direct string comparison
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
                    Text = mtext.Text,
                    EntityType = TextEntityType.MText
                };
            }
            return null;
        }

        private void ApplyTextEdits(List<TextEntity> textEntities, AdvancedEditDialog editForm, Editor ed)
        {
            ed.WriteMessage($"\n=== DEBUG: ApplyTextEdits called with {textEntities.Count} entities ===");
            ed.WriteMessage($"\nDialog values - Pattern: '{editForm.PatternText}', Find: '{editForm.FindText}', Replace: '{editForm.ReplaceText}', Math: '{editForm.MathOperationText}'");

            int modifiedCount = 0;
            var docs = AcadApp.DocumentManager;

            // Group by document for efficient processing
            var groupedEntities = textEntities.GroupBy(t => t.DocumentPath).ToList();
            ed.WriteMessage($"\nDEBUG: Grouped into {groupedEntities.Count} document groups");

            foreach (var docGroup in groupedEntities)
            {
                ed.WriteMessage($"\nDEBUG: Processing document group: '{docGroup.Key}' with {docGroup.Count()} entities");
                var targetDoc = FindDocumentByPath(docGroup.Key);
                if (targetDoc == null)
                {
                    ed.WriteMessage($"\nDEBUG: ERROR - Could not find document for path: '{docGroup.Key}'");
                    continue;
                }
                ed.WriteMessage($"\nDEBUG: Found target document: '{targetDoc.Name}'");

                // Apply transformations and collect results
                var transformedTexts = new List<string>();
                ed.WriteMessage($"\nDEBUG: Applying transformations to {docGroup.Count()} entities...");
                foreach (var textEntity in docGroup)
                {
                    ed.WriteMessage($"\nDEBUG: Transforming text: '{textEntity.Text}'");
                    var transformedText = ApplyTransformation(textEntity.Text, editForm);
                    transformedTexts.Add(transformedText);
                    ed.WriteMessage($"\nDEBUG: Result: '{transformedText}'");
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

                            // Debug output
                            ed.WriteMessage($"\nProcessing entity - Original: '{textEntity.Text}', New: '{newText}'");

                            if (newText != textEntity.Text)
                            {
                                ed.WriteMessage($"\n  Text changed, applying update...");
                                var objectId = targetDoc.Database.GetObjectId(false, new Handle(textEntity.Handle), 0);
                                var entity = tr.GetObject(objectId, OpenMode.ForWrite) as Entity;

                                if (entity is DBText dbText)
                                {
                                    ed.WriteMessage($"\n  Updating DBText from '{dbText.TextString}' to '{newText}'");
                                    dbText.TextString = newText;
                                    modifiedCount++;
                                }
                                else if (entity is MText mtext)
                                {
                                    // Preserve formatting by replacing only text content, not formatting codes
                                    string formattedText = ReplaceTextPreservingFormatting(mtext.Contents, textEntity.Text, newText);
                                    ed.WriteMessage($"\n  Updating MText from '{mtext.Contents}' to '{formattedText}'");
                                    mtext.Contents = formattedText;
                                    modifiedCount++;
                                }
                            }
                            else
                            {
                                ed.WriteMessage($"\n  No change detected, skipping...");
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
            var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            string result = originalText;
            ed.WriteMessage($"\n    ApplyTransformation - Input: '{originalText}'");
            ed.WriteMessage($"\n    Dialog values - Pattern: '{editForm.PatternText}', Find: '{editForm.FindText}', Replace: '{editForm.ReplaceText}', Math: '{editForm.MathOperationText}'");

            // Use the same transformation logic as AdvancedEditDialog.TransformValue
            // Find/Replace and Math operations take precedence over Pattern operations

            // 1. Pattern transformation (applied first as base)
            if (!string.IsNullOrEmpty(editForm.PatternText))
            {
                result = editForm.PatternText;
                // Replace {} with original value
                result = result.Replace("{}", originalText);
                ed.WriteMessage($"\n    After Pattern: '{result}'");
            }

            // 2. Find/Replace transformation (takes precedence - applied to pattern result)
            if (!string.IsNullOrEmpty(editForm.FindText))
            {
                string beforeReplace = result;
                result = result.Replace(editForm.FindText, editForm.ReplaceText ?? "");
                ed.WriteMessage($"\n    After Find/Replace ('{editForm.FindText}' -> '{editForm.ReplaceText ?? ""}'): '{beforeReplace}' -> '{result}'");
            }

            // 3. Math transformation (takes precedence - applied to result of pattern + find/replace, if result is numeric)
            if (!string.IsNullOrEmpty(editForm.MathOperationText))
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
                            string beforeMath = result;
                            result = computedValue.ToString();
                            ed.WriteMessage($"\n    After Math ('{editForm.MathOperationText}'): '{beforeMath}' -> '{result}'");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n    Math evaluation failed: {ex.Message}");
                    }
                }
                else
                {
                    ed.WriteMessage($"\n    Math skipped - '{result}' is not numeric");
                }
            }

            ed.WriteMessage($"\n    ApplyTransformation - Final result: '{result}'");
            return result;
        }

        private string ReplaceTextPreservingFormatting(string originalContents, string oldText, string newText)
        {
            // If the old text and new text are identical, no change needed
            if (oldText == newText)
            {
                return originalContents;
            }

            // If no formatting codes present, simple replacement
            if (!originalContents.Contains("\\"))
            {
                return originalContents.Replace(oldText, newText);
            }

            // SAFE APPROACH: Only attempt formatting preservation for same-length replacements
            // For different lengths, fall back to simple replacement (may lose formatting but is safe)
            if (oldText.Length != newText.Length)
            {
                return originalContents.Replace(oldText, newText);
            }

            // Same length replacement - try to preserve formatting with character-by-character replacement
            // This handles most common formatting codes safely
            string result = originalContents;

            // Find the start of the actual text content (skip initial formatting)
            int contentStart = 0;
            while (contentStart < result.Length)
            {
                if (result[contentStart] == '\\')
                {
                    // Skip formatting code - find the end (either ';' or space or end of non-alphabetic chars)
                    contentStart++;
                    while (contentStart < result.Length &&
                           result[contentStart] != ';' &&
                           result[contentStart] != ' ' &&
                           result[contentStart] != '\\')
                    {
                        contentStart++;
                    }
                    if (contentStart < result.Length && result[contentStart] == ';')
                    {
                        contentStart++; // Skip the semicolon
                    }
                }
                else
                {
                    break; // Found start of actual text
                }
            }

            // Simple character replacement starting from content
            var chars = result.ToCharArray();
            bool inFormattingCode = false;
            int textCharIndex = 0;

            for (int i = contentStart; i < chars.Length; i++)
            {
                if (chars[i] == '\\' && !inFormattingCode)
                {
                    inFormattingCode = true;
                }
                else if (inFormattingCode && (chars[i] == ';' || chars[i] == ' '))
                {
                    inFormattingCode = false;
                    if (chars[i] == ';') continue; // Skip semicolon
                }

                if (!inFormattingCode && textCharIndex < oldText.Length && chars[i] == oldText[textCharIndex])
                {
                    chars[i] = newText[textCharIndex];
                    textCharIndex++;
                }
                else if (!inFormattingCode && textCharIndex > 0)
                {
                    textCharIndex = 0; // Reset if match breaks
                }
            }

            return new string(chars);
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