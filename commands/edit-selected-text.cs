using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using AutoCADCommands;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;
using System.IO;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADCommands
{
    public class EditSelectedText
    {
        public static void ExecuteViewScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;

            List<TextEntity> textEntities = new List<TextEntity>();

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
                        {
                            textEntities.Add(textEntity);
                        }
                    }
                }
                tr.Commit();
            }

            ShowTextEditingDialog(textEntities, ed);
        }

        public static void ExecuteDocumentScope(Editor ed, Database db)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var docName = Path.GetFileName(doc.Name);
            var storedSelection = SelectionStorage.LoadSelection(docName);
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage($"\nNo stored selection found for current document '{docName}'. Use commands like 'select-by-categories-in-document' first.\n");
                return;
            }

            List<TextEntity> textEntities = new List<TextEntity>();

            // Process stored selection items from current document only
            foreach (var item in storedSelection)
            {
                try
                {
                    var handle = Convert.ToInt64(item.Handle, 16);
                    var objectId = db.GetObjectId(false, new Handle(handle), 0);

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                        if (IsTextEntity(entity))
                        {
                            var textEntity = CreateTextEntity(entity, doc.Name, objectId.Handle.Value);
                            if (textEntity != null)
                                textEntities.Add(textEntity);
                        }
                        tr.Commit();
                    }
                }
                catch (System.Exception)
                {
                    // Skip invalid handles/documents
                    continue;
                }
            }

            ShowTextEditingDialog(textEntities, ed);
        }

        public static void ExecuteApplicationScope(Editor ed)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var storedSelection = SelectionStorage.LoadSelectionFromAllDocuments();
            if (storedSelection == null || storedSelection.Count == 0)
            {
                ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-categories-in-session' first.\n");
                return;
            }

            List<TextEntity> textEntities = new List<TextEntity>();

            // Process stored selection items
            ProcessStoredSelection(storedSelection, textEntities, doc);

            ShowTextEditingDialog(textEntities, ed);
        }

        private static void ShowTextEditingDialog(List<TextEntity> textEntities, Editor ed)
        {
            if (textEntities.Count == 0)
            {
                ed.WriteMessage("\nNo text/mtext entities found in selection.\n");
                return;
            }

            // Show text editing dialog
            var originalTexts = textEntities.Select(t => t.Text).ToList();

            using (var editForm = new AdvancedEditDialog(originalTexts, null, "Edit Selected Text"))
            {
                var dialogResult = editForm.ShowDialog();

                if (dialogResult == WinForms.DialogResult.OK)
                {
                    ApplyTextEdits(textEntities, editForm, ed);
                }
                else
                {
                    ed.WriteMessage("\nText editing cancelled.\n");
                }
            }
        }

        private static void ProcessStoredSelection(List<SelectionItem> storedSelection, List<TextEntity> textEntities, Document currentDoc)
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

        private static Document FindDocumentByPath(string documentPath)
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

        private static bool IsTextEntity(Entity entity)
        {
            return entity is DBText || entity is MText;
        }

        private static TextEntity CreateTextEntity(Entity entity, string documentPath, long handle)
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
                // For multiline MText, show the raw Contents (with formatting codes) instead of plain Text
                // This preserves all formatting and lets users edit the actual AutoCAD formatting
                string textContent = mtext.Text.Contains("\n") ? mtext.Contents : mtext.Text;

                return new TextEntity
                {
                    DocumentPath = documentPath,
                    Handle = handle,
                    Text = textContent,
                    EntityType = TextEntityType.MText
                };
            }
            return null;
        }

        private static void ApplyTextEdits(List<TextEntity> textEntities, AdvancedEditDialog editForm, Editor ed)
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
                                    ed.WriteMessage($"\n  Updating MText from '{mtext.Contents}' to '{newText}'");
                                    // Direct contents replacement - newText already contains proper formatting codes
                                    mtext.Contents = newText;
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

        private static string ApplyTransformation(string originalText, AdvancedEditDialog editForm)
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

            // 3. Math transformation (takes precedence - applied to result of pattern + find/replace)
            if (!string.IsNullOrEmpty(editForm.MathOperationText))
            {
                string beforeMath = result;
                result = ApplyMathToAlphanumeric(result, editForm.MathOperationText);
                ed.WriteMessage($"\n    After Math ('{editForm.MathOperationText}'): '{beforeMath}' -> '{result}'");
            }

            ed.WriteMessage($"\n    ApplyTransformation - Final result: '{result}'");
            return result;
        }

        /// <summary>
        /// Apply math operations to alphanumeric strings by extracting and modifying numeric parts.
        /// Examples: "W1" + "x+3" → "W4", "Room12" + "x*2" → "Room24"
        /// </summary>
        private static string ApplyMathToAlphanumeric(string input, string mathExpression)
        {
            var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // First try pure numeric approach (existing behavior)
                if (double.TryParse(input, out double numericValue))
                {
                    string expr = mathExpression.Replace("x", numericValue.ToString());
                    var dataTable = new System.Data.DataTable();
                    var computedValue = dataTable.Compute(expr, null);
                    if (computedValue != DBNull.Value)
                    {
                        ed.WriteMessage($"\n      Math on pure number: {numericValue} -> {computedValue}");
                        return computedValue.ToString();
                    }
                    return input;
                }

                // Handle alphanumeric strings - extract numeric parts
                var matches = System.Text.RegularExpressions.Regex.Matches(input, @"\d+");
                if (matches.Count == 0)
                {
                    ed.WriteMessage($"\n      Math skipped - no numbers found in '{input}'");
                    return input;
                }

                ed.WriteMessage($"\n      Found {matches.Count} numeric parts in '{input}'");

                // Apply math to each numeric part
                string result = input;
                for (int i = matches.Count - 1; i >= 0; i--) // Process in reverse to maintain positions
                {
                    var match = matches[i];
                    if (double.TryParse(match.Value, out double numberValue))
                    {
                        string expr = mathExpression.Replace("x", numberValue.ToString());
                        var dataTable = new System.Data.DataTable();
                        var computedValue = dataTable.Compute(expr, null);
                        if (computedValue != DBNull.Value)
                        {
                            // Format the result to remove unnecessary decimal places for integers
                            string newNumberStr = FormatNumericResult(computedValue);
                            ed.WriteMessage($"\n      Math on part '{match.Value}': {numberValue} -> {computedValue} -> '{newNumberStr}'");
                            result = result.Substring(0, match.Index) + newNumberStr + result.Substring(match.Index + match.Length);
                        }
                    }
                }
                return result;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n      Math evaluation failed: {ex.Message}");
                return input; // If anything fails, return original value
            }
        }

        /// <summary>
        /// Format numeric result to avoid unnecessary decimal places for whole numbers
        /// </summary>
        private static string FormatNumericResult(object computedValue)
        {
            if (computedValue is double doubleVal)
            {
                // If it's a whole number, format as integer
                if (doubleVal == Math.Floor(doubleVal))
                {
                    return ((long)doubleVal).ToString();
                }
                return doubleVal.ToString();
            }
            return computedValue.ToString();
        }

        private static string ReplaceTextPreservingFormatting(string originalContents, string oldText, string newText)
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