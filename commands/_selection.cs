using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;

namespace AutoCADBallet
{
    public class SelectionItem
    {
        public string DocumentPath { get; set; }
        public string Handle { get; set; }
        public string SessionId { get; set; }
    }

    public static class SelectionStorage
    {
        private static string GetSessionId()
        {
            // Generate unique session identifier combining process ID and session ID
            int processId = System.Diagnostics.Process.GetCurrentProcess().Id;
            int sessionId = System.Diagnostics.Process.GetCurrentProcess().SessionId;
            return $"{processId}_{sessionId}";
        }

        private static string GetSelectionFilePath(string documentName = null)
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var runtimeDir = Path.Combine(appDataPath, "autocad-ballet", "runtime");

            if (!string.IsNullOrEmpty(documentName))
            {
                // Create per-document selection file in subdirectory
                var selectionDir = Path.Combine(runtimeDir, "selection");
                if (!Directory.Exists(selectionDir))
                    Directory.CreateDirectory(selectionDir);

                var fileName = Path.GetFileNameWithoutExtension(documentName);
                // Sanitize filename to avoid invalid characters
                fileName = string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));
                return Path.Combine(selectionDir, fileName);
            }

            // Legacy path - EXACT same path as original for backward compatibility
            if (!Directory.Exists(runtimeDir))
                Directory.CreateDirectory(runtimeDir);
            return Path.Combine(runtimeDir, "selection");
        }

        public static void SaveSelection(List<SelectionItem> items, string documentName = null)
        {
            if (items == null || items.Count == 0)
            {
                // For empty selections, we may want to clear the selection file
                if (!string.IsNullOrEmpty(documentName))
                {
                    var emptyFilePath = GetSelectionFilePath(documentName);
                    if (File.Exists(emptyFilePath))
                    {
                        File.Delete(emptyFilePath);
                    }
                }
                else
                {
                    // Legacy behavior for global selection
                    var filePath = GetSelectionFilePath();
                    File.WriteAllLines(filePath, new string[0]);
                }
                return;
            }

            // Group items by document path to create separate files
            var groupedByDocument = items.GroupBy(item => item.DocumentPath);

            foreach (var docGroup in groupedByDocument)
            {
                var docPath = docGroup.Key;
                var docName = !string.IsNullOrEmpty(documentName) ? documentName : Path.GetFileName(docPath);
                var filePath = GetSelectionFilePath(docName);
                var lines = new List<string>();

                // Group items by session within the document
                var groupedItems = docGroup.GroupBy(item => item.SessionId);

                foreach (var group in groupedItems)
                {
                    var sessionId = group.Key ?? GetSessionId();
                    // Format: "DocumentPath",SessionId (quote document path to handle special characters)
                    var quotedPath = docPath.Contains(",") || docPath.Contains("\"")
                        ? $"\"{docPath.Replace("\"", "\"\"")}\""
                        : docPath;
                    lines.Add($"{quotedPath},{sessionId}");

                    // Format: comma-separated handles
                    var handles = group.Select(item => item.Handle);
                    lines.Add(string.Join(",", handles));
                }

                File.WriteAllLines(filePath, lines);
            }
        }

        public static List<SelectionItem> LoadSelection(string documentName = null)
        {
            var items = new List<SelectionItem>();

            if (!string.IsNullOrEmpty(documentName))
            {
                // Load selection for specific document
                var filePath = GetSelectionFilePath(documentName);
                if (File.Exists(filePath))
                {
                    items.AddRange(LoadSelectionFromFile(filePath));
                }
            }
            else
            {
                // Backward compatibility: Load from legacy global file only
                // This preserves the exact old behavior for existing commands
                var legacyFilePath = GetSelectionFilePath();
                if (File.Exists(legacyFilePath))
                {
                    items.AddRange(LoadSelectionFromFile(legacyFilePath));
                }
            }

            return items;
        }

        // New method for explicitly loading from all documents (session scope)
        public static List<SelectionItem> LoadSelectionFromAllDocuments()
        {
            var items = new List<SelectionItem>();

            // Load from all per-document selection files
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var selectionDir = Path.Combine(appDataPath, "autocad-ballet", "runtime", "selection");
            if (Directory.Exists(selectionDir))
            {
                foreach (var file in Directory.GetFiles(selectionDir))
                {
                    items.AddRange(LoadSelectionFromFile(file));
                }
            }

            // Also check legacy global file for backward compatibility
            var legacyFilePath = GetSelectionFilePath();
            if (File.Exists(legacyFilePath))
            {
                items.AddRange(LoadSelectionFromFile(legacyFilePath));
            }

            return items;
        }

        private static List<SelectionItem> LoadSelectionFromFile(string filePath)
        {
            var items = new List<SelectionItem>();

            if (!File.Exists(filePath))
                return items;

            var lines = File.ReadAllLines(filePath);
            string currentDocumentPath = null;
            string currentSessionId = null;

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrEmpty(line))
                    continue;

                // Check if this line contains a document path (contains comma but not just handles)
                if (line.Contains(",") && !IsHandlesLine(line))
                {
                    // This is a document path line: DocumentPath,SessionId or "DocumentPath",SessionId
                    if (line.StartsWith("\""))
                    {
                        // Handle quoted path: "C:\path with spaces\file.dwg",sessionId
                        var quoteEnd = line.IndexOf("\"", 1);
                        if (quoteEnd > 1 && line.Length > quoteEnd + 1 && line[quoteEnd + 1] == ',')
                        {
                            currentDocumentPath = line.Substring(1, quoteEnd - 1).Replace("\"\"", "\""); // Unescape quotes
                            currentSessionId = line.Substring(quoteEnd + 2);
                        }
                    }
                    else
                    {
                        // Handle unquoted path: DocumentPath,SessionId
                        var parts = line.Split(',');
                        if (parts.Length >= 2)
                        {
                            currentDocumentPath = parts[0];
                            currentSessionId = parts[1];
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(currentDocumentPath))
                {
                    // This is a handles line: comma-separated handles
                    var handles = line.Split(',');
                    foreach (var handle in handles)
                    {
                        if (!string.IsNullOrWhiteSpace(handle))
                        {
                            items.Add(new SelectionItem
                            {
                                DocumentPath = currentDocumentPath,
                                SessionId = currentSessionId,
                                Handle = handle.Trim()
                            });
                        }
                    }
                }
            }

            return items;
        }

        private static bool IsHandlesLine(string line)
        {
            // Handles are typically hexadecimal strings
            // If all comma-separated parts look like handles (hex), it's a handles line
            var parts = line.Split(',');
            if (parts.Length == 0)
                return false;

            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (string.IsNullOrEmpty(trimmed))
                    continue;

                // Check if it looks like a hex handle (only contains hex characters)
                if (!IsHexString(trimmed))
                    return false;
            }
            return true;
        }

        private static bool IsHexString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return false;

            foreach (char c in input)
            {
                if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')))
                    return false;
            }
            return true;
        }
    }

    // Extension method for Editor to support legacy SetImpliedSelectionEx calls
    public static class EditorExtensions
    {
        public static void SetImpliedSelectionEx(this Editor ed, ObjectId[] objectIds)
        {
            // Simply delegate to the standard SetImpliedSelection
            ed.SetImpliedSelection(objectIds);
        }
    }
}
