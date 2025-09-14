using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

        private static string GetSelectionFilePath()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(appDataPath, "autocad-ballet", "runtime");
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return Path.Combine(dir, "selection");
        }

        public static void SaveSelection(List<SelectionItem> items)
        {
            var filePath = GetSelectionFilePath();
            var lines = new List<string>();

            // Group items by document path and session
            var groupedItems = items.GroupBy(item => new { item.DocumentPath, item.SessionId });

            foreach (var group in groupedItems)
            {
                var sessionId = group.Key.SessionId ?? GetSessionId();
                // Format: "DocumentPath",SessionId (quote document path to handle special characters)
                var quotedPath = group.Key.DocumentPath.Contains(",") || group.Key.DocumentPath.Contains("\"")
                    ? $"\"{group.Key.DocumentPath.Replace("\"", "\"\"")}\""
                    : group.Key.DocumentPath;
                lines.Add($"{quotedPath},{sessionId}");

                // Format: comma-separated handles
                var handles = group.Select(item => item.Handle);
                lines.Add(string.Join(",", handles));
            }

            File.WriteAllLines(filePath, lines);
        }

        public static List<SelectionItem> LoadSelection()
        {
            var filePath = GetSelectionFilePath();
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

}
