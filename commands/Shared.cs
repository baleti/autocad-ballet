using System;
using System.Collections.Generic;
using System.IO;

namespace AutoCADBallet
{
    public class SelectionItem
    {
        public string DocumentPath { get; set; }
        public string Handle { get; set; }
    }

    public static class SelectionStorage
    {
        private static string GetSelectionFilePath()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(appDataPath, "autocad-ballet");
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return Path.Combine(dir, "selection");
        }

        public static void SaveSelection(List<SelectionItem> items)
        {
            var filePath = GetSelectionFilePath();
            var lines = new List<string>();

            foreach (var item in items)
            {
                // Format: DocumentPath|Handle
                lines.Add($"{item.DocumentPath}|{item.Handle}");
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
            foreach (var line in lines)
            {
                var parts = line.Split('|');
                if (parts.Length == 2)
                {
                    items.Add(new SelectionItem
                    {
                        DocumentPath = parts[0],
                        Handle = parts[1]
                    });
                }
            }

            return items;
        }
    }
}
