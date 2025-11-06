using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoCADBallet
{
    public class TransientGraphicsGroup
    {
        public string Description { get; set; } // Free-form description like "viewport: Work, scale: 1:20, marker count: 5"
        public List<int> Markers { get; set; }

        public TransientGraphicsGroup()
        {
            Markers = new List<int>();
        }
    }

    public static class TransientGraphicsStorage
    {
        private static string GetStorageDir()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(appData, "autocad-ballet", "runtime", "transient-graphics-tracker");
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            return dir;
        }

        private static string GetStorageFile(string documentId)
        {
            return Path.Combine(GetStorageDir(), documentId);
        }

        public static void SaveGroups(string documentId, List<TransientGraphicsGroup> groups)
        {
            try
            {
                var filePath = GetStorageFile(documentId);
                var lines = new List<string>();

                foreach (var group in groups)
                {
                    // Format: Description|marker1,marker2,marker3
                    var markerString = string.Join(",", group.Markers);
                    lines.Add($"{group.Description}|{markerString}");
                }

                File.WriteAllLines(filePath, lines);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to save groups: {ex.Message}");
            }
        }

        public static List<TransientGraphicsGroup> LoadGroups(string documentId)
        {
            var groups = new List<TransientGraphicsGroup>();
            try
            {
                var filePath = GetStorageFile(documentId);
                if (File.Exists(filePath))
                {
                    var lines = File.ReadAllLines(filePath);
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var parts = line.Split('|');
                        if (parts.Length == 2)
                        {
                            var group = new TransientGraphicsGroup
                            {
                                Description = parts[0],
                                Markers = parts[1].Split(',')
                                    .Where(m => !string.IsNullOrWhiteSpace(m))
                                    .Select(m => int.Parse(m.Trim()))
                                    .ToList()
                            };
                            groups.Add(group);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load groups: {ex.Message}");
            }
            return groups;
        }

        public static void ClearGroups(string documentId)
        {
            try
            {
                var filePath = GetStorageFile(documentId);
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to clear groups: {ex.Message}");
            }
        }
    }
}
