using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoCADBallet
{
    /// <summary>
    /// Represents a stored selection filter with a name and query string
    /// </summary>
    public class StoredSelectionFilter
    {
        public string Name { get; set; }
        public string QueryString { get; set; }
        public string SourceDocumentPath { get; set; }
    }

    /// <summary>
    /// Manages storage and retrieval of selection filters
    /// </summary>
    public static class SelectionFilterStorage
    {
        private static string GetStorageFilePath()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dirPath = Path.Combine(appDataPath, "autocad-ballet", "runtime");
            Directory.CreateDirectory(dirPath);
            return Path.Combine(dirPath, "selection-filters");
        }

        /// <summary>
        /// Loads all selection filters from storage
        /// </summary>
        public static List<StoredSelectionFilter> LoadFilters()
        {
            var filters = new List<StoredSelectionFilter>();
            string filePath = GetStorageFilePath();

            if (!File.Exists(filePath))
            {
                return filters;
            }

            try
            {
                var lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // Format: Name|||QueryString|||SourceDocumentPath (v2)
                    // Legacy format: Name|||QueryString (v1)
                    var parts = line.Split(new[] { "|||" }, StringSplitOptions.None);
                    if (parts.Length >= 2)
                    {
                        filters.Add(new StoredSelectionFilter
                        {
                            Name = parts[0],
                            QueryString = parts[1],
                            SourceDocumentPath = parts.Length >= 3 ? parts[2] : null
                        });
                    }
                }
            }
            catch
            {
                // Return empty list if file can't be read
            }

            return filters;
        }

        /// <summary>
        /// Saves all selection filters to storage
        /// </summary>
        public static void SaveFilters(List<StoredSelectionFilter> filters)
        {
            string filePath = GetStorageFilePath();

            try
            {
                var lines = filters.Select(f => $"{f.Name}|||{f.QueryString}|||{f.SourceDocumentPath ?? ""}");
                File.WriteAllLines(filePath, lines);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save selection filters: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Adds a new filter or updates existing one with the same name
        /// </summary>
        public static void AddOrUpdateFilter(string name, string queryString, string sourceDocumentPath = null)
        {
            var filters = LoadFilters();

            // Remove existing filter with same name (case-insensitive)
            filters.RemoveAll(f => string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase));

            // Add new filter
            filters.Add(new StoredSelectionFilter
            {
                Name = name,
                QueryString = queryString,
                SourceDocumentPath = sourceDocumentPath
            });

            SaveFilters(filters);
        }

        /// <summary>
        /// Deletes filters by name
        /// </summary>
        public static void DeleteFilters(List<string> names)
        {
            var filters = LoadFilters();

            // Remove filters with matching names (case-insensitive)
            filters.RemoveAll(f => names.Any(n => string.Equals(f.Name, n, StringComparison.OrdinalIgnoreCase)));

            SaveFilters(filters);
        }

        /// <summary>
        /// Deletes a single filter by name
        /// </summary>
        public static void DeleteFilter(string name)
        {
            DeleteFilters(new List<string> { name });
        }
    }
}
