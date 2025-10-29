using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCADBallet
{
    /// <summary>
    /// Type-safe, extensible attribute storage for AutoCAD entities
    /// Stores arbitrary key-value attributes in entity extension dictionaries
    /// Supports multiple data types with compile-time safety
    /// </summary>
    public static class EntityAttributes
    {
        private const string AttributeDictionaryKey = "AUTOCAD_BALLET_ATTRIBUTES";

        /// <summary>
        /// Supported attribute value types
        /// </summary>
        public enum AttributeType
        {
            Text = 1,      // String values
            Integer = 2,   // Integer values
            Real = 3,      // Double/decimal values
            Boolean = 4,   // True/false
            Date = 5       // DateTime as text (ISO 8601)
        }

        /// <summary>
        /// Represents a typed attribute value
        /// </summary>
        public class AttributeValue
        {
            public string Key { get; set; }
            public AttributeType Type { get; set; }
            public object Value { get; set; }

            public string AsString() => Value?.ToString() ?? "";
            public int AsInt() => Value is int i ? i : 0;
            public double AsDouble() => Value is double d ? d : 0.0;
            public bool AsBool() => Value is bool b && b;
            public DateTime AsDate() => Value is DateTime dt ? dt : DateTime.MinValue;
        }

        /// <summary>
        /// Gets all attributes for an entity
        /// </summary>
        public static Dictionary<string, AttributeValue> GetEntityAttributes(ObjectId entityId, Database db)
        {
            var attributes = new Dictionary<string, AttributeValue>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var entity = tr.GetObject(entityId, OpenMode.ForRead) as DBObject;
                    if (entity == null || entity.ExtensionDictionary == ObjectId.Null)
                    {
                        tr.Commit();
                        return attributes;
                    }

                    var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    if (extDict == null || !extDict.Contains(AttributeDictionaryKey))
                    {
                        tr.Commit();
                        return attributes;
                    }

                    var xrecId = extDict.GetAt(AttributeDictionaryKey);
                    var xrec = tr.GetObject(xrecId, OpenMode.ForRead) as Xrecord;
                    if (xrec != null && xrec.Data != null)
                    {
                        string currentKey = null;
                        AttributeType currentType = AttributeType.Text;

                        foreach (TypedValue tv in xrec.Data)
                        {
                            // Format: [Text:Key, Int16:Type, Value:TypedValue]
                            if (tv.TypeCode == (int)DxfCode.Text && tv.Value != null)
                            {
                                currentKey = tv.Value.ToString();
                            }
                            else if (tv.TypeCode == (int)DxfCode.Int16 && currentKey != null)
                            {
                                currentType = (AttributeType)(short)tv.Value;
                            }
                            else if (currentKey != null)
                            {
                                // Parse value based on type
                                object parsedValue = ParseValue(tv, currentType);
                                attributes[currentKey] = new AttributeValue
                                {
                                    Key = currentKey,
                                    Type = currentType,
                                    Value = parsedValue
                                };
                                currentKey = null;
                            }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            return attributes;
        }

        /// <summary>
        /// Sets attributes for an entity (replaces existing)
        /// </summary>
        public static void SetEntityAttributes(ObjectId entityId, Database db, Dictionary<string, AttributeValue> attributes)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var entity = tr.GetObject(entityId, OpenMode.ForRead) as DBObject;
                    if (entity == null)
                    {
                        tr.Abort();
                        return;
                    }

                    // Create extension dictionary if it doesn't exist
                    if (entity.ExtensionDictionary == ObjectId.Null)
                    {
                        entity.UpgradeOpen();
                        entity.CreateExtensionDictionary();
                        entity.DowngradeOpen();
                    }

                    var extDict = tr.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
                    if (extDict == null)
                    {
                        tr.Abort();
                        return;
                    }

                    // Remove existing XRecord if present
                    if (extDict.Contains(AttributeDictionaryKey))
                    {
                        var oldXrecId = extDict.GetAt(AttributeDictionaryKey);
                        var oldXrec = tr.GetObject(oldXrecId, OpenMode.ForWrite);
                        oldXrec.Erase();
                        extDict.Remove(AttributeDictionaryKey);
                    }

                    // Create new XRecord with attributes
                    if (attributes != null && attributes.Count > 0)
                    {
                        var xrec = new Xrecord();
                        var rbValues = new List<TypedValue>();

                        foreach (var attr in attributes.Values.OrderBy(a => a.Key))
                        {
                            // Store as: [Key, Type, Value]
                            rbValues.Add(new TypedValue((int)DxfCode.Text, attr.Key));
                            rbValues.Add(new TypedValue((int)DxfCode.Int16, (short)attr.Type));
                            rbValues.Add(CreateTypedValue(attr));
                        }

                        xrec.Data = new ResultBuffer(rbValues.ToArray());
                        extDict.SetAt(AttributeDictionaryKey, xrec);
                        tr.AddNewlyCreatedDBObject(xrec, true);
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                    throw;
                }
            }
        }

        /// <summary>
        /// Gets a single attribute value
        /// </summary>
        public static AttributeValue GetAttribute(ObjectId entityId, Database db, string key)
        {
            var attributes = GetEntityAttributes(entityId, db);
            return attributes.TryGetValue(key, out var value) ? value : null;
        }

        /// <summary>
        /// Sets a single attribute (merges with existing)
        /// </summary>
        public static void SetAttribute(ObjectId entityId, Database db, string key, AttributeType type, object value)
        {
            var attributes = GetEntityAttributes(entityId, db);
            attributes[key] = new AttributeValue { Key = key, Type = type, Value = value };
            SetEntityAttributes(entityId, db, attributes);
        }

        /// <summary>
        /// Removes an attribute
        /// </summary>
        public static void RemoveAttribute(ObjectId entityId, Database db, string key)
        {
            var attributes = GetEntityAttributes(entityId, db);
            attributes.Remove(key);
            SetEntityAttributes(entityId, db, attributes);
        }

        /// <summary>
        /// Removes all attributes
        /// </summary>
        public static void ClearAttributes(ObjectId entityId, Database db)
        {
            SetEntityAttributes(entityId, db, new Dictionary<string, AttributeValue>());
        }

        /// <summary>
        /// Gets all unique attribute keys across all entities in database
        /// </summary>
        public static HashSet<string> GetAllAttributeKeys(Database db)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        foreach (ObjectId entityId in btr)
                        {
                            try
                            {
                                var attrs = GetEntityAttributes(entityId, db);
                                foreach (var key in attrs.Keys)
                                {
                                    keys.Add(key);
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            return keys;
        }

        /// <summary>
        /// Finds entities with a specific attribute key
        /// </summary>
        public static List<ObjectId> FindEntitiesWithAttribute(Database db, string key)
        {
            var entities = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        foreach (ObjectId entityId in btr)
                        {
                            try
                            {
                                var attrs = GetEntityAttributes(entityId, db);
                                if (attrs.ContainsKey(key))
                                {
                                    entities.Add(entityId);
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            return entities;
        }

        /// <summary>
        /// Finds entities with a specific attribute value
        /// </summary>
        public static List<ObjectId> FindEntitiesWithAttributeValue(Database db, string key, object value)
        {
            var entities = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        foreach (ObjectId entityId in btr)
                        {
                            try
                            {
                                var attr = GetAttribute(entityId, db, key);
                                if (attr != null && attr.Value?.ToString() == value?.ToString())
                                {
                                    entities.Add(entityId);
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            return entities;
        }

        // Helper methods for type conversion
        private static TypedValue CreateTypedValue(AttributeValue attr)
        {
            switch (attr.Type)
            {
                case AttributeType.Text:
                    return new TypedValue((int)DxfCode.ExtendedDataAsciiString, attr.Value?.ToString() ?? "");

                case AttributeType.Integer:
                    return new TypedValue((int)DxfCode.ExtendedDataInteger32, Convert.ToInt32(attr.Value ?? 0));

                case AttributeType.Real:
                    return new TypedValue((int)DxfCode.ExtendedDataReal, Convert.ToDouble(attr.Value ?? 0.0));

                case AttributeType.Boolean:
                    return new TypedValue((int)DxfCode.ExtendedDataInteger32, Convert.ToBoolean(attr.Value) ? 1 : 0);

                case AttributeType.Date:
                    var dateStr = attr.Value is DateTime dt ? dt.ToString("o") : attr.Value?.ToString() ?? "";
                    return new TypedValue((int)DxfCode.ExtendedDataAsciiString, dateStr);

                default:
                    return new TypedValue((int)DxfCode.ExtendedDataAsciiString, attr.Value?.ToString() ?? "");
            }
        }

        private static object ParseValue(TypedValue tv, AttributeType type)
        {
            switch (type)
            {
                case AttributeType.Text:
                    return tv.Value?.ToString() ?? "";

                case AttributeType.Integer:
                    return tv.Value != null ? Convert.ToInt32(tv.Value) : 0;

                case AttributeType.Real:
                    return tv.Value != null ? Convert.ToDouble(tv.Value) : 0.0;

                case AttributeType.Boolean:
                    return tv.Value != null && Convert.ToInt32(tv.Value) != 0;

                case AttributeType.Date:
                    var dateStr = tv.Value?.ToString();
                    return DateTime.TryParse(dateStr, out var dt) ? dt : DateTime.MinValue;

                default:
                    return tv.Value?.ToString() ?? "";
            }
        }
    }

    /// <summary>
    /// Tag management helpers - Tags are just a special text attribute with multi-value support
    /// </summary>
    public static class TagHelpers
    {
        private const string TagsAttributeKey = "Tags";

        /// <summary>
        /// Gets all tags for an entity (tags are stored as comma-separated text attribute)
        /// </summary>
        public static List<string> GetTags(this ObjectId entityId, Database db)
        {
            var tagString = entityId.GetTextAttribute(db, TagsAttributeKey, "");
            if (string.IsNullOrWhiteSpace(tagString))
                return new List<string>();

            return tagString.Split(',')
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();
        }

        /// <summary>
        /// Adds tags to an entity (merges with existing tags)
        /// </summary>
        public static void AddTags(this ObjectId entityId, Database db, params string[] tags)
        {
            if (tags == null || tags.Length == 0)
                return;

            var existingTags = entityId.GetTags(db);
            var newTags = tags.Where(t => !string.IsNullOrWhiteSpace(t)).Select(t => t.Trim());
            var mergedTags = existingTags.Union(newTags, StringComparer.OrdinalIgnoreCase).ToList();

            entityId.SetTextAttribute(db, TagsAttributeKey, string.Join(", ", mergedTags));
        }

        /// <summary>
        /// Sets tags for an entity (replaces existing tags)
        /// </summary>
        public static void SetTags(this ObjectId entityId, Database db, params string[] tags)
        {
            if (tags == null || tags.Length == 0)
            {
                entityId.SetTextAttribute(db, TagsAttributeKey, "");
                return;
            }

            var cleanTags = tags.Where(t => !string.IsNullOrWhiteSpace(t))
                .Select(t => t.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            entityId.SetTextAttribute(db, TagsAttributeKey, string.Join(", ", cleanTags));
        }

        /// <summary>
        /// Removes specific tags from an entity
        /// </summary>
        public static void RemoveTags(this ObjectId entityId, Database db, params string[] tagsToRemove)
        {
            if (tagsToRemove == null || tagsToRemove.Length == 0)
                return;

            var existingTags = entityId.GetTags(db);
            var remainingTags = existingTags.Except(tagsToRemove, StringComparer.OrdinalIgnoreCase).ToList();

            entityId.SetTextAttribute(db, TagsAttributeKey, string.Join(", ", remainingTags));
        }

        /// <summary>
        /// Removes all tags from an entity
        /// </summary>
        public static void ClearTags(this ObjectId entityId, Database db)
        {
            entityId.SetTextAttribute(db, TagsAttributeKey, "");
        }

        /// <summary>
        /// Checks if an entity has a specific tag
        /// </summary>
        public static bool HasTag(this ObjectId entityId, Database db, string tag)
        {
            var tags = entityId.GetTags(db);
            return tags.Any(t => string.Equals(t, tag, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if an entity has any tags
        /// </summary>
        public static bool HasAnyTags(this ObjectId entityId, Database db)
        {
            return entityId.GetTags(db).Count > 0;
        }

        /// <summary>
        /// Gets all unique tags from entities in the current space with usage counts
        /// </summary>
        public static List<TagInfo> GetAllTagsInCurrentSpace(Database db)
        {
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                    foreach (ObjectId entityId in currentSpace)
                    {
                        try
                        {
                            var tags = entityId.GetTags(db);
                            foreach (var tag in tags)
                            {
                                if (tagCounts.ContainsKey(tag))
                                    tagCounts[tag]++;
                                else
                                    tagCounts[tag] = 1;
                            }
                        }
                        catch { }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            return tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select((kvp, index) => new TagInfo
                {
                    Id = index,
                    Name = kvp.Key,
                    UsageCount = kvp.Value
                })
                .ToList();
        }

        /// <summary>
        /// Gets all unique tags from entities in the database with usage counts
        /// </summary>
        public static List<TagInfo> GetAllTagsInDatabase(Database db)
        {
            var diagnosticTimer = System.Diagnostics.Stopwatch.StartNew();
            int blockTableRecordsChecked = 0;
            int entitiesChecked = 0;
            int entitiesWithTags = 0;

            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        blockTableRecordsChecked++;

                        foreach (ObjectId entityId in btr)
                        {
                            entitiesChecked++;
                            try
                            {
                                var tags = entityId.GetTags(db);
                                if (tags.Count > 0)
                                {
                                    entitiesWithTags++;
                                }
                                foreach (var tag in tags)
                                {
                                    if (tagCounts.ContainsKey(tag))
                                        tagCounts[tag]++;
                                    else
                                        tagCounts[tag] = 1;
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            diagnosticTimer.Stop();

            // Write diagnostics to AutoCAD command line
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage($"\n[TAG DIAGNOSTIC] GetAllTagsInDatabase completed in {diagnosticTimer.ElapsedMilliseconds}ms");
                doc.Editor.WriteMessage($"\n[TAG DIAGNOSTIC] Block table records checked: {blockTableRecordsChecked}");
                doc.Editor.WriteMessage($"\n[TAG DIAGNOSTIC] Total entities checked: {entitiesChecked}");
                doc.Editor.WriteMessage($"\n[TAG DIAGNOSTIC] Entities with tags: {entitiesWithTags}");
                doc.Editor.WriteMessage($"\n[TAG DIAGNOSTIC] Unique tags found: {tagCounts.Count}");
            }

            return tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select((kvp, index) => new TagInfo
                {
                    Id = index,
                    Name = kvp.Key,
                    UsageCount = kvp.Value
                })
                .ToList();
        }

        /// <summary>
        /// Gets all unique tags from entities in all layouts (Model + Paper Spaces) of a database
        /// Does not scan block definitions
        /// </summary>
        public static List<TagInfo> GetAllTagsInLayouts(Database db)
        {
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get all layouts (Model + Paper Spaces)
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                        foreach (ObjectId entityId in btr)
                        {
                            try
                            {
                                var tags = entityId.GetTags(db);
                                foreach (var tag in tags)
                                {
                                    if (tagCounts.ContainsKey(tag))
                                        tagCounts[tag]++;
                                    else
                                        tagCounts[tag] = 1;
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            return tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select((kvp, index) => new TagInfo
                {
                    Id = index,
                    Name = kvp.Key,
                    UsageCount = kvp.Value
                })
                .ToList();
        }

        /// <summary>
        /// Gets all unique tags from layouts across multiple databases
        /// Does not scan block definitions
        /// </summary>
        public static List<TagInfo> GetAllTagsInLayoutsAcrossDocuments(IEnumerable<Database> databases)
        {
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var db in databases)
            {
                var dbTags = GetAllTagsInLayouts(db);
                foreach (var tagInfo in dbTags)
                {
                    if (tagCounts.ContainsKey(tagInfo.Name))
                        tagCounts[tagInfo.Name] += tagInfo.UsageCount;
                    else
                        tagCounts[tagInfo.Name] = tagInfo.UsageCount;
                }
            }

            return tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select((kvp, index) => new TagInfo
                {
                    Id = index,
                    Name = kvp.Key,
                    UsageCount = kvp.Value
                })
                .ToList();
        }

        /// <summary>
        /// Gets all unique tags from multiple databases
        /// </summary>
        public static List<TagInfo> GetAllTagsInDocuments(IEnumerable<Database> databases)
        {
            var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var db in databases)
            {
                var dbTags = GetAllTagsInDatabase(db);
                foreach (var tagInfo in dbTags)
                {
                    if (tagCounts.ContainsKey(tagInfo.Name))
                        tagCounts[tagInfo.Name] += tagInfo.UsageCount;
                    else
                        tagCounts[tagInfo.Name] = tagInfo.UsageCount;
                }
            }

            return tagCounts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key)
                .Select((kvp, index) => new TagInfo
                {
                    Id = index,
                    Name = kvp.Key,
                    UsageCount = kvp.Value
                })
                .ToList();
        }

        /// <summary>
        /// Finds all entities with a specific tag
        /// </summary>
        public static List<ObjectId> FindEntitiesWithTag(Database db, string tag)
        {
            var entities = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        foreach (ObjectId entityId in btr)
                        {
                            try
                            {
                                if (entityId.HasTag(db, tag))
                                {
                                    entities.Add(entityId);
                                }
                            }
                            catch { }
                        }
                    }

                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

            return entities;
        }
    }

    /// <summary>
    /// Information about a tag
    /// </summary>
    public class TagInfo
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public int UsageCount { get; set; }
    }

    /// <summary>
    /// Convenience extension methods for common attribute operations
    /// </summary>
    public static class EntityAttributeExtensions
    {
        public static void SetTextAttribute(this ObjectId entityId, Database db, string key, string value)
        {
            EntityAttributes.SetAttribute(entityId, db, key, EntityAttributes.AttributeType.Text, value);
        }

        public static void SetIntAttribute(this ObjectId entityId, Database db, string key, int value)
        {
            EntityAttributes.SetAttribute(entityId, db, key, EntityAttributes.AttributeType.Integer, value);
        }

        public static void SetRealAttribute(this ObjectId entityId, Database db, string key, double value)
        {
            EntityAttributes.SetAttribute(entityId, db, key, EntityAttributes.AttributeType.Real, value);
        }

        public static void SetBoolAttribute(this ObjectId entityId, Database db, string key, bool value)
        {
            EntityAttributes.SetAttribute(entityId, db, key, EntityAttributes.AttributeType.Boolean, value);
        }

        public static void SetDateAttribute(this ObjectId entityId, Database db, string key, DateTime value)
        {
            EntityAttributes.SetAttribute(entityId, db, key, EntityAttributes.AttributeType.Date, value);
        }

        public static string GetTextAttribute(this ObjectId entityId, Database db, string key, string defaultValue = "")
        {
            var attr = EntityAttributes.GetAttribute(entityId, db, key);
            return attr?.AsString() ?? defaultValue;
        }

        public static int GetIntAttribute(this ObjectId entityId, Database db, string key, int defaultValue = 0)
        {
            var attr = EntityAttributes.GetAttribute(entityId, db, key);
            return attr?.AsInt() ?? defaultValue;
        }

        public static double GetRealAttribute(this ObjectId entityId, Database db, string key, double defaultValue = 0.0)
        {
            var attr = EntityAttributes.GetAttribute(entityId, db, key);
            return attr?.AsDouble() ?? defaultValue;
        }

        public static bool GetBoolAttribute(this ObjectId entityId, Database db, string key, bool defaultValue = false)
        {
            var attr = EntityAttributes.GetAttribute(entityId, db, key);
            return attr?.AsBool() ?? defaultValue;
        }

        public static DateTime GetDateAttribute(this ObjectId entityId, Database db, string key, DateTime? defaultValue = null)
        {
            var attr = EntityAttributes.GetAttribute(entityId, db, key);
            return attr?.AsDate() ?? defaultValue ?? DateTime.MinValue;
        }
    }
}
