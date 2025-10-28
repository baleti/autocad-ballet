using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Linq;

namespace AutoCADBallet
{
    /// <summary>
    /// Utilities for managing XData on AutoCAD entities
    /// Used to attach persistent GUIDs that link to external database records
    /// </summary>
    public static class XDataUtilities
    {
        private const string AppName = "AUTOCAD_BALLET_TAG";
        private const int MaxGuidRetries = 10;

        /// <summary>
        /// Ensures the application is registered in the current database
        /// </summary>
        private static void EnsureAppRegistered(Database db)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var regAppTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);

                if (!regAppTable.Has(AppName))
                {
                    regAppTable.UpgradeOpen();
                    var regApp = new RegAppTableRecord();
                    regApp.Name = AppName;
                    regAppTable.Add(regApp);
                    tr.AddNewlyCreatedDBObject(regApp, true);
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// Gets or creates a GUID for an entity
        /// If the entity already has a GUID in XData, returns it
        /// Otherwise, creates a new GUID, stores it in XData, and returns it
        /// </summary>
        public static string GetOrCreateEntityGuid(ObjectId objectId, Database db, string documentPath)
        {
            EnsureAppRegistered(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = (Entity)tr.GetObject(objectId, OpenMode.ForRead);

                // Check if entity already has a GUID
                var xdata = entity.GetXDataForApplication(AppName);
                if (xdata != null && xdata.AsArray().Length >= 2)
                {
                    var guidValue = xdata.AsArray()[1];
                    if (guidValue.TypeCode == (short)DxfCode.ExtendedDataAsciiString)
                    {
                        var existingGuid = (string)guidValue.Value;
                        tr.Commit();
                        return existingGuid;
                    }
                }

                // Create new GUID for this entity
                var newGuid = Guid.NewGuid().ToString();

                // Store in XData
                entity.UpgradeOpen();
                var rb = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, AppName),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, newGuid)
                );
                entity.XData = rb;
                rb.Dispose();

                tr.Commit();
                return newGuid;
            }
        }

        /// <summary>
        /// Gets the GUID from an entity's XData, if it exists
        /// Returns null if no GUID is found
        /// </summary>
        public static string GetEntityGuid(ObjectId objectId, Database db)
        {
            EnsureAppRegistered(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = (Entity)tr.GetObject(objectId, OpenMode.ForRead);

                var xdata = entity.GetXDataForApplication(AppName);
                if (xdata != null && xdata.AsArray().Length >= 2)
                {
                    var guidValue = xdata.AsArray()[1];
                    if (guidValue.TypeCode == (short)DxfCode.ExtendedDataAsciiString)
                    {
                        tr.Commit();
                        return (string)guidValue.Value;
                    }
                }

                tr.Commit();
                return null;
            }
        }

        /// <summary>
        /// Sets a GUID on an entity's XData
        /// Overwrites any existing GUID
        /// </summary>
        public static void SetEntityGuid(ObjectId objectId, Database db, string guid, string documentPath)
        {
            EnsureAppRegistered(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);

                var rb = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, AppName),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, guid)
                );
                entity.XData = rb;
                rb.Dispose();

                tr.Commit();
            }
        }

        /// <summary>
        /// Removes the GUID from an entity's XData
        /// </summary>
        public static void RemoveEntityGuid(ObjectId objectId, Database db)
        {
            EnsureAppRegistered(db);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var entity = (Entity)tr.GetObject(objectId, OpenMode.ForWrite);

                // Remove XData by setting to null
                var xdata = entity.GetXDataForApplication(AppName);
                if (xdata != null)
                {
                    entity.XData = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, AppName)
                    );
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// Finds all entities in the current database that have GUIDs
        /// Returns a dictionary mapping ObjectId to GUID
        /// </summary>
        public static System.Collections.Generic.Dictionary<ObjectId, string> GetAllEntitiesWithGuids(Database db)
        {
            EnsureAppRegistered(db);

            var result = new System.Collections.Generic.Dictionary<ObjectId, string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in blockTable)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    foreach (ObjectId entityId in btr)
                    {
                        var entity = (Entity)tr.GetObject(entityId, OpenMode.ForRead);

                        var xdata = entity.GetXDataForApplication(AppName);
                        if (xdata != null && xdata.AsArray().Length >= 2)
                        {
                            var guidValue = xdata.AsArray()[1];
                            if (guidValue.TypeCode == (short)DxfCode.ExtendedDataAsciiString)
                            {
                                result[entityId] = (string)guidValue.Value;
                            }
                        }
                    }
                }

                tr.Commit();
            }

            return result;
        }
    }
}
