using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.SwitchLayoutCommand))]

namespace AutoCADBallet
{
    public class SwitchLayoutCommand
    {
        [CommandMethod("switch-layout")]
        public void SwitchLayout()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;

            string projectName = Path.GetFileNameWithoutExtension(doc.Name) ?? "UnknownProject";

            var availableLayouts = new List<Dictionary<string, object>>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                var layoutsWithOrder = new List<Dictionary<string, object>>();
                foreach (DictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                    if (layout != null)
                    {
                        layoutsWithOrder.Add(new Dictionary<string, object>
                        {
                            ["LayoutName"] = layout.LayoutName,
                            ["TabOrder"] = layout.TabOrder,
                            ["ObjectId"] = (ObjectId)entry.Value,
                            ["Handle"] = layout.Handle.ToString()
                        });
                    }
                }
                tr.Commit();

                // Sort by tab order, then add to availableLayouts in that order
                layoutsWithOrder = layoutsWithOrder.OrderBy(l => (int)l["TabOrder"]).ToList();
                foreach (var layoutInfo in layoutsWithOrder)
                {
                    availableLayouts.Add(new Dictionary<string, object>
                    {
                        ["layout"] = layoutInfo["LayoutName"],
                        ["ObjectId"] = layoutInfo["ObjectId"],
                        ["Handle"] = layoutInfo["Handle"],
                        ["DocumentPath"] = doc.Name,
                        ["DocumentObject"] = doc
                    });
                }
            }

            if (availableLayouts.Count == 0)
                return;

            int selectedIndex = -1;
            string currentLayoutName = LayoutManager.Current.CurrentLayout;
            selectedIndex = availableLayouts.FindIndex(l => l["layout"].ToString() == currentLayoutName);

            var propertyNames = new List<string> { "layout" };
            var initialSelectionIndices = selectedIndex >= 0
                                            ? new List<int> { selectedIndex }
                                            : new List<int>();

            List<Dictionary<string, object>> chosen = CustomGUIs.DataGrid(availableLayouts, propertyNames, false, initialSelectionIndices);

            if (chosen != null && chosen.Count > 0)
            {
                var selected = chosen.First();
                string chosenLayoutName = selected["layout"].ToString();
                Document targetDoc = selected["DocumentObject"] as Document;

                if (targetDoc != null)
                {
                    try
                    {
                        using (DocumentLock docLock = targetDoc.LockDocument())
                        {
                            LayoutManager.Current.CurrentLayout = chosenLayoutName;
                        }
                        SwitchViewLogging.LogLayoutChange(projectName, chosenLayoutName, true);
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception)
                    {
                        ed.WriteMessage($"\nFailed to switch to layout: {chosenLayoutName}\n");
                    }
                }
            }
        }
    }
}