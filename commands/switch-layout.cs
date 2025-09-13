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

                var layoutsWithOrder = new List<(string name, int order)>();
                foreach (DictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject((ObjectId)entry.Value, OpenMode.ForRead) as Layout;
                    if (layout != null)
                    {
                        layoutsWithOrder.Add((layout.LayoutName, layout.TabOrder));
                    }
                }
                tr.Commit();

                // Sort by tab order, then add to availableLayouts in that order
                layoutsWithOrder = layoutsWithOrder.OrderBy(l => l.order).ToList();
                foreach (var (name, order) in layoutsWithOrder)
                {
                    availableLayouts.Add(new Dictionary<string, object>
                    {
                        ["layout"] = name
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
                string chosenLayoutName = chosen.First()["layout"].ToString();
                LayoutManager.Current.CurrentLayout = chosenLayoutName;
                LayoutChangeTracker.LogLayoutChange(projectName, chosenLayoutName, true);
            }
        }
    }
}