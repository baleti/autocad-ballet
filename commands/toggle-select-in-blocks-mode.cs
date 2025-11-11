using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.ToggleSelectInBlocksMode))]

namespace AutoCADBallet
{
    /// <summary>
    /// Utility class to manage the select-in-blocks mode state.
    /// This controls whether selection commands should search within blocks (block references, dynamic blocks, xrefs).
    /// </summary>
    public static class SelectInBlocksMode
    {
        private static readonly string RuntimeDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "autocad-ballet",
            "runtime"
        );

        private static readonly string StateFilePath = Path.Combine(RuntimeDir, "select-in-blocks-mode");

        /// <summary>
        /// Gets the current select-in-blocks mode state.
        /// </summary>
        /// <returns>True if selection commands should search within blocks, false otherwise. Default is false.</returns>
        public static bool IsEnabled()
        {
            try
            {
                if (!File.Exists(StateFilePath))
                {
                    return false; // Default: don't search within blocks
                }

                string content = File.ReadAllText(StateFilePath).Trim().ToLower();
                return content == "true";
            }
            catch
            {
                return false; // Default on error
            }
        }

        /// <summary>
        /// Sets the select-in-blocks mode state.
        /// </summary>
        /// <param name="enabled">True to enable searching within blocks, false to disable.</param>
        public static void SetEnabled(bool enabled)
        {
            try
            {
                // Ensure runtime directory exists
                if (!Directory.Exists(RuntimeDir))
                {
                    Directory.CreateDirectory(RuntimeDir);
                }

                // Write state to file
                File.WriteAllText(StateFilePath, enabled ? "true" : "false");
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Failed to save select-in-blocks mode state: {ex.Message}", ex);
            }
        }
    }

    public class ToggleSelectInBlocksMode
    {
        [CommandMethod("toggle-select-in-blocks-mode", CommandFlags.Modal)]
        public void ToggleSelectInBlocksModeCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                // Get current state
                bool currentState = SelectInBlocksMode.IsEnabled();

                // Build DataGrid entries for the two states
                var entries = new List<Dictionary<string, object>>();

                entries.Add(new Dictionary<string, object>
                {
                    { "State", "Enabled" },
                    { "Description", "Selection commands will search within blocks/xrefs" },
                    { "IsEnabled", true }
                });

                entries.Add(new Dictionary<string, object>
                {
                    { "State", "Disabled" },
                    { "Description", "Selection commands will NOT search within blocks/xrefs" },
                    { "IsEnabled", false }
                });

                // Determine initial selection based on current state
                var initialSelection = new List<int> { currentState ? 0 : 1 };

                // Show DataGrid
                var propertyNames = new List<string> { "State", "Description" };
                var chosenRows = CustomGUIs.DataGrid(
                    entries,
                    propertyNames,
                    spanAllScreens: false,
                    initialSelectionIndices: initialSelection);

                if (chosenRows == null || chosenRows.Count == 0)
                {
                    ed.WriteMessage("\nSelect-in-blocks mode not changed.\n");
                    return;
                }

                // Get the selected state
                var selectedRow = chosenRows[0];
                bool newState = (bool)selectedRow["IsEnabled"];

                // Only update if state changed
                if (newState != currentState)
                {
                    SelectInBlocksMode.SetEnabled(newState);
                    string modeDescription = newState
                        ? "enabled (selection commands will search within blocks/xrefs)"
                        : "disabled (selection commands will NOT search within blocks/xrefs)";
                    ed.WriteMessage($"\nSelect-in-blocks mode {modeDescription}\n");
                }
                else
                {
                    ed.WriteMessage("\nSelect-in-blocks mode unchanged.\n");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError changing select-in-blocks mode: {ex.Message}\n");
            }
        }
    }
}
