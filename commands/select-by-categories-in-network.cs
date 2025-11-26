using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
#if NET48 || NET47 || NET46
using Newtonsoft.Json;
#endif

[assembly: CommandClass(typeof(AutoCADBallet.SelectByCategoriesInNetwork))]

namespace AutoCADBallet
{
    public class SelectByCategoriesInNetwork
    {
        [CommandMethod("select-by-categories-in-network", CommandFlags.Modal)]
        public void SelectByCategoriesInNetworkCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                ed.WriteMessage("\nNetwork Mode: Gathering entities from all sessions in network...\n");

                // Get active network sessions
                var sessions = AutoCADBalletServer.GetActiveNetworkSessions();
                if (sessions == null || sessions.Count == 0)
                {
                    ed.WriteMessage("\nNo active network sessions found. Start server in other AutoCAD instances first.\n");
                    return;
                }

                // Get shared auth token
                var authToken = AutoCADBalletServer.GetSharedAuthToken();
                if (string.IsNullOrEmpty(authToken))
                {
                    ed.WriteMessage("\nNo shared authentication token found. Network authentication failed.\n");
                    return;
                }

                ed.WriteMessage($"\nFound {sessions.Count} active sessions in network.\n");

                // Query each session for categories
                var allCategoryReferences = new List<CategoryEntityReference>();
                var categoryGroups = new Dictionary<string, List<CategoryEntityReference>>();

                foreach (var session in sessions)
                {
                    ed.WriteMessage($"\nQuerying session {session.SessionId.Substring(0, 8)}... (Port {session.Port})");

                    try
                    {
                        var categories = QuerySessionCategories(session, authToken);
                        if (categories != null && categories.Count > 0)
                        {
                            ed.WriteMessage($" Found {categories.Count} entity references.");
                            allCategoryReferences.AddRange(categories);

                            foreach (var categoryRef in categories)
                            {
                                if (!categoryGroups.ContainsKey(categoryRef.Category))
                                    categoryGroups[categoryRef.Category] = new List<CategoryEntityReference>();

                                categoryGroups[categoryRef.Category].Add(categoryRef);
                            }
                        }
                        else
                        {
                            ed.WriteMessage(" No entities found.");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n  Error querying session: {ex.Message}");
                    }
                }

                if (categoryGroups.Count == 0)
                {
                    ed.WriteMessage("\nNo entities found across network sessions.\n");
                    return;
                }

                ed.WriteMessage($"\n\nTotal categories found: {categoryGroups.Count}\n");

                // Show selection dialog with session information
                ShowSelectionDialogForNetwork(ed, db, categoryGroups, sessions);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in select-by-categories-in-network: {ex.Message}\n");
            }
        }

        private List<CategoryEntityReference> QuerySessionCategories(NetworkEntry session, string authToken)
        {
            var references = new List<CategoryEntityReference>();

            try
            {
                var url = $"https://127.0.0.1:{session.Port}/select-by-categories";

#if NET48 || NET47 || NET46
                // Use WebRequest for .NET Framework
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.Headers.Add("X-Auth-Token", authToken);
                request.ContentType = "application/json";

                // Accept self-signed certificates (localhost only)
                ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

                // Empty body for this endpoint (it returns all categories from the session)
                var requestBody = "{}";
                var requestBytes = Encoding.UTF8.GetBytes(requestBody);
                request.ContentLength = requestBytes.Length;

                using (var requestStream = request.GetRequestStream())
                {
                    requestStream.Write(requestBytes, 0, requestBytes.Length);
                }

                // Get response
                using (var response = (HttpWebResponse)request.GetResponse())
                using (var responseStream = response.GetResponseStream())
                using (var reader = new StreamReader(responseStream))
                {
                    var responseText = reader.ReadToEnd();
                    var scriptResponse = JsonConvert.DeserializeObject<ScriptResponse>(responseText);

                    if (scriptResponse != null && scriptResponse.Success && !string.IsNullOrEmpty(scriptResponse.Output))
                    {
                        references = JsonConvert.DeserializeObject<List<CategoryEntityReference>>(scriptResponse.Output);
                    }
                }
#else
                // Use HttpClient for .NET 8
                using (var handler = new System.Net.Http.HttpClientHandler())
                {
                    handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;

                    using (var client = new System.Net.Http.HttpClient(handler))
                    {
                        client.DefaultRequestHeaders.Add("X-Auth-Token", authToken);

                        var content = new System.Net.Http.StringContent("{}", Encoding.UTF8, "application/json");
                        var response = client.PostAsync(url, content).Result;

                        if (response.IsSuccessStatusCode)
                        {
                            var responseText = response.Content.ReadAsStringAsync().Result;
                            var scriptResponse = System.Text.Json.JsonSerializer.Deserialize<ScriptResponse>(responseText);

                            if (scriptResponse != null && scriptResponse.Success && !string.IsNullOrEmpty(scriptResponse.Output))
                            {
                                references = System.Text.Json.JsonSerializer.Deserialize<List<CategoryEntityReference>>(scriptResponse.Output);
                            }
                        }
                    }
                }
#endif
            }
            catch
            {
                // Return empty list on error
            }

            return references ?? new List<CategoryEntityReference>();
        }

        private void ShowSelectionDialogForNetwork(Editor ed, Database db, Dictionary<string, List<CategoryEntityReference>> categoryGroups, List<NetworkEntry> sessions)
        {
            if (categoryGroups.Count == 0)
            {
                ed.WriteMessage("\nNo entities found across network sessions.\n");
                return;
            }

            var entries = new List<Dictionary<string, object>>();
            foreach (var cat in categoryGroups.OrderBy(c => c.Key))
            {
                // Group by session for display
                var sessionCounts = cat.Value
                    .GroupBy(e => e.DocumentName)
                    .Select(g => $"{g.Key}: {g.Count()}")
                    .ToList();

                var sessionIds = cat.Value
                    .Select(e => sessions.FirstOrDefault(s => s.Documents.Contains(e.DocumentName))?.SessionId ?? "Unknown")
                    .Distinct()
                    .Select(sid => sid.Substring(0, 8))
                    .ToList();

                entries.Add(new Dictionary<string, object>
                {
                    { "Category", cat.Key },
                    { "Total Count", cat.Value.Count },
                    { "Sessions", string.Join(", ", sessionIds) },
                    { "Documents", string.Join(", ", sessionCounts) }
                });
            }

            var propertyNames = new List<string> { "Category", "Total Count", "Sessions", "Documents" };
            var chosenRows = CustomGUIs.DataGrid(entries, propertyNames, false, null);

            if (chosenRows.Count == 0)
            {
                ed.WriteMessage("\nNo categories selected.\n");
                return;
            }

            var selectedCategoryNames = new HashSet<string>(chosenRows.Select(row => row["Category"].ToString()));
            var selectedEntities = new List<CategoryEntityReference>();

            foreach (var categoryName in selectedCategoryNames)
            {
                if (categoryGroups.ContainsKey(categoryName))
                {
                    selectedEntities.AddRange(categoryGroups[categoryName]);
                }
            }

            var selectionItems = selectedEntities.Select(entityRef => new SelectionItem
            {
                DocumentPath = entityRef.DocumentPath,
                Handle = entityRef.Handle,
                SessionId = null
            }).ToList();

            // Save network selection (replaces previous selection)
            try
            {
                SelectionStorage.SaveSelection(selectionItems);
                ed.WriteMessage($"\nNetwork selection saved.\n");
                ed.WriteMessage($"\nSummary:\n");
                ed.WriteMessage($"  Total entities: {selectedEntities.Count}\n");
                ed.WriteMessage($"  Categories: {selectedCategoryNames.Count}\n");
                ed.WriteMessage($"  Sessions: {sessions.Count}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError saving selection: {ex.Message}\n");
            }

            // Set current view selection to subset of selected entities in current document's current space
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var currentDocPath = Path.GetFullPath(doc.Name);
            var currentViewIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var entityRef in selectedEntities)
                {
                    try
                    {
                        // Only process entities from current document
                        var entityDocPath = Path.GetFullPath(entityRef.DocumentPath);
                        if (!string.Equals(entityDocPath, currentDocPath, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var handle = Convert.ToInt64(entityRef.Handle, 16);
                        var objectId = db.GetObjectId(false, new Handle(handle), 0);

                        if (!objectId.IsNull && objectId.IsValid)
                        {
                            var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                            if (entity != null && entity.BlockId == db.CurrentSpaceId)
                            {
                                currentViewIds.Add(objectId);
                            }
                        }
                    }
                    catch
                    {
                        // Skip invalid entities
                    }
                }
                tr.Commit();
            }

            if (currentViewIds.Count > 0)
            {
                try
                {
                    ed.SetImpliedSelection(currentViewIds.ToArray());
                    ed.WriteMessage($"  Selected {currentViewIds.Count} entities in current view.\n");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError setting current view selection: {ex.Message}\n");
                }
            }
        }

        // Helper class to deserialize script response
        private class ScriptResponse
        {
            public bool Success { get; set; }
            public string Output { get; set; }
            public string Error { get; set; }
            public List<string> Diagnostics { get; set; }
        }
    }
}
