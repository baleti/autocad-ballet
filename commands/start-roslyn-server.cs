using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(AutoCADBallet.StartRoslynServer))]

namespace AutoCADBallet
{
    /// <summary>
    /// Global context accessible from executed scripts
    /// </summary>
    public class ScriptGlobals
    {
        public Document Doc { get; set; }
        public Database Db { get; set; }
        public Editor Ed { get; set; }
    }

    /// <summary>
    /// Response structure for script execution results
    /// </summary>
    public class ScriptResponse
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
        public List<string> Diagnostics { get; set; }

        public ScriptResponse()
        {
            Diagnostics = new List<string>();
        }

        public string ToJson()
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append($"\"success\": {Success.ToString().ToLower()},");
            sb.Append($"\"output\": {EscapeJson(Output ?? "")},");
            sb.Append($"\"error\": {EscapeJson(Error ?? "")},");
            sb.Append("\"diagnostics\": [");
            for (int i = 0; i < Diagnostics.Count; i++)
            {
                if (i > 0) sb.Append(",");
                sb.Append(EscapeJson(Diagnostics[i]));
            }
            sb.Append("]");
            sb.Append("}");
            return sb.ToString();
        }

        private string EscapeJson(string str)
        {
            if (string.IsNullOrEmpty(str)) return "\"\"";

            var sb = new StringBuilder();
            sb.Append("\"");
            foreach (char c in str)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(c); break;
                }
            }
            sb.Append("\"");
            return sb.ToString();
        }
    }

    public class StartRoslynServer
    {
        internal const int PORT = 23714;
        internal const string URL = "http://127.0.0.1:23714/";

        [CommandMethod("start-roslyn-server", CommandFlags.Session)]
        public void StartRoslynServerCommand()
        {
            var dialog = new RoslynServerDialog();
            dialog.ShowDialog();
        }
    }

    public class RoslynServerDialog : Form
    {
        private TextBox logTextBox;
        private Label statusLabel;
        private TcpListener listener;
        private CancellationTokenSource cancellationTokenSource;
        private bool isRunning = false;

        public RoslynServerDialog()
        {
            InitializeComponents();
            StartServer();
        }

        private void InitializeComponents()
        {
            this.Text = "Roslyn Server - AutoCAD Ballet";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.KeyPreview = true;
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    StopServer();
                    this.Close();
                }
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10)
            };

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            statusLabel = new Label
            {
                Text = "Initializing server...",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            };

            logTextBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.LightGreen
            };

            layout.Controls.Add(statusLabel, 0, 0);
            layout.Controls.Add(logTextBox, 0, 1);

            this.Controls.Add(layout);

            this.FormClosing += (s, e) =>
            {
                StopServer();
            };
        }

        private void StartServer()
        {
            try
            {
                cancellationTokenSource = new CancellationTokenSource();
                listener = new TcpListener(IPAddress.Parse("127.0.0.1"), StartRoslynServer.PORT);
                listener.Start();
                isRunning = true;

                UpdateStatus($"Listening on http://127.0.0.1:{StartRoslynServer.PORT}/");
                LogMessage($"[{DateTime.Now:HH:mm:ss}] Press ESC to stop the server");

                Task.Run(() => AcceptRequestsLoop(cancellationTokenSource.Token));
            }
            catch (System.Exception ex)
            {
                UpdateStatus($"Error starting server: {ex.Message}");
                LogMessage($"[{DateTime.Now:HH:mm:ss}] ERROR: {ex.Message}");
            }
        }

        private async Task AcceptRequestsLoop(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested && isRunning)
            {
                try
                {
                    // Check for pending connections
                    if (listener.Pending())
                    {
                        var client = await listener.AcceptTcpClientAsync();
                        var clientEndpoint = client.Client.RemoteEndPoint.ToString();

                        // Handle request in background
                        _ = Task.Run(async () => await HandleHttpRequest(client, clientEndpoint, cancellationToken));
                    }
                    else
                    {
                        // Small delay to prevent busy waiting
                        await Task.Delay(100, cancellationToken);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (System.Exception ex)
                {
                    LogMessage($"[{DateTime.Now:HH:mm:ss}] Error accepting request: {ex.Message}");
                }
            }
        }

        private async Task HandleHttpRequest(TcpClient client, string clientEndpoint, CancellationToken cancellationToken)
        {
            try
            {
                using (client)
                using (var stream = client.GetStream())
                using (var reader = new StreamReader(stream, Encoding.UTF8, leaveOpen: true))
                {
                    // Read HTTP request line and headers
                    var requestLine = await reader.ReadLineAsync();

                    var headers = new Dictionary<string, string>();
                    string line;
                    while ((line = await reader.ReadLineAsync()) != null && line.Length > 0)
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                        {
                            var key = line.Substring(0, colonIndex).Trim();
                            var value = line.Substring(colonIndex + 1).Trim();
                            headers[key] = value;
                        }
                    }

                    // Read body based on Content-Length
                    string code = "";
                    if (headers.ContainsKey("Content-Length"))
                    {
                        var contentLength = int.Parse(headers["Content-Length"]);
                        var buffer = new char[contentLength];
                        var totalRead = 0;

                        while (totalRead < contentLength)
                        {
                            var bytesRead = await reader.ReadAsync(buffer, totalRead, contentLength - totalRead);
                            if (bytesRead == 0) break;
                            totalRead += bytesRead;
                        }

                        code = new string(buffer, 0, totalRead);
                    }

                    LogMessage($"[{DateTime.Now:HH:mm:ss}] Request body received ({code.Length} bytes)");

                    if (string.IsNullOrWhiteSpace(code))
                    {
                        var emptyResponse = new ScriptResponse
                        {
                            Success = false,
                            Error = "Empty request body"
                        };
                        await SendHttpResponse(stream, emptyResponse);
                        return;
                    }

                    LogMessage($"[{DateTime.Now:HH:mm:ss}] Compiling and executing...");

                    // Execute the script
                    var scriptResponse = await ExecuteScript(code);

                    LogMessage($"[{DateTime.Now:HH:mm:ss}] Execution completed - Success: {scriptResponse.Success}");

                    if (scriptResponse.Success && !string.IsNullOrEmpty(scriptResponse.Output))
                    {
                        var firstLine = scriptResponse.Output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
                        var preview = firstLine.Length > 100
                            ? firstLine.Substring(0, 100) + "..."
                            : firstLine;
                        if (!string.IsNullOrEmpty(preview))
                        {
                            LogMessage($"[{DateTime.Now:HH:mm:ss}] Output preview: {preview}");
                        }
                    }

                    // Send HTTP response
                    await SendHttpResponse(stream, scriptResponse);
                }
            }
            catch (System.Exception ex)
            {
                LogMessage($"[{DateTime.Now:HH:mm:ss}] Error handling request from {clientEndpoint}: {ex.Message}");
            }
        }

        private async Task SendHttpResponse(NetworkStream stream, ScriptResponse scriptResponse)
        {
            var jsonResponse = scriptResponse.ToJson();
            var responseBytes = Encoding.UTF8.GetBytes(jsonResponse);

            var statusCode = scriptResponse.Success ? 200 : 400;
            var statusText = scriptResponse.Success ? "OK" : "Bad Request";

            var httpResponse = $"HTTP/1.1 {statusCode} {statusText}\r\n" +
                             $"Content-Type: application/json\r\n" +
                             $"Content-Length: {responseBytes.Length}\r\n" +
                             $"\r\n";

            var headerBytes = Encoding.UTF8.GetBytes(httpResponse);

            await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
            await stream.WriteAsync(responseBytes, 0, responseBytes.Length);
            await stream.FlushAsync();
        }

        private async Task<ScriptResponse> ExecuteScript(string code)
        {
            var response = new ScriptResponse();
            var outputCapture = new StringWriter();
            var originalOut = Console.Out;

            try
            {
                // Redirect console output
                Console.SetOut(outputCapture);

                // Get AutoCAD context
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                var db = doc?.Database;
                var ed = doc?.Editor;

                if (doc == null || db == null)
                {
                    response.Success = false;
                    response.Error = "No active AutoCAD document";
                    return response;
                }

                // Create script globals with AutoCAD context
                var globals = new ScriptGlobals
                {
                    Doc = doc,
                    Db = db,
                    Ed = ed
                };

                // Configure script options with necessary references
                var scriptOptions = ScriptOptions.Default;

                // Add reference to the currently loaded autocad-ballet.dll assembly
                // This works whether the assembly was loaded from file or from bytes
                // Using the same assembly instance ensures type identity consistency
                var scriptGlobalsAsm = typeof(ScriptGlobals).Assembly;
                scriptOptions = scriptOptions
                    .AddReferences(scriptGlobalsAsm)
                    .AddReferences(typeof(Document).Assembly)        // Autodesk.AutoCAD.ApplicationServices
                    .AddReferences(typeof(Database).Assembly)        // Autodesk.AutoCAD.DatabaseServices
                    .AddReferences(typeof(Editor).Assembly)          // Autodesk.AutoCAD.EditorInput
                    .AddReferences(typeof(System.Linq.Enumerable).Assembly)
                    .AddImports("System")
                    .AddImports("System.Linq")
                    .AddImports("System.Collections.Generic")
                    .AddImports("Autodesk.AutoCAD.ApplicationServices")
                    .AddImports("Autodesk.AutoCAD.DatabaseServices")
                    .AddImports("Autodesk.AutoCAD.EditorInput")
                    .AddImports("Autodesk.AutoCAD.Geometry");

                // Compile the script first to catch compilation errors
                var script = CSharpScript.Create(code, scriptOptions, typeof(ScriptGlobals));
                var compilation = script.GetCompilation();
                var diagnostics = compilation.GetDiagnostics();

                // Check for compilation errors
                var errors = diagnostics.Where(d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error).ToList();
                if (errors.Any())
                {
                    response.Success = false;
                    response.Error = "Compilation failed";
                    response.Diagnostics = errors.Select(e => e.ToString()).ToList();
                    return response;
                }

                // Add warnings to diagnostics
                var warnings = diagnostics.Where(d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Warning).ToList();
                if (warnings.Any())
                {
                    response.Diagnostics = warnings.Select(w => w.ToString()).ToList();
                }

                // Execute the script
                var result = await script.RunAsync(globals);

                response.Success = true;
                response.Output = outputCapture.ToString();

                // If script returned a value, append it to output
                if (result.ReturnValue != null)
                {
                    if (!string.IsNullOrEmpty(response.Output))
                        response.Output += "\n";
                    response.Output += $"Return value: {result.ReturnValue}";
                }
            }
            catch (CompilationErrorException ex)
            {
                response.Success = false;
                response.Error = "Compilation error";
                response.Diagnostics = ex.Diagnostics.Select(d => d.ToString()).ToList();
            }
            catch (System.Exception ex)
            {
                response.Success = false;
                response.Error = ex.Message;
                response.Output = outputCapture.ToString();
            }
            finally
            {
                Console.SetOut(originalOut);
            }

            return response;
        }

        private void StopServer()
        {
            if (!isRunning) return;

            isRunning = false;
            cancellationTokenSource?.Cancel();
            listener?.Stop();

            UpdateStatus("Server stopped");
            LogMessage($"[{DateTime.Now:HH:mm:ss}] Server stopped");
        }

        private void UpdateStatus(string message)
        {
            if (statusLabel.InvokeRequired)
            {
                statusLabel.Invoke(new Action(() => statusLabel.Text = message));
            }
            else
            {
                statusLabel.Text = message;
            }
        }

        private void LogMessage(string message)
        {
            if (logTextBox.InvokeRequired)
            {
                logTextBox.Invoke(new Action(() =>
                {
                    logTextBox.AppendText(message + Environment.NewLine);
                }));
            }
            else
            {
                logTextBox.AppendText(message + Environment.NewLine);
            }
        }
    }
}
