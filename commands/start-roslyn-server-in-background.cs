using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
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

[assembly: CommandClass(typeof(AutoCADBallet.StartRoslynServerInBackground))]

namespace AutoCADBallet
{
    public class StartRoslynServerInBackground
    {
        internal const int PORT = 34157;
        internal const string URL = "http://127.0.0.1:34157/";

        // Global instance to manage the background server
        private static BackgroundRoslynServer serverInstance = null;
        private static readonly object lockObject = new object();

        [CommandMethod("start-roslyn-server-in-background", CommandFlags.Session)]
        public void StartRoslynServerInBackgroundCommand()
        {
            lock (lockObject)
            {
                if (serverInstance != null && serverInstance.IsRunning)
                {
                    var doc = AcadApp.DocumentManager.MdiActiveDocument;
                    var ed = doc?.Editor;
                    ed?.WriteMessage("\nRoslyn server is already running. Use 'stop-roslyn-server' to stop it.\n");

                    // Show the existing dialog
                    if (serverInstance.Dialog != null && !serverInstance.Dialog.IsDisposed)
                    {
                        serverInstance.Dialog.Invoke(new Action(() =>
                        {
                            serverInstance.Dialog.BringToFront();
                            serverInstance.Dialog.Activate();
                        }));
                    }
                    return;
                }

                try
                {
                    serverInstance = new BackgroundRoslynServer();
                    serverInstance.Start();
                }
                catch (System.Exception ex)
                {
                    var doc = AcadApp.DocumentManager.MdiActiveDocument;
                    var ed = doc?.Editor;
                    ed?.WriteMessage($"\nError starting background Roslyn server: {ex.Message}\n");
                    serverInstance = null;
                }
            }
        }

        internal static void StopServer()
        {
            lock (lockObject)
            {
                if (serverInstance != null)
                {
                    serverInstance.Stop();
                    serverInstance = null;
                }
            }
        }

        internal static BackgroundRoslynServer GetServerInstance()
        {
            lock (lockObject)
            {
                return serverInstance;
            }
        }
    }

    public class BackgroundRoslynServer
    {
        private TcpListener listener;
        private CancellationTokenSource cancellationTokenSource;
        private bool isRunning = false;
        private string authToken;
        private BackgroundRoslynServerDialog dialog;

        public bool IsRunning => isRunning;
        public string AuthToken => authToken;
        public BackgroundRoslynServerDialog Dialog => dialog;

        public void Start()
        {
            if (isRunning)
                return;

            // Generate secure random password (32 bytes = 256 bits)
            authToken = GenerateSecureToken();

            cancellationTokenSource = new CancellationTokenSource();
            listener = new TcpListener(IPAddress.Parse("127.0.0.1"), StartRoslynServerInBackground.PORT);
            listener.Start();
            isRunning = true;

            // Start background listener thread
            Task.Run(() => AcceptRequestsLoop(cancellationTokenSource.Token));

            // Show non-blocking status dialog
            dialog = new BackgroundRoslynServerDialog(this);
            dialog.Show();

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            ed?.WriteMessage($"\nRoslyn server started in background on {StartRoslynServerInBackground.URL}\n");
            ed?.WriteMessage($"Authentication token: {authToken}\n");
        }

        public void Stop()
        {
            if (!isRunning)
                return;

            isRunning = false;
            cancellationTokenSource?.Cancel();

            try
            {
                listener?.Stop();
            }
            catch { }

            if (dialog != null && !dialog.IsDisposed)
            {
                dialog.Invoke(new Action(() => dialog.Close()));
            }

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            ed?.WriteMessage("\nRoslyn server stopped.\n");
        }

        private string GenerateSecureToken()
        {
            var tokenBytes = new byte[32]; // 256 bits
#if NET8_0_OR_GREATER
            RandomNumberGenerator.Fill(tokenBytes);
#else
            using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
            {
                rng.GetBytes(tokenBytes);
            }
#endif
            return Convert.ToBase64String(tokenBytes);
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

                        LogMessage($"[{DateTime.Now:HH:mm:ss}] Connection from {clientEndpoint}");

                        // Handle request in background thread
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
                using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
                {
                    // Parse HTTP request on background thread (no AutoCAD interaction)
                    var requestLine = await reader.ReadLineAsync();

                    var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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

                    // Check authentication
                    string providedToken = null;
                    if (headers.ContainsKey("X-Auth-Token"))
                    {
                        providedToken = headers["X-Auth-Token"];
                    }
                    else if (headers.ContainsKey("Authorization"))
                    {
                        // Support Bearer token format
                        var authHeader = headers["Authorization"];
                        if (authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
                        {
                            providedToken = authHeader.Substring(7).Trim();
                        }
                    }

                    if (string.IsNullOrEmpty(providedToken) || providedToken != authToken)
                    {
                        LogMessage($"[{DateTime.Now:HH:mm:ss}] Authentication failed from {clientEndpoint}");

                        var errorResponse = new ScriptResponse
                        {
                            Success = false,
                            Error = "Authentication required. Provide X-Auth-Token header or Authorization: Bearer <token> header."
                        };
                        await SendHttpResponse(stream, errorResponse, 401, "Unauthorized");
                        return;
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

                    LogMessage($"[{DateTime.Now:HH:mm:ss}] Authenticated request received ({code.Length} bytes)");

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

                    LogMessage($"[{DateTime.Now:HH:mm:ss}] Compiling script on background thread...");

                    // Compile the script on background thread (no AutoCAD interaction needed)
                    var compilationResult = await CompileScript(code);

                    if (!compilationResult.Success)
                    {
                        LogMessage($"[{DateTime.Now:HH:mm:ss}] Compilation failed");
                        await SendHttpResponse(stream, compilationResult.ErrorResponse);
                        return;
                    }

                    LogMessage($"[{DateTime.Now:HH:mm:ss}] Script compiled successfully. Executing on UI thread...");

                    // Execute script on UI thread with document lock
                    var scriptResponse = await ExecuteScriptOnUIThread(compilationResult.CompiledScript);

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

        private class CompilationResult
        {
            public bool Success { get; set; }
            public Script<object> CompiledScript { get; set; }
            public ScriptResponse ErrorResponse { get; set; }
        }

        private Task<CompilationResult> CompileScript(string code)
        {
            var result = new CompilationResult();

            try
            {
                // Configure script options with necessary references
                var scriptOptions = ScriptOptions.Default;

                // Add reference to autocad-ballet.dll for ScriptGlobals
                var scriptGlobalsAsm = typeof(ScriptGlobals).Assembly;

                if (string.IsNullOrEmpty(scriptGlobalsAsm.Location))
                {
                    result.Success = false;
                    result.ErrorResponse = new ScriptResponse
                    {
                        Success = false,
                        Error = $"Cannot reference assembly without location. Assembly: {scriptGlobalsAsm.FullName}"
                    };
                    return Task.FromResult(result);
                }

                scriptOptions = scriptOptions
                    .AddReferences(scriptGlobalsAsm)
                    .AddReferences(typeof(Document).Assembly)
                    .AddReferences(typeof(Database).Assembly)
                    .AddReferences(typeof(Editor).Assembly)
                    .AddReferences(typeof(System.Linq.Enumerable).Assembly)
                    .AddImports("System")
                    .AddImports("System.Linq")
                    .AddImports("System.Collections.Generic")
                    .AddImports("Autodesk.AutoCAD.ApplicationServices")
                    .AddImports("Autodesk.AutoCAD.DatabaseServices")
                    .AddImports("Autodesk.AutoCAD.EditorInput")
                    .AddImports("Autodesk.AutoCAD.Geometry");

                // Compile the script
                var script = CSharpScript.Create(code, scriptOptions, typeof(ScriptGlobals));
                var compilation = script.GetCompilation();
                var diagnostics = compilation.GetDiagnostics();

                // Check for compilation errors
                var errors = diagnostics.Where(d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error).ToList();
                if (errors.Any())
                {
                    result.Success = false;
                    result.ErrorResponse = new ScriptResponse
                    {
                        Success = false,
                        Error = "Compilation failed",
                        Diagnostics = errors.Select(e => e.ToString()).ToList()
                    };
                    return Task.FromResult(result);
                }

                // Success - script compiled
                result.Success = true;
                result.CompiledScript = script;

                // Add warnings to response if any
                var warnings = diagnostics.Where(d => d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Warning).ToList();
                if (warnings.Any())
                {
                    result.ErrorResponse = new ScriptResponse
                    {
                        Success = true,
                        Diagnostics = warnings.Select(w => w.ToString()).ToList()
                    };
                }
            }
            catch (CompilationErrorException ex)
            {
                result.Success = false;
                result.ErrorResponse = new ScriptResponse
                {
                    Success = false,
                    Error = "Compilation error",
                    Diagnostics = ex.Diagnostics.Select(d => d.ToString()).ToList()
                };
            }
            catch (System.Exception ex)
            {
                result.Success = false;
                result.ErrorResponse = new ScriptResponse
                {
                    Success = false,
                    Error = ex.Message
                };
            }

            return Task.FromResult(result);
        }

        private Task<ScriptResponse> ExecuteScriptOnUIThread(Script<object> script)
        {
            var tcs = new TaskCompletionSource<ScriptResponse>();

            // Marshal to AutoCAD UI thread using Idle event
            EventHandler idleHandler = null;
            idleHandler = (sender, e) =>
            {
                AcadApp.Idle -= idleHandler;

                // Execute on UI thread
                Task.Run(async () =>
                {
                    var response = await ExecuteCompiledScript(script);
                    tcs.SetResult(response);
                });
            };

            AcadApp.Idle += idleHandler;

            return tcs.Task;
        }

        private async Task<ScriptResponse> ExecuteCompiledScript(Script<object> script)
        {
            var response = new ScriptResponse();
            var outputCapture = new StringWriter();
            var originalOut = Console.Out;

            // Show visual feedback dialog
            ExecutionDialog executionDialog = null;

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

                // Show execution dialog on UI thread
                AcadApp.Idle += ShowExecutionDialog;

                void ShowExecutionDialog(object sender, EventArgs e)
                {
                    AcadApp.Idle -= ShowExecutionDialog;
                    executionDialog = new ExecutionDialog();
                    executionDialog.Show();
                }

                // Give dialog a moment to show
                await Task.Delay(100);

                // Create script globals with AutoCAD context
                var globals = new ScriptGlobals
                {
                    Doc = doc,
                    Db = db,
                    Ed = ed
                };

                // Execute the script with document lock
                using (var docLock = doc.LockDocument())
                {
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

                // Close execution dialog
                if (executionDialog != null && !executionDialog.IsDisposed)
                {
                    executionDialog.Invoke(new Action(() => executionDialog.Close()));
                }
            }

            return response;
        }

        private async Task SendHttpResponse(NetworkStream stream, ScriptResponse scriptResponse, int statusCode = 0, string statusText = null)
        {
            var jsonResponse = scriptResponse.ToJson();
            var responseBytes = Encoding.UTF8.GetBytes(jsonResponse);

            if (statusCode == 0)
            {
                statusCode = scriptResponse.Success ? 200 : 400;
            }

            if (statusText == null)
            {
                statusText = scriptResponse.Success ? "OK" : "Bad Request";
            }

            var httpResponse = $"HTTP/1.1 {statusCode} {statusText}\r\n" +
                             $"Content-Type: application/json\r\n" +
                             $"Content-Length: {responseBytes.Length}\r\n" +
                             $"\r\n";

            var headerBytes = Encoding.UTF8.GetBytes(httpResponse);

            await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
            await stream.WriteAsync(responseBytes, 0, responseBytes.Length);
            await stream.FlushAsync();
        }

        internal void LogMessage(string message)
        {
            if (dialog != null && !dialog.IsDisposed)
            {
                dialog.LogMessage(message);
            }
        }
    }

    public class BackgroundRoslynServerDialog : Form
    {
        private Label statusLabel;
        private TextBox tokenTextBox;
        private TextBox curlTextBox;
        private TextBox logTextBox;
        private Button stopButton;
        private BackgroundRoslynServer server;

        public BackgroundRoslynServerDialog(BackgroundRoslynServer server)
        {
            this.server = server;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Roslyn Server (Background) - AutoCAD Ballet";
            this.Size = new System.Drawing.Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new System.Drawing.Size(700, 500);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(15)
            };

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // Status
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));  // Token section
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 150)); // Curl example
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));  // Log header
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // Log area
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));  // Buttons

            // Status label
            statusLabel = new Label
            {
                Text = $"✓ Server running on {StartRoslynServerInBackground.URL}",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.DarkGreen
            };

            // Token section
            var tokenPanel = new Panel { Dock = DockStyle.Fill };
            var tokenLabel = new Label
            {
                Text = "Authentication Token:",
                Location = new System.Drawing.Point(0, 0),
                Size = new System.Drawing.Size(200, 20),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            };
            tokenTextBox = new TextBox
            {
                Text = server.AuthToken,
                Location = new System.Drawing.Point(0, 25),
                Width = 850,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.LightYellow,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            var copyTokenButton = new Button
            {
                Text = "Copy",
                Location = new System.Drawing.Point(0, 52),
                Size = new System.Drawing.Size(80, 25),
                Font = new System.Drawing.Font("Segoe UI", 9)
            };
            copyTokenButton.Click += (s, e) =>
            {
                Clipboard.SetText(server.AuthToken);
                copyTokenButton.Text = "Copied!";
                Task.Delay(2000).ContinueWith(_ =>
                {
                    if (!copyTokenButton.IsDisposed)
                    {
                        copyTokenButton.Invoke(new Action(() => copyTokenButton.Text = "Copy"));
                    }
                });
            };

            tokenPanel.Controls.Add(tokenLabel);
            tokenPanel.Controls.Add(tokenTextBox);
            tokenPanel.Controls.Add(copyTokenButton);

            // Curl example section
            var curlPanel = new Panel { Dock = DockStyle.Fill };
            var curlLabel = new Label
            {
                Text = "Sample curl command:",
                Location = new System.Drawing.Point(0, 0),
                Size = new System.Drawing.Size(200, 20),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            };

            var curlExample = $"curl -X POST {StartRoslynServerInBackground.URL} \\\n" +
                            $"  -H \"X-Auth-Token: {server.AuthToken}\" \\\n" +
                            "  -d 'Console.WriteLine(\"Hello from AutoCAD!\");'";

            curlTextBox = new TextBox
            {
                Text = curlExample,
                Location = new System.Drawing.Point(0, 25),
                Width = 850,
                Height = 100,
                Multiline = true,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 8),
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ScrollBars = ScrollBars.Vertical
            };

            var copyCurlButton = new Button
            {
                Text = "Copy",
                Location = new System.Drawing.Point(0, 128),
                Size = new System.Drawing.Size(80, 25),
                Font = new System.Drawing.Font("Segoe UI", 9),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            copyCurlButton.Click += (s, e) =>
            {
                Clipboard.SetText(curlTextBox.Text);
                copyCurlButton.Text = "Copied!";
                Task.Delay(2000).ContinueWith(_ =>
                {
                    if (!copyCurlButton.IsDisposed)
                    {
                        copyCurlButton.Invoke(new Action(() => copyCurlButton.Text = "Copy"));
                    }
                });
            };

            curlPanel.Controls.Add(curlLabel);
            curlPanel.Controls.Add(curlTextBox);
            curlPanel.Controls.Add(copyCurlButton);

            // Log header
            var logHeaderLabel = new Label
            {
                Text = "Activity Log:",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.BottomLeft,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            };

            // Log area
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

            // Button panel
            var buttonPanel = new Panel { Dock = DockStyle.Fill };
            stopButton = new Button
            {
                Text = "Stop Server",
                Size = new System.Drawing.Size(120, 35),
                Location = new System.Drawing.Point(0, 8),
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.FromArgb(220, 53, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            stopButton.FlatAppearance.BorderSize = 0;
            stopButton.Click += (s, e) =>
            {
                StartRoslynServerInBackground.StopServer();
                this.Close();
            };

            buttonPanel.Controls.Add(stopButton);

            // Add all to layout
            layout.Controls.Add(statusLabel, 0, 0);
            layout.Controls.Add(tokenPanel, 0, 1);
            layout.Controls.Add(curlPanel, 0, 2);
            layout.Controls.Add(logHeaderLabel, 0, 3);
            layout.Controls.Add(logTextBox, 0, 4);
            layout.Controls.Add(buttonPanel, 0, 5);

            this.Controls.Add(layout);

            // Log initial message
            LogMessage($"[{DateTime.Now:HH:mm:ss}] Server started and waiting for requests...");
            LogMessage($"[{DateTime.Now:HH:mm:ss}] All requests must include authentication token");

            this.FormClosing += (s, e) =>
            {
                // Don't stop server when dialog is closed, only when button is clicked
                // This allows the dialog to be closed while server continues running
            };
        }

        internal void LogMessage(string message)
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

    /// <summary>
    /// Brief dialog shown during script execution to provide visual feedback
    /// </summary>
    public class ExecutionDialog : Form
    {
        public ExecutionDialog()
        {
            this.Text = "AutoCAD Ballet - Executing Script";
            this.Size = new System.Drawing.Size(400, 120);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;

            var label = new Label
            {
                Text = "Executing Roslyn script...\n\nPlease wait.",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI", 11)
            };

            this.Controls.Add(label);
        }
    }
}
