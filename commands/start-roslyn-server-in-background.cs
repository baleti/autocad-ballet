using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Reflection;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
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
        internal const string URL = "https://127.0.0.1:34157/";

        // Global instance - note: this will be reset when assembly is hot-reloaded
        // We rely on port conflict detection and HTTP shutdown endpoint for cross-reload control
        private static BackgroundRoslynServer serverInstance = null;
        private static readonly object lockObject = new object();

        [CommandMethod("start-roslyn-server-in-background", CommandFlags.Session)]
        public void StartRoslynServerInBackgroundCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            lock (lockObject)
            {
                if (serverInstance != null && serverInstance.IsRunning)
                {
                    ed?.WriteMessage("\nRoslyn server is already running. Use 'stop-roslyn-server' to stop it.\n");
                    return;
                }

                try
                {
                    serverInstance = new BackgroundRoslynServer();
                    serverInstance.Start();
                }
                catch (SocketException sockEx) when (sockEx.SocketErrorCode == SocketError.AddressAlreadyInUse)
                {
                    // Port is already in use - server is running from a previous assembly load
                    ed?.WriteMessage($"\nRoslyn server is already running on port {PORT}. Use 'stop-roslyn-server' to stop it.\n");
                    serverInstance = null;
                }
                catch (System.Exception ex)
                {
                    ed?.WriteMessage($"\nError starting background Roslyn server: {ex.Message}\n");
                    serverInstance = null;
                }
            }
        }

        internal static void StopServer()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            // Send HTTPS shutdown request to the server
            // This works even when the server was started from a different assembly instance
            try
            {
                // For HTTPS with self-signed certificate, we need to bypass certificate validation
                var handler = new System.Net.Http.HttpClientHandler();
                handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;

                using (var client = new System.Net.Http.HttpClient(handler))
                {
                    client.Timeout = TimeSpan.FromSeconds(5);
                    client.DefaultRequestHeaders.Add("X-Shutdown-Key", "AUTOCAD_BALLET_SHUTDOWN");

                    ed?.WriteMessage("\nSending shutdown request to Roslyn server...\n");

                    var task = client.PostAsync(URL + "shutdown", new System.Net.Http.StringContent(""));
                    task.Wait(); // Block until complete

                    var response = task.Result;
                    response.EnsureSuccessStatusCode();

                    ed?.WriteMessage("\n========================================\n");
                    ed?.WriteMessage("Roslyn server stopped.\n");
                    ed?.WriteMessage("========================================\n");
                }
            }
            catch (System.AggregateException aggEx)
            {
                var innerEx = aggEx.InnerException;
                if (innerEx is System.Net.Http.HttpRequestException || innerEx is System.Net.Sockets.SocketException)
                {
                    // Server not running or already stopped
                    ed?.WriteMessage("\nNo Roslyn server is currently running.\n");
                }
                else
                {
                    ed?.WriteMessage($"\nError stopping server: {innerEx?.Message ?? aggEx.Message}\n");
                    ed?.WriteMessage($"Exception type: {innerEx?.GetType().Name ?? aggEx.GetType().Name}\n");
                }
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\nError stopping server: {ex.Message}\n");
                ed?.WriteMessage($"Exception type: {ex.GetType().Name}\n");
                ed?.WriteMessage($"Stack trace: {ex.StackTrace}\n");
            }

            // Clean up local instance if it exists
            lock (lockObject)
            {
                if (serverInstance != null)
                {
                    try
                    {
                        serverInstance.Stop();
                    }
                    catch { }
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
        private X509Certificate2 serverCertificate;
        private Assembly scriptGlobalsAssembly;  // Cached assembly with valid Location
        private Type scriptGlobalsType;           // Cached ScriptGlobals type from that assembly

        public bool IsRunning => isRunning;
        public string AuthToken => authToken;

        public void Start()
        {
            if (isRunning)
                return;

            // CRITICAL: Assembly Location Resolution for Roslyn Compilation
            // ============================================================
            // Problem: In .NET 8 (AutoCAD 2025/2026), assemblies loaded from modified bytes
            // (with rewritten identity via Mono.Cecil) do not have a file Location property.
            // Roslyn's ScriptOptions.AddReferences() requires assemblies with valid file paths.
            //
            // Solution: Search for any loaded copy of autocad-ballet.dll that HAS a Location.
            // Cache both the assembly AND the ScriptGlobals type from it. Use these cached
            // references for BOTH compilation (AddReferences) and execution (CreateInstance).
            //
            // Why this matters: .NET treats types from different assembly instances as incompatible
            // even if they have the same definition. Using the SAME assembly instance for both
            // compilation and execution prevents "Type A cannot be cast to Type B" errors.
            //
            // Future: When implementing a central server hub, this logic will need to ensure
            // all connected AutoCAD sessions use compatible assembly versions.

            var currentAsm = typeof(ScriptGlobals).Assembly;
            if (!string.IsNullOrEmpty(currentAsm.Location))
            {
                // Current assembly has a file location - use it directly
                scriptGlobalsAssembly = currentAsm;
            }
            else
            {
                // Current assembly was loaded from bytes - search AppDomain for a copy with Location
                var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
                foreach (var asm in loadedAssemblies)
                {
                    if (!string.IsNullOrEmpty(asm.Location))
                    {
                        try
                        {
                            var type = asm.GetType("AutoCADBallet.ScriptGlobals");
                            if (type != null)
                            {
                                scriptGlobalsAssembly = asm;
                                break;
                            }
                        }
                        catch { }
                    }
                }
            }

            if (scriptGlobalsAssembly == null)
            {
                throw new InvalidOperationException("Cannot find autocad-ballet assembly with valid file location for Roslyn compilation");
            }

            scriptGlobalsType = scriptGlobalsAssembly.GetType("AutoCADBallet.ScriptGlobals");
            if (scriptGlobalsType == null)
            {
                throw new InvalidOperationException("Cannot find ScriptGlobals type in assembly");
            }

            // Load or generate self-signed certificate for HTTPS
            try
            {
                serverCertificate = SslCertificateHelper.GetOrCreateCertificate();
            }
            catch (NotSupportedException ex)
            {
                // Certificate generation not supported on older .NET Framework
                throw new InvalidOperationException(
                    "SSL/TLS certificate generation failed.\n" + ex.Message);
            }

            // Generate secure random password (32 bytes = 256 bits)
            authToken = GenerateSecureToken();

            // Copy token to clipboard for user convenience
            try
            {
                Clipboard.SetText(authToken);
            }
            catch (System.Exception clipEx)
            {
                // Clipboard access may fail in some environments, but don't block server startup
                // Will notify user in the startup message if needed
                System.Diagnostics.Debug.WriteLine($"Failed to copy token to clipboard: {clipEx.Message}");
            }

            cancellationTokenSource = new CancellationTokenSource();
            listener = new TcpListener(IPAddress.Parse("127.0.0.1"), StartRoslynServerInBackground.PORT);
            listener.Start();
            isRunning = true;

            // Start background listener thread
            Task.Run(() => AcceptRequestsLoop(cancellationTokenSource.Token));

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            ed?.WriteMessage($"\n========================================\n");
            ed?.WriteMessage($"Roslyn server started in background (HTTPS)\n");
            ed?.WriteMessage($"URL: {StartRoslynServerInBackground.URL}\n");
            ed?.WriteMessage($"Authentication token: {authToken}\n");
            ed?.WriteMessage($"(Token copied to clipboard)\n");
            ed?.WriteMessage($"\nSample curl command:\n");
            ed?.WriteMessage($"curl --insecure -X POST {StartRoslynServerInBackground.URL} \\\n");
            ed?.WriteMessage($"  -H \"X-Auth-Token: {authToken}\" \\\n");
            ed?.WriteMessage($"  -d 'Console.WriteLine(\"Hello from AutoCAD!\");'\n");
            ed?.WriteMessage($"\nNote: Using self-signed certificate. Use --insecure (curl) or equivalent.\n");
            ed?.WriteMessage($"Use 'stop-roslyn-server' command to stop the server.\n");
            ed?.WriteMessage($"========================================\n");
        }

        public void Stop()
        {
            if (!isRunning)
                return;

            isRunning = false;

            // Cancel any pending operations
            try
            {
                cancellationTokenSource?.Cancel();
            }
            catch { }

            // Stop the listener
            try
            {
                listener?.Stop();
            }
            catch { }

            // Dispose of certificate
            try
            {
                serverCertificate?.Dispose();
                serverCertificate = null;
            }
            catch { }

            // Dispose of resources
            try
            {
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
            catch { }

            try
            {
                // Note: TcpListener doesn't implement IDisposable, but we null it out
                listener = null;
            }
            catch { }

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            ed?.WriteMessage("\n========================================\n");
            ed?.WriteMessage("Roslyn server stopped.\n");
            ed?.WriteMessage("========================================\n");
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
                        _ = Task.Run(async () => await HandleHttpsRequest(client, clientEndpoint, cancellationToken));
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

        private async Task HandleHttpsRequest(TcpClient client, string clientEndpoint, CancellationToken cancellationToken)
        {
            try
            {
                using (client)
                {
                    // Wrap TCP stream with SSL/TLS encryption
                    var networkStream = client.GetStream();
                    var sslStream = new SslStream(networkStream, false);

                    try
                    {
                        LogMessage($"[{DateTime.Now:HH:mm:ss}] Starting SSL/TLS handshake with {clientEndpoint}");

                        // Verify certificate has private key
                        if (!serverCertificate.HasPrivateKey)
                        {
                            LogMessage($"[{DateTime.Now:HH:mm:ss}] ERROR: Server certificate does not have a private key!");
                            return;
                        }

                        // Authenticate as server with the self-signed certificate
                        // Use TLS 1.2 only for maximum compatibility
#if NET8_0_OR_GREATER
                        // .NET 8 supports TLS 1.3
                        var protocols = SslProtocols.Tls12 | SslProtocols.Tls13;
#else
                        // .NET Framework 4.8 and earlier - use TLS 1.2 only
                        var protocols = SslProtocols.Tls12;
#endif
                        await sslStream.AuthenticateAsServerAsync(
                            serverCertificate,
                            clientCertificateRequired: false,
                            enabledSslProtocols: protocols,
                            checkCertificateRevocation: false
                        );

                        LogMessage($"[{DateTime.Now:HH:mm:ss}] SSL/TLS handshake completed with {clientEndpoint}");

                        // Now handle the HTTP request over the encrypted stream
                        await HandleHttpRequest(sslStream, clientEndpoint, cancellationToken);
                    }
                    catch (AuthenticationException authEx)
                    {
                        LogMessage($"[{DateTime.Now:HH:mm:ss}] SSL/TLS authentication failed: {authEx.Message}");
                        LogMessage($"[{DateTime.Now:HH:mm:ss}] Stack trace: {authEx.StackTrace}");
                    }
                    catch (System.Exception innerEx)
                    {
                        LogMessage($"[{DateTime.Now:HH:mm:ss}] SSL handshake error: {innerEx.GetType().Name}: {innerEx.Message}");
                        LogMessage($"[{DateTime.Now:HH:mm:ss}] Stack trace: {innerEx.StackTrace}");
                    }
                    finally
                    {
                        try
                        {
                            sslStream?.Dispose();
                        }
                        catch { }
                    }
                }
            }
            catch (System.Exception ex)
            {
                LogMessage($"[{DateTime.Now:HH:mm:ss}] Error handling HTTPS request from {clientEndpoint}: {ex.Message}");
                LogMessage($"[{DateTime.Now:HH:mm:ss}] Stack trace: {ex.StackTrace}");
            }
        }

        private async Task HandleHttpRequest(Stream stream, string clientEndpoint, CancellationToken cancellationToken)
        {
            try
            {
                using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
                {
                    // Parse HTTP request on background thread (no AutoCAD interaction)
                    var requestLine = await reader.ReadLineAsync();

                    // Extract the request path from the request line (e.g., "POST /shutdown HTTP/1.1")
                    string requestPath = "/";
                    if (!string.IsNullOrEmpty(requestLine))
                    {
                        var parts = requestLine.Split(' ');
                        if (parts.Length >= 2)
                        {
                            requestPath = parts[1];
                        }
                    }

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

                    // Handle shutdown endpoint (no authentication required for shutdown from localhost)
                    if (requestPath.TrimEnd('/').Equals("/shutdown", StringComparison.OrdinalIgnoreCase))
                    {
                        if (headers.ContainsKey("X-Shutdown-Key") && headers["X-Shutdown-Key"] == "AUTOCAD_BALLET_SHUTDOWN")
                        {
                            LogMessage($"[{DateTime.Now:HH:mm:ss}] Shutdown request received from {clientEndpoint}");

                            var shutdownResponse = new ScriptResponse
                            {
                                Success = true,
                                Output = "Server shutting down..."
                            };
                            await SendHttpResponse(stream, shutdownResponse);

                            // Stop the server on a background task to allow response to be sent
                            _ = Task.Run(() =>
                            {
                                System.Threading.Thread.Sleep(100); // Give time for response to be sent
                                Stop();
                            });
                            return;
                        }
                        else
                        {
                            var errorResponse = new ScriptResponse
                            {
                                Success = false,
                                Error = "Invalid shutdown key"
                            };
                            await SendHttpResponse(stream, errorResponse, 403, "Forbidden");
                            return;
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
                // Configure Roslyn script compilation options
                // CRITICAL: Use cached scriptGlobalsAssembly (has valid Location) instead of
                // typeof(ScriptGlobals).Assembly (may lack Location in .NET 8 with identity rewriting)
                var scriptOptions = ScriptOptions.Default
                    .AddReferences(scriptGlobalsAssembly)             // autocad-ballet.dll for ScriptGlobals
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

                // Compile the script using the cached ScriptGlobals type
                // This type MUST match the one used in ExecuteCompiledScript() for type identity
                var script = CSharpScript.Create(code, scriptOptions, scriptGlobalsType);
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

                // CRITICAL: Create ScriptGlobals instance using cached type
                // =========================================================
                // Why use reflection instead of "new ScriptGlobals()"?
                // - The cached scriptGlobalsType is from an assembly with a valid Location
                // - Using "new ScriptGlobals()" would create an instance from a DIFFERENT assembly
                //   (the current one, which may lack a Location)
                // - Roslyn compiles the script against scriptGlobalsType (from cached assembly)
                // - If we pass an instance from a different assembly, .NET throws:
                //   "Type A cannot be cast to Type B" even though they're the same type
                // - Solution: Use Activator.CreateInstance + reflection to ensure type identity match
                var globals = Activator.CreateInstance(scriptGlobalsType);

                // Set AutoCAD context properties via reflection
                var docProp = scriptGlobalsType.GetProperty("Doc");
                var dbProp = scriptGlobalsType.GetProperty("Db");
                var edProp = scriptGlobalsType.GetProperty("Ed");

                docProp.SetValue(globals, doc);
                dbProp.SetValue(globals, db);
                edProp.SetValue(globals, ed);

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

        private async Task SendHttpResponse(Stream stream, ScriptResponse scriptResponse, int statusCode = 0, string statusText = null)
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
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            ed?.WriteMessage($"{message}\n");
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

    /// <summary>
    /// Helper class for generating and managing self-signed SSL certificates for HTTPS server
    /// </summary>
    internal static class SslCertificateHelper
    {
        private static readonly string CertificateStorePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "autocad-ballet",
            "runtime",
            "roslyn-server-cert.pfx"
        );

        private const string CertificatePassword = "autocad-ballet-localhost";

        /// <summary>
        /// Gets or creates a self-signed certificate for localhost HTTPS
        /// Certificate is cached in %APPDATA%/autocad-ballet/runtime/ for reuse
        /// </summary>
        public static X509Certificate2 GetOrCreateCertificate()
        {
            // Try to load existing certificate from disk
            if (File.Exists(CertificateStorePath))
            {
                try
                {
                    var cert = new X509Certificate2(CertificateStorePath, CertificatePassword,
                        X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);

                    // Check if certificate is still valid (not expired)
                    if (cert.NotAfter > DateTime.Now.AddDays(30))
                    {
                        return cert;
                    }

                    // Certificate is expiring soon, generate a new one
                    cert.Dispose();
                }
                catch
                {
                    // Failed to load cert, will generate new one
                }
            }

            // Generate new self-signed certificate
            var certificate = GenerateSelfSignedCertificate();

            // Save to disk for future use
            try
            {
                var certDirectory = Path.GetDirectoryName(CertificateStorePath);
                if (!Directory.Exists(certDirectory))
                {
                    Directory.CreateDirectory(certDirectory);
                }

                var certBytes = certificate.Export(X509ContentType.Pfx, CertificatePassword);
                File.WriteAllBytes(CertificateStorePath, certBytes);
            }
            catch
            {
                // Failed to save, but we can still use the in-memory certificate
            }

            return certificate;
        }

        private static X509Certificate2 GenerateSelfSignedCertificate()
        {
#if NET48 || NET8_0_OR_GREATER
            // Use CertificateRequest API (available in .NET Framework 4.7.1+, .NET Core 2.0+)
            // For AutoCAD 2021-2024 (net48) and 2025-2026 (net8.0)

            // IMPORTANT: Don't use 'using' on RSA - we need to keep it alive through export/import
            var rsa = RSA.Create(2048);
            try
            {
                var request = new CertificateRequest(
                    "CN=localhost",
                    rsa,
                    HashAlgorithmName.SHA256,
                    RSASignaturePadding.Pkcs1
                );

                // Add Subject Alternative Name for localhost
                var sanBuilder = new SubjectAlternativeNameBuilder();
                sanBuilder.AddDnsName("localhost");
                sanBuilder.AddIpAddress(System.Net.IPAddress.Loopback);
                sanBuilder.AddIpAddress(System.Net.IPAddress.Parse("127.0.0.1"));
                request.CertificateExtensions.Add(sanBuilder.Build());

                // Add Key Usage extension
                request.CertificateExtensions.Add(
                    new X509KeyUsageExtension(
                        X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
                        critical: true
                    )
                );

                // Add Enhanced Key Usage for Server Authentication
                request.CertificateExtensions.Add(
                    new X509EnhancedKeyUsageExtension(
                        new OidCollection { new Oid("1.3.6.1.5.5.7.3.1") }, // Server Authentication
                        critical: true
                    )
                );

                // Create self-signed certificate (valid for 365 days)
                // The certificate will contain a copy of the private key
                var certificate = request.CreateSelfSigned(
                    DateTimeOffset.Now.AddDays(-1),
                    DateTimeOffset.Now.AddDays(365)
                );

                // Export to PFX and re-import to ensure private key is fully embedded and exportable
                // This is necessary for the certificate to work with SslStream
                var certBytes = certificate.Export(X509ContentType.Pfx, CertificatePassword);

                // Dispose the original certificate before creating the new one
                certificate.Dispose();

                // Import with flags that ensure private key is accessible
                var finalCertificate = new X509Certificate2(certBytes, CertificatePassword,
                    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.MachineKeySet);

                return finalCertificate;
            }
            finally
            {
                // Now we can safely dispose the RSA key
                rsa?.Dispose();
            }
#else
            // Fallback for older .NET Framework versions (net46, net47)
            // For AutoCAD 2017-2020

            // Since CertificateRequest is not available, we'll use a simpler approach:
            // Create a certificate using makecert-style parameters with X509Certificate2

            throw new NotSupportedException(
                "Self-signed certificate generation is not supported on .NET Framework 4.6/4.7.\n" +
                "For AutoCAD 2017-2020, please generate a certificate manually using:\n" +
                "  PowerShell: New-SelfSignedCertificate -DnsName 'localhost' -CertStoreLocation 'Cert:\\CurrentUser\\My'\n" +
                "Then export it to: " + CertificateStorePath
            );
#endif
        }

        /// <summary>
        /// Deletes the cached certificate file
        /// </summary>
        public static void DeleteCachedCertificate()
        {
            try
            {
                if (File.Exists(CertificateStorePath))
                {
                    File.Delete(CertificateStorePath);
                }
            }
            catch
            {
                // Ignore errors
            }
        }
    }
}
