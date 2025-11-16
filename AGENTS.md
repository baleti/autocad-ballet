# Roslyn Server for AI Agent Integration

> **Documentation Size Limit**: This file must stay under 20KB. Keep changes concise and avoid approaching this limit to allow for future additions.

AutoCAD Ballet provides a Roslyn compiler-as-a-service that allows AI agents (like Claude Code) to execute C# code within the AutoCAD session context.

## Purpose

- Query and analyze current AutoCAD session programmatically
- Perform validation tasks (layer naming, entity placement checks)
- Inspect document state and entity properties dynamically
- Support automated drawing quality checks and audits
- Enable cross-session network queries for multi-AutoCAD workflows

## Architecture

AutoCAD Ballet uses a **peer-to-peer network architecture**:

- Each AutoCAD instance runs its own HTTPS server
- Servers auto-start when AutoCAD loads (no manual command needed)
- Port allocation: 34157-34257 (automatic first-available selection)
- Shared token authentication across all sessions
- File-based service discovery via `%appdata%/autocad-ballet/runtime/network/`

### Network Registry Structure

```
%appdata%/autocad-ballet/runtime/network/
  ├── sessions      # CSV file with session info (SessionId, Port, Hostname, ProcessId, etc.)
  └── token         # Shared authentication token used by all sessions
```

**Design Decision: Shared Token**
- All AutoCAD sessions in the network share a single authentication token
- First session to start generates the token and stores it in `network/token`
- Subsequent sessions read and reuse this token
- Token regenerates when all sessions close and a new one starts
- Simplifies authentication: only one token to manage
- Enables seamless cross-session communication

## Quick Start

### 1. Server Auto-Start

The server **automatically starts** when AutoCAD loads:
- No manual command needed
- Finds first available port in range 34157-34257
- Registers in network registry (`%appdata%/autocad-ballet/runtime/network/sessions`)
- Generates or reads shared authentication token

Check server status in AutoCAD command line:
```
[AutoCAD Ballet] Server started on port 34157 (Session: 12ab34cd)
```

### 2. Get Authentication Token

The shared authentication token is stored at:
```
%appdata%/autocad-ballet/runtime/network/token
```

Read it with:
```bash
cat %appdata%/autocad-ballet/runtime/network/token
```

Or in the repository:
```bash
cat runtime/network/token
```

Example token:
```
K1FcjG0+tWv+eydypV8j4ZPupcOoOfDm34i9aPsE7vU=
```

### 3. Send Queries

The server provides four main endpoints:
- **`/roslyn`** - Execute C# scripts in AutoCAD context
- **`/select-by-categories`** - Query entity categories (used by `select-by-categories-in-network`)
- **`/screenshot`** - Capture screenshot of AutoCAD window
- **`/invoke-addin-command`** - Hot-reload and test modified commands without restarting AutoCAD

**Execute Roslyn Script:**
```bash
# Using X-Auth-Token header
curl -k -X POST https://127.0.0.1:34157/roslyn \
  -H "X-Auth-Token: K1FcjG0+tWv+eydypV8j4ZPupcOoOfDm34i9aPsE7vU=" \
  -d 'Console.WriteLine("Hello from AutoCAD!");'

# Or using Authorization Bearer header
curl -k -X POST https://127.0.0.1:34157/roslyn \
  -H "Authorization: Bearer K1FcjG0+tWv+eydypV8j4ZPupcOoOfDm34i9aPsE7vU=" \
  -d 'Console.WriteLine("Hello from AutoCAD!");'
```

**Note:** Use `-k` flag with curl to accept self-signed SSL certificates (localhost only).

### 4. Parse Response

The server returns JSON with this structure:

```json
{
  "success": true,
  "output": "Hello from AutoCAD!\n",
  "error": "",
  "diagnostics": []
}
```

## Available Context

Scripts have access to these global variables:

- **`Doc`** - Current active AutoCAD document
- **`Db`** - Current document's database
- **`Ed`** - Current document's editor

Pre-imported namespaces:
- `System`, `System.Linq`, `System.Collections.Generic`
- `Autodesk.AutoCAD.ApplicationServices`
- `Autodesk.AutoCAD.DatabaseServices`
- `Autodesk.AutoCAD.EditorInput`
- `Autodesk.AutoCAD.Geometry`

## Common Query Patterns

### Listing Information

**Get All Layer Names:**
```bash
TOKEN=$(cat runtime/network/token)
curl -k -X POST https://127.0.0.1:34157/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var layerTable = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
    foreach (ObjectId layerId in layerTable)
    {
        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
        Console.WriteLine(layer.Name);
    }
    tr.Commit();
}'
```

**Count Entities by Type:**
```bash
TOKEN=$(cat runtime/network/token)
curl -k -X POST https://127.0.0.1:34157/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var blockTable = (BlockTable)tr.GetObject(Db.BlockTableId, OpenMode.ForRead);
    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
    var typeCounts = new Dictionary<string, int>();

    foreach (ObjectId entId in modelSpace)
    {
        var entity = tr.GetObject(entId, OpenMode.ForRead) as Entity;
        var typeName = entity.GetType().Name;
        if (!typeCounts.ContainsKey(typeName))
            typeCounts[typeName] = 0;
        typeCounts[typeName]++;
    }

    foreach (var kvp in typeCounts.OrderByDescending(x => x.Value))
    {
        Console.WriteLine($"{kvp.Key}: {kvp.Value}");
    }
    tr.Commit();
}'
```

### Validation Queries

**Check Layer Naming Compliance:**
```bash
TOKEN=$(cat runtime/network/token)
curl -k -X POST https://127.0.0.1:34157/roslyn \
  -H "X-Auth-Token: $TOKEN"  -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var layerTable = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
    var nonCompliant = new List<string>();
    var validPrefixes = new[] { "A-", "S-", "E-", "M-", "P-", "FP-", "C-", "L-", "I-" };

    foreach (ObjectId layerId in layerTable)
    {
        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
        if (!validPrefixes.Any(prefix => layer.Name.StartsWith(prefix)))
        {
            nonCompliant.Add(layer.Name);
        }
    }

    if (nonCompliant.Count > 0)
    {
        Console.WriteLine($"Found {nonCompliant.Count} non-compliant layers:");
        foreach (var name in nonCompliant)
        {
            Console.WriteLine($"  - {name}");
        }
    }
    else
    {
        Console.WriteLine("All layers follow naming conventions!");
    }
    tr.Commit();
}'
```

**Check for Entities on Layer 0:**
```bash
TOKEN=$(cat runtime/network/token)
curl -k -X POST https://127.0.0.1:34157/roslyn \
  -H "X-Auth-Token: $TOKEN"  -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var blockTable = (BlockTable)tr.GetObject(Db.BlockTableId, OpenMode.ForRead);
    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
    var count = 0;

    foreach (ObjectId entId in modelSpace)
    {
        var entity = tr.GetObject(entId, OpenMode.ForRead) as Entity;
        if (entity.Layer == "0")
        {
            Console.WriteLine($"{entity.GetType().Name} at ({entity.GeometricExtents.MinPoint.X:F2}, {entity.GeometricExtents.MinPoint.Y:F2})");
            count++;
        }
    }

    Console.WriteLine(count > 0 ? $"\nWarning: {count} entities on Layer 0" : "No entities on Layer 0");
    tr.Commit();
}'
```

### Statistical Queries

**Get Drawing Statistics:**
```bash
TOKEN=$(cat runtime/network/token)
curl -k -X POST https://127.0.0.1:34157/roslyn \
  -H "X-Auth-Token: $TOKEN"  -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var blockTable = (BlockTable)tr.GetObject(Db.BlockTableId, OpenMode.ForRead);
    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
    var layerTable = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);

    Console.WriteLine("Drawing Statistics:");
    Console.WriteLine($"  Total Layers: {layerTable.Cast<ObjectId>().Count()}");
    Console.WriteLine($"  Total Blocks: {blockTable.Cast<ObjectId>().Count(id => !((BlockTableRecord)tr.GetObject(id, OpenMode.ForRead)).IsLayout)}");
    Console.WriteLine($"  Model Space Entities: {modelSpace.Cast<ObjectId>().Count()}");

    var layoutDict = (DBDictionary)tr.GetObject(Db.LayoutDictionaryId, OpenMode.ForRead);
    var layoutCount = layoutDict.Cast<DBDictionaryEntry>().Count(e => ((Layout)tr.GetObject(e.Value, OpenMode.ForRead)).LayoutName != "Model");
    Console.WriteLine($"  Paper Space Layouts: {layoutCount}");

    tr.Commit();
}'
```

### Screenshot Capture

**Capture AutoCAD Window Screenshot:**
```bash
TOKEN=$(cat runtime/network/token)
curl -k -X POST https://127.0.0.1:34157/screenshot \
  -H "X-Auth-Token: $TOKEN"
```

**Response:**
```json
{
  "Success": true,
  "Output": "C:\\Users\\YourName\\AppData\\Roaming\\autocad-ballet\\runtime\\screenshots\\session-id_20231116_143022_456.png",
  "Error": null,
  "Diagnostics": []
}
```

**Features:**
- Captures entire AutoCAD window including toolbars, command line, and drawing area
- Screenshots saved to `%appdata%/autocad-ballet/runtime/screenshots/`
- Filename format: `{sessionId}_{timestamp}.png`
- Auto-cleanup: Keeps last 20 screenshots, older ones deleted automatically
- Returns absolute file path for immediate use

**Use Cases:**
- Visual debugging of command execution
- Verify layout/space switching
- Inspect entity visibility and zoom levels
- Capture error dialogs or UI state
- Document drawing state for analysis

### Hot-Reload Command Testing

**Invoke Modified Command (Hot-Reload):**
```bash
TOKEN=$(cat runtime/network/token)

# Invoke specific command (hot-reloads from DLL if modified)
curl -k -X POST https://127.0.0.1:34157/invoke-addin-command \
  -H "X-Auth-Token: $TOKEN" \
  -d 'my-command-name'

# Invoke last command (uses runtime/invoke-addin-command-last)
curl -k -X POST https://127.0.0.1:34157/invoke-addin-command \
  -H "X-Auth-Token: $TOKEN" \
  -d ''
```

**Response:**
```json
{
  "Success": true,
  "Output": "Invoked command: my-command-name (as my-command-name-20231116143022)",
  "Error": null,
  "Diagnostics": []
}
```

**How It Works:**
- Loads command from latest `autocad-ballet.dll` with unique suffix
- Caches loaded commands (same timestamp/size = same cache)
- Executes via AutoCAD's command pipeline
- No AutoCAD restart needed - rebuild DLL and invoke immediately

**Use Cases:**
- Test command changes without restarting AutoCAD
- Automated testing of modified commands
- Rapid iteration during development
- Verify fixes applied correctly

## Error Handling

### Compilation Errors

If code has syntax errors, the server returns them:

```json
{
  "success": false,
  "output": "",
  "error": "Compilation failed",
  "diagnostics": [
    "(1,23): error CS1002: ; expected"
  ]
}
```

The server **continues listening** after errors, so you can send corrected code.

### Runtime Errors

If code compiles but fails during execution:

```json
{
  "success": false,
  "output": "Any output before the error\n",
  "error": "System.NullReferenceException: Object reference not set...",
  "diagnostics": []
}
```

## Best Practices

1. **Always Use Transactions**: AutoCAD database access requires transactions
   ```csharp
   using (var tr = Db.TransactionManager.StartTransaction())
   {
       // Your code here
       tr.Commit();
   }
   ```

2. **Use Console.WriteLine for Output**: All output should go through `Console.WriteLine()`

3. **Handle Null Values**: Always check for null when working with AutoCAD objects

4. **Format Output Consistently**: Make output easy to parse programmatically
   ```csharp
   Console.WriteLine($"LAYER|{layer.Name}|{layer.Color.ColorIndex}");
   ```

5. **Return Values**: Scripts can return values that will be appended to output
   ```csharp
   return layerTable.Cast<ObjectId>().Count();
   ```

## Integration with Claude Code

When Claude Code needs to query the AutoCAD session:

1. **Read shared token**: Load token from `%appdata%/autocad-ballet/runtime/network/token`
2. **Discover active sessions**: Read network registry from `%appdata%/autocad-ballet/runtime/network/sessions`
3. **Construct query**: Build appropriate C# code based on task
4. **Send request**: POST to `/roslyn` endpoint with authentication
5. **Parse JSON response**: Extract `success`, `output`, `error`, and `diagnostics`
6. **Iterate if needed**: If compilation fails, fix errors and retry

Example workflow:
```bash
# Read authentication token
TOKEN_FILE="runtime/network/token"
if [ ! -f "$TOKEN_FILE" ]; then
    echo "No authentication token found. AutoCAD Ballet server may not be running."
    exit 1
fi
TOKEN=$(cat "$TOKEN_FILE")

# Find first available session port from registry
SESSIONS_FILE="runtime/network/sessions"
if [ ! -f "$SESSIONS_FILE" ]; then
    echo "No active sessions found."
    exit 1
fi

# Extract first port from CSV (skip comment lines)
PORT=$(grep -v '^#' "$SESSIONS_FILE" | head -1 | cut -d',' -f2)

# Send query and parse response (requires jq for JSON parsing)
RESPONSE=$(curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
  -H "X-Auth-Token: $TOKEN" \
  -d 'Console.WriteLine("Test");')
SUCCESS=$(echo $RESPONSE | jq -r '.success')
OUTPUT=$(echo $RESPONSE | jq -r '.output')

if [ "$SUCCESS" == "true" ]; then
    echo "Success: $OUTPUT"
fi
```

## Network Queries

The `select-by-categories-in-network` command demonstrates cross-session queries:

1. Reads network registry to find all active sessions
2. Queries each peer's `/select-by-categories` endpoint
3. Combines results in DataGrid with session column
4. Saves combined selection for cross-document workflows

**Example usage in AutoCAD:**
```
select-by-categories-in-network
```

This enables powerful multi-session workflows where selections span multiple AutoCAD instances and documents.

## Security Considerations

- **HTTPS with TLS 1.2/1.3**: All communication is encrypted using self-signed certificates
- **Localhost-only binding**: Servers only listen on 127.0.0.1 - not accessible from network
- **Shared token authentication**: 256-bit random token shared across all sessions
- **Token storage**: `%appdata%/autocad-ballet/runtime/network/token` (plain text file)
- **Code execution**: Scripts have **full AutoCAD API access** - can read, modify, and delete drawing data
- **Intended use**: Local development and automation only
- **Do not**:
  - Expose this service to untrusted networks
  - Share the authentication token with untrusted parties
  - Run AutoCAD with elevated privileges when using the server
- **Token visibility**: First session to start displays token in command line
- **Token lifecycle**: Regenerates when all sessions close and a new session starts

## Type Identity Issue in .NET 8 (AutoCAD 2025-2026)

The auto-started server addresses a critical .NET 8 type identity issue that occurs when assemblies are loaded with rewritten identities for hot-reloading support.

**Problem:**
- Assemblies loaded from modified bytes (via Mono.Cecil) don't have a file `Location` property
- Roslyn's `AddReferences()` requires assemblies with valid file paths
- .NET treats types from different assembly instances as incompatible, causing cast exceptions

**Symptoms:**
```
Cannot reference assembly without location
Type A cannot be cast to Type B (both are ScriptGlobals)
```

**Solution:**
The server uses a unified type approach:
1. At startup, search `AppDomain` for ANY loaded `autocad-ballet.dll` with a valid `Location`
2. Cache both the assembly AND the `ScriptGlobals` type from that instance
3. Use the cached assembly for Roslyn compilation references
4. Use the cached type via reflection (`Activator.CreateInstance`) to create globals instance
5. This ensures type identity consistency across compilation and execution

See `commands/_server.cs` for detailed implementation.

## Implementation Details

- **Auto-start mechanism**: Initialized by `AutoCADBalletStartup` extension application in `commands/_startup.cs`
- **TCP/TLS server**: Runs on background thread with cancellation token
- **Thread safety**: Marshals AutoCAD API access to UI thread via `Application.Idle` event
- **Output capture**: Uses `StringWriter` to capture `Console.WriteLine()` output
- **Error handling**: Returns both compilation diagnostics and runtime errors
- **Continuous operation**: Keeps listening even after compilation errors (allows correction)
- **Document locking**: Uses `doc.LockDocument()` for thread-safe database operations
- **Network registry**: Heartbeat every 30 seconds, 2-minute timeout for dead session cleanup
- **Port allocation**: Searches range 34157-34257 for first available port

For general AutoCAD Ballet documentation, see `CLAUDE.md`.
