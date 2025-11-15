# Roslyn Server for AI Agent Integration

AutoCAD Ballet provides a Roslyn compiler-as-a-service that allows AI agents (like Claude Code) to execute C# code within the AutoCAD session context.

## Purpose

- Query and analyze current AutoCAD session programmatically
- Perform validation tasks (layer naming, entity placement checks)
- Inspect document state and entity properties dynamically
- Support automated drawing quality checks and audits

## Quick Start

### 1. Start the Server

**Modal Server (blocks AutoCAD UI):**
```
start-roslyn-server
```
- Opens modal dialog showing server activity
- AutoCAD UI unavailable while server runs
- No authentication required
- Press ESC to stop
- Best for: Quick interactive sessions

**Background Server (non-blocking):**
```
start-roslyn-server-in-background
```
- Opens non-blocking dialog (can be closed)
- AutoCAD remains fully functional
- Requires authentication token (generated on startup)
- Server persists across document switches
- Use `stop-roslyn-server` command to stop
- Best for: Long-running automation, CI/CD integration

Both servers listen on `http://127.0.0.1:34157/`

### 2. Get Authentication Token (Background Server Only)

When starting the background server, AutoCAD displays:
```
Roslyn server started in background on http://127.0.0.1:34157/
Authentication token: K1FcjG0+tWv+eydypV8j4ZPupcOoOfDm34i9aPsE7vU=
```

Copy this token for use in your requests. The token is also displayed in the server dialog.

### 3. Send Queries

**To Modal Server (no authentication):**
```bash
curl -X POST http://127.0.0.1:34157/ -d 'Console.WriteLine("Hello from AutoCAD!");'
```

**To Background Server (with authentication):**
```bash
# Using X-Auth-Token header
curl -X POST http://127.0.0.1:34157/ \
  -H "X-Auth-Token: K1FcjG0+tWv+eydypV8j4ZPupcOoOfDm34i9aPsE7vU=" \
  -d 'Console.WriteLine("Hello from AutoCAD!");'

# Or using Authorization Bearer header
curl -X POST http://127.0.0.1:34157/ \
  -H "Authorization: Bearer K1FcjG0+tWv+eydypV8j4ZPupcOoOfDm34i9aPsE7vU=" \
  -d 'Console.WriteLine("Hello from AutoCAD!");'
```

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
curl -X POST http://127.0.0.1:34157/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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
curl -X POST http://127.0.0.1:34157/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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
curl -X POST http://127.0.0.1:34157/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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
curl -X POST http://127.0.0.1:34157/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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
curl -X POST http://127.0.0.1:34157/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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

1. **Verify server is running**: Check if port 34157 is listening
2. **Construct query**: Build appropriate C# code based on task
3. **Send request**: Use curl or equivalent HTTP client
4. **Parse JSON response**: Extract `success`, `output`, `error`, and `diagnostics`
5. **Iterate if needed**: If compilation fails, fix errors and retry

Example workflow:
```bash
# Check if server is running
if ! curl -s --connect-timeout 1 http://127.0.0.1:34157 > /dev/null 2>&1; then
    echo "Roslyn server not running. Run 'start-roslyn-server' in AutoCAD."
    exit 1
fi

# Send query and parse response (requires jq for JSON parsing)
RESPONSE=$(curl -s -X POST http://127.0.0.1:34157/ -d 'Console.WriteLine("Test");')
SUCCESS=$(echo $RESPONSE | jq -r '.success')
OUTPUT=$(echo $RESPONSE | jq -r '.output')

if [ "$SUCCESS" == "true" ]; then
    echo "Success: $OUTPUT"
fi
```

## Security Considerations

- Server only listens on **localhost (127.0.0.1)** - not accessible from network
- **Modal server**: No authentication (relies on localhost-only binding and blocking dialog)
- **Background server**: Requires authentication token (256-bit random, generated per session)
- Code executes with **full AutoCAD API access** - can read, modify, and delete drawing data
- Intended for **local development and automation only**
- Do not expose this service to untrusted networks
- Token is visible in AutoCAD command line and server dialog - keep AutoCAD session secure

## Type Identity Issue in .NET 8 (AutoCAD 2025-2026)

The background server (`start-roslyn-server-in-background`) addresses a critical .NET 8 type identity issue that occurs when assemblies are loaded with rewritten identities for hot-reloading support.

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
The background server uses a unified type approach:
1. At startup, search `AppDomain` for ANY loaded `autocad-ballet.dll` with a valid `Location`
2. Cache both the assembly AND the `ScriptGlobals` type from that instance
3. Use the cached assembly for Roslyn compilation references
4. Use the cached type via reflection (`Activator.CreateInstance`) to create globals instance
5. This ensures type identity consistency across compilation and execution

See `start-roslyn-server-in-background.cs` for detailed implementation.

## Implementation Details

- Uses `CommandFlags.Session` to work at application level
- Runs TCP server on background thread with cancellation token
- Marshals AutoCAD API access appropriately (executes on UI thread via `Application.Idle` event)
- Captures `Console.WriteLine()` output using `StringWriter`
- Returns both compilation diagnostics and runtime errors
- Keeps listening even after compilation errors (allows correction)
- Background server uses document locking for thread-safe database operations

For general AutoCAD Ballet documentation, see `CLAUDE.md`.
