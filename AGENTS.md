# Repository Guidelines

## Project Structure & Module Organization
- `commands/` — C# AutoCAD plugin and AutoLISP utilities (`*.lsp`).
  - Multi-version targeting via `AutoCADYear` in `autocad-ballet.csproj`.
- `installer/` — Windows Forms installer that bundles built artifacts.
- `runtime/` — **IMPORTANT**: This codebase IS located at `%APPDATA%/autocad-ballet/`, so the `runtime/` folder contains **live usage data** and can be examined to verify code behavior and debug issues. Use `ls -lt runtime/switch-view-logs/ | head` to check recent log activity.
- Key files: `commands/autocad-ballet.csproj`, `commands/_selection.cs`, `commands/_switch-view-logging.cs`, `commands/_reopen-documents.cs`, `commands/aliases.lsp`, `installer/installer.csproj`.

## Build, Test, and Development Commands
- Build plugin (default 2026): `dotnet build commands/autocad-ballet.csproj`
- Build for a specific AutoCAD version: `dotnet build commands/autocad-ballet.csproj -p:AutoCADYear=2024`
  - Output: `commands/bin/{AutoCADYear}/autocad-ballet.dll`
- Build installer (Windows only): `dotnet build installer/installer.csproj`
  - Output: `installer/bin/Release/installer.exe` (or `Debug/installer.exe` for debug builds)
- Manual smoke test in AutoCAD: copy `.lsp` to `%APPDATA%/autocad-ballet/`, NETLOAD the DLL, run aliases from `aliases.lsp`.

## Coding Style & Naming Conventions
- C#: 4-space indent; PascalCase for types/methods; camelCase for locals/fields.
- Commands: one command per file (kebab-case naming); create under `commands/` with `[CommandMethod]` attribute.
- **CRITICAL**: Every command class MUST have `[assembly: CommandClass(typeof(YourClassName))]` at the top of the file (after `using` statements, before `namespace`).
  - Without this assembly-level attribute, AutoCAD will not register the command and it will be invisible.
  - Example: `[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedNamedInView))]`
- Shared utilities: files prefixed with underscore (e.g., `_selection.cs`, `_edit-dialog.cs`) may contain multiple related classes.
- AutoLISP: filenames kebab-case (e.g., `select-current-layer.lsp`); define commands as `c:command-name`; maintain shortcuts in `aliases.lsp`.
- Keep logic small and composable; avoid UI in core command classes.
- Scope-based commands: Main implementation file with `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()` static methods; separate stub files for each scope (`{command}-in-view.cs`, etc.) delegate to main implementation.

## DataGrid UI Component
**CRITICAL**: When asked to "use datagrid" or "show a datagrid", ALWAYS use the built-in `CustomGUIs.DataGrid()` method from `commands/datagrid/` — **DO NOT** create custom Windows Forms with DataGridView manually.

**Usage Pattern:**
```csharp
// 1. Build data as List<Dictionary<string, object>>
var data = new List<Dictionary<string, object>>();
data.Add(new Dictionary<string, object>
{
    ["PropertyName"] = value,
    ["AnotherProperty"] = anotherValue
});

// 2. Define columns to display
var columns = new List<string> { "PropertyName", "AnotherProperty" };

// 3. Show DataGrid and get user selection
var selected = CustomGUIs.DataGrid(data, columns, spanAllScreens: false);

// 4. Process selected items
if (selected != null && selected.Count > 0)
{
    foreach (var item in selected)
    {
        var value = item["PropertyName"];
        // ... process selection
    }
}
```

**Key Features:**
- Multi-select with Ctrl/Shift+Click, search/filter, column sorting, F2 edit mode
- Automatic column header formatting (PascalCase → "lowercase with spaces")
- Supports `allowCreateFromSearch: true` for creating new items from search box
- Virtual scrolling for large datasets, multi-monitor spanning

**Example Commands:**
- `switch-view.cs` — Layout switching with DataGrid
- `filter-selection.cs` — Entity filtering
- `tag-selected-in-view.cs` — Tag management with create-from-search

See `CLAUDE.md` DataGrid System section for full API documentation.

## CommandFlags.Session and LISP Integration
- **Use `CommandFlags.Session`** for commands that open/close/switch documents or modify DocumentCollection.
- **CRITICAL**: LISP wrappers for Session commands MUST end with `(princ)` to prevent command context leaks.
  - Pattern: `(defun c:alias () (command "command-name") (princ))`
  - Without `(princ)`, commands remain stuck on documents causing "Drawing is busy" errors.
- Examples: `open-documents-recent-read-write`, `open-documents-recent-read-only`, `close-all-without-saving`, `switch-view`, `switch-view-last`, `switch-view-recent`.

## Testing Guidelines
- No formal test suite yet. Validate by:
  - Running commands on sample drawings and checking behavior.
  - **Debugging workflow**: Since this codebase is at `%APPDATA%/autocad-ballet/`, examine `runtime/` logs directly:
    - `runtime/switch-view-logs/` — Layout switching history (use `ls -lt runtime/switch-view-logs/ | head` for recent activity)
    - `runtime/selection/` — Per-document selection persistence
    - `runtime/selection-logs/` — Selection change history
    - `runtime/invoke-addin-command-last` — Last invoked addin command
  - NETLOAD checks across targeted AutoCAD years (2017-2026).
- For bug fixes, include repro steps and expected vs. actual results in the PR.

## Commit & Pull Request Guidelines
- Commits: imperative, concise, lowercase (e.g., `fix startup loader path match`).
- PRs: clear description, scope, targeted AutoCAD years, manual test notes, and screenshots for UI/installer changes.
- Link issues and note when porting from revit-ballet.

## Security & Configuration Tips
- Never hardcode user paths; use `%APPDATA%/autocad-ballet/…` for data and runtime files.
- Do not commit runtime artifacts; the `runtime/` directory in the repository is gitignored and contains only local development/test data.

## Persistent Storage Best Practices

**CRITICAL**: Do not store persistent data in static fields or class-level variables in plugin code - these values will be lost when the DLL is reloaded via `invoke-addin-command` or hot-reload mechanisms.

**Problem**:
- Static fields are reset when AutoCAD reloads the plugin DLL
- Transient graphics and other runtime state persists in AutoCAD's memory but tracking data in static variables is lost
- This causes commands to fail after reload (e.g., clear commands can't find what to clear)

**Solution**:
- Store all persistent data in `%APPDATA%/autocad-ballet/runtime/` subdirectories
- Use file-based storage for tracking information that must survive DLL reloads
- Examples of proper storage locations:
  - `runtime/transient-graphics-tracker/` - Transient graphics marker tracking
  - `runtime/selection/` - Cross-document selection persistence
  - `runtime/switch-view-logs/` - Layout switching history
  - `runtime/selection-logs/` - Selection change history

**Example Pattern**:
```csharp
// BAD - Data lost on DLL reload
private static List<int> _markers = new List<int>();

// GOOD - Data persists across reloads
public static void SaveMarkers(string documentId, List<int> markers)
{
    var filePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "autocad-ballet", "runtime", "transient-graphics-tracker", documentId);
    File.WriteAllLines(filePath, markers.Select(m => m.ToString()));
}
```

**When to Use Runtime Storage**:
- Transient graphics tracking (markers must persist to clear them later)
- Cross-session data (selection sets, command history)
- Document-specific state that survives plugin reloads
- Any data that needs to outlive the current plugin load session

**See Also**:
- `commands/transient-graphics-storage.cs` - Reference implementation
- `commands/_selection.cs` - Selection persistence example

## Entity Attributes and Tagging System
AutoCAD Ballet provides a type-safe, extensible attribute storage system for entities (`commands/_entity-attributes.cs`).

### Overview
Tags are a specialized form of entity attributes stored in extension dictionaries. They provide a flexible way to organize, categorize, and select entities across views, documents, and sessions.

### Architecture
- **Storage**: Tags stored in entity extension dictionaries under key "AUTOCAD_BALLET_ATTRIBUTES"
- **Format**: Comma-separated strings in the "Tags" attribute key
- **Case sensitivity**: All tag operations are case-insensitive
- **Persistence**: Tags persist in the drawing file and survive save/load cycles

### Core API (Extension Methods on ObjectId)
```csharp
// Getting tags
List<string> tags = entityId.GetTags(db);
bool hasTags = entityId.HasAnyTags(db);
bool hasSpecific = entityId.HasTag(db, "floor1");

// Setting tags
entityId.AddTags(db, "floor1", "electrical");      // Merges with existing
entityId.SetTags(db, "floor1", "electrical");      // Replaces all tags
entityId.RemoveTags(db, "electrical");             // Removes specific tags
entityId.ClearTags(db);                            // Removes all tags
```

### Tag Discovery (Static Methods on TagHelpers)
```csharp
// Get all tags in different scopes (returns List<TagInfo>)
var tags = TagHelpers.GetAllTagsInCurrentSpace(db);           // Current Model/Paper space
var tags = TagHelpers.GetAllTagsInLayouts(db);                // All layouts in document
var tags = TagHelpers.GetAllTagsInDatabase(db);               // Entire database + blocks
var tags = TagHelpers.GetAllTagsInLayoutsAcrossDocuments(dbs);// Multiple documents

// Find entities with specific tags
List<ObjectId> entities = TagHelpers.FindEntitiesWithTag(db, "floor1");

// TagInfo structure
public class TagInfo {
    public long Id { get; set; }
    public string Name { get; set; }
    public int UsageCount { get; set; }  // Number of entities with this tag
}
```

### Tag Commands
All tag commands follow the scope-based pattern with `-in-{view,document,session}` variants:

**`tag-selected-in-{view,document,session}`**
- Interactive tagging with DataGrid UI
- Shows existing tags with usage counts
- Allows creating new tags via search box (allowCreateFromSearch: true)
- Merges selected/new tags with existing entity tags

**`select-by-sibling-tags-of-selected-in-{view,document,session}`**
- Finds entities with **exactly** the same tag combination (siblings)
- Use case: Find entities that were tagged together as a group
- Example: Entity with "floor1, electrical" → finds all entities with exactly those two tags

**`select-by-parent-tags-of-selected-in-{view,document,session}`**
- Shows DataGrid to select which tags to use for selection
- Finds entities with **at least one** of the selected tags (parents)
- Use case: Find broader categories or related entities
- Example: Selecting tags "floor1, floor2" → finds all entities with either tag

### Implementation Patterns

**When creating new tag-based commands:**
1. Use extension methods from TagHelpers for tag operations
2. Follow scope-based pattern: `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()`
3. For UI, use `CustomGUIs.DataGrid()` with tag entries
4. Tag discovery methods return `List<TagInfo>` sorted by UsageCount (descending) then Name

**Example tag usage in a command:**
```csharp
// Get tags from selected entities
var tagCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
foreach (var objId in selectedIds)
{
    var tags = objId.GetTags(db);
    foreach (var tag in tags)
    {
        if (tagCounts.ContainsKey(tag))
            tagCounts[tag]++;
        else
            tagCounts[tag] = 1;
    }
}

// Show in DataGrid
var entries = tagCounts
    .OrderByDescending(kvp => kvp.Value)
    .Select(kvp => new Dictionary<string, object>
    {
        { "TagName", kvp.Key },
        { "UsageCount", kvp.Value }
    })
    .ToList();

var selectedTags = CustomGUIs.DataGrid(entries, new List<string> { "TagName", "UsageCount" });
```

### Use Cases
- **Organization**: Tag by category, phase, discipline, responsibility
- **Selection workflows**: Build complex selections based on tag criteria
- **Cross-document coordination**: Tags work across documents in session
- **Filtering**: Combine with other selection commands for powerful filtering
- **Documentation**: Track metadata without modifying drawing properties

### Performance Considerations
- Tag discovery includes diagnostic timing (logged to command line)
- `GetAllTagsInDatabase()` scans all block definitions (slower)
- `GetAllTagsInLayouts()` scans only layouts (faster, recommended)
- `GetAllTagsInCurrentSpace()` is fastest (current space only)
- Tag operations on individual entities are fast (extension dictionary lookups)

### Best Practices
- Use descriptive, hierarchical tag names (e.g., "floor1-electrical", "phase2-demo")
- Leverage usage counts in DataGrid to show popularity
- For cross-document operations, use session scope commands
- Tags are case-insensitive but preserve original casing on first use
- Empty/whitespace tags are automatically filtered out

## Roslyn Server for AI Agent Integration

AutoCAD Ballet provides a Roslyn compiler-as-a-service that enables AI agents (like Claude Code) to execute C# code within the AutoCAD session context. This allows for dynamic inspection, validation, and automation of AutoCAD workflows.

### Quick Start

#### 1. Start the Server

In AutoCAD, run the command:
```
start-roslyn-server
```

This will:
- Open a monitoring dialog showing server activity
- Start HTTP server on `http://127.0.0.1:23714/`
- Display HTTP request/response activity in real-time
- Continue running until you press ESC

#### 2. Send Queries

From your terminal or AI agent, send C# code via HTTP POST:

```bash
# Simple query
curl -X POST http://127.0.0.1:23714/ -d 'Console.WriteLine("Hello from AutoCAD!");'

# Multi-line code with proper quoting
curl -X POST http://127.0.0.1:23714/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var layerTable = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
    Console.WriteLine($"Total layers: {layerTable.Cast<ObjectId>().Count()}");
    tr.Commit();
}'

# From a file
curl -X POST http://127.0.0.1:23714/ -d @script.cs
```

#### 3. Parse Response

The server returns JSON with this structure:

```json
{
  "success": true,
  "output": "Hello from AutoCAD!\n",
  "error": "",
  "diagnostics": []
}
```

### Available Context

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

### Common Query Patterns

#### Listing Information

**Get All Layer Names:**
```bash
curl -X POST http://127.0.0.1:23714/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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
curl -X POST http://127.0.0.1:23714/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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

#### Validation Queries

**Check Layer Naming Compliance:**
```bash
curl -X POST http://127.0.0.1:23714/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var layerTable = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
    var nonCompliant = new List<string>();

    // Check for AIA CAD Layer Guidelines (layers should start with discipline code)
    var validPrefixes = new[] { "A-", "S-", "E-", "M-", "P-", "FP-", "C-", "L-", "I-" };

    foreach (ObjectId layerId in layerTable)
    {
        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
        var layerName = layer.Name;

        if (!validPrefixes.Any(prefix => layerName.StartsWith(prefix)))
        {
            nonCompliant.Add(layerName);
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

**Find Window Blocks on Wrong Layers:**
```bash
curl -X POST http://127.0.0.1:23714/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var blockTable = (BlockTable)tr.GetObject(Db.BlockTableId, OpenMode.ForRead);
    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
    var violations = 0;

    foreach (ObjectId entId in modelSpace)
    {
        var entity = tr.GetObject(entId, OpenMode.ForRead) as Entity;
        if (entity is BlockReference blockRef)
        {
            var blockName = blockRef.Name.ToLower();
            var layerName = entity.Layer.ToLower();

            if ((blockName.Contains("window") || blockName.Contains("wndw")) &&
                !layerName.Contains("window") && !layerName.Contains("wndw") && !layerName.StartsWith("a-glaz"))
            {
                Console.WriteLine($"Window block \"{blockRef.Name}\" on layer \"{entity.Layer}\"");
                violations++;
            }
        }
    }

    Console.WriteLine($"\nTotal violations: {violations}");
    tr.Commit();
}'
```

**Check for Entities on Layer 0:**
```bash
curl -X POST http://127.0.0.1:23714/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var blockTable = (BlockTable)tr.GetObject(Db.BlockTableId, OpenMode.ForRead);
    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
    var count = 0;

    foreach (ObjectId entId in modelSpace)
    {
        var entity = tr.GetObject(entId, OpenMode.ForRead) as Entity;
        if (entity.Layer == "0")
        {
            var extents = entity.GeometricExtents;
            Console.WriteLine($"{entity.GetType().Name} at ({extents.MinPoint.X:F2}, {extents.MinPoint.Y:F2})");
            count++;
        }
    }

    if (count > 0)
    {
        Console.WriteLine($"\nWarning: {count} entities found on Layer 0");
    }
    else
    {
        Console.WriteLine("No entities on Layer 0");
    }
    tr.Commit();
}'
```

#### Statistical Queries

**Get Drawing Statistics:**
```bash
curl -X POST http://127.0.0.1:23714/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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

### Error Handling

#### Compilation Errors

If your code has syntax errors, the server returns them:

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

#### Runtime Errors

If code compiles but fails during execution:

```json
{
  "success": false,
  "output": "Any output before the error\n",
  "error": "System.NullReferenceException: Object reference not set...",
  "diagnostics": []
}
```

### Best Practices for Roslyn Server

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

### Integration with Claude Code

When Claude Code needs to query the AutoCAD session:

1. **Verify server is running**: Check if port 23714 is listening
2. **Construct query**: Build appropriate C# code based on task
3. **Send request**: Use curl or equivalent HTTP client
4. **Parse JSON response**: Extract `success`, `output`, `error`, and `diagnostics`
5. **Iterate if needed**: If compilation fails, fix errors and retry

Example workflow:
```bash
# Check if server is running
if ! curl -s --connect-timeout 1 http://127.0.0.1:23714 > /dev/null 2>&1; then
    echo "Roslyn server not running. Run 'start-roslyn-server' in AutoCAD."
    exit 1
fi

# Send query and parse response (requires jq for JSON parsing)
RESPONSE=$(curl -s -X POST http://127.0.0.1:23714/ -d 'Console.WriteLine("Test");')
SUCCESS=$(echo $RESPONSE | jq -r '.success')
OUTPUT=$(echo $RESPONSE | jq -r '.output')

if [ "$SUCCESS" == "true" ]; then
    echo "Success: $OUTPUT"
fi
```

### Security Considerations

- Server only listens on **localhost (127.0.0.1)** - not accessible from network
- No authentication - relies on localhost-only binding
- Code executes with **full AutoCAD API access**
- Intended for **local development and automation only**
- Do not expose this service to untrusted networks

### Future Capabilities

The Roslyn server will eventually support:
- **Vision Integration**: Send screenshots for visual validation
- **Persistent Scripts**: Save commonly-used validation scripts
- **Batch Execution**: Run multiple queries in one request
- **CI/CD Integration**: Automated drawing validation in build pipelines

For more implementation details, see the "Roslyn Server for AI Agents" section in `CLAUDE.md`.
