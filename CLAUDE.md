# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AutoCAD Ballet is a collection of custom commands for AutoCAD that provides enhanced selection management and productivity tools. The project consists of two main components:

1. **Commands** (`/commands/`) - C# plugin with custom AutoCAD commands and LISP utilities
2. **Installer** (`/installer/`) - Windows Forms installer application

## Development Environment

**IMPORTANT**: This codebase is located in the same directory as `%APPDATA%/autocad-ballet`, which means:
- The `runtime/` folder in this repository reflects **live usage** of the plugin
- You can examine runtime logs and data to verify code behavior and debug issues
- **ALL runtime data MUST be stored in `runtime/` subdirectory** - this includes logs, temporary data, caches, and user-generated data files
- Runtime data includes:
  - `runtime/switch-view-logs/` - Layout switching history logs (one file per document)
  - `runtime/selection/` - Per-document selection storage files
  - `runtime/selection-logs/` - Selection change history logs
  - `runtime/selection-filters` - Saved selection filter queries with names and source documents
  - `runtime/document-open-times.txt` - Document open/close timestamps
  - `runtime/invoke-addin-command-last` - Last invoked addin command (DLL path and command name)

**Debugging Workflow**: When investigating issues:
1. Check the most recent files in `runtime/switch-view-logs/` to verify logging behavior
2. Examine selection storage in `runtime/selection/` to confirm cross-document selection is working
3. Review last command invocation in `runtime/invoke-addin-command-last` to trace command executions
4. Use `ls -lt runtime/switch-view-logs/ | head` to find recently modified log files for active documents

## Architecture

### Multi-Version AutoCAD Support
The project supports AutoCAD versions 2017-2026 through conditional compilation:
- Uses `AutoCADYear` property to target specific versions
- Different .NET Framework targets based on AutoCAD version:
  - 2017-2018: .NET Framework 4.6  
  - 2019-2020: .NET Framework 4.7
  - 2021-2024: .NET Framework 4.8
  - 2025-2026: .NET 8.0 Windows
- AutoCAD.NET package versions are automatically selected based on target year

### Core Components

**Selection Management System** (`_selection.cs`):
- Provides cross-document selection capabilities
- SelectionStorage class for persisting selections across sessions
- SelectionItem class with DocumentPath, Handle, and SessionId properties
- Persistent selection storage using handles in `%APPDATA%/autocad-ballet/runtime/selection/`
- Per-document selection files with format: `DocumentPath,SessionId` followed by comma-separated handles
- Extension methods for Editor to support legacy `SetImpliedSelectionEx()` calls

**Entity Attributes and Tag System** (`_entity-attributes.cs`):
AutoCAD Ballet provides a type-safe, extensible attribute storage system for AutoCAD entities using extension dictionaries.

*Architecture:*
- **EntityAttributes class**: Generic key-value attribute storage with multiple data types (Text, Integer, Real, Boolean, Date)
- **TagHelpers class**: Specialized tag management where tags are stored as comma-separated text attributes
- **Storage mechanism**: Attributes are stored in entity extension dictionaries under the key "AUTOCAD_BALLET_ATTRIBUTES"
- **Tag storage**: Tags use the "Tags" attribute key and are stored as comma-separated strings (case-insensitive)

*Core Tag Operations (Extension Methods):*
- `GetTags(entityId, db)` - Returns List<string> of all tags on an entity
- `AddTags(entityId, db, params string[] tags)` - Adds tags to entity (merges with existing, no duplicates)
- `SetTags(entityId, db, params string[] tags)` - Replaces all tags on entity
- `RemoveTags(entityId, db, params string[] tags)` - Removes specific tags from entity
- `ClearTags(entityId, db)` - Removes all tags from entity
- `HasTag(entityId, db, string tag)` - Checks if entity has specific tag
- `HasAnyTags(entityId, db)` - Checks if entity has any tags

*Tag Discovery Methods (Static):*
- `TagHelpers.GetAllTagsInCurrentSpace(db)` - Returns List<TagInfo> from current space (Model/Paper)
- `TagHelpers.GetAllTagsInLayouts(db)` - Returns List<TagInfo> from all layouts (Model + Paper Spaces) in database
- `TagHelpers.GetAllTagsInLayoutsAcrossDocuments(databases)` - Returns List<TagInfo> across multiple databases
- `TagHelpers.GetAllTagsInDatabase(db)` - Returns List<TagInfo> from entire database including block definitions
- `TagHelpers.FindEntitiesWithTag(db, string tag)` - Returns List<ObjectId> of all entities with specific tag
- TagInfo contains: Id, Name, UsageCount (number of entities with that tag)

*Tag Commands:*
- `tag-selected-in-{view,document,session}` - Tag entities interactively
  - Shows DataGrid with existing tags and usage counts
  - Allows creating new tags by typing in search box (allowCreateFromSearch: true)
  - User can select multiple existing tags or create new ones
  - Adds selected/new tags to selected entities (merges with existing)

- `select-by-sibling-tags-of-selected-in-{view,document,session}` - Select sibling entities
  - Collects all unique tag sets from selected entities
  - Finds all entities with **exactly** the same combination of tags
  - Useful for finding entities that were tagged together as a group
  - Example: If entity has tags "floor1, electrical", finds all entities with exactly those two tags

- `select-by-parent-tags-of-selected-in-{view,document,session}` - Select parent entities by tag
  - Collects all unique tags from selected entities with usage counts
  - Shows DataGrid for user to select which tags to use for selection
  - Finds all entities with **at least one** of the selected tags
  - Useful for finding broader categories or related entities
  - Example: If selecting tags "floor1, floor2", finds all entities with either tag

*Use Cases:*
- **Organization**: Tag entities by category, phase, responsibility, etc.
- **Selection workflows**: Build complex selections based on tag criteria
- **Cross-document coordination**: Tags persist with entities and work across documents
- **Filtering**: Combine with other selection commands for powerful filtering
- **Documentation**: Track entity metadata without modifying drawing properties

*Implementation Notes:*
- All tag operations are case-insensitive
- Tags are trimmed of whitespace automatically
- Empty or whitespace-only tags are ignored
- Tag data persists in the drawing file (stored in extension dictionaries)
- No limit on number of tags per entity or number of unique tags in a document
- Performance: Tag discovery methods include diagnostic timing information

**Roslyn Server for AI Agents** (`start-roslyn-server.cs`):
AutoCAD Ballet provides a Roslyn compiler-as-a-service that allows external tools (like Claude Code) to execute C# code within the AutoCAD session context.

*Purpose:*
- Enable AI agents to query and analyze the current AutoCAD session programmatically
- Perform validation tasks (layer naming compliance, entity placement checks, etc.)
- Inspect document state and entity properties dynamically
- Support automated drawing quality checks and audits

*Architecture:*
- **HTTP REST API**: Listens on `http://127.0.0.1:34157/` for incoming POST requests
- **Roslyn Compilation**: Compiles received C# code using Microsoft.CodeAnalysis.CSharp.Scripting
- **Execution Context**: Provides access to AutoCAD API through `ScriptGlobals` class
- **Output Capture**: Redirects `Console.WriteLine()` output and returns to client
- **Error Handling**: Returns compilation errors and warnings back to client
- **Live Monitoring**: Shows real-time HTTP request/response activity in dialog

*Available Context in Scripts:*
Scripts executed by the server have access to these globals:
- `Doc` - Current active AutoCAD document (`Autodesk.AutoCAD.ApplicationServices.Document`)
- `Db` - Current document's database (`Autodesk.AutoCAD.DatabaseServices.Database`)
- `Ed` - Current document's editor for user interaction (`Autodesk.AutoCAD.EditorInput.Editor`)

Note: Application-level features can be accessed through the `Autodesk.AutoCAD.ApplicationServices.Core.Application` namespace (pre-imported)

*Pre-imported Namespaces:*
- `System`, `System.Linq`, `System.Collections.Generic`
- `Autodesk.AutoCAD.ApplicationServices`
- `Autodesk.AutoCAD.DatabaseServices`
- `Autodesk.AutoCAD.EditorInput`
- `Autodesk.AutoCAD.Geometry`

*Protocol:*
- **Request**: HTTP POST to `http://127.0.0.1:34157/` with C# code in request body
- **Response**: JSON object with structure:
  ```json
  {
    "success": true/false,
    "output": "captured console output",
    "error": "error message if failed",
    "diagnostics": ["warning/error messages from compilation"]
  }
  ```

*Usage from Command Line (curl):*
```bash
# Example: Get all layer names
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

# Example: Get document name
curl -X POST http://127.0.0.1:34157/ -d 'Console.WriteLine($"Document: {Doc.Name}");'

# Or using a file
curl -X POST http://127.0.0.1:34157/ -d @script.cs

# Example: Check layer naming compliance
curl -X POST http://127.0.0.1:34157/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
{
    var layerTable = (LayerTable)tr.GetObject(Db.LayerTableId, OpenMode.ForRead);
    var nonCompliant = new List<string>();
    foreach (ObjectId layerId in layerTable)
    {
        var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
        if (!layer.Name.StartsWith("A-") && !layer.Name.StartsWith("S-"))
        {
            nonCompliant.Add(layer.Name);
        }
    }
    if (nonCompliant.Count > 0)
    {
        Console.WriteLine("Non-compliant layers:");
        foreach (var name in nonCompliant)
        {
            Console.WriteLine("  - " + name);
        }
    }
    else
    {
        Console.WriteLine("All layers are compliant");
    }
    tr.Commit();
}'

# Example: Find entities on wrong layers
curl -X POST http://127.0.0.1:34157/ -d 'using (var tr = Db.TransactionManager.StartTransaction())
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
            if (blockName.Contains("window") && !layerName.Contains("window"))
            {
                Console.WriteLine($"Window block on wrong layer: {blockRef.Name} on {entity.Layer}");
                violations++;
            }
        }
    }
    Console.WriteLine($"Total violations: {violations}");
    tr.Commit();
}'
```

*Commands:*
- `start-roslyn-server` - Opens modal dialog and starts HTTP server on `http://127.0.0.1:34157/`
  - Modal blocking dialog - AutoCAD UI is unavailable while server runs
  - Press ESC to stop the server and close dialog
  - Dialog shows connection events and request/response activity in real-time
  - Server continues accepting connections until stopped
  - No authentication required (localhost-only)

- `start-roslyn-server-in-background` - Starts HTTP server in background with non-blocking dialog
  - Non-blocking dialog - AutoCAD remains fully functional while server runs
  - Dialog can be closed while server continues running
  - Authentication required via `X-Auth-Token` header or `Authorization: Bearer <token>` header
  - Server generates secure random token (256-bit) on startup
  - Token displayed in dialog and written to AutoCAD command line
  - Use `stop-roslyn-server` command or dialog button to stop server
  - Server persists across document switches and stays active until explicitly stopped

*Authentication (background server only):*
```bash
# Using X-Auth-Token header
curl -X POST http://127.0.0.1:34157/ \
  -H "X-Auth-Token: YOUR_TOKEN_HERE" \
  -d 'Console.WriteLine("Hello");'

# Using Authorization Bearer header
curl -X POST http://127.0.0.1:34157/ \
  -H "Authorization: Bearer YOUR_TOKEN_HERE" \
  -d 'Console.WriteLine("Hello");'
```

*Security Notes:*
- Server only listens on localhost (127.0.0.1)
- Intended for local development and automation only
- Code execution happens within AutoCAD process with full API access
- Background server requires authentication token to prevent unauthorized access
- Foreground server has no authentication (relies on localhost-only binding and modal dialog)

*Type Identity Issue (.NET 8 / AutoCAD 2025-2026):*
The background server implementation addresses a critical .NET 8 type identity issue:

**Problem**: In .NET 8, assemblies loaded with rewritten identities (via Mono.Cecil for hot-reloading) do not have a file `Location` property. Roslyn's `ScriptOptions.AddReferences()` requires assemblies with valid file paths. Additionally, .NET treats types from different assembly instances as incompatible even if they have identical definitions.

**Symptoms**:
- `Cannot reference assembly without location` error during compilation
- `Type A cannot be cast to Type B` error during execution (both Type A and Type B are `ScriptGlobals`)

**Solution** (implemented in `start-roslyn-server-in-background.cs`):
1. **Assembly Location Resolution**: On server startup, search `AppDomain.CurrentDomain.GetAssemblies()` to find ANY loaded copy of `autocad-ballet.dll` that has a valid `Location` property
2. **Type Caching**: Cache both the assembly reference AND the `ScriptGlobals` type from that specific assembly instance
3. **Unified Type Usage**:
   - Use cached assembly for Roslyn's `AddReferences()`
   - Use cached `ScriptGlobals` type for `CSharpScript.Create()`
   - Use cached type with `Activator.CreateInstance()` + reflection to create globals instance
4. **Result**: All operations use the SAME assembly instance, ensuring type identity consistency

**Code Pattern**:
```csharp
// Cache assembly and type at startup
Assembly scriptGlobalsAssembly = FindAssemblyWithLocation();
Type scriptGlobalsType = scriptGlobalsAssembly.GetType("AutoCADBallet.ScriptGlobals");

// Use cached assembly for Roslyn references
var scriptOptions = ScriptOptions.Default.AddReferences(scriptGlobalsAssembly);

// Use cached type for script creation
var script = CSharpScript.Create(code, scriptOptions, scriptGlobalsType);

// Use cached type for globals instance (via reflection, not "new ScriptGlobals()")
var globals = Activator.CreateInstance(scriptGlobalsType);
scriptGlobalsType.GetProperty("Doc").SetValue(globals, doc);
```

This pattern ensures type identity across compilation and execution, preventing cast exceptions.

*Future: Central Server Hub Architecture:*
The current implementation provides one server per AutoCAD session. Future enhancements will support:
- **Central Hub Server**: Single server process managing multiple AutoCAD sessions
- **Session Routing**: Clients specify target session ID in requests
- **Cross-Session Queries**: Execute scripts across multiple AutoCAD instances
- **Session Discovery**: API to enumerate available AutoCAD sessions
- **Distributed Validation**: Batch validation across entire project drawing sets
- **Type Version Compatibility**: Ensure ScriptGlobals types are compatible across sessions

*Implementation Details:*
- Uses `CommandFlags.Session` to work at application level
- Runs TCP server on background thread with cancellation token
- Marshals AutoCAD API access appropriately (executes on UI thread via `Application.Idle` event)
- Captures `Console.WriteLine()` output using `StringWriter`
- Returns both compilation diagnostics and runtime errors
- Keeps listening even after compilation errors (allows correction)
- Background server uses document locking for thread-safe database operations

For detailed usage examples and integration with AI agents, see `AGENTS.md`.

**Command Structure**:
- C# commands use `[CommandMethod]` attributes with kebab-case names (e.g., "delete-selected")
- **CRITICAL**: Every command class MUST have the assembly-level `[assembly: CommandClass(typeof(YourClassName))]` attribute at the top of the file, outside the namespace declaration
  - Without this attribute, AutoCAD will not register the commands and they will be invisible
  - Example: `[assembly: CommandClass(typeof(AutoCADBallet.TagSelectedNamedInView))]`
  - This tells AutoCAD to scan the class for `[CommandMethod]` attributes
  - Place this line after `using` statements but before the `namespace` declaration
- LISP files provide aliases and utility functions
- Main command invocation through `invoke-addin-command.cs` with history tracking via `invoke-last-addin-command.cs`
- **IMPORTANT**: Do not add command aliases to `aliases.lsp` - leave alias creation to the project owner

**CommandFlags.Session Usage**:
- **ALWAYS use `CommandFlags.Session`** for commands that:
  - Open, close, or switch between documents
  - Modify the DocumentCollection (opening/closing documents)
  - Need to work when NO document is active (application-level context)
  - Switch between views/layouts across documents
- Commands with `CommandFlags.Session` run at **application level**, not document level
- **CRITICAL**: LISP wrappers for Session commands MUST include `(princ)` to prevent command context leaks
- Without `CommandFlags.Session`, cross-document operations will fail or leave commands stuck
- Examples requiring Session flag: `open-documents-recent-read-write`, `open-documents-recent-read-only`, `close-all-without-saving`, `switch-view`, `switch-view-last`, `switch-view-recent`

**LISP Command Wrappers**:
- All LISP command wrappers in `aliases.lsp` MUST end with `(princ)` to prevent command context leaks
- Pattern: `(defun c:alias () (command "command-name") (princ))`
- Without `(princ)`, Session-level commands will leave LISP command contexts stuck on documents
- This causes "Drawing is busy" errors and prevents documents from closing
- Example: `(defun c:ww () (command "switch-view") (princ))`

**DataGrid System** (`commands/datagrid/`):
AutoCAD Ballet provides a sophisticated, reusable DataGrid UI component for displaying and interacting with tabular data.

*CRITICAL: When instructed to "use datagrid" or "show a datagrid", ALWAYS use the built-in `CustomGUIs.DataGrid()` method - DO NOT create custom Windows Forms DataGridView implementations.*

*Architecture:*
- Modular implementation split across multiple files:
  - `datagrid.Main.cs` - Core DataGrid functionality and API
  - `datagrid.Filtering.cs` - Column filtering and search
  - `datagrid.Sorting.cs` - Multi-column sorting
  - `datagrid.Virtualization.cs` - Virtual scrolling for large datasets
  - `datagrid.EditMode.cs` - In-place cell editing (Excel-like)
  - `datagrid.EditApply.cs` - Apply edits to AutoCAD entities
  - `datagrid.CacheIndex.cs` - Search indexing for performance
  - `datagrid.SearchHistory.cs` - Search query history (F4 dropdown)
  - `datagrid.Types.cs` - Shared type definitions
  - `datagrid.Utils.cs` - Utility functions

*API Signature:*
```csharp
public static List<Dictionary<string, object>> DataGrid(
    List<Dictionary<string, object>> entries,        // Data rows
    List<string> propertyNames,                      // Columns to display
    bool spanAllScreens,                             // Multi-monitor support
    List<int> initialSelectionIndices = null,        // Pre-select rows
    Func<List<Dictionary<string, object>>, bool> onDeleteEntries = null,  // Delete callback
    bool allowCreateFromSearch = false,              // Allow creating new items from search box
    string commandName = null)                       // Optional command name (auto-inferred from stack)
```

*Usage Pattern:*
```csharp
// 1. Build data as List<Dictionary<string, object>>
var data = new List<Dictionary<string, object>>();
data.Add(new Dictionary<string, object>
{
    ["LayoutName"] = "Sheet1",
    ["ViewportHandle"] = "A3F",
    ["Width"] = 10.5,
    ["IsActive"] = true
});

// 2. Define columns to display
var columns = new List<string> { "LayoutName", "ViewportHandle", "Width" };

// 3. Show DataGrid and get user selection
var selected = CustomGUIs.DataGrid(data, columns, spanAllScreens: false);

// 4. Process selected items
if (selected != null && selected.Count > 0)
{
    foreach (var item in selected)
    {
        string layoutName = item["LayoutName"].ToString();
        // ... process selection
    }
}
```

*Key Features:*
- **Multi-select**: Ctrl+Click, Shift+Click, Ctrl+A for all
- **Search**: Type in search box to filter (F4 for history dropdown)
- **Column filtering**: Click column headers, type filter query
- **Multi-column sorting**: Shift+Click column headers
- **Edit mode**: F2 to enter Excel-like editing mode
- **Keyboard navigation**: Arrow keys, Enter to confirm, Escape to cancel
- **Multi-monitor**: Span across screens with `spanAllScreens: true`
- **Create from search**: `allowCreateFromSearch: true` lets users type new values
- **Auto-formatting**: Column headers converted to "lowercase with spaces" automatically

*Display Conventions:*
- Column headers are automatically formatted to lowercase with spaces between words
- Example: "DocumentName" becomes "document name", "LayoutType" becomes "layout type"
- This is handled by the `FormatColumnHeader` method in `datagrid/datagrid.Main.cs`
- Actual property names remain unchanged (e.g., "DocumentName"), only display headers are formatted

*Example Commands Using DataGrid:*
- `switch-view.cs` - Layout switching with multi-document support
- `filter-selection.cs` - Entity filtering with complex queries
- `tag-selected-in-view.cs` - Tag selection with create-from-search
- `select-by-parent-tags-of-selected-in-view.cs` - Tag-based selection
- `delete-layouts.cs` - Layout deletion with confirmation
- `list-system-variables.cs` - System variable browsing and editing

*DO NOT:*
- Create custom Windows Forms with DataGridView manually
- Implement your own search, filter, or sorting logic
- Build custom column header formatting
- Use `ShowDataGrid` or similar custom methods

*DO:*
- Use `CustomGUIs.DataGrid()` for all tabular data display needs
- Build data as `List<Dictionary<string, object>>`
- Store complex objects in dictionary values for later use
- Let the DataGrid handle all UI interactions automatically

## Building

### Commands Plugin
```bash
# Build for specific AutoCAD version (default: 2026)
dotnet build commands/autocad-ballet.csproj

# Build for specific version
dotnet build commands/autocad-ballet.csproj -p:AutoCADYear=2024

# Output location: commands/bin/{AutoCADYear}/autocad-ballet.dll
```

### Installer
```bash
# Build Windows Forms installer
dotnet build installer/installer.csproj

# Output: installer/bin/Release/installer.exe (or Debug/installer.exe for debug builds)
```

## Key Files

- `commands/autocad-ballet.csproj` - Main plugin project with multi-version support
- `commands/_selection.cs` - Core selection storage system with SelectionItem and SelectionStorage classes
- `commands/_edit-dialog.cs` - Shared text editing dialog (internal utility)
- `commands/_reopen-documents.cs` - Shared implementation for document reopening functionality
- `commands/_switch-view-logging.cs` - Layout/view switching history logging utilities
- `commands/start-roslyn-server.cs` - Roslyn compiler-as-a-service for AI agent integration
- `commands/filter-selection.cs` - Entity filtering with data grid UI (main implementation)
- `commands/datagrid/` - Modular data grid implementation with separate files for:
  - `datagrid.Main.cs` - Core data grid functionality
  - `datagrid.Filtering.cs` - Column filtering logic
  - `datagrid.Sorting.cs` - Column sorting functionality
  - `datagrid.EditMode.cs` - In-place editing support
  - `datagrid.EditApply.cs` - Apply edits to entities
  - `datagrid.Virtualization.cs` - Virtual scrolling for large datasets
  - `datagrid.CacheIndex.cs` - Entity cache indexing
  - `datagrid.Types.cs` - Shared type definitions
  - `datagrid.Utils.cs` - Utility functions
- `commands/invoke-addin-command.cs` - Command invocation with grid selection
- `commands/invoke-last-addin-command.cs` - Re-invoke last command with history tracking
- `commands/aliases.lsp` - LISP command aliases
- `installer/installer.csproj` - Installer project that embeds plugin DLLs

## File Naming Conventions

**Command Files**: Use kebab-case for AutoCAD command files (e.g., `edit-selected-text.cs`, `switch-view-last.cs`)

**One Command Per File**: Each AutoCAD command should be in its own file named after the command. The file should contain either:
- The complete implementation of the command
- A stub that delegates to shared implementation code in a related file

**Internal/Utility Files**: Use underscore prefix with kebab-case for files that are not actual commands but provide internal functionality (e.g., `_edit-dialog.cs`, `_utilities.cs`)

This convention helps distinguish between:
- **Commands**: Files that define AutoCAD commands with `[CommandMethod]` attributes (one command per file)
- **Internal utilities**: Shared classes, dialogs, and helper functionality used by commands

## Installation Structure

The plugin installs to `%APPDATA%/autocad-ballet/` with:
- Plugin DLLs for each AutoCAD version
- LISP files for command aliases and utilities

Runtime data is stored in `%APPDATA%/autocad-ballet/runtime/`:
- `selection/` - Per-document selection storage files
- `switch-view-logs/` - Layout switching history by project name
- `selection-logs/` - Selection change history by project name

## Command Porting

Many commands in AutoCAD Ballet are ported from the revit-ballet repository. When porting commands from Revit to AutoCAD:

- **Views → Layouts**: Revit views correspond to AutoCAD layouts
- **ElementId → Layout Names**: Revit uses numeric ElementIds, AutoCAD uses string layout names
- **View Change Logging**: Both systems use similar logging patterns:
  - Revit: `%APPDATA%/revit-scripts/LogViewChanges/{ProjectName}`
  - AutoCAD: `%APPDATA%/autocad-ballet/runtime/switch-view-logs/{ProjectName}`
- **API Differences**:
  - Revit: `uidoc.ActiveView = view`
  - AutoCAD: `LayoutManager.Current.CurrentLayout = layoutName`
- **Document Access**: Both use similar document manager patterns but with different namespaces
- **Runtime Data Storage**: All runtime information (logs, temporary data) should be stored in `%APPDATA%/autocad-ballet/runtime/`
  - `runtime/selection/` - Per-document selection files (from SelectionStorage)
  - `runtime/switch-view-logs/` - Layout switching history logs (from SwitchViewLogging)
  - `runtime/selection-logs/` - Selection change logs (from SwitchViewLogging)

Commands successfully ported from revit-ballet include:
- `switch-view-last.cs` - Switches to the most recent layout from log history
- `switch-view-recent.cs` - Shows layouts ordered by usage history for selection
- `switch-view.cs` - Interactive layout switching with grid selection

## AutoCAD .NET API Patterns

### Cross-Document Operations
When working with multiple documents in AutoCAD, certain operations require proper timing and context management:

**Document Switching Pattern**: Use the `DocumentActivated` event for reliable cross-document operations:

```csharp
// INCORRECT - Timing issue
docs.MdiActiveDocument = targetDoc;
LayoutManager.Current.CurrentLayout = targetLayout;  // May fail

// CORRECT - Event-driven approach
DocumentCollectionEventHandler handler = null;
handler = (sender, e) => {
    if (e.Document == targetDoc) {
        docs.DocumentActivated -= handler;  // Unsubscribe immediately
        // Now safely perform operations on the activated document
        LayoutManager.Current.CurrentLayout = targetLayout;
    }
};
docs.DocumentActivated += handler;
docs.MdiActiveDocument = targetDoc;
```

**Key Principles**:
- Document switching (`docs.MdiActiveDocument = targetDoc`) is asynchronous
- Use `DocumentActivated` event to know when switching is complete
- Always unsubscribe from events to prevent memory leaks
- Use the target document's `Editor` context, not the original document's

### Document Locking for Database Operations
**CRITICAL**: AutoCAD database modifications require explicit document locking to prevent `eLockViolation` exceptions and ensure reliable operation.

**The Problem**:
- Database operations like `blockTable.UpgradeOpen()`, entity creation, modification, and deletion can fail with `eLockViolation`
- This occurs when trying to modify the database without proper document locking
- Layout switching (`LayoutManager.Current.CurrentLayout = layoutName`) can fail with `eSetFailed` in cross-document scenarios

**The Solution**: Always use `DocumentLock` when performing database modifications:

```csharp
// INCORRECT - May fail with eLockViolation or eSetFailed
using (var tr = db.TransactionManager.StartTransaction())
{
    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
    blockTable.UpgradeOpen(); // eLockViolation!
    // ... database modifications
    tr.Commit();
}

// CORRECT - Reliable with document lock
using (var docLock = doc.LockDocument())
{
    using (var tr = db.TransactionManager.StartTransaction())
    {
        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        blockTable.UpgradeOpen(); // Safe!
        // ... database modifications
        tr.Commit();
    }
}

// CORRECT - Layout operations with document lock
using (DocumentLock docLock = targetDoc.LockDocument())
{
    LayoutManager.Current.CurrentLayout = targetLayout;
}
```

**Key Benefits**:
- **Eliminates `eLockViolation` exceptions** for database modifications
- **Eliminates `eSetFailed` exceptions** in cross-document layout switching
- **Works directly from Model Space** - no need for intermediate "priming" layouts
- **Prevents document context conflicts** between multiple operations
- **Essential for event handlers** that operate on non-active documents

**Document Locking Best Practice**: Always lock the document when:
- **Creating, modifying, or deleting entities** in the database
- **Upgrading database objects to write mode** (e.g., `blockTable.UpgradeOpen()`)
- **Adding new objects to symbol tables** (blocks, layers, styles, etc.)
- **Switching layouts** after document activation
- **Performing layout operations** in `DocumentActivated` event handlers
- **Working with non-active documents**
- **Any operation that modifies document state** across document boundaries

**Pattern for Database Modifications**:
```csharp
// Always wrap database modification transactions with document locking
using (var docLock = doc.LockDocument())
{
    using (var tr = db.TransactionManager.StartTransaction())
    {
        // Safe to upgrade objects to write mode and modify database
        var someTable = (SomeTable)tr.GetObject(db.SomeTableId, OpenMode.ForRead);
        someTable.UpgradeOpen();

        // Create new objects, modify entities, etc.

        tr.Commit();
    }
}
```

This pattern is essential for reliable AutoCAD .NET development and prevents the most common database access exceptions.

### Selection Handling Patterns
**CRITICAL**: AutoCAD commands should use existing selection (pickfirst set) by default. This is the expected AutoCAD workflow.

**The Problem**: Commands that always prompt for selection instead of using pre-selected objects break AutoCAD's standard workflow where users select first, then run commands.

**The Solution**: Always use `CommandFlags.UsePickSet` and proper selection handling:

```csharp
[CommandMethod("my-command", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
public void MyCommand()
{
    var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

    // Get pickfirst set (pre-selected objects)
    var selResult = ed.SelectImplied();

    // If there is no pickfirst set, request user to select objects
    if (selResult.Status == PromptStatus.Error)
    {
        var selectionOpts = new PromptSelectionOptions();
        selectionOpts.MessageForAdding = "\nSelect objects: ";
        selResult = ed.GetSelection(selectionOpts);
    }
    else if (selResult.Status == PromptStatus.OK)
    {
        // Clear the pickfirst set since we're consuming it
        ed.SetImpliedSelection(new ObjectId[0]);
    }

    if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
    {
        ed.WriteMessage("\nNo objects selected.\n");
        return;
    }

    var selectedObjects = selResult.Value.GetObjectIds();
    // Process selected objects...
}
```

**Command Flags Explained**:
- **`CommandFlags.UsePickSet`**: **REQUIRED** - Makes pickfirst selection available to the command
- **`CommandFlags.Redraw`**: Updates display after command completion
- **`CommandFlags.Modal`**: Standard modal command behavior

**Key Principles**:
- **Always use `CommandFlags.UsePickSet`** - Without this flag, `ed.SelectImplied()` will always return `PromptStatus.Error`
- **Prefer existing selection** - Check for pickfirst set first, only prompt if nothing is selected
- **Clear pickfirst after use** - Call `ed.SetImpliedSelection(new ObjectId[0])` after consuming the selection
- **Follow AutoCAD conventions** - Users expect to select objects first, then run commands

**Examples in Codebase**:
- `filter-selection.cs` - Perfect example of proper selection handling
- `copy-selection-to-views-in-session.cs` - Uses this pattern correctly (note: renamed from "process" to "session" scope)

### Selection Scope Handling
**CRITICAL**: AutoCAD Ballet provides three selection scopes through separate command variants: view, document, and session.

**The Three Selection Scopes**:
1. **`view`** - Current active space/layout only (default AutoCAD behavior, uses pickfirst set)
2. **`document`** - Entire current document (all layouts, uses stored selection)
3. **`session`** - All opened documents in current AutoCAD process (uses stored selection)

**Architecture**: Commands use a **scope-based command pattern**:
- Each command has three variants: `{command}-in-view`, `{command}-in-document`, `{command}-in-session`
- Main implementation file contains separate `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()` methods
- Scope-specific command files are stubs that call the appropriate method

**The Problem**: Commands that only handle view scope don't provide the cross-document capabilities users expect from AutoCAD Ballet.

**Note on Terminology**: The third scope is called "session" in command names (e.g., `select-by-categories-in-session`) but the implementation method is still called `ExecuteApplicationScope()` for backwards compatibility.

**The Solution**: Implement all three scope methods in the main command file:

```csharp
// Main implementation file: my-command.cs
public static class MyCommand
{
    // View scope: Use pickfirst set or prompt for selection
    public static void ExecuteViewScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        var selResult = ed.SelectImplied();

        if (selResult.Status == PromptStatus.Error)
        {
            var selectionOpts = new PromptSelectionOptions();
            selectionOpts.MessageForAdding = "\nSelect objects: ";
            selResult = ed.GetSelection(selectionOpts);
        }
        else if (selResult.Status == PromptStatus.OK)
        {
            ed.SetImpliedSelection(new ObjectId[0]);
        }

        if (selResult.Status != PromptStatus.OK || selResult.Value.Count == 0)
        {
            ed.WriteMessage("\nNo objects selected.\n");
            return;
        }

        var selectedIds = selResult.Value.GetObjectIds();
        // Process current view selection...
    }

    // Document scope: Use stored selection filtered to current document
    public static void ExecuteDocumentScope(Editor ed, Database db)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var docName = Path.GetFileName(doc.Name);

        var storedSelection = SelectionStorage.LoadSelection(docName);
        if (storedSelection == null || storedSelection.Count == 0)
        {
            ed.WriteMessage($"\nNo stored selection found for document. Use 'select-by-categories-in-document' first.\n");
            return;
        }

        // Filter to current document only
        var currentDocPath = Path.GetFullPath(doc.Name);
        storedSelection = storedSelection.Where(item =>
        {
            try
            {
                var itemPath = Path.GetFullPath(item.DocumentPath);
                return string.Equals(itemPath, currentDocPath, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(item.DocumentPath, doc.Name, StringComparison.OrdinalIgnoreCase);
            }
        }).ToList();

        // Process document-wide selection...
    }

    // Session scope: Use stored selection from all open documents
    public static void ExecuteApplicationScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        var storedSelection = SelectionStorage.LoadSelectionFromOpenDocuments();
        if (storedSelection == null || storedSelection.Count == 0)
        {
            ed.WriteMessage("\nNo stored selection found. Use 'select-by-categories-in-session' first.\n");
            return;
        }

        // Separate current document from external documents
        var currentDocEntities = new List<ObjectId>();
        var externalDocuments = new Dictionary<string, List<SelectionItem>>();

        foreach (var item in storedSelection)
        {
            var itemPath = Path.GetFullPath(item.DocumentPath);
            var currentPath = Path.GetFullPath(doc.Name);

            if (string.Equals(itemPath, currentPath, StringComparison.OrdinalIgnoreCase))
            {
                var handle = Convert.ToInt64(item.Handle, 16);
                var objectId = db.GetObjectId(false, new Handle(handle), 0);
                if (objectId != ObjectId.Null)
                {
                    currentDocEntities.Add(objectId);
                }
            }
            else
            {
                if (!externalDocuments.ContainsKey(item.DocumentPath))
                {
                    externalDocuments[item.DocumentPath] = new List<SelectionItem>();
                }
                externalDocuments[item.DocumentPath].Add(item);
            }
        }

        // Process current document entities...
        // Process external documents (may need to open them)...
    }
}

// Stub file: my-command-in-view.cs
public class MyCommandInView
{
    [CommandMethod("my-command-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void MyCommandInViewCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            MyCommand.ExecuteViewScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in my-command-in-view: {ex.Message}\n");
        }
    }
}

// Stub file: my-command-in-document.cs
public class MyCommandInDocument
{
    [CommandMethod("my-command-in-document", CommandFlags.Modal)]
    public void MyCommandInDocumentCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        try
        {
            MyCommand.ExecuteDocumentScope(ed, db);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in my-command-in-document: {ex.Message}\n");
        }
    }
}

// Stub file: my-command-in-session.cs
public class MyCommandInSession
{
    [CommandMethod("my-command-in-session", CommandFlags.Modal)]
    public void MyCommandInSessionCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            MyCommand.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in my-command-in-session: {ex.Message}\n");
        }
    }
}
```

**Key Selection Scope Principles**:
- **Implement three separate methods**: `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()`
- **View scope**: Use pickfirst set or prompt for selection (standard AutoCAD behavior with `CommandFlags.UsePickSet`)
- **Document scope**: Use `SelectionStorage.LoadSelection(docName)` and filter to current document
- **Session scope**: Use `SelectionStorage.LoadSelectionFromOpenDocuments()` for cross-document selections (only includes currently open documents)
- **Cross-document limitations**: Many operations can only be performed on current document entities - may need to open external documents
- **Update stored selection**: Use `SelectionStorage.SaveSelection()` to update stored selection after modifications
- **Clear error messages**: Guide users to use appropriate `select-by-categories-in-{scope}` command first

**Required Imports for Selection Scope Support**:
```csharp
using AutoCADBallet;  // For SelectionStorage, SelectionItem
using System.IO;      // For Path operations in cross-document scenarios
using System.Linq;    // For LINQ filtering operations
```

**Examples in Codebase**:
- `filter-selection.cs` - Perfect reference implementation for scope-aware selection
- `delete-selected.cs` - Main implementation with `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()` methods
  - Stub files: `delete-selected-in-view.cs`, `delete-selected-in-document.cs`, `delete-selected-in-session.cs`
- `select-by-categories.cs` - Selection by entity category with scope variants
  - Stub files: `select-by-categories-in-view.cs`, `select-by-categories-in-document.cs`, `select-by-categories-in-session.cs`
- `duplicate-blocks.cs` - Block duplication with all selection scopes
  - Stub files: `duplicate-blocks-in-view.cs`, `duplicate-blocks-in-session.cs`
