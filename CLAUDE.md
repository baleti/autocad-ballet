# CLAUDE.md

> **Documentation Size Limit**: This file must stay under 20KB. Keep changes concise and avoid approaching this limit to allow for future additions.

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AutoCAD Ballet is a collection of custom commands for AutoCAD that provides enhanced selection management and productivity tools. The project consists of two main components:

1. **Commands** (`/commands/`) - C# plugin with custom AutoCAD commands and LISP utilities
2. **Installer** (`/installer/`) - Windows Forms installer application

## Development Environment

**IMPORTANT**: This codebase is located in the same directory as `%APPDATA%/autocad-ballet`, which means:
- The `runtime/` folder in this repository reflects **live usage** of the plugin
- You can examine runtime logs and data to verify code behavior and debug issues
- **ALL runtime data MUST be stored in `runtime/` subdirectory** - includes logs, selection storage, filters, document timestamps
- Use `ls -lt runtime/switch-view-logs/ | head` to find recently modified log files for active documents

### Querying and Testing AutoCAD Sessions

**When you need to query AutoCAD state, test commands, or inspect drawings:**

1. **Use the Roslyn Server** - Every AutoCAD session runs an HTTPS server that executes C# scripts
2. **Location**: Token at `runtime/network/token`, sessions at `runtime/network/sessions`
3. **Find active port**: `grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2`

**Quick example - List layers:**
```bash
TOKEN=$(cat runtime/network/token)
PORT=$(grep -v '^#' runtime/network/sessions | head -1 | cut -d',' -f2)

curl -k -s -X POST https://127.0.0.1:$PORT/roslyn \
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

**Available in scripts:**
- `Doc` - Current document
- `Db` - Database
- `Ed` - Editor
- Pre-imported: System, LINQ, AutoCAD.DatabaseServices, AutoCAD.EditorInput, etc.

See **AGENTS.md** for comprehensive examples and query patterns.

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

**Entity Attributes and Tag System** (`_entity-attributes.cs`):
- **EntityAttributes class**: Generic key-value attribute storage (Text, Integer, Real, Boolean, Date)
- **TagHelpers class**: Specialized tag management (comma-separated, case-insensitive)
- **Storage**: Extension dictionaries under key "AUTOCAD_BALLET_ATTRIBUTES"

*Core Tag Operations:*
- `GetTags(entityId, db)`, `AddTags()`, `SetTags()`, `RemoveTags()`, `ClearTags()`, `HasTag()`, `HasAnyTags()`
- `TagHelpers.GetAllTagsInCurrentSpace(db)`, `GetAllTagsInLayouts()`, `GetAllTagsInDatabase()`, `FindEntitiesWithTag()`

*Tag Commands:*
- `tag-selected-in-{view,document,session}` - Interactive tagging with DataGrid
- `select-by-sibling-tags-of-selected-in-{view,document,session}` - Exact tag match
- `select-by-parent-tags-of-selected-in-{view,document,session}` - At least one tag match

*Use Cases:* Organization, selection workflows, cross-document coordination, filtering, metadata tracking

**Roslyn Server for AI Agents** (`commands/_server.cs`):
AutoCAD Ballet provides a Roslyn compiler-as-a-service with peer-to-peer network architecture for cross-session queries.

*Purpose:*
- Query and analyze AutoCAD sessions programmatically
- Perform validation tasks (layer naming, entity placement checks)
- Inspect document state and entity properties dynamically
- Support cross-session network queries for multi-AutoCAD workflows

*Architecture:*
- **Peer-to-peer network**: Each AutoCAD instance runs its own HTTPS server
- **Auto-start**: Servers start automatically when AutoCAD loads (no manual command)
- **Port allocation**: 34157-34257 range, automatic first-available selection
- **Shared token**: Single authentication token shared across all sessions
- **Service discovery**: File-based registry at `%appdata%/autocad-ballet/runtime/network/`

*Network Registry Structure:*
```
%appdata%/autocad-ballet/runtime/network/
  ├── sessions      # CSV: SessionId,Port,Hostname,ProcessId,RegisteredAt,LastHeartbeat,Documents
  └── token         # Shared 256-bit authentication token
```

*Endpoints:*
- **`/roslyn`** - Execute C# scripts in AutoCAD context
- **`/select-by-categories`** - Query entity categories (for network queries)
- **`/screenshot`** - Capture screenshot of AutoCAD window, saved to `runtime/screenshots/`
- **`/invoke-addin-command`** - Hot-reload and invoke modified commands (body: command name, or empty for last command)

*Execution Context:*
- **Globals**: `Doc`, `Db`, `Ed` available to scripts
- **Pre-imported namespaces**: System, LINQ, AutoCAD API
- **Output**: JSON with `success`, `output`, `error`, `diagnostics` fields
- **IMPORTANT**: Perform all operations within Roslyn scripts. Do NOT use `Doc.SendStringToExecute()` as console output blocks the command queue - instead execute operations directly via the AutoCAD API

*Usage Example:*
```bash
# Read shared token
TOKEN=$(cat runtime/network/token)

# Execute Roslyn script
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

*Network Commands:*
- `select-by-categories-in-network` - Query all sessions, combine results in DataGrid with session column

*Security:*
- HTTPS with TLS 1.2/1.3 (self-signed certificates)
- Localhost-only binding (127.0.0.1)
- Shared token stored in plain text at `%appdata%/autocad-ballet/runtime/network/token`
- Full AutoCAD API access - intended for local development only

*Type Identity Issue (.NET 8 / AutoCAD 2025-2026):*
In .NET 8, assemblies loaded with rewritten identities lack `Location` property. Solution: Cache assembly with valid Location at startup, use cached type for all Roslyn operations via reflection. See `commands/_server.cs` for implementation.

*Implementation Details:*
- Auto-initialized by `AutoCADBalletStartup` extension application in `commands/_startup.cs`
- Background thread with UI marshalling via `Application.Idle` event
- Document locking for thread-safe database operations
- Network registry heartbeat every 30 seconds, 2-minute timeout for dead sessions

*Testing:*
- **Verify server status**: Check `runtime/network/sessions` for active sessions
- **Test connectivity**: Use curl with token from `runtime/network/token`
- **AI agent integration**: Use Roslyn endpoint to validate commands and inspect drawings
- **See `AGENTS.md`** for comprehensive testing guide, query patterns, and troubleshooting

For detailed usage examples, testing procedures, and API documentation, see `AGENTS.md`.

**Command Structure**:
- C# commands use `[CommandMethod]` attributes with kebab-case names (e.g., "delete-selected")
- **CRITICAL**: Every command class MUST have `[assembly: CommandClass(typeof(YourClassName))]` attribute at top of file, outside namespace
  - Without this, AutoCAD will not register commands
  - Place after `using` statements, before `namespace` declaration
- **IMPORTANT**: Do not add command aliases to `aliases.lsp` - leave alias creation to project owner

**CommandFlags.Session Usage**:
- **ALWAYS use `CommandFlags.Session`** for commands that:
  - Open, close, or switch between documents
  - Modify DocumentCollection
  - Need to work when NO document is active
  - Switch between views/layouts across documents
- **CRITICAL**: LISP wrappers for Session commands MUST include `(princ)` to prevent command context leaks
- Pattern: `(defun c:alias () (command "command-name") (princ))`

**DataGrid System** (`commands/datagrid/`):
*CRITICAL: When instructed to "use datagrid", ALWAYS use `CustomGUIs.DataGrid()` - DO NOT create custom Windows Forms DataGridView.*

*Usage Pattern:*
```csharp
var data = new List<Dictionary<string, object>>();
data.Add(new Dictionary<string, object> { ["Name"] = "Value" });
var columns = new List<string> { "Name" };
var selected = CustomGUIs.DataGrid(data, columns, spanAllScreens: false);
```

*Key Features:* Multi-select, search, column filtering/sorting, F2 edit mode, auto-formatting (PascalCase → lowercase with spaces)

*Example Commands:* `switch-view.cs`, `filter-selection.cs`, `tag-selected-in-view.cs`

## Building

```bash
# Build plugin (default 2026)
dotnet build commands/autocad-ballet.csproj

# Build for specific version
dotnet build commands/autocad-ballet.csproj -p:AutoCADYear=2024

# Build installer
dotnet build installer/installer.csproj
```

## Key Files

- `commands/autocad-ballet.csproj` - Main plugin project
- `commands/_startup.cs` - Extension application entry point (initializes server and logging)
- `commands/_server.cs` - Auto-started Roslyn server for AI agents and network queries
- `commands/_switch-view-logging.cs` - View/layout switch logging system
- `commands/_selection.cs` - Selection storage system
- `commands/_entity-attributes.cs` - Entity attributes and tag system
- `commands/datagrid/` - DataGrid UI component
- `commands/select-by-categories-in-network.cs` - Cross-session category selection

## File Naming Conventions

- **Command Files**: kebab-case (e.g., `edit-selected-text.cs`)
- **One Command Per File**: Complete implementation or stub delegating to shared code
- **Internal/Utility Files**: Underscore prefix (e.g., `_selection.cs`)

## AutoCAD .NET API Patterns

### Cross-Document Operations
Use `DocumentActivated` event for reliable cross-document operations:

```csharp
DocumentCollectionEventHandler handler = null;
handler = (sender, e) => {
    if (e.Document == targetDoc) {
        docs.DocumentActivated -= handler;
        LayoutManager.Current.CurrentLayout = targetLayout;
    }
};
docs.DocumentActivated += handler;
docs.MdiActiveDocument = targetDoc;
```

### Document Locking for Database Operations
**CRITICAL**: Always use `DocumentLock` when performing database modifications:

```csharp
using (var docLock = doc.LockDocument())
{
    using (var tr = db.TransactionManager.StartTransaction())
    {
        var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        blockTable.UpgradeOpen(); // Safe with document lock
        // ... database modifications
        tr.Commit();
    }
}
```

**When to lock:** Creating/modifying/deleting entities, upgrading to write mode, adding to symbol tables, switching layouts, working with non-active documents.

### Selection Handling Patterns
**CRITICAL**: Use `CommandFlags.UsePickSet` and check pickfirst set before prompting:

```csharp
[CommandMethod("my-command", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
public void MyCommand()
{
    var ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
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

    if (selResult.Status != PromptStatus.OK) return;
    var selectedObjects = selResult.Value.GetObjectIds();
}
```

### Selection Scope Handling
**Three Selection Scopes:**
1. **`view`** - Current space/layout (uses pickfirst set)
2. **`document`** - Entire document (uses `SelectionStorage.LoadSelection()`)
3. **`session`** - All open documents (uses `SelectionStorage.LoadSelectionFromOpenDocuments()`)

**Pattern:** Main file with `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()` methods. Stub files delegate to main implementation.

**Examples:** `filter-selection.cs`, `delete-selected.cs`, `select-by-categories.cs`, `duplicate-blocks.cs`
