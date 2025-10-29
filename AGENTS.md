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
    - `runtime/InvokeAddinCommand-history` — Command invocation history
  - NETLOAD checks across targeted AutoCAD years (2017-2026).
- For bug fixes, include repro steps and expected vs. actual results in the PR.

## Commit & Pull Request Guidelines
- Commits: imperative, concise, lowercase (e.g., `fix startup loader path match`).
- PRs: clear description, scope, targeted AutoCAD years, manual test notes, and screenshots for UI/installer changes.
- Link issues and note when porting from revit-ballet.

## Security & Configuration Tips
- Never hardcode user paths; use `%APPDATA%/autocad-ballet/…` for data and runtime files.
- Do not commit runtime artifacts; the `runtime/` directory in the repository is gitignored and contains only local development/test data.

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
