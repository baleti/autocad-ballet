# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AutoCAD Ballet is a collection of custom commands for AutoCAD that provides enhanced selection management and productivity tools. The project consists of two main components:

1. **Commands** (`/commands/`) - C# plugin with custom AutoCAD commands and LISP utilities
2. **Installer** (`/installer/`) - Windows Forms installer application

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

**Selection Management System** (`set-selection-scope.cs`):
- Provides cross-document selection capabilities  
- Five selection scopes: SpaceLayout, Drawing, Process, Desktop, Network
- Persistent selection storage using handles in `%APPDATA%/autocad-ballet/`
- Extension methods for Editor to support enhanced selection operations

**Shared Utilities** (`Shared.cs`):
- SelectionStorage class for persisting selections across sessions
- Uses `%APPDATA%/autocad-ballet/selection` file format: `DocumentPath|Handle`

**Command Structure**:
- C# commands use `[CommandMethod]` attributes with kebab-case names (e.g., "set-scope")
- LISP files provide aliases and utility functions
- Main command invocation through `InvokeAddinCommand.cs`
- **IMPORTANT**: Do not add command aliases to `aliases.lsp` - leave alias creation to the project owner

**DataGrid Display Conventions**:
- Column headers are automatically formatted to lowercase with spaces between words
- Example: "DocumentName" becomes "document name", "LayoutType" becomes "layout type"
- This is handled by the `FormatColumnHeader` method in `DataGrid2_Main.cs`
- Actual property names remain unchanged (e.g., "DocumentName"), only display headers are formatted

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

# Output: installer/bin/installer.exe
```

## Key Files

- `commands/autocad-ballet.csproj` - Main plugin project with multi-version support
- `commands/set-selection-scope.cs` - Core selection management system
- `commands/Shared.cs` - Utility classes for selection persistence
- `commands/FilterSelected.cs` - Entity filtering and data grid functionality
- `commands/_edit-dialog.cs` - Shared text editing dialog (internal utility)
- `commands/aliases.lsp` - LISP command aliases
- `installer/installer.csproj` - Installer project that embeds plugin DLLs

## File Naming Conventions

**Command Files**: Use kebab-case for AutoCAD command files (e.g., `edit-selected-text.cs`, `switch-view-last.cs`)

**Internal/Utility Files**: Use underscore prefix with kebab-case for files that are not actual commands but provide internal functionality (e.g., `_edit-dialog.cs`, `_utilities.cs`)

This convention helps distinguish between:
- **Commands**: Files that define AutoCAD commands with `[CommandMethod]` attributes
- **Internal utilities**: Shared classes, dialogs, and helper functionality used by commands

## Installation Structure

The plugin installs to `%APPDATA%/autocad-ballet/` with:
- Plugin DLLs for each AutoCAD version
- LISP files for command aliases
- Selection persistence files
- Mode configuration files

Runtime data is stored in `%APPDATA%/autocad-ballet/runtime` for selection scope management.

## Command Porting

Many commands in AutoCAD Ballet are ported from the revit-ballet repository. When porting commands from Revit to AutoCAD:

- **Views → Layouts**: Revit views correspond to AutoCAD layouts
- **ElementId → Layout Names**: Revit uses numeric ElementIds, AutoCAD uses string layout names
- **View Change Logging**: Both systems use similar logging patterns:
  - Revit: `%APPDATA%/revit-scripts/LogViewChanges/{ProjectName}`
  - AutoCAD: `%APPDATA%/autocad-ballet/runtime/LogLayoutChanges/{ProjectName}`
- **API Differences**:
  - Revit: `uidoc.ActiveView = view`
  - AutoCAD: `LayoutManager.Current.CurrentLayout = layoutName`
- **Document Access**: Both use similar document manager patterns but with different namespaces
- **Runtime Data Storage**: All runtime information (logs, temporary data) should be stored in `%APPDATA%/autocad-ballet/runtime/`

Commands successfully ported from revit-ballet include:
- `switch-view-last` - Switches to the most recent layout from log history
- `switch-view-recent` - Shows layouts ordered by usage history for selection

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
- `copy-selection-to-views-in-process.cs` - Uses this pattern correctly

### Selection Scope Handling
**CRITICAL**: All AutoCAD Ballet commands must properly handle the five selection scopes. Commands should check `SelectionScopeManager.CurrentScope` and adapt their selection behavior accordingly.

**The Five Selection Scopes**:
1. **`view`** - Current active space/layout only (default AutoCAD behavior)
2. **`document`** - Entire current document (all layouts)
3. **`process`** - All opened documents in current AutoCAD process
4. **`desktop`** - All documents in all AutoCAD instances on desktop
5. **`network`** - Across network (future implementation)

**The Problem**: Commands that only handle view scope break the enhanced selection workflow that users expect from AutoCAD Ballet.

**The Solution**: Always implement scope-aware selection handling:

```csharp
[CommandMethod("my-command", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
public void MyCommand()
{
    var doc = AcadApp.DocumentManager.MdiActiveDocument;
    var ed = doc.Editor;
    var currentScope = SelectionScopeManager.CurrentScope;

    if (currentScope == SelectionScopeManager.SelectionScope.view)
    {
        // Handle view scope: Use pickfirst set or prompt for selection
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

        // Process selection...
    }
    else
    {
        // Handle document, process, desktop, network scopes: Use stored selection
        var storedSelection = SelectionStorage.LoadSelection();
        if (storedSelection == null || storedSelection.Count == 0)
        {
            ed.WriteMessage("\nNo stored selection found. Use commands like 'select-by-categories' first or switch to 'view' scope.\n");
            return;
        }

        // Filter to current document if in document scope
        if (currentScope == SelectionScopeManager.SelectionScope.document)
        {
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
        }

        if (storedSelection.Count == 0)
        {
            ed.WriteMessage($"\nNo stored selection found for current scope '{currentScope}'.\n");
            return;
        }

        // Process stored selection items...
        foreach (var item in storedSelection)
        {
            // Handle current document items
            if (string.Equals(Path.GetFullPath(item.DocumentPath), Path.GetFullPath(doc.Name), StringComparison.OrdinalIgnoreCase))
            {
                var handle = Convert.ToInt64(item.Handle, 16);
                var objectId = db.GetObjectId(false, new Handle(handle), 0);
                // Process current document entity...
            }
            else
            {
                // Handle external document items (report or skip as appropriate)
                ed.WriteMessage($"\nNote: External entity found but cannot be processed: {item.DocumentPath}\n");
            }
        }

        // Update stored selection with results if needed
        // ed.SetImpliedSelectionEx(newObjectIds); // Use extension method for scope-aware selection setting
    }
}
```

**Key Selection Scope Principles**:
- **Always check `SelectionScopeManager.CurrentScope`** before processing selection
- **View scope**: Use pickfirst set or prompt for selection (standard AutoCAD behavior)
- **Other scopes**: Use `SelectionStorage.LoadSelection()` to get cross-document selections
- **Document scope filtering**: Filter stored selection to current document only
- **Cross-document limitations**: Many operations can only be performed on current document entities
- **Update stored selection**: Use `ed.SetImpliedSelectionEx()` to update selection storage in non-view scopes
- **Clear error messages**: Guide users to use `select-by-categories` or switch scope when no stored selection exists

**Required Imports for Selection Scope Support**:
```csharp
using AutoCADBallet;  // For SelectionScopeManager, SelectionStorage, SelectionItem
using System.IO;      // For Path operations in cross-document scenarios
```

**Examples in Codebase**:
- `filter-selection.cs` lines 37-86 - Perfect reference implementation
- `duplicate-block.cs` - Updated to handle all selection scopes properly
- `set-scope.cs` - Core selection scope management system
