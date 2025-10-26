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

**Selection Management System** (`_selection.cs`):
- Provides cross-document selection capabilities
- SelectionStorage class for persisting selections across sessions
- SelectionItem class with DocumentPath, Handle, and SessionId properties
- Persistent selection storage using handles in `%APPDATA%/autocad-ballet/runtime/selection/`
- Per-document selection files with format: `DocumentPath,SessionId` followed by comma-separated handles
- Extension methods for Editor to support legacy `SetImpliedSelectionEx()` calls

**Command Structure**:
- C# commands use `[CommandMethod]` attributes with kebab-case names (e.g., "delete-selected")
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

**DataGrid Display Conventions**:
- Column headers are automatically formatted to lowercase with spaces between words
- Example: "DocumentName" becomes "document name", "LayoutType" becomes "layout type"
- This is handled by the `FormatColumnHeader` method in `datagrid/datagrid.Main.cs`
- Actual property names remain unchanged (e.g., "DocumentName"), only display headers are formatted
- DataGrid implementation is modular with separate files for filtering, sorting, virtualization, edit mode, and cache indexing

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
- `copy-selection-to-views-in-application.cs` - Uses this pattern correctly (note: renamed from "process" to "application" scope)

### Selection Scope Handling
**CRITICAL**: AutoCAD Ballet provides three selection scopes through separate command variants: view, document, and application.

**The Three Selection Scopes**:
1. **`view`** - Current active space/layout only (default AutoCAD behavior, uses pickfirst set)
2. **`document`** - Entire current document (all layouts, uses stored selection)
3. **`application`** - All opened documents in current AutoCAD process (uses stored selection)

**Architecture**: Commands use a **scope-based command pattern**:
- Each command has three variants: `{command}-in-view`, `{command}-in-document`, `{command}-in-application`
- Main implementation file contains separate `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()` methods
- Scope-specific command files are stubs that call the appropriate method

**The Problem**: Commands that only handle view scope don't provide the cross-document capabilities users expect from AutoCAD Ballet.

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

    // Application scope: Use stored selection from all documents
    public static void ExecuteApplicationScope(Editor ed)
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var db = doc.Database;

        var storedSelection = SelectionStorage.LoadSelectionFromAllDocuments();
        if (storedSelection == null || storedSelection.Count == 0)
        {
            ed.WriteMessage("\nNo stored selection found. Use 'select-by-categories-in-application' first.\n");
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

// Stub file: my-command-in-application.cs
public class MyCommandInApplication
{
    [CommandMethod("my-command-in-application", CommandFlags.Modal)]
    public void MyCommandInApplicationCommand()
    {
        var doc = AcadApp.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            MyCommand.ExecuteApplicationScope(ed);
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError in my-command-in-application: {ex.Message}\n");
        }
    }
}
```

**Key Selection Scope Principles**:
- **Implement three separate methods**: `ExecuteViewScope()`, `ExecuteDocumentScope()`, `ExecuteApplicationScope()`
- **View scope**: Use pickfirst set or prompt for selection (standard AutoCAD behavior with `CommandFlags.UsePickSet`)
- **Document scope**: Use `SelectionStorage.LoadSelection(docName)` and filter to current document
- **Application scope**: Use `SelectionStorage.LoadSelectionFromAllDocuments()` for cross-document selections
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
  - Stub files: `delete-selected-in-view.cs`, `delete-selected-in-document.cs`, `delete-selected-in-application.cs`
- `select-by-categories.cs` - Selection by entity category with scope variants
  - Stub files: `select-by-categories-in-view.cs`, `select-by-categories-in-document.cs`, `select-by-categories-in-application.cs`
- `duplicate-blocks.cs` - Block duplication with all selection scopes
  - Stub files: `duplicate-blocks-in-view.cs`, `duplicate-blocks-in-application.cs`
