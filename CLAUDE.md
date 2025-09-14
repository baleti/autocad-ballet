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
- `commands/aliases.lsp` - LISP command aliases
- `installer/installer.csproj` - Installer project that embeds plugin DLLs

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

### Document Locking for Layout Operations
**CRITICAL**: AutoCAD layout switching requires explicit document locking for reliable operation, especially in cross-document scenarios.

**The Problem**:
- `LayoutManager.Current.CurrentLayout = layoutName` can fail with `eSetFailed` if the document context isn't properly established
- This is especially common after document switching, even when using the `DocumentActivated` event

**The Solution**: Always use `DocumentLock` when performing layout operations:

```csharp
// INCORRECT - May fail with eSetFailed
LayoutManager.Current.CurrentLayout = targetLayout;

// CORRECT - Reliable with document lock
using (DocumentLock docLock = targetDoc.LockDocument())
{
    LayoutManager.Current.CurrentLayout = targetLayout;
}
```

**Key Benefits**:
- **Eliminates `eSetFailed` exceptions** in cross-document layout switching
- **Works directly from Model Space** - no need for intermediate "priming" layouts
- **Prevents document context conflicts** between multiple operations
- **Essential for event handlers** that operate on non-active documents

**Document Locking Best Practice**: Always lock the document when:
- Switching layouts after document activation
- Performing layout operations in `DocumentActivated` event handlers
- Working with non-active documents
- Any operation that modifies document state across document boundaries

This pattern is essential for reliable AutoCAD .NET development and prevents many common layout switching issues.

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
