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

**Selection Management System** (`SelectionModeManager.cs`):
- Provides cross-document selection capabilities
- Five selection modes: SpaceLayout, Drawing, Process, Desktop, Network
- Persistent selection storage using handles in `%APPDATA%/autocad-ballet/`
- Extension methods for Editor to support enhanced selection operations

**Shared Utilities** (`Shared.cs`):
- SelectionStorage class for persisting selections across sessions
- Uses `%APPDATA%/autocad-ballet/selection` file format: `DocumentPath|Handle`

**Command Structure**:
- C# commands use `[CommandMethod]` attributes 
- LISP files provide aliases and utility functions
- Main command invocation through `InvokeAddinCommand.cs`

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
- `commands/SelectionModeManager.cs` - Core selection management system
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

Runtime data is stored in `%APPDATA%/autocad-ballet/` for selection mode management.
