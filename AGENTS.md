# Repository Guidelines

## Project Structure & Module Organization
- `commands/` — C# AutoCAD plugin and AutoLISP utilities (`*.lsp`).
  - Multi-version targeting via `AutoCADYear` in `autocad-ballet.csproj`.
- `installer/` — Windows Forms installer that bundles built artifacts.
- `runtime/` — Local development/test runtime data (gitignored); production data lives in `%APPDATA%/autocad-ballet/runtime/`.
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
  - Verifying selection persistence under `%APPDATA%/autocad-ballet/runtime/selection/`.
  - Checking layout history logs under `%APPDATA%/autocad-ballet/runtime/switch-view-logs/`.
  - NETLOAD checks across targeted AutoCAD years (2017-2026).
- For bug fixes, include repro steps and expected vs. actual results in the PR.

## Commit & Pull Request Guidelines
- Commits: imperative, concise, lowercase (e.g., `fix startup loader path match`).
- PRs: clear description, scope, targeted AutoCAD years, manual test notes, and screenshots for UI/installer changes.
- Link issues and note when porting from revit-ballet.

## Security & Configuration Tips
- Never hardcode user paths; use `%APPDATA%/autocad-ballet/…` for data and runtime files.
- Do not commit runtime artifacts; the `runtime/` directory in the repository is gitignored and contains only local development/test data.
