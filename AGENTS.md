# Repository Guidelines

## Project Structure & Module Organization
- `commands/` — C# AutoCAD plugin and AutoLISP utilities (`*.lsp`).
  - Multi-version targeting via `AutoCADYear` in `autocad-ballet.csproj`.
- `installer/` — Windows Forms installer that bundles built artifacts.
- `runtime/` — Example runtime folders; real data lives in `%APPDATA%/autocad-ballet/`.
- Key files: `commands/autocad-ballet.csproj`, `commands/Shared.cs`, `commands/set-selection-scope.cs`, `commands/aliases.lsp`, `installer/installer.csproj`.

## Build, Test, and Development Commands
- Build plugin (default 2026): `dotnet build commands/autocad-ballet.csproj`
- Build for a specific AutoCAD version: `dotnet build commands/autocad-ballet.csproj -p:AutoCADYear=2024`
  - Output: `commands/bin/{AutoCADYear}/autocad-ballet.dll`
- Build installer (Windows only): `dotnet build installer/installer.csproj`
  - Output: `installer/bin/installer.exe`
- Manual smoke test in AutoCAD: copy `.lsp` to `%APPDATA%/autocad-ballet/`, NETLOAD the DLL, run aliases from `aliases.lsp`.

## Coding Style & Naming Conventions
- C#: 4-space indent; PascalCase for types/methods; camelCase for locals/fields; one class per file when feasible.
- Commands: create under `commands/` with `[CommandMethod]`; share helpers in `Shared.cs`.
- AutoLISP: filenames kebab-case (e.g., `select-current-layer.lsp`); define commands as `c:command-name`; maintain shortcuts in `aliases.lsp`.
- Keep logic small and composable; avoid UI in core command classes.

## Testing Guidelines
- No formal test suite yet. Validate by:
  - Running commands on sample drawings and checking behavior.
  - Verifying selection persistence under `%APPDATA%/autocad-ballet/runtime`.
  - NETLOAD checks across targeted AutoCAD years.
- For bug fixes, include repro steps and expected vs. actual results in the PR.

## Commit & Pull Request Guidelines
- Commits: imperative, concise, lowercase (e.g., `fix startup loader path match`).
- PRs: clear description, scope, targeted AutoCAD years, manual test notes, and screenshots for UI/installer changes.
- Link issues and note when porting from revit-ballet.

## Security & Configuration Tips
- Never hardcode user paths; use `%APPDATA%/autocad-ballet/…` for data and runtime files.
- Do not commit runtime artifacts; keep writes inside the user profile directory.
