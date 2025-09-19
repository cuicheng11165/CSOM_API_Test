# Copilot Instructions for CSOM_API_Test

## Project Overview
This repository contains demo and utility projects for working with SharePoint Online using the CSOM (Client-Side Object Model) API. Each subfolder is a self-contained C# project targeting a specific SharePoint scenario, such as authentication, file operations, taxonomy, permissions, and more.

## Key Components & Structure
- **Authentication & Token Export**: `ExportCsomToken/`, `ExportGraphToken/`, `ExportGraphDelegateToken/` handle authentication and token export for CSOM and Graph APIs. Config values are in `Config/`.
- **Configuration**: `Configuration/` contains shared config logic and project references.
- **Scenario Demos**: Each folder (e.g., `CSOM File Add Test/`, `CSOM_CheckPermission/`, `CSOM_View_Test/`, `UpdateConlictSample/`) demonstrates a specific SharePoint CSOM use case.
- **Test Projects**: Some folders (e.g., `CSOM_ExceptionHandlingScope_Test/`, `TimeZoneTest/`) are for testing or illustrating API behaviors.

## Developer Workflows
- **Build**: Use Visual Studio or `dotnet build` on the individual `.csproj` files. Solution files (`*.sln`) group related projects but are not always required.
- **Run**: Most projects have a `Program.cs` entry point. Run with Visual Studio or `dotnet run --project <ProjectFolder>/<ProjectName>.csproj`.
- **Configuration**: Secrets and environment-specific values are stored in `Config/` as plain text files (e.g., `ClientId.txt`, `TenantId.txt`).
- **Dependencies**: Managed via `packages.config` in each project. Use NuGet restore if needed.

## Project-Specific Conventions
- **No shared library**: Most code is duplicated or copy-pasted between projects for demo isolation.
- **Config values**: Read from `Config/` at runtime. Do not hardcode secrets.
- **Chinese comments**: Some files contain Chinese-language comments for context.
- **Minimal error handling**: Demos focus on API usage, not production robustness.

## Integration & Patterns
- **SharePoint Online**: All code targets SharePoint Online via CSOM or Graph API.
- **DLLs**: Some projects reference local or NuGet-provided SharePoint DLLs.
- **No cross-project dependencies**: Each project is standalone unless explicitly referencing `Configuration/`.

## Examples
- To test file upload: see `CSOM File Add Test/Program.cs`.
- To test taxonomy: see `CSOM Taxonomy Field Update/`.
- To export tokens: see `ExportCsomToken/` and `ExportGraphToken/`.

## Tips for AI Agents
- Always check `Config/` for required runtime values.
- When adding a new scenario, copy an existing project as a template.
- Keep each demo self-contained unless sharing config logic.
- Use English for code, but Chinese comments are acceptable for explanations.

---
For more details, see the `README.md` or the entry point in each project folder.
