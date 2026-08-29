# Outlook Google Calendar Sync Developer Agent Instructions

This file contains the generic instructions and guidelines for any AI developer agent working on the Outlook Google Calendar Sync (OGCS) project.

## Core Directives
When assisting with code modifications, debugging, or documentation:
- **Be Minimalist**: Implement precisely what is requested with the fewest lines of code possible while meeting all requirements.
- **Consult Architecture**: Before touching Outlook or Google Calendar sync code, review [.ai/architecture.md](architecture.md).
- **Adhere to Coding Standards**: Refer to [.ai/coding-standards.md](coding-standards.md) for C# coding style, WinForms/thread safety, naming conventions, logging standards, and line ending requirements.
- **Follow Naming and Error Handling Rules**: Use camelCase for private methods and handlers, and use `Analyse()` with a contextual message for exception handling rather than bare exception logging.
- **Prefer Explicit Types**: Avoid using `var` for local variables unless the type is immediately obvious; prefer explicit type declarations to keep code clear.
- **Use CRLF Line Endings**: Save all repository source code and text files with CRLF line endings.
- **Maintain UI Responsiveness**: Since OGCS is a WinForms tray application, always avoid blocking the UI thread. Use background tasks or async patterns carefully.

## Agent Architecture
- **Agnostic Instructions**: All rich, platform-independent agent instructions and workflows must reside in `.ai/agents/` (e.g., `.ai/agents/ogcs-release-preparer.md`), global guidance in `.github/instructions/`, and skills in `.ai/skills/` (e.g., `.ai/skills/update-ai-files/SKILL.md`).
- **Platform Redirections (Shims)**: Platform-specific agent and skill definitions must be minimalist "pointers" or redirections in standard locations.
  - **Custom Agents**: Pointers reside at `.github/agents/ogcs-*.agent.md`. They require YAML frontmatter including `name`, `description`, and `tools`.
  - **Global Guidance**: Workspace-wide guidance resides at `.github/instructions/ogcs-*.instructions.md` with `applyTo: "**"`; do not represent this guidance as a selectable agent.
  - **Custom Skills**: Pointers reside at `.github/skills/<name>/SKILL.md`. They require YAML frontmatter including `name`, `description`, and `user-invocable: true`.
- **Naming Consistency**: Selectable custom agents and prompts must use the `ogcs-` prefix, and their YAML `name` must exactly match the filename. Skills retain unprefixed names unless they are OGCS-specific.
- **UI Discovery**: VS Code requires these shims in the `.github/` hierarchy for UI integration. A "Reload Window" should not be required after adding or modifying these shims.

## Key Files to Monitor
- **Memory Log**: Always refer to the root-level [AGENTS.md](../AGENTS.md) to understand current progress, milestones, and active tasks, and keep it updated as changes are completed.
- **Rules Config**: Ensure actions adhere to the root-level [.clinerules](../.clinerules).

## Development Environment Specifics

This project is developed under a specific cross-platform setup that requires careful consideration:

- **Host OS**: Linux (VS Code runs on this host).
- **Development VM**: Windows 10 VirtualBox guest (Visual Studio Community 2019 is used here for building, debugging, and running tests).
- **Target Framework**: .NET Framework 4.7.2 (for the main application and test project).

### **VS Code (Linux Host) Setup for C#**

- **.NET SDK Requirement**: The VS Code C# extensions (including the Language Server for IntelliSense and code analysis) require **.NET SDK 8.0** to be installed on the Linux host. This is purely for editor functionality and does not affect the target framework of the project building on Windows.
  - *Rationale*: The Language Server itself is a .NET Core/.NET application and needs a compatible runtime to execute. It then uses .NET Framework reference assemblies (if configured via `Microsoft.NETFramework.ReferenceAssemblies` NuGet package) to provide correct IntelliSense for .NET Framework projects.
- **`dotnet` executable**: Ensure the `dotnet` executable is in the system's PATH or explicitly configured in VS Code settings (`dotnet.dotnetPath`).

### **Visual Studio (Windows VM) Workflow**

- **Building and Testing**: All actual compilation, execution, and test running must be performed within Visual Studio Community 2019 on the Windows 10 VirtualBox guest.
- **MSTest Framework**: Use the "Unit Test Project (.Net Framework)" template when creating new test projects, as this aligns with the existing project's target framework and leverages built-in Visual Studio features.
- **Assembly Reference Management**:
    - **Project References**: For custom code within the solution (e.g., `OutlookGoogleCalendarSync` project), add a direct Project Reference from the test project.
    - **NuGet Package References**: For external libraries (e.g., Google APIs, Microsoft Kiota, Microsoft Graph SDK), explicitly install the *same NuGet packages with matching versions* into the test project.
    - **Linked Files**: Avoid adding source files from the main project as "linked files" in the test project (`<Compile Include="..." />` with `<Link>...</Link>` for external files). This causes `CS0436` type conflicts due to duplicate type definitions across assemblies. Such links should be removed from the `.csproj` file.
    - **Assembly Binding Redirects (`app.config`)**: Ensure that the `app.config` file in the test project contains `assemblyBinding` redirects that are **identical** to the main application's `app.config`. This is critical for resolving conflicts between different versions of shared assemblies (e.g., `System.Diagnostics.DiagnosticSource`, `System.Memory`, `System.Buffers`) at runtime. Manually synchronize these if necessary.