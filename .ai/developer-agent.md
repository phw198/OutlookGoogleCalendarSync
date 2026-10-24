# Outlook Google Calendar Sync Developer Agent Instructions

This file contains the generic instructions and guidelines for any AI developer agent working on the Outlook Google Calendar Sync (OGCS) project.

## Core Directives
When assisting with code modifications, debugging, or documentation:
- **Be Minimalist**: Implement precisely what is requested with the fewest lines of code possible while meeting all requirements.
- **Consult Architecture**: Before touching Outlook or Google Calendar sync code, review [.ai/architecture.md](architecture.md).
- **Adhere to Coding Standards**: Refer to [.ai/coding-standards.md](coding-standards.md) for C# coding style, WinForms/thread safety, naming conventions, logging standards, and line ending requirements.
- **Follow Naming and Error Handling Rules**: Use camelCase for private methods and handlers, and use `Analyse()` with a contextual message for exception handling rather than bare exception logging.
- **Prefer Explicit Types**: Avoid using `var` for local variables unless the type is immediately obvious; prefer explicit type declarations to keep code clear.
- **Avoid Redundant Locals**: Do not assign a value to a local variable when it is used only once and the variable does not meaningfully improve readability.
- **Manage Dependencies Deliberately**: Prefer `packages.config` and references under `..\packages\<package>.<version>\...` for standard third-party libraries. Keep `lib` for bespoke, patched, generated, or configuration-specific assemblies, including selected Outlook interop variants. Migrate incrementally, validate restore/build/runtime/package output, and resolve NuGet conflicts rather than hiding them with manually copied DLLs. Keep compile-only `.winmd` metadata references set to `<Private>False</Private>`.
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
- **Memory Log**: Always refer to the branch-scoped WIP note for current progress and active tasks. Use the root-level [AGENTS.md](../AGENTS.md) for significant project-level milestones and durable guidance, and update it only when such a milestone is established.
- **Knowledge Base**: Consult [.ai/knowledge/index.md](knowledge/index.md) when an issue may match stored domain knowledge. Add concise, source-linked case notes there when an investigation establishes durable context that is not specific to one branch's current progress.
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

### **Linux Compile-Only Validation**

The .NET 8 SDK can be used on Linux to compile-check this .NET Framework 4.7.2 solution when the targeting references and resource-generation dependency are supplied temporarily. This does not make the application or its full test suite runnable on Linux.

With the repository's NuGet dependencies already restored, prepare temporary references outside the repository:

```bash
mkdir -p /tmp/ogcs-net472-ref /tmp/ogcs-resources-ref
curl -fLsS -o /tmp/ogcs-net472.nupkg \
  https://api.nuget.org/v3-flatcontainer/microsoft.netframework.referenceassemblies.net472/1.0.3/microsoft.netframework.referenceassemblies.net472.1.0.3.nupkg
unzip -q -o /tmp/ogcs-net472.nupkg 'build/.NETFramework/*' -d /tmp/ogcs-net472-ref
curl -fLsS -o /tmp/ogcs-resources.nupkg \
  https://api.nuget.org/v3-flatcontainer/system.resources.extensions/4.7.1/system.resources.extensions.4.7.1.nupkg
unzip -q -o /tmp/ogcs-resources.nupkg 'lib/net461/System.Resources.Extensions.dll' -d /tmp/ogcs-resources-ref
```

Create `/tmp/ogcs-build-targets.targets` with the resource assembly reference:

```xml
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <ItemGroup>
    <Reference Include="System.Resources.Extensions">
      <HintPath>/tmp/ogcs-resources-ref/lib/net461/System.Resources.Extensions.dll</HintPath>
    </Reference>
  </ItemGroup>
</Project>
```

Build the test project and its app project reference:

```bash
dotnet build src/OutlookGoogleCalendarSync.Tests/OutlookGoogleCalendarSync.Tests.csproj --no-restore -v:q \
  -p:DefineConstants=DEVELOP_AGAINST_2007 \
  -p:TargetFrameworkRootPath=/tmp/ogcs-net472-ref/build/ \
  -p:GenerateResourceUsePreserializedResources=true \
  -p:CustomAfterMicrosoftCommonTargets=/tmp/ogcs-build-targets.targets \
  -p:GenerateClickOnceManifests=false -p:PostBuildEvent=
```

The C# compile can complete before the build fails in the app's Windows-only PowerShell post-build script (for example, `fixValueTuple.ps1`). Treat compiler diagnostics such as `error CS...` as compile failures; a PowerShell/post-build failure still means the overall build did not succeed. Once the app assembly has been produced, compile-check the test assembly separately:

```bash
dotnet build src/OutlookGoogleCalendarSync.Tests/OutlookGoogleCalendarSync.Tests.csproj --no-restore -v:q \
  -p:BuildProjectReferences=false \
  -p:DefineConstants=DEVELOP_AGAINST_2007 \
  -p:TargetFrameworkRootPath=/tmp/ogcs-net472-ref/build/ \
  -p:GenerateResourceUsePreserializedResources=true \
  -p:CustomAfterMicrosoftCommonTargets=/tmp/ogcs-build-targets.targets \
  -p:GenerateClickOnceManifests=false -p:PostBuildEvent=
```

Do not use this workflow to claim that tests passed. The Linux `dotnet vstest` installation may also lack `testhost.net472.exe`; the full .NET Framework test runner is not available there. There is no value in repeatedly attempting to run the full .NET Framework test suite on Linux: execute it on Windows with the .NET Framework 4.7.2 Developer Pack and a compatible test host. Outlook COM integration and UI behavior must be validated on Windows with classic Outlook installed.

### **Visual Studio (Windows VM) Workflow**

- **Full Builds and Testing**: Full builds, application execution, and test running must be performed within Visual Studio Community 2019 on the Windows 10 VirtualBox guest. The Linux procedure above is supplementary compile-only validation.
- **MSTest Framework**: Use the "Unit Test Project (.Net Framework)" template when creating new test projects, as this aligns with the existing project's target framework and leverages built-in Visual Studio features.
- **Assembly Reference Management**:
    - **Project References**: For custom code within the solution (e.g., `OutlookGoogleCalendarSync` project), add a direct Project Reference from the test project.
    - **NuGet Package References**: For external libraries (e.g., Google APIs, Microsoft Kiota, Microsoft Graph SDK), explicitly install the *same NuGet packages with matching versions* into the test project.
    - **Linked Files**: Avoid adding source files from the main project as "linked files" in the test project (`<Compile Include="..." />` with `<Link>...</Link>` for external files). This causes `CS0436` type conflicts due to duplicate type definitions across assemblies. Such links should be removed from the `.csproj` file.
    - **Assembly Binding Redirects (`app.config`)**: Ensure that the `app.config` file in the test project contains `assemblyBinding` redirects that are **identical** to the main application's `app.config`. This is critical for resolving conflicts between different versions of shared assemblies (e.g., `System.Diagnostics.DiagnosticSource`, `System.Memory`, `System.Buffers`) at runtime. Manually synchronize these if necessary.