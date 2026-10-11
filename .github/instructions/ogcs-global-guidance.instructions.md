---
name: ogcs-global-guidance
description: "Project-wide operational guidance for Outlook Google Calendar Sync work."
applyTo: "**"
---

You are a senior developer working on the Outlook Google Calendar Sync project.

## Project Memory Protocol
1. When updating WIP notes, always check the current branch with `git branch --show-current` before editing anything and make sure you are updating the correct branch-scoped notes file.
2. Resolve the current branch name, then open the matching branch-scoped WIP note at `.ai/work-in-progress/<branch>.md`. If the branch contains `/`, keep that path as nested folders so `dev/ai` maps to `.ai/work-in-progress/dev/ai.md`.
3. If the branch note does not exist yet, create it and keep that branch�s notes isolated from the active work on other branches.
4. Before declaring the task complete, update the current branch note with the latest working notes. Do not add progress entries to the tracked `AGENTS.md`; branch progress belongs in the ignored `.ai/work-in-progress/` notes.
5. Keep the active issue branch limited to the work for that specific issue. Do not carry unrelated changes from other sessions, other branches, opportunistic refactors, release prep, or AI-guidance edits into the branch unless the user explicitly authorizes them.
6. If unrelated work is already present on the branch, stop, call it out, and ask for explicit direction before merging or preserving it. Do not silently fold unrelated edits into the issue fix.

## UI vs Worker Thread Blocking Rule
- The background sync worker may block on Graph or Outlook calls because it does not service WinForms UI interaction.
- Any method reached directly from a WinForms UI event handler, form callback, or other UI-thread path must not block; it must remain async or must marshal onto a background worker before waiting.
- Blocking calls such as `Result`, `Wait()`, and `GetAwaiter().GetResult()` are acceptable only when the call is guaranteed to run on the background sync worker rather than on the UI thread.
- Treat any UI-thread sync wait on Graph/Outlook APIs as a bug unless it is explicitly isolated to a non-UI background worker.

## Outlook COM Leak Review
When any code related to the Outlook client is changed, explicitly review it for COM lifecycle leaks before finishing the task.

Checklist:
- Inspect every new or modified Outlook access path for COM object creation, including `Outlook.Calendar.Instance`, `Outlook.Graph.Calendar.Instance`, `GetActiveObject()`, `new Outlook.Application`, `Categories.Add`, `MAPIFolder`, `AppointmentItem`, `Category`, `Store`, `Items`, and other Office Interop objects.
- Ensure each created COM object has a matching cleanup path, typically via `Outlook.Calendar.Disconnect(...)`, `Calendar.ReleaseObject(...)`, or a `try/finally` block that runs on every exit path.
- Check UI handlers and other non-sync code paths for singleton auto-connects that trigger Outlook without a follow-up disconnect.
- Treat missing `finally` cleanup, early returns, and exception paths as leaks until proven otherwise.
- If a change introduces an Outlook COM dependency, call out in the completion notes whether cleanup was verified, and note any remaining Windows-only validation required.

## Code Simplicity
- Avoid assigning a value to a local variable when it is used only once and the variable does not meaningfully improve readability.

## Dependency and Assembly Reference Management
- Prefer NuGet-managed references for standard third-party dependencies. In this legacy project, use `packages.config` and references resolving from `..\packages\<package>.<version>\...` rather than manually copied DLLs under `lib`.
- Keep `lib` for bespoke, patched, generated, or configuration-specific assemblies that are not suitable for ordinary NuGet management, including the explicitly selected Outlook interop variants.
- Migrate references incrementally. When moving a dependency, verify restore, clean Debug and Release builds, runtime output, and release packaging before removing its old `lib` copy.
- Treat NuGet version conflicts and restore failures as useful dependency-management feedback to resolve deliberately; do not preserve manual DLL copies merely to hide them.
- Keep compile-only metadata such as Windows SDK `.winmd` files as explicit non-copying references (`Private=False`); do not ship them as runtime dependencies.

## Release Website Documentation Review
- Every release-preparation task must compare the release's user-facing changes with the current `gh-pages` User Guide, including settings, workflows, notifications, integrations, and troubleshooting guidance.
- Every release-preparation task must also update `gh-pages/release-notes.md` by copying the nuspec release notes into a new section above the previous release. Keep the standard advert section between releases, and add `----` immediately below beta release sections.
- Fast-forward the local `gh-pages` branch before editing it, preserve unrelated existing work, and keep User Guide changes isolated to that branch.
- If the review finds no required changes, record that conclusion and its rationale in the branch-scoped WIP note and final release summary.