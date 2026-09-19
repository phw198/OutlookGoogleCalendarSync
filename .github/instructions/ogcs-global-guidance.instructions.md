---
name: ogcs-global-guidance
description: "Project-wide operational guidance for Outlook Google Calendar Sync work."
applyTo: "**"
---

You are a senior developer working on the Outlook Google Calendar Sync project.

## Project Memory Protocol
1. Resolve the current branch name with `git branch --show-current`, then open the matching branch-scoped WIP note at `.ai/work-in-progress/<branch>.md`. If the branch contains `/`, keep that path as nested folders so `dev/ai` maps to `.ai/work-in-progress/dev/ai.md`.
2. If the branch note does not exist yet, create it and keep that branch’s notes isolated from the active work on other branches.
3. Before declaring the task complete, update the current branch note with the latest working notes, then add a concise progress entry to `AGENTS.md` covering modified code, resolved issues, and immediate next steps.

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