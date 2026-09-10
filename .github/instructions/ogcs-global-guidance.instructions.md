---
name: ogcs-global-guidance
description: "Project-wide operational guidance for Outlook Google Calendar Sync work."
applyTo: "**"
---

You are a senior developer working on the Outlook Google Calendar Sync project.

## Project Memory Protocol
1. At the start of every task, read `AGENTS.md` to understand the current architecture, recent changes, and project trajectory.
2. Before declaring the task complete, update `AGENTS.md` with a concise progress entry covering modified code, resolved issues, and immediate next steps.

## Outlook COM Leak Review
When any code related to the Outlook client is changed, explicitly review it for COM lifecycle leaks before finishing the task.

Checklist:
- Inspect every new or modified Outlook access path for COM object creation, including `Outlook.Calendar.Instance`, `Outlook.Graph.Calendar.Instance`, `GetActiveObject()`, `new Outlook.Application`, `Categories.Add`, `MAPIFolder`, `AppointmentItem`, `Category`, `Store`, `Items`, and other Office Interop objects.
- Ensure each created COM object has a matching cleanup path, typically via `Outlook.Calendar.Disconnect(...)`, `Calendar.ReleaseObject(...)`, or a `try/finally` block that runs on every exit path.
- Check UI handlers and other non-sync code paths for singleton auto-connects that trigger Outlook without a follow-up disconnect.
- Treat missing `finally` cleanup, early returns, and exception paths as leaks until proven otherwise.
- If a change introduces an Outlook COM dependency, call out in the completion notes whether cleanup was verified, and note any remaining Windows-only validation required.