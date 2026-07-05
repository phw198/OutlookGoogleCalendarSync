# Outlook Google Calendar Sync Developer Agent Instructions

This file contains the generic instructions and guidelines for any AI developer agent working on the Outlook Google Calendar Sync (OGCS) project.

## Core Directives
When assisting with code modifications, debugging, or documentation:
- **Be Minimalist**: Implement precisely what is requested with the fewest lines of code possible while meeting all requirements.
- **Consult Architecture**: Before touching Outlook or Google Calendar sync code, review [.ai/architecture.md](architecture.md).
- **Adhere to Coding Standards**: Refer to [.ai/coding-standards.md](coding-standards.md) for C# coding style, WinForms/thread safety, and error logging standards.
- **Maintain UI Responsiveness**: Since OGCS is a WinForms tray application, always avoid blocking the UI thread. Use background tasks or async patterns carefully.

## Key Files to Monitor
- **Memory Log**: Always refer to the root-level [AGENTS.md](../AGENTS.md) to understand current progress, milestones, and active tasks, and keep it updated as changes are completed.
- **Rules Config**: Ensure actions adhere to the root-level [.clinerules](../.clinerules).
