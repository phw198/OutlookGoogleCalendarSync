# Project Memory: Outlook Google Calendar Sync

## Architecture Overview
- C# WinForms background tray application. See [.ai/architecture.md](.ai/architecture.md) for full architectural mapping.

## Recent Milestone
- Established the Universal AI Root (`.ai/`) structure with generic AI agent developer instructions, project architecture overview (using Mermaid diagrams), specific C# and WinForms coding standards (such as OTBS/K&R style brace placement and `Analyse()` logging), and reusable prompt templates.
- Reverted to a single global NotifyIcon wrapper to handle Windows notification routing cleanly.

## Active Task
- Testing native VS Code Chat agent execution and memory file syncing.

## Testing Guidance
- Before creating or amending tests, inspect nearby and related existing tests for conflicting expectations or duplicate coverage.
- Raise any conflict in test logic or ambiguity in the expected behavior before encoding it in a new test.
- Keep [docs/testing.md](docs/testing.md) current as a user-focused reference for application behavior covered by automated tests. Do not include implementation, source-code, test-framework, or test-location details.

## C# Naming
- Use PascalCase only for public members and types. Private and internal members must begin with a lower-case character.
- Use `using Ogcs = OutlookGoogleCalendarSync;` as the only import for the OGCS project namespace. Reference OGCS types with their fully qualified namespace from `Ogcs`.
- For a function whose parameter list spans multiple lines, append `//` to the final parameter line and put the opening brace on the following line.

## File Format
- Use Windows CRLF line endings for all text files, including source, project, test, and documentation files.
