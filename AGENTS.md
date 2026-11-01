# Project Memory: Outlook Google Calendar Sync

## Architecture Overview
- C# WinForms background tray application. See [.ai/architecture.md](.ai/architecture.md) for full architectural mapping.

## Local WIP Memory
- The active work-in-progress note is branch-scoped: resolve the current branch name, then read the matching note under [.ai/work-in-progress](.ai/work-in-progress) using that branch name as the category. For example, `dev/ai` maps to [.ai/work-in-progress/dev/ai.md](.ai/work-in-progress/dev/ai.md).
- The shared index at [.ai/work-in-progress.md](.ai/work-in-progress.md) helps discover the active branch note, but each branch keeps its own isolated task log so parallel development does not overwrite the same notes.

## Release Build Note
- Excluded the test project from Release solution builds so the app can keep the required embedded Outlook interop settings in Release without tripping the `CS1769` generic interop-type boundary error when the tests are compiled in the same solution.

## Release Website Documentation Review
- Always compare release changes against the current `gh-pages` User Guide and update relevant pages for user-facing behavior or UX changes. If no update is needed, record the review and rationale in the branch-scoped WIP note and release summary.

## Testing Guidance
- Before creating or amending tests, inspect nearby and related existing tests for conflicting expectations or duplicate coverage.
- Raise any conflict in test logic or ambiguity in the expected behavior before encoding it in a new test.
- Keep [docs/testing.md](docs/testing.md) current as a user-focused reference for application behavior covered by automated tests. Do not include implementation, source-code, test-framework, or test-location details.

## C# Naming
- Use PascalCase for public and internal members and types. Private methods use lower camelCase.
- Use `using Ogcs = OutlookGoogleCalendarSync;` as the only import for the OGCS project namespace. Reference OGCS types with their fully qualified namespace from `Ogcs`.
- For a function whose parameter list spans multiple lines, append `//` to the final parameter line and put the opening brace on the following line.

## File Format
- Use Windows CRLF line endings for all text files, including source, project, test, and documentation files.
