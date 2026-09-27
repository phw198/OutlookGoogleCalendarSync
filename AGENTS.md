# Project Memory: Outlook Google Calendar Sync

## Architecture Overview
- C# WinForms background tray application. See [.ai/architecture.md](.ai/architecture.md) for full architectural mapping.

## Local WIP Memory
- The active work-in-progress note is branch-scoped: resolve the current branch name, then read the matching note under [.ai/work-in-progress](.ai/work-in-progress) using that branch name as the category. For example, `dev/ai` maps to [.ai/work-in-progress/dev/ai.md](.ai/work-in-progress/dev/ai.md).
- The shared index at [.ai/work-in-progress.md](.ai/work-in-progress.md) helps discover the active branch note, but each branch keeps its own isolated task log so parallel development does not overwrite the same notes.

## Recent Milestone
- Established the Universal AI Root (`.ai/`) structure with generic AI agent developer instructions, project architecture overview (using Mermaid diagrams), specific C# and WinForms coding standards (such as OTBS/K&R style brace placement and `Analyse()` logging), and reusable prompt templates.
- Reverted to a single global NotifyIcon wrapper to handle Windows notification routing cleanly.
- Bumped version variables and references across `nuget-build.bat`, `docs/latest_zip_release.md`, and `src/OutlookGoogleCalendarSync/OutlookGoogleCalendarSync.nuspec` from 3.0.2 to 3.0.3.
- Populated the `nuspec` release notes with missing GitHub issue #1870 (Modern Toast notifications) by identifying merge commits from the active branch since the last master branch commit.
- Initialised the custom agent `.ai/agents/ogcs-release-preparer.md` to automate future release script updates, documentation updates, and git commit-based release note generation.
- Implemented an agnostic agent architecture (Pattern 1): rich instructions reside in `.ai/` while `.github/` contains thin pointers (shims) for VS Code UI integration.
- Successfully established the shim pattern for both Custom Agents (`.github/agents/*.agent.md`) and Custom Skills (`.github/skills/<name>/SKILL.md`), enabling their discovery in VS Code chat.
- Cleaned up redundant `.ai/.github/` directory.
- Standardised selectable workspace customizations with the `ogcs-` prefix: the `ogcs-release-preparer` agent and `/ogcs-code-review` prompt.
- Replaced the selectable `sync-dev` agent with always-applied `ogcs-global-guidance` project instructions.
- Added a mandatory COM lifetime review to the project guidance: any Outlook client change must check for new COM object leaks, singleton auto-connects, and matching `Disconnect`/`ReleaseObject` cleanup before completion.

## Release Build Note
- Excluded the test project from Release solution builds so the app can keep the required embedded Outlook interop settings in Release without tripping the `CS1769` generic interop-type boundary error when the tests are compiled in the same solution.

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
