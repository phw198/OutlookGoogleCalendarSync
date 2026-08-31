# Project Memory: Outlook Google Calendar Sync

## Architecture Overview
- C# WinForms background tray application. See [.ai/architecture.md](.ai/architecture.md) for full architectural mapping.

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

## Active Task
- Refining release automation prompts and agent skills.
- Documented shim discovery findings in `.ai/developer-agent.md`.
- Added optional `BASE_REF` handling to `/ogcs-code-review`: explicit bases resolve without confirmation, while inferred bases require confirmation or a valid override before review.
- Updated `/ogcs-code-review` to accept validated named `BRANCH`, `ISSUE`, and `BASE_REF` inputs in any order; its explicit and inferred comparison-base behaviour remains unchanged.
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
