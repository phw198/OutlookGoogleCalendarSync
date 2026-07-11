# Project Memory: Outlook Google Calendar Sync

## Architecture Overview
- C# WinForms background tray application. See [.ai/architecture.md](.ai/architecture.md) for full architectural mapping.

## Recent Milestone
- Established the Universal AI Root (`.ai/`) structure with generic AI agent developer instructions, project architecture overview (using Mermaid diagrams), specific C# and WinForms coding standards (such as OTBS/K&R style brace placement and `Analyse()` logging), and reusable prompt templates.
- Reverted to a single global NotifyIcon wrapper to handle Windows notification routing cleanly.
- Bumped version variables and references across `nuget-build.bat`, `docs/latest_zip_release.md`, and `src/OutlookGoogleCalendarSync/OutlookGoogleCalendarSync.nuspec` from 3.0.2 to 3.0.3.
- Populated the `nuspec` release notes with missing GitHub issue #1870 (Modern Toast notifications) by identifying merge commits from the active branch since the last master branch commit.
- Initialised the custom agent `.ai/agents/release-preparer.md` to automate future release script updates, documentation updates, and git commit-based release note generation.
- Implemented an agnostic agent architecture (Pattern 1): rich instructions reside in `.ai/agents/` while `.github/agents/` contains thin pointers for VS Code UI integration.
- Cleaned up redundant `.ai/.github/` directory.

## Active Task
- Refining release automation prompts and agent skills.
