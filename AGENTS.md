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
- Added a mandatory COM lifetime review to the project guidance: any Outlook client change must check for new COM object leaks, singleton auto-connects, and matching `Disconnect`/`ReleaseObject` cleanup before completion.

## Active Task
- Updated the alpha release metadata for 3.0.4 across the build script, docs, and package metadata while preserving 3.0.3 as the current released baseline and 3.0.2 as the prior release reference.
- Mined the v3 release branch history to reconcile the 3.0.4 changelog against actual issue branches and add the missing Git-derived entries for the OOO sync enhancement and the SafeDateTime bugfix.
- Refined release automation to keep the new version bump aligned with the project’s existing alpha packaging and ZIP naming conventions, and to validate the changelog against Git history even when the Nuspec header has already been bumped.
- Diagnosed issue #2233 (`Outlook.Calendar.FilterCalendarEntries`, "An item with the same key has already been added."): Outlook's live `[Start]`-sorted `Items` enumeration can re-deliver the same appointment mid-loop if any item's `Start` changes concurrently (background Exchange sync, another session, etc.) - not a genuine duplicate `EntryID`. Replaced the exception-driven diagnostic (full-folder re-query, CSV export, rethrow that aborted the whole calendar filter) with a simple `ExcludedByCategory.ContainsKey()` guard that skips the redundant re-delivery instead of crashing the sync.
- Implemented a registry-based `HasOutlookProfile()` check in `OutlookFactory` so OGCS only probes `Outlook.Application` after confirming a valid Outlook MAPI profile exists, preventing COM activation when Outlook is installed but entirely unconfigured.

## Immediate Next Step
- Validate the targeted fix with the lightest available build or project-level check for the Outlook factory logic, and then confirm whether any broader Outlook profile fallback behavior should be exercised in a Windows-specific environment.

## Progress Update
- Confirmed the UI freeze was caused by a deadlock in the Graph pagination path: the first page of calendars was awaited correctly, but the second page used a blocking Result call while the WinForms UI thread was still active, which only reproduced for accounts with multiple calendar pages.
- Reproduced the issue by forcing Graph page size to 1 and observing the freeze on the second-page fetch in [src/OutlookGoogleCalendarSync/Outlook.Graph/O365Calendar.cs](src/OutlookGoogleCalendarSync/Outlook.Graph/O365Calendar.cs).
- Replaced the blocking next-page fetch with an awaited pagination path to keep the UI responsive; the fix matches the observed reproduction and resolves the freeze for multi-page calendar sets.
- Converted the remaining Graph read/write blockers in [src/OutlookGoogleCalendarSync/Outlook.Graph/O365Calendar.cs](src/OutlookGoogleCalendarSync/Outlook.Graph/O365Calendar.cs), [src/OutlookGoogleCalendarSync/Outlook.Graph/O365Recurrence.cs](src/OutlookGoogleCalendarSync/Outlook.Graph/O365Recurrence.cs), and [src/OutlookGoogleCalendarSync/Outlook.Graph/O365CustomProperty.cs](src/OutlookGoogleCalendarSync/Outlook.Graph/O365CustomProperty.cs) to async await-based paths with compatibility wrappers for the existing sync call sites.
- Restored the worker-thread sync wrapper pattern for the Graph calendar calls: the background sync worker may block on Graph/Outlook calls, but UI-invoked methods must not block; the wrapper approach keeps the async helpers available while preventing UI deadlocks.
- Added explicit project guidance to distinguish UI-thread safety from background-worker blocking, so future Graph/O365 changes are checked for the correct execution context before converting methods to async or sync.
- Left the separate MSAL auth bootstrap and Google API blocking paths alone for now because they are provider-specific and outside the Graph calendar deadlock fix; a fuller async cleanup pass should be done in a follow-up if those flows are later moved onto the same await-based pattern.
- Restored the single blocking `GetCalendarEntry` path in [src/OutlookGoogleCalendarSync/Outlook.Graph/O365Calendar.cs](src/OutlookGoogleCalendarSync/Outlook.Graph/O365Calendar.cs) and removed the stale duplicate method block that introduced the syntax breakage during the Graph pagination cleanup.
- Kept the final pattern aligned to the project rule: background sync worker code may block on Graph/Outlook calls, while any method reached directly from the WinForms UI must remain non-blocking.

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
