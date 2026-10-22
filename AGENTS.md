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
- Clarified the review prompt’s issue-resolution rule: explicit `ISSUE=` values override any issue number inferred from the branch name, while the branch convention still provides a fallback when no explicit issue is given.
- Enforced issue isolation in the project guidance: each issue branch must stay scoped to that issue only and must not absorb unrelated work from other sessions or branches unless the user explicitly authorizes it.
- Confirmed the sync-window semantics are date-only boundaries intended to be treated as UTC midnight, and documented the clearer explicit UTC-midnight pattern to avoid local-time ambiguity.

## Release Build Note
- Excluded the test project from Release solution builds so the app can keep the required embedded Outlook interop settings in Release without tripping the `CS1769` generic interop-type boundary error when the tests are compiled in the same solution.

## Recent Progress
- Fixed the null-category guard in [src/OutlookGoogleCalendarSync/Outlook/OutlookCalendar.cs](src/OutlookGoogleCalendarSync/Outlook/OutlookCalendar.cs) so `updateCategoriesForGoogleColour()` no longer dereferences a null Outlook category cache during unit tests, while still preserving non-OGCS categories and stripping OGCS entries as intended.
- Verified the focused Outlook category sync tests now pass with exit code 0.
- Fixed the Windows toast activation gap in [src/OutlookGoogleCalendarSync/Notifications/ModernToastNotificationHandler.cs](src/OutlookGoogleCalendarSync/Notifications/ModernToastNotificationHandler.cs): the app was only listening to the live toast instance, while the Action Center item uses a COM/AUMID activation callback that survives the original toast disappearing. Registered the app-level activator and routed both activation paths through the same callback so flyout clicks restore focus the same way as the initial toast.
- Removed the obsolete multi-process activation guard: direct `ToastNotificationManagerCompat.OnActivated` subscriptions are process-local, so each OGCS instance can handle clicks for notifications it created while `activationCallbacksRegistered` still prevents duplicate listeners within that process.
- Made toast action extraction order-independent: the activation argument parser now locates the `action` key rather than assuming it is the first serialized argument.
- Added per-process toast routing in [src/OutlookGoogleCalendarSync/Notifications/ModernToastNotificationHandler.cs](src/OutlookGoogleCalendarSync/Notifications/ModernToastNotificationHandler.cs): each process embeds a random instance token in its toast, listens on a matching named pipe, and forwards activations received by another process to the owning instance. Notifications from a restarted process have no live pipe and are ignored.
- Added an early startup activation gate: a process launched by a toast forwards the activation to the live owner pipe or exits silently when the owner has gone, preventing orphaned Action Center clicks from leaving a new OGCS instance running.
- Orphaned `OOMsecurity` toast activations now open the security FAQ directly when the creator process is unavailable, while other orphaned actions remain ignored.
- Generalized browser-opening notifications: optional HTTP(S) URLs now travel through both legacy and modern notification handlers; modern toast URLs are base64-encoded and forwarded with the action payload so query-string delimiters are preserved.
- Documented the modern toast lifecycle, per-process routing, and orphaned activation handling in [.ai/architecture.md](.ai/architecture.md).
- Corrected the toast manager event hookup to the installed toolkit contract (`OnActivated` takes a single `ToastNotificationActivatedEventArgsCompat` parameter), matching the actual 7.1.3 delegate shape and clearing the compile-time mismatch. Immediate next step: validate this on a Windows/ .NET Framework 4.7.2 environment because the Linux build here cannot restore the required targeting pack.
- Added graceful modern-toast shutdown cleanup: modern notifications now use process-specific tags, and tray exit unsubscribes activation, cancels the named-pipe listener, and removes only that process's tracked toast history. Legacy notification bubbles remain unchanged. Diagnostics pass on Linux; Windows/.NET Framework validation is still required.
- Updated [src/OutlookGoogleCalendarSync/Sync/Calendar.cs](src/OutlookGoogleCalendarSync/Sync/Calendar.cs) so only user-cancelled `TaskCanceledException` paths produce `SyncResult.UserCancelled`, while preserving OGCS error context output. Added focused exception-helper and transient-network coverage; the test project builds successfully, but the legacy `dotnet test` invocation emits no test execution summary.

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
