# Private Backlog

## DLL Management Migration

**Status:** Planned
**Origin:** Follow-up from `v3/development` and issue `#2392` DLL-signing work
**Visibility:** Internal project backlog; not a public GitHub issue

### Goal
Gradually replace manually copied standard third-party DLL references from `lib` with assemblies restored through `packages.config` and referenced from `..\\packages\\<package>.<version>\\...`.

### Scope
- Migrate one standard third-party dependency at a time.
- Keep `lib` for bespoke, patched, generated, configuration-specific, and explicitly selected Outlook interop assemblies.
- Resolve NuGet version conflicts explicitly rather than preserving copied DLLs to hide them.
- Leave compile-only Windows SDK `.winmd` metadata references non-copying with `<Private>False</Private>`.

### Release and validation requirements
- Validate package restore before removing each corresponding `lib` reference or copy.
- Validate clean Debug and Release builds, runtime output, NuGet/Squirrel packaging, and the portable ZIP for each migration.
- Sign any unsigned DLLs in the Release output with Authenticode before creating the nupkg, Squirrel installer, or portable ZIP.
- Fail the release packaging step if any shipped DLL does not have a valid Authenticode signature.
- Do not modify dependency files in `src/OutlookGoogleCalendarSync/lib` or the restored `src/packages` cache as part of signing; sign only the Release build output.
