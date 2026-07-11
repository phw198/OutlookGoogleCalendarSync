# Coding Standards - Outlook Google Calendar Sync (OGCS)

Please adhere to these coding standards when developing or modifying code for OGCS.

## C# Coding Conventions
- **Naming Conventions**:
  - Use PascalCase for class names, methods, and public properties (e.g., `NotificationTray`, `SyncEngine`).
  - Use camelCase for local variables, method parameters, and private methods (e.g., `eventItem`, `syncDirection`, `notificationClicked`).
  - Prefix private class fields with an underscore (e.g., `_notifyIcon`, `_settings`).
- **Explicit Types**:
  - Prefer explicit type declarations over `var` for local variables. Only use `var` when the type is obvious from the right-hand side (e.g., `var x = new Exception("...")`) or the type is extremely verbose; explicit types improve readability in this codebase.
- **Line Endings**:
  - All repository source code and text files must use CRLF line endings. Ensure `.cs`, `.md`, `.txt`, and other project files are saved with CRLF.
- **Bracing and Spacing**:
  - Use K&R / OTBS brace style: **Open brace is on the same line** as the statement/declaration (e.g., `if (condition) {`).
  - Indent with spaces (4 spaces standard).

## WinForms & UI Responsiveness
- **Never Block the UI Thread**: Long-running operations like API requests, filesystem tasks, and calendar parsing must run asynchronously or in background threads (e.g., using `Sync/AbortableBackgroundWorker.cs` or async tasks).
- **Control Access from Non-UI Threads**: Use `Control.Invoke` or `Control.BeginInvoke` when updating WinForms controls or triggering UI notifications from a background thread.
- **Resource Cleanup**: Always explicitly dispose of `IDisposable` resources (especially Outlook COM objects, streams, or graphics brushes/icons) using `using` blocks or explicit `.Dispose()` calls.

## Logging & Telemetry
- **Error Logging**: Always use the log extension function `Analyse()` when logging exceptions/errors, passing in a contextual description of what caused the error (which is logged at `WARN` level). Avoid bare `ex.Analyse()` calls when a message is available; include a short, specific reason such as `ex.Analyse("Unable to show toast notification.")`.
- Log meaningful info, warnings, and errors with details of what was happening at that moment. Do not output raw passwords or sensitive auth tokens to log files.

## Build and Testing
- **Build Scripts**:
  - Release packages and NuGet builds are managed via the root-level scripts such as `nuget-build.bat`.
- **Testing**:
  - Always verify code compiles and runs locally before committing.
  - Test all sync variations (Outlook -> Google, Google -> Outlook, Bidirectional) when changing the underlying `Engine.cs` or provider mechanics.

## DateTime & TimeZone Handling
- **Prefer `DateTimeOffset`**: Use `DateTimeOffset` for all timestamps, especially when interacting with external APIs (Google, Graph).
- **Avoid Implicit Casts**: Do not implicitly cast `DateTime` to `DateTimeOffset`. This uses the local system timezone and can cause equality comparisons to fail against UTC values from APIs.
- **Clock-Time Equality**: For comparing "wall clock" time (same hour/minute on the dial), explicitly use `.DateTime` from the `DateTimeOffset` (e.g., `dto.DateTime == dt`).
- **Date-Only Equality**: For comparing the calendar date only, use the `.Date` property (e.g., `dto.Date == dt.Date`).
- **Instant Equality**: For comparing the exact universal instant, compare `.UtcDateTime` or direct `DateTimeOffset` if offsets are known.
- **Safe Helpers**: Always use the `SafeDateTimeOffset()` extension methods. The older `SafeDateTime()` is deprecated and must not be used.
