# Coding Standards - Outlook Google Calendar Sync (OGCS)

Please adhere to these coding standards when developing or modifying code for OGCS.

## C# Coding Conventions
- **Naming Conventions**:
  - Use PascalCase for class names, methods, and public properties (e.g., `NotificationTray`, `SyncEngine`).
  - Use camelCase for local variables and method parameters (e.g., `eventItem`, `syncDirection`).
  - Prefix private class fields with an underscore (e.g., `_notifyIcon`, `_settings`).
- **Bracing and Spacing**:
  - Use K&R / OTBS brace style: **Open brace is on the same line** as the statement/declaration (e.g., `if (condition) {`).
  - Indent with spaces (4 spaces standard).

## WinForms & UI Responsiveness
- **Never Block the UI Thread**: Long-running operations like API requests, filesystem tasks, and calendar parsing must run asynchronously or in background threads (e.g., using `Sync/AbortableBackgroundWorker.cs` or async tasks).
- **Control Access from Non-UI Threads**: Use `Control.Invoke` or `Control.BeginInvoke` when updating WinForms controls or triggering UI notifications from a background thread.
- **Resource Cleanup**: Always explicitly dispose of `IDisposable` resources (especially Outlook COM objects, streams, or graphics brushes/icons) using `using` blocks or explicit `.Dispose()` calls.

## Logging & Telemetry
- **Error Logging**: Always use the log extension function `Analyse()` when logging exceptions/errors, passing in a contextual description of what caused the error (which is logged at `WARN` level).
- Log meaningful info, warnings, and errors with details of what was happening at that moment. Do not output raw passwords or sensitive auth tokens to log files.

## Build and Testing
- **Build Scripts**:
  - Release packages and NuGet builds are managed via the root-level scripts such as `nuget-build.bat`.
- **Testing**:
  - Always verify code compiles and runs locally before committing.
  - Test all sync variations (Outlook -> Google, Google -> Outlook, Bidirectional) when changing the underlying `Engine.cs` or provider mechanics.
