---
name: ogcs-build
description: "Build the OGCS solution or run its tests when explicitly requested, with accurate .NET Framework platform caveats."
user-invocable: true
disable-model-invocation: true
---

# OGCS Build

Use this skill only when the user invokes `/ogcs-build` or explicitly asks to build or test OGCS. Run only the requested operation; do not run both a build and tests unless asked.

## Build

- The application is a legacy WinForms solution targeting .NET Framework 4.7.2. On Windows, use a Visual Studio Developer Command Prompt with the .NET Framework 4.7.2 targeting pack:

  ```powershell
  msbuild src\OutlookGoogleCalendarSync.sln /t:Build /p:Configuration=Debug
  ```

- On Linux, the legacy application build is blocked in this environment by the missing `Microsoft.CSharp.Targets`. If the user requests a Linux build, run the requested command once and report that blocker; do not cycle through alternative build tools.
- Do not use `nuget-build.bat` for an ordinary build; it performs release signing and packaging.

## Tests

- Run the requested test project or filter. The full test project command is:

  ```sh
  dotnet test src/OutlookGoogleCalendarSync.Tests/OutlookGoogleCalendarSync.Tests.csproj --no-restore --verbosity normal
  ```

- Require a test-discovery/execution summary before reporting tests as passed. In the current Linux environment, `dotnet test` may report a successful build without discovering or executing MSTest tests; report that as build-only validation.
