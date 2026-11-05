---
name: ogcs-release-rollout
description: 'Use when gradually rolling out an OGCS Squirrel release, checking rollout percentages, or updating both full and delta NUPKG entries in the GitHub RELEASES asset through the 10%, 25%, 50%, 75%, and 100% stages.'
user-invocable: true
disable-model-invocation: true
argument-hint: "VERSION=<release-tag> (optional)"
---

# OGCS Squirrel Release Rollout

Use this workflow to inspect and advance the staged rollout recorded in a GitHub release's `RELEASES` asset. The workflow changes a published release asset, so never upload a replacement without showing the exact proposed change and getting explicit user approval.

## Inputs

- Optional `VERSION=<release-tag>`: target that exact GitHub release. When omitted, select the most recently published, non-draft release for `phw198/OutlookGoogleCalendarSync`, including prereleases.
- If an explicit release identifier is supplied in the user's message, use it instead of finding the newest release. If identifiers conflict or the requested release cannot be resolved exactly, stop and clarify.

## Workflow

1. **Resolve the target release.**
   - Use GitHub release/asset tools if available and capable of inspecting and replacing release assets. Otherwise use the authenticated GitHub CLI (`gh`).
   - Confirm the target repository is `phw198/OutlookGoogleCalendarSync`; do not silently operate on a fork or a different remote.
   - When no version was supplied, inspect the releases ordered by `published_at` and select the newest non-draft release, including prereleases. Do not assume GitHub's "latest" release endpoint includes prereleases.
   - Report the selected release's tag, title, publication date, and prerelease status. If the selection is ambiguous or the release is a draft, stop and ask.
2. **Inspect the asset without changing it.**
   - Locate the exact asset named `RELEASES` on that release. If it is missing, duplicated, or cannot be downloaded, stop and report the problem.
   - Download it to a temporary working directory and retain an untouched local copy as a backup.
   - Inspect the raw bytes and identify exactly one full NUPKG line and one delta NUPKG line for the target version. Do not assume the final line is the only one to update. Use the existing asset's exact format and line endings; never infer or rewrite either package's version, hash, or size.
   - The rollout stages are `10`, `25`, `50`, `75`, and `100` percent. The full and delta lines must always have matching rollout percentages. A specified suffix less than 100 indicates an in-progress rollout; no percentage suffix on either line means both are complete. If the lines disagree, report both values and ask which target stage to apply to both. A literal `100` suffix is noncanonical; stop and clarify rather than guessing.
3. **Confirm the current stage and ask for the next stage.**
   - Tell the user which release was inspected and state the current percentage on both package lines.
   - If the lines match, ask the user to confirm the detected current stage before proceeding. If they disagree, stop without editing.
   - If the lines disagree, report each line's current value and ask which target stage to apply to both; do not infer which line is authoritative.
   - Ask which higher supported stage to use next. Recommend the next sequential band (10 -> 25 -> 50 -> 75 -> 100), but accept another higher listed band if the user explicitly chooses it. Do not accept a lower, equal, or unsupported percentage.
   - Choosing 100% means the rollout is complete and must be represented by removing the percentage suffix from both lines, not by writing `100` or `100%`.
4. **Prepare and review the candidate.**
   - Apply the selected stage to both the full and delta lines. For a selected stage of 100%, remove both percentage suffixes entirely. Preserve all other bytes, entries, spacing, versions, hashes, sizes, and line endings.
   - Show a focused before/after comparison for both lines. Confirm their resulting percentages match and no other entry or byte was modified.
   - Ask for explicit approval to replace the `RELEASES` asset on the named release. A choice of percentage alone is not approval to upload.
5. **Replace and verify the asset.**
   - If the user approves, replace the asset using the authenticated GitHub release tools or `gh release upload <tag> <candidate-path> --repo phw198/OutlookGoogleCalendarSync --clobber`, which handles replacement under the exact name `RELEASES`. Keep the downloaded backup until verification succeeds.
   - If neither replacement method is available and the GitHub web UI is used, keep a local backup, then in one edit remove the existing `RELEASES` asset, upload the candidate whose local basename is exactly `RELEASES`, and save once. The browser uploads using the selected local filename immediately; do not upload a differently named temporary file and try to rename it in the filename textbox, because that can trigger a duplicate-name validation error. If the single-save replacement fails, stop and report the current remote asset state rather than retrying with a different sequence without approval.
   - Download the final asset again from GitHub and compare it byte-for-byte with the approved candidate. Verify that exactly one asset named `RELEASES` remains and both package lines have matching percentages. Report success only if all checks pass.
   - If upload or verification fails, report the exact failure, preserve the local backup and candidate, and do not claim that the release was updated. Do not retry a destructive replacement without explaining the failure and obtaining approval.

## GitHub CLI fallback

If GitHub release tools are unavailable, use `gh release list` and `gh release view` to identify and inspect the release, then `gh release download` to fetch its asset. After explicit approval, use `gh release upload <tag> <candidate-path> --repo phw198/OutlookGoogleCalendarSync --clobber` to replace the existing asset. If the authenticated CLI is unavailable, follow the single-edit web UI procedure above: remove the existing asset, attach the candidate with the local filename `RELEASES`, and save once. Never upload a differently named temporary candidate and then try to rename it in the release form. Verify by downloading the asset again, confirming exactly one `RELEASES` asset exists, and comparing bytes. If neither route is available, stop and report the inspected state; do not claim to have edited the remote asset.
