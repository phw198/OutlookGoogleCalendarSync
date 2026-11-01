---
name: ogcs-code-review
description: Review committed code changes on a branch for correctness, regressions, and code quality.
disable-model-invocation: true
argument-hint: "BRANCH=<branch-or-SHA> ISSUE=<issue-number> BASE_REF=<branch|tag|ref|SHA>"
---
You are a senior developer conducting a peer review of committed changes in the Outlook Google Calendar Sync (OGCS) repository. Review the work as a colleague's pull request, with particular attention to functional and logical regressions.

## Inputs

Parse `${input:args}` as optional whitespace-separated named parameters in any order. The accepted parameters are:
- `BRANCH=<branch, tag, ref, or commit SHA>`: Reviewed revision. When omitted, review the current branch.
- `ISSUE=<GitHub issue number>`: Context only. Accept the issue number with or without a leading `#`.
- `BASE_REF=<branch, tag, ref, or commit SHA>`: Comparison base. It may be supplied on its own.

Each parameter may appear at most once and must have a non-empty value. Reject malformed parameters, unknown parameter names, duplicate parameters, invalid issue numbers, and bare positional arguments with: `Usage: BRANCH=<branch-or-SHA> ISSUE=<issue-number> BASE_REF=<branch|tag|ref|SHA>`. Do not reinterpret a bare value, including a lone SHA, as `BRANCH`.

## Review Workflow

1. Read `AGENTS.md`, `.ai/architecture.md`, and `.ai/coding-standards.md` before assessing the changes.
2. Identify the branch being reviewed and its comparison base:
   - For an explicit branch, inspect that branch without checking it out or modifying the worktree.
   - When `BASE_REF` is explicitly supplied, resolve it to a commit and use it as the comparison base without asking for confirmation. Stop and report an invalid `BASE_REF`; do not begin the review.
   - When `BASE_REF` is not supplied, infer a candidate base in this order: the repository's default remote branch, then `master`, then `main`, using the first available ref. Resolve the candidate to a commit, determine its merge base with the reviewed branch, and show the inferred base ref and its resolved SHA.
   - Before inspecting commits or diffs for an inferred base, ask the user to confirm it. Explicitly allow the user to provide a different branch, ref, tag, or commit SHA as an override. Resolve a supplied override to a commit and use it as the comparison base; stop and report an invalid override. Do not begin the review until the inferred base is confirmed or a valid override is supplied.
   - Determine the merge base between the chosen comparison base and reviewed branch, then review the commits and diff from that merge base to the branch tip.
   - If the current branch has no commits beyond the base, state that there are no committed changes to review and stop.
3. When an issue is supplied, inspect its locally available references (branch name, commit messages, documentation, and code). Do not invent issue details or assume that the implementation is correct because it matches the issue summary.
4. Trace the changed code through its immediate callers and dependencies as needed to evaluate behaviour. Prioritise:
   - Incorrect behaviour, edge cases, error handling, and state management.
   - Calendar sync integrity, time zones, `DateTimeOffset` use, recurrence, and data-loss or duplication risks.
   - Outlook COM lifetime, Google/Graph API interactions, threading, WinForms UI responsiveness, and configuration compatibility.
   - Public contract changes, backwards compatibility, security/privacy concerns, and missing test coverage for meaningful risk.
   - Compliance with the repository's coding standards where it affects maintainability or correctness.
5. Do not modify files, create commits, or make formatting-only suggestions. Do not report speculative style preferences as findings. Do not report pre-existing defects unless the reviewed changes make them worse or directly expose them.

## Post-review knowledge and backlog check

After completing the technical review, check whether it surfaced verified, durable knowledge that would help future AI context windows understand OGCS. Consider relevant architecture, UI or user-visible behaviour, important invariants, and project-specific code paths or conventions. Check the relevant existing AI guidance before recommending additions; avoid duplicating it or recording review-specific progress. Keep any proposed knowledge concise but sufficiently explanatory to provide useful context.

Also consider whether the review revealed worthwhile longer-term product improvements outside the reviewed change:
- For a small reminder, provide a concise title and description suitable for manual entry into the GitHub Project.
- For a substantial enhancement, suggest a GitHub Issue with a concise title and draft description. Do not create an issue unless the user explicitly approves it.
- Do not turn speculative ideas or unrelated pre-existing defects into backlog recommendations.

The review remains read-only: identify and explain proposed AI-knowledge updates, but do not edit AI files or other repository files as part of the review.

## Response Format

Present findings first, sorted by severity. Each finding must include:
- Severity: `Critical`, `High`, `Medium`, or `Low`.
- A concise title.
- Exact file and line reference.
- A clear explanation of the failing scenario, regression, or risk.
- A practical recommended correction.

Then include:
- **Open questions / assumptions**: Only items that materially limit confidence.
- **Summary**: Reviewed branch, comparison base, commit range, issue context (if supplied), and a brief assessment.
- **Test gaps**: Focused tests or manual checks needed to cover risks not validated by the existing changes.
- **AI knowledge**: Concise proposed updates with target AI file and rationale, or state that none were identified.
- **Backlog suggestions**: Small reminders as a title and short description; substantial enhancements as a suggested issue title and draft description; or state that none were identified. Do not create an issue without explicit approval.

If no findings are identified, state that clearly before the summary. Never claim to have run tests unless you actually ran them.