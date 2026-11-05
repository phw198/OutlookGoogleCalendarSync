---
name: ogcs-code-review
description: Review committed code changes on a branch for correctness, regressions, and code quality.
user-invocable: true
disable-model-invocation: true
argument-hint: "BRANCH=<branch-or-SHA> ISSUE=<issue-number> BASE_REF=<branch|tag|ref|SHA>"
---
Always load and follow the platform-agnostic workflow in `.ai/skills/ogcs-code-review/SKILL.md`.

Pass `${input:args}` to the source file's **Inputs** section and follow its parsing and validation rules exactly.