# Skill: Update AI Configuration Files

## Description
This skill allows the AI to request updates to its own configuration files (e.g., instructions, prompts, agent definitions, skill definitions, or repository-level AI guidance) from any chat context. It leverages the underlying `agent-customization` skill to perform the actual file modifications.

## When to Use This Skill
Use this skill when the AI needs to:
- Modify its `.instructions.md` to refine its behavior.
- Update `.prompt.md` files for specific agent roles.
- Adjust `SKILL.md` files to define new capabilities or refine existing ones.
- Amend repository-level AI guidance files like `.ai/architecture.md`, `.ai/coding-standards.md`, or `.ai/developer-agent.md`.
- Incorporate user feedback or new instructions directly into its own operational files.

## How to Use This Skill
To invoke this skill, simply state your intention to update an AI configuration file, mentioning the specific file or type of file you wish to modify, and the content or changes you want to apply. The skill will then engage the `agent-customization` agent to perform the update.

## Examples
- "Update my `.instructions.md` to prioritize user feedback more strongly."
- "Add a new section to `.ai/coding-standards.md` about asynchronous programming."
- "Fix a typo in the `.prompt.md` for the 'release-preparer' agent."
- "Ensure the `.ai/developer-agent.md` file reflects the current VM setup details."

## Underlying Tools
- `runSubagent` with `agentName: 'agent-customization'`
