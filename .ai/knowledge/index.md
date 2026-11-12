# OGCS Knowledge Base

Curated, issue- and workflow-derived knowledge that is likely to help with future investigations and repository operations. Entries preserve useful context and lessons, not a substitute for checking current code, dependencies, or external data.

The knowledge notes use the [Open Knowledge Format (OKF)](https://github.com/GoogleCloudPlatform/open-knowledge-format) as a lightweight template: Markdown with concise YAML metadata, source provenance, and links. This repository does not currently require OKF-specific tooling or optional metadata.

## Topics

- [Git branch cleanup](git/remote-tracking-branch-cleanup.md) — removing local remote-tracking refs, pruning stale refs, and avoiding remote branch deletion.
- [Timezone data and mapping](timezones/issue-2377.md) — Windows timezone IDs, IANA tzdb rules, CLDR mappings, and the OGCS/NodaTime update path.
