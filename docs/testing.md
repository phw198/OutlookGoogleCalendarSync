# Behaviour Reference

This document records the calendar-sync behaviour users can expect.

## Google To Outlook Category Sync

Category and colour synchronization applies to Outlook events, including recurring events.

- When the Google colour maps to the first Outlook category already on an event, synchronization leaves the event unchanged.
- When the mapped category differs and the option to use a single Outlook category is enabled, synchronization replaces every existing Outlook category with the mapped category.
- When multiple Outlook categories are allowed, synchronization adds the mapped category while retaining categories users have added themselves.
- A configured category or colour override replaces the Google mapping. It is applied to new and existing Outlook events unless synchronization is set to update newly created events only.
- A category without an Outlook colour remains a valid category and keeps its configured name.
- Selecting `<No category assigned>` as either a mapping or an override removes categories. With one category per event enabled, all categories are removed. When multiple categories are allowed, user-added categories remain.

## Outlook Category Filtering

The `<No category assigned>` category can be used in either category filter mode.

- In exclude mode, events with no category are excluded and events with categories are retained.
- In include mode, events with no category are retained and events with categories are excluded.

## Exception Failure Logging

When an error is explicitly marked to log as a failure, the failure status is
applied to the aggregate exception and all of its inner exceptions, including
nested aggregate exceptions.

## Classic Outlook Copy Protection

When OGCS saves Google IDs to an Outlook appointment, it fills missing source EntryIDs for the appointment's existing calendar key sets without replacing populated markers. This happens during bidirectional sync and during one-way sync when the Outlook item already contains Google metadata, helping later syncs distinguish future copies from their originals.

Outlook-to-Google-only sync leaves previously untagged Outlook appointments free of OGCS Google and source EntryID properties. Existing copied appointments that already carry OGCS metadata continue to have those IDs refreshed when a separate Google event is created. If the profile later switches to bidirectional sync, the normal sync matches existing Google events to Outlook appointments using the Outlook IDs stored in Google, then adds missing Outlook metadata to successfully matched items. This repair is automatic and progressive within the configured sync window; items outside the window are repaired when a later sync includes them.

The full-folder, three-attempt backfill is limited to bidirectional sync.

- The full-folder backfill makes at most three attempts for a profile and calendar pair. If it cannot complete, OGCS reports the problem and stops automatic rescans.
- To retry after resolving the issue, close OGCS, back up `settings.xml`, and clear `SourceEntryIDBackfillProfile` only in the affected calendar profile. The settings file location depends on the installation.
- If a marker is missing on an appointment that was already copied, its original EntryID cannot be recovered; OGCS records the current EntryID when filling that marker.