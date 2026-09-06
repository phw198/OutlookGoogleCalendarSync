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