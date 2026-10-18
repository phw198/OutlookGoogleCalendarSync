# Behaviour Reference

This document records the calendar-sync behaviour users can expect.

## Reminder Sync

The Reminders option has three states:

- Clicking or activating the option from the keyboard cycles through
  **Unchecked**, **No Reminders**, and **Reminders** in that order.
- **Reminders** syncs reminder times according to the selected Google and
  Outlook default-reminder and do-not-disturb settings.
- **Unchecked** leaves reminder settings untouched during synchronization. New
  events inherit the default configured by the destination calendar.
- **No Reminders** prevents reminders from being synchronized and removes
  explicit reminders and calendar defaults from synchronized events.

Existing profiles from v3.0.3 and earlier are migrated to **No Reminders** to
preserve their previous behavior. New profiles start unchecked so the calendar
platform defaults apply.

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