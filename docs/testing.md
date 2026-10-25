# Behaviour Reference

This document records the calendar-sync behaviour users can expect.

## Google To Outlook Recurring Events

- When a timed recurring event crosses midnight in its Outlook timezone, its series begins on the date of the converted local start time, even if that is the previous or next calendar day.
- A timed event's start and end can use different timezones; each instant is converted using the timezone rules at the event date, including daylight-saving transitions.
- The series' final date follows the recurrence rule, not the end date of its first occurrence.

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