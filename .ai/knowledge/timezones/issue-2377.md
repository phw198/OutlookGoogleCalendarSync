---
type: Case Study
title: "Issue #2377: Windows timezone arrives before its IANA mapping"
description: "A new Windows timezone can be available to Outlook while OGCS lacks the Windows-to-IANA mapping needed to sync it."
resource: https://github.com/phw198/OutlookGoogleCalendarSync/issues/2377
tags:
  - timezone
  - Windows
  - IANA
  - tzdb
  - CLDR
sources:
  - resource: https://github.com/phw198/OutlookGoogleCalendarSync/issues/2377
    title: "New time zone British Columbia Standard Time won't sync to Google"
---

# Summary

Issue #2377 reported that the new `British Columbia Standard Time` Windows zone could be selected after a Windows update, but Outlook-to-Google sync could not map it. The underlying rules and the identifier mapping were not delivered as one atomic update. The issue discussion also noted that recurring events had inconsistent results when some had been created with a fixed offset and others with a named timezone.

## Why a Windows timezone can be unmappable

Several independently maintained data sources participate:

1. A Windows update adds the Windows timezone ID and its rules.
2. IANA publishes tzdb rules for the region.
3. Unicode CLDR's `windowsZones.xml` maps the Windows ID to an IANA ID.
4. NodaTime packages tzdb and CLDR mapping data for applications to consume.

The presence of a zone in Windows therefore does not establish that an application's timezone library can map that zone to IANA. The mapping dataset can lag the Windows registry and tzdb rule updates.

In the September 2026 issue discussion, CLDR had not yet added the new British Columbia ID and a NodaTime package containing the later CLDR mapping was still pending. Those release versions and dates are historical; check the current CLDR and NodaTime data before applying this explanation to a later report.

## Where to check for updates

Check each layer independently; an update in one does not mean it has propagated to the next.

1. **Windows timezone IDs and rules**
   - [Microsoft's Windows time zone list](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-time-zones?view=windows-11) documents available Windows zones and the `tzutil /l` command for the machine's installed list.
   - [Windows release health](https://learn.microsoft.com/en-us/windows/release-health/) tracks OS update information. For a specific timezone change, follow the relevant Windows KB or [Microsoft Daylight Saving Time guidance](https://techcommunity.microsoft.com/blog/dstblog/interim-guidance-for-british-columbia-time-zone-changes-2026/4544988/) where available.
2. **IANA tzdb rules**
   - [IANA Time Zone Database](https://www.iana.org/time-zones) reports the current release and changes.
   - [IANA release files](https://data.iana.org/time-zones/releases/) provide versioned source archives.
3. **Unicode CLDR Windows-to-IANA mapping**
   - [CLDR release/download page](https://cldr.unicode.org/index/downloads) identifies stable releases.
   - [`windowsZones.xml` on CLDR main](https://github.com/unicode-org/cldr/blob/main/common/supplemental/windowsZones.xml) shows the current development mapping. Compare the corresponding file in the stable release being used; `main` may include changes not yet released or consumed by NodaTime.
4. **NodaTime data consumed by OGCS**
   - [NodaTime's tzdb data directory](https://nodatime.org/tzdb/) explains its downloadable NZD files.
   - [`latest.txt`](https://nodatime.org/tzdb/latest.txt) points to the latest NZD file and is the endpoint OGCS checks for data updates. Verify which data version OGCS actually loaded rather than assuming the latest download has already been applied.

## OGCS-specific path

- OGCS uses NodaTime for timezone identifiers and mappings. `TimezoneDB` loads `tzdb.nzd` beside the application, falls back to NodaTime's built-in source if that file cannot be read, and checks for newer timezone data separately from an application release.
- Outlook and Graph conversion paths consult NodaTime's Windows-to-IANA mapping. A missing mapping can leave OGCS unable to translate a Windows timezone into a Google-compatible IANA timezone.
- Settings > Outlook > Timezone Mapping provides a user-defined Windows/system timezone mapping when automatic mapping cannot resolve an ID. This can be a temporary workaround, not a replacement for the missing authoritative mapping.

## Investigation guidance

1. Record the exact Windows timezone ID, not just its display name or current UTC offset.
2. Check which timezone source OGCS loaded and its tzdb version. Do not assume that updating Windows also updated OGCS's `tzdb.nzd` or NodaTime's built-in data.
3. Check whether the loaded CLDR mapping contains the Windows ID and which IANA ID it maps to. Distinguish absent mapping data from an absent or incorrect timezone rule.
4. Verify the proposed IANA/system-zone mapping against the affected dates, including historical and future transitions and recurring-event behavior.
5. If suggesting a custom mapping, choose a zone with equivalent rules over the relevant dates. A zone with the same current UTC offset may observe different daylight-saving or future rules. Remove the workaround and retest after updated mapping data is available.

Do not assume that any one timezone layer is the sole cause until the effective IDs, data versions, and event behavior have been checked. Consult the [issue discussion](https://github.com/phw198/OutlookGoogleCalendarSync/issues/2377) for the original report and the temporary workaround discussion.
