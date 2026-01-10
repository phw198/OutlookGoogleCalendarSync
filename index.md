---
layout: page
---
<div align="center" style="font-size:18px; line-height:18px; font-family:'Architects Daughter', 'Helvetica Neue', Helvetica, Arial, serif;">
  <p>:small_blue_diamond: Sync your Outlook and Google calendars securely, including meeting attendees, reminders, full description and more.</p>
  <p>:small_blue_diamond: Ideal for liberating your corporate Exchange calendar, making it available on any of your devices with access to Google Calendar.</p>
  <p>:small_blue_diamond: No install necessary, works behind web proxies and actively developed.<br/>Get <a href="guide">syncing in minutes</a>.</p>
</div>

{% include carousel.html height="90" unit="px" duration="10" number="1" %}

### :eyes: The New "Outlook for Windows" Application
<div style="background-color: #ff6a001c; border: red; border-style: dashed; border-width: thin; padding: 5px; padding-left: 10px;"><font color="darkred">
   The "<a href="https://www.microsoft.com/en-us/microsoft-365/outlook/outlook-for-windows">New Outlook for Windows</a>" application is not compatible with OGCS, as it no longer provides the COM interoperability that OGCS v2 was built upon.<br/>
   <a href="blog/2025/06/17/v3-release-candidate-available.html" style="background-color: yellow; text-decoration: underline;">A new v3 release of OGCS is now available</a> that no longer requires the Outlook client at all, and connects directly to Microsoft 365 cloud accounts.
</font></div>

## Functionality

<style> ul { margin-bottom: 2px; } </style>
- Supports all versions of Outlook from 2003 to 2024 64-bit 
   - Including Microsoft365 releases from the [General Availability](https://learn.microsoft.com/en-us/windows/deployment/update/get-started-updates-channels-tools#general-availability-channel) channel
   - For "New Outlook", [check the latest developments](https://github.com/phw198/OutlookGoogleCalendarSync/issues/1888)
- Installable and portable options - even runs from a USB thumbdrive
- Synchronises items in any calendar folder, including those shared with you, from
   - Outlook :arrow_right: Google
   - Outlook :arrow_left: Google
   - Outlook :left_right_arrow: Google (two-way/bidirectional sync)
- Includes the following event attributes:
   - Subject
   - Description
   - Location
   - Attendees (including whether required or optional)
   - Reminder events
   - Availability (free/busy)
   - Privacy (public/private)
   - Categories/colours
- Differential comparison updates only attributes that have changed
- Customisable date range to synchronise, past and future
- Frequency of automatic syncs, including push-sync from Outlook
- Configurable proxy settings, or use Internet Explorer's
- Merge new events into existing on destination calendar
- Prompt on deletion of items
- Ability to obfuscate custom words for privacy/security
- Option to force items in target calendar
   - as private
   - as available
- Syncs recurring items properly as a series
- Can run unobtrusively in the system tray, with bubble notifications on sync
- Application can start on login, with delay if required
</span>

## Minimal Requirements
- Any version of Windows with .Net Framework 4.5 installed*
- Outlook 2003 to 2019/Microsoft 365, 32 or 64-bit

\* Installed by OGCS Setup.exe if necessary.

### *Not* Required :wink:
- Install or local administrator privileges
- Direct internet connection (proxy aware)



<img src="{{ site.github-repo }}/raw/master/docs/images/home_screen1.png" width="450px" />
<img src="{{ site.github-repo }}/raw/master/docs/images/home_screen2.png" width="450px" />
<img src="{{ site.github-repo }}/raw/master/docs/images/home_screen3.png" width="450px" />
