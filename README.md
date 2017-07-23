# <img src="https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/ogcs128x128.png" valign="middle"> Outlook Google Calendar Sync

> Offers calendar synchronisation between Outlook and Google, including attendees and reminders.  
> Completely free, no install necessary, works behind web proxies and actively developed.

<p align="center">:sparkles: We're new here, having just moved over <a href="https://outlookgooglecalendarsync.codeplex.com" target="_blank">from Codeplex</a>. Hi! :sparkles:<br/>
Watch this space for updates as we busily get things moved over.</p>

### Continually Improving
<img src="https://campbowiedistrict.com/wp-content/uploads/2016/08/page0-under-construction1.png" v
align="left" width="100px"/> 
A lot of work has gone in to this project, aiming to cater for many different versions of Outlook and types of usage.  
Beta releases should now be pretty stable, but if you find a bug [you can help](https://github.com/phw198/OutlookGoogleCalendarSync/wiki/Reporting-Problems) squash it! :beetle:  
If you would like to support this tool and its further development please [![donate](https://www.paypalobjects.com/en_GB/i/btn/btn_donate_SM.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=RT46CXQDSSYWJ&item_name=Outlook%20Google%20Calendar%20Sync%20donation.%20For%20splash%20screen%20hiding,%20enter%20your%20Gmail%20address%20in%20comment%20section)

<a href="https://plus.google.com/communities/114412828247015553563"><img src="https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/home_google_community.png" align="center"></a> <a href="http://www.twitter.com/OGcalsync"><img src="https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/home_twitter_follow.png" align="center"></a> <a href="https://twitter.com/intent/tweet?original_referer=https%3A%2F%2Fabout.twitter.com%2Fresources%2Fbuttons&text=I%20just%20found%20this%20amazing%20free%20tool%20to%20sync%20Outlook%20and%20Google%20calendars&tw_p=tweetbutton&url=http%3A%2F%2Fbit.ly%2FOGcalsync&via=OGcalsync"><img src="https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/home_tweet.png" align="center"></a>


### Latest Release: [![Latest Release](https://img.shields.io/github/release/phw198/OutlookGoogleCalendarSync.svg)](https://github.com/phw198/OutlookGoogleCalendarSync/releases/latest) [![Latest Release downloads](https://img.shields.io/github/downloads/phw198/outlookgooglecalendarsync/latest/total.svg)](https://github.com/phw198/OutlookGoogleCalendarSync/releases/latest)

:floppy_disk: [Installer](https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.5.3-alpha/Setup.exe)  
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[![](https://img.shields.io/github/downloads/phw198/outlookgooglecalendarsync/v2.5.3-alpha/Setup.exe.svg)](https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.5.3-alpha/Setup.exe)
 
:package: [Portable ZIP](https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.5.3-alpha/Portable_OGCS_v2.5.3.zip)  
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[![](https://img.shields.io/github/downloads/phw198/outlookgooglecalendarsync/v2.5.3-alpha/Portable_OGCS_v2.5.3.zip.svg)](https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.5.3-alpha/Portable_OGCS_v2.5.3.zip)

:information_source: Upgrades to this release  
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;![](https://img.shields.io/github/downloads/phw198/outlookgooglecalendarsync/v2.5.3-alpha/OutlookGoogleCalendarSync-2.5.3-alpha-full.nupkg.svg)  
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;![](https://img.shields.io/github/downloads/phw198/outlookgooglecalendarsync/v2.5.3-alpha/OutlookGoogleCalendarSync-2.5.3-alpha-delta.nupkg.svg)


### Current Beta Release: [![Beta Release)](https://img.shields.io/github/downloads/phw198/outlookgooglecalendarsync/v2.5.0-beta/total.svg)](https://github.com/phw198/OutlookGoogleCalendarSync/releases/tag/v2.5.0-beta)

:package: [CodePlex ClickOnce Installer](https://outlookgooglecalendarsync.codeplex.com/downloads/get/clickOnce/OutlookGoogleCalendarSync.application)  
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[![](https://img.shields.io/badge/downloads-45976-green.svg)](https://outlookgooglecalendarsync.codeplex.com/downloads/get/clickOnce/OutlookGoogleCalendarSync.application)
 
:package: [Portable ZIP](https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.5.0-beta/Portable_OGCS_v2.5.0.zip)  
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[![](https://img.shields.io/github/downloads/phw198/outlookgooglecalendarsync/v2.5.0-beta/Portable_OGCS_v2.5.0.zip.svg)](https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v2.5.0-beta/Portable_OGCS_v2.5.0.zip)
 
## Functionality

- Supports all versions of Outlook from 2003 to 2016 64-bit!
- Installable and portable options - even runs from a USB thumbdrive
- Synchronises items in any calendar folder from
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
- Customisable date range to synchronise, past and future
- Frequency of automatic syncs, including push from Outlook
- Configurable proxy settings, or use Internet Explorer's
- Merge new events into existing on destination calendar
- Prompt on deletion of items
- Can run unobtrusively in the system tray, with bubble notifications on sync
- Ability to obfuscate custom words for privacy/security
- Syncs recurring items properly as a series

**Improvements**
- Match entries on multiple ID keys (not a just a "signature" string)
- Differential comparison and update only attributes that have changed.
- Sync non-default Outlook calendars
- Keep application responsive whilst synchronising
- Full CSV exports

![](https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/home_screen1.png)
![](https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/home_screen2.png)
![](https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/home_screen3.png)
