---
layout: page
next: Settings
next-url: settings
---
{% include navigation-buttons.html %}

# Sync

![Sync Screenshot](https://github.com/phw198/OutlookGoogleCalendarSync/raw/master/docs/images/home_screen1.png)

**Last Successful:** The last time a sync was successfully completed, whether it be an automated/scheduled sync, or a manual.  
**Next Scheduled:** If you have set an sync interval, this will display the next time an automated sync is scheduled.

**Verbose Output:** When checked, details of the items being synced will show in the main output pane above.  
**Mute Clicks:** On some systems, when the output pane is updated during sync, Windows produces the sound of a mouse clicking. To prevent this, check the box.

**Start Sync:** Perform a manual, on-demand sync. The default will be a differential sync - only changes since the last sync.
<div class="tip">:bulb:To perform a full sync, press and hold a shift key whilst clicking the button. You may never need to do this, or very rarely, but it may be helpful after making a change to OGCS settings. For example, if you decide to sync meeting attendees this will only start syncing them from that point onwards. To add attendees to previously synced meetings, perform a full sync.</div>


<p>&nbsp;</p>
{% include navigation-buttons.html %}
<p>&nbsp;</p>
