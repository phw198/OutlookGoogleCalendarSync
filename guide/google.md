---
layout: page
title: User Guide | Settings | Google
previous: Outlook
previous-url: outlook
next: Sync Options
next-url: syncoptions
---
{% include navigation-buttons.html %}

# Google Settings

This is where you configure the Google calendar you wish to sync. 

![Google Settings Screenshot](google.png)

**Connected Account:** Once OGCS has been authorised to connect to a Google calendar, the associated account will appear here. Donations are associated with this account, where ever OGCS is installed or used.

**Retrieve calendars:** Until clicked, a calendar cannot be selected. If this is the first use of OGCS, it will step through the process of authorising the application to then access the calendars.  
<br/>
<div class="tip">:memo: At no point is the Google account password used or revealed. Technically, a process called OAuth is used - from v2.6.1, once logged into Google, a special authentication code will automatically be provided to OGCS. In earlier versions, a code needed to be manually copy and pasted in to OGCS.</div>
<br/>
<div class="tip">:warning: If any other third party software required entry of a Google password directly into the application, it is strongly recommended you change your password immediately.

**Disconnect Account:** To go through the authentication process again, or change the Google account previously configured for OGCS to work with, then click this button.</div>
<br/>
To revoke authorisation from OGCS continuing to access a Google account, navigate to [https://myaccount.google.com/permissions](https://myaccount.google.com/permissions), click on `Outlook Google Calendar Sync` and then `Remove Access`.

**Select calendar:** Once successfully authenticated, chose the calendar to sync with.

**Exclude invitations I have declined:** Do not sync Google events that have been declined.  
**Exclude "Goal" events from sync:** Do not sync Google _Goals_ (a now [deprecated Google feature](https://support.google.com/calendar/answer/12207659), replaced by _focus time_ for those with work or school accounts)

### Advanced/developer options
This section allows users to utilise their own personal Google API quota. You will need to know the client ID and secret as provided within your Google developer console. Google Plus and Calendar APIs must be enabled for OGCS to work.


<p>&nbsp;</p>
{% include navigation-buttons.html %}
<p>&nbsp;</p>
