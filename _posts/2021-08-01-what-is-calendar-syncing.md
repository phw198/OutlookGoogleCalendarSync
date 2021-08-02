---
layout: post
title:  "What is (and isn't) calendar syncing"
date:   2021-08-01
categories: guide blog
---

The internet is awash with articles, guidance and videos on how to "sync" your calendars, when what they actually talk about is __sharing__ a calendar through the use of an `.ics` file or URL.
This is **not** synchronisation in its true sense, and will most likely lead to confusion, wasted time and just general frustration. Let's try and clear some things up...

### Importing an ICS file

Outlook and Google calendars allow you to export some or all of their appointments to an `ics` format, which [follows a standard](https://en.wikipedia.org/wiki/ICalendar) understood across different calendar vendors. However, this should be seen as a one-off exercise that you might want to do if you are migrating from one calendar to another. It also works well if simply importing a one-off calendar item, perhaps from an event, hotel or flight you have booked - often these websites will provide a link to download an `ics` file that contains your booking details, so you can save them in your calendar.

![google ics import]({{ site.baseurl }}/images/posts/import-ics.png)

This isn't synchronisation.

### Subscribing to an ICS URL

An `ics` file is generally a static file residing on your computer or in an email attachment - it doesn't change. However, these files can also be hosted on a website and updated, for example [https://fixtur.es](https://fixtur.es) hosts `ics` files for football teams all around the world, regularly updating them with past results and upcoming fixtures. These can then be subscribed to and will show in your calendar.

![google ics subscribe]({{ site.baseurl }}/images/posts/subscribe-ics.png)

This works well for slowly changing feeds, as calendar like Google will only check the subscribed link for updates every 6 to 48 hours. Despite this, many guides will go through the process of making your Google calendar accessible via an `ics` URL, and then subscribing to it in Outlook calendar. This is unsatisfactory for two reasons:
1. All your calendar data is now available publically.
1. The update frequency is too low - most people want a meeting invite to appear almost immediately in their synced calendar

This isn't synchronisation either.

### True synchronisation

Google and Outlook don't like to let their calendars "talk" with each other, so true synchronisation is much harder to achieve. To do the job properly, it therefore requires a dedicated application to translate changes happening in one calendar and pushing them across to the other. Google used to provide a calendar sync application of their own, but they stopped making it available several years back and this became the reason for OGCS.

Having a dedicated sync application also allows users to configure exactly what syncs, when and how. For example, setting a sync to run every hour, only for items in the next month that have a specific category assigned and for which the invitation has been accepted. Similarly, the appointment title could be included in the sync, but the description omitted whilst also forcing all items to show as "private" in the synced calendar, so anyone else with view rights on it cannot see the detail.

Outlook Google Calendar Sync offers all this and more!

#### Microsoft Flow / Microsoft Power Automate

Some guides suggest using Microsoft's cloud offering for integration and synchronisation. Whilst this could work for very simple calendars with single appointments, it is doomed to fail when considering the complexities in recurring series appointments, for which various occurrences have been cancelled or moved. Microsoft and Google handle recurring series very differently and the average user cannot be expected to understand all this when trying to write their own sync process in something as simplistic as Power Automate.


### So how do I sync my Google and Outlook calendars properly then?

Using something like OGCS, of course! For a step-by-step guide, see my other blog posts:
* [How to sync calendars from Outlook to Google]({{ site.baseurl }}{% post_url 2021-07-24-how-to-sync-calendars-from-outlook-to-google %}) 
