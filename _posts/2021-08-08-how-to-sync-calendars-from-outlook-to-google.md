---
layout: post
title:  "How to sync calendars from Outlook to Google"
date:   2021-08-08
categories: blog
---

How often have you found you've missed an important meeting, or discovered too late that you are in fact double booked? Or perhaps, it would be helpful to share your free/busy status with your significant other, so they can plan around your work commitments. Or maybe you simply want to be able to ask your Google assistant what your schedule is for the day, and it be able to know about your work calendar, not just your personal appointments? No doubt this is why you're here.

It's a common problem - your corporate calendar, often hosted on Microsoft Exchange, is configured in such a way that you cannot get to it without signing in on your work laptop or (if you're really desperate!) having to allow the IT department full control of your personal device before you can view emails and calendars on it. This is likely true even if you use Office365 (O365/M365). No thank you.

It would make life so much easier if you only had one calendar with all your appointments and meetings...but figuring out how to get both Google and Microsoft calendars to talk to each other might be proving rather more difficult than you anticipated!

### A Dedicated Sync Application: OGCS

You may have already read other guides and by now worked out [calendar __sharing__ won't cut it]({{ site.baseurl }}{% post_url 2021-08-01-what-is-calendar-syncing %}) - you need an application to properly sync your calendars.

OGCS is free, so get started by downloading it with one of the buttons on the right of the page. [Not sure which version?]({{ site.baseurl }}/guide/install)

#### Video Tutorial

Here's a YouTube how-to tutorial as created by [Andy](https://ekiwi-blog.de/9306/synchronize-outlook-calendar-with-google-calendar-english/) from the user community:
<iframe width="560" height="315" src="https://www.youtube.com/embed/h5BDx-9UP3Y" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>

#### Select Outlook Calendar
Once OGCS is running, you'll need to configure it to sync calendar items from the correct source Outlook calendar. Shown below is an Outlook calendar conveniently named `OGCS` - select the same from the drop down on the `Settings > Outlook` tab. If you are syncing your default Outlook calendar, this step can be skipped.

![Select Outlook calendar]({{ site.baseurl }}/images/posts/pick-outlook-cal.png)

#### Select Google Calendar
Next up is to choose the Google calendar to sync to. First you will need to authorise OGCS to access your Google calendar, which is easily done by clicking the `Retrieve Calendars` button. A browser page will open where you need select the calendar with which you want OGCS to sync:

![Google OAuth Step 1]({{ site.baseurl }}/images/posts/google-oauth-1.png)

Review the permissions OGCS require, and if happy, click `Allow`:

![Google OAuth Step 2]({{ site.baseurl }}/images/posts/google-oauth-2.png)

Close the browser window once you have confirmation you can do so, then return to OGCS which should now show your Google calendars in the dropdown:

![Select Google calendar]({{ site.baseurl }}/images/posts/pick-google-cal.png)

### Synchronisation Options

Lastly, click the `Sync Options` tab to configure the direction of sync to Outlook to Google. A couple of other settings you'll likely want to change are:
* Uncheck disable deletions (after making sure sync is running OK)
* Change the value from zero in the `Schedule every...` in order for syncs to run automatically

![Sync options]({{ site.baseurl }}/images/posts/sync-options.png)

### Run a Manual Sync

Now you're ready to run your first sync! On the `Sync` tab, click `Start Sync` - and watch your calendar items sync across:

![Run Sync]({{ site.baseurl }}/images/posts/outlook-to-google.png)
