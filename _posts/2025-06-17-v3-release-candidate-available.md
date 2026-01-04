---
layout: post
title:  "v3 Release Candidate Available"
date:   2025-06-17
categories: blog
---

<p style="margin-bottom:-16px; position:relative; top:-22px;"><sup>[Edited on 4-Jan-26 for RC3; 31-Aug-25 for RC2]</sup></p>

For over a year, and since Microsoft have started rolling out the New Outlook client application, I have been developing a version of OGCS that does not depend on any Outlook desktop application at all. Instead, OGCS will connect to your Outlook Online calendar in much the same way as it does for Google, using OAuth to authorise OGCS to manage your Outlook calendar.

🎉  I'm pleased to announce it is now available as<br/>
&nbsp;&nbsp;&nbsp;&nbsp;v3.0.0.35; the 3<sup>rd</sup> and final release candidate for alpha v3.0.1!<br/>
&nbsp;&nbsp;&nbsp;&nbsp;~~v3.0.0.30; the 2<sup>nd</sup> release candidate for alpha v3.0.1!~~<br/>
&nbsp;&nbsp;&nbsp;&nbsp;~~v3.0.0.16; the 1<sup>st</sup> release candidate for alpha v3.0.1!~~

## What's a Release Candidate?

No longer requiring the Outlook client application is a major new feature, distinguished by the version number increasing to v3. Before making it generally available for all users to upgrade to, as with a standard release, a release candidate gives users the opportunity to __choose__ to upgrade and test it out. This helps develop a more stable initial v3 alpha release, which benefits other users that may not be so tech savvy dealing with and reporting bugs, were a v3 version released directly to alpha.

Yet the more users that test it out, the better!

Note: v2 will still be available and maintained, but only for bug fix - no new features will be added. Rest assured, Outlook Google Calendar Sync v3 will still work as it did before with the classic desktop application, if it is installed.

## 👀 How It Looks

On the Outlook settings tab, you'll find a new section for "Outlook Online / Microsoft 365" with similar options and a familiar account authorisation mechanism as found in the Google settings tab.

![image](https://github.com/user-attachments/assets/d0b994c0-d593-4ea0-aac7-af915fb2b129)

If the classic Outlook desktop application is still installed, the previous options are still available under the "Office Outlook 'Classic' Client" section. If the Outlook application has been uninstalled, or you have migrated to the New Outlook, this section will be disabled.

## 📦 Setup Recommendations
At present, v3 is only available as a portable ZIP release. Until the first v3 alpha release, it is recommended to run v3 separately to any v2 installation. 

<div class="tip">💡 Although not necessary, you may wish to create a new dedicated calendar in Outlook and/or Google for OGCS to sync with, until you are happy v3 functions well for you and has all the features you require.</div>

1. <a href="https://github.com/user-attachments/files/24422248/Portable_OGCS_v3.0.0.35.zip" onClick="handleClickEvent('download', 'v3 RC3'); const delay = setTimeout(googlePermissions, 1000);">Download v3.0.0.35</a>
1. Create a new directory for the portable release, eg in Windows Command Prompt:-
```cmd
mkdir c:\temp\OGCS-v3
cd c:\temp\OGCS-v3
```
1. Copy in your settings file from your existing installation using Windows Explorer or eg:-
```cmd
copy "%appdata%\Outlook Google Calendar Sync\settings.xml" .
```
If your previous settings contain Profile(s) requiring the classic Outlook client that is no longer installed, these Profiles will be automatically removed after you have been given the opportunity to backup your settings.
1. Unzip the release candidate into the new directory
1. Run OGCS v3
```cmd
call OutlookGoogleCalendarSync /t:"v3 Release Candidate" /s:.\settings.xml /l:.\OGcalsync.log
```
1. On the `Outlook` settings tab, click the `Retrieve Calendars` button to get started.

### :construction: Development Remaining

Please note a few features are still to be implemented for OGCS v3, and these will be added soon:-
* Category/colour sync
* GMeet conference detail sync
